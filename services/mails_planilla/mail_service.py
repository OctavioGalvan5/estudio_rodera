import base64
import os
import pathlib
import smtplib
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from io import BytesIO

import msal
import requests
from dotenv import load_dotenv
from pypdf import PdfReader

load_dotenv(pathlib.Path(__file__).resolve().parents[3] / '.env')

CACHE_FILE = pathlib.Path(__file__).parent / 'ms_token_cache.json'

CUERPO_TEMPLATE = """\
Estimados:

Por la presente, me dirijo a ustedes en mi carácter de letrada apoderada de {nombre}, con CUIL {cuil},  a fin de remitir en archivo adjunto la planilla correspondiente para el sorteo de la demanda, solicitando que se proceda a su realización y se me informe oportunamente la radicación de la misma.
La presente corresponde a un juicio de reajuste contra ANSES.
Quedo a disposición para cualquier información adicional que resulte necesaria y aguardo su respuesta

Atentamente,

Dra. María Agustina Rodera
DNI 26.575.948
T 600 - F 190 CAFLP
Celular: 0221-155767603"""


def _leer_campo(fields, nombre):
    obj = fields.get(nombre)
    if obj is None:
        return ''
    if hasattr(obj, 'value'):
        return str(obj.value or '').strip()
    if isinstance(obj, dict):
        val = obj.get('/V', '')
        return str(val or '').strip()
    return str(obj).strip()


def extraer_datos_pdf(pdf_bytes):
    """Extrae nombre y CUIL del primer interviniente de una planilla de demanda.
    Devuelve (nombre, cuil). Retorna ('', '') si los campos no son legibles."""
    try:
        reader = PdfReader(BytesIO(pdf_bytes))
        fields = reader.get_fields() or {}
        nombre = _leer_campo(fields, 'Apellido y nombre Interviniente')
        cuil   = _leer_campo(fields, 'Doc Interviniente')
        return nombre, cuil
    except Exception:
        return '', ''


# ── Gmail SMTP ────────────────────────────────────────────────────────────────

def _enviar_gmail(nombre, cuil, pdf_bytes, pdf_filename, destinatario):
    mail_user = os.getenv('GMAIL_USER', '').strip()
    mail_pass = os.getenv('GMAIL_APP_PASSWORD', '').strip()

    msg = MIMEMultipart()
    msg['From']    = mail_user
    msg['To']      = destinatario
    msg['Subject'] = f'Solicito sorteo de demanda {nombre}'
    msg.attach(MIMEText(CUERPO_TEMPLATE.format(nombre=nombre, cuil=cuil), 'plain', 'utf-8'))

    adjunto = MIMEApplication(pdf_bytes, _subtype='pdf')
    adjunto.add_header('Content-Disposition', 'attachment', filename=pdf_filename)
    msg.attach(adjunto)

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(mail_user, mail_pass)
        smtp.send_message(msg)


# ── Microsoft Graph API ───────────────────────────────────────────────────────

def _cargar_cache():
    cache = msal.SerializableTokenCache()
    cache_b64 = os.getenv('MS_TOKEN_CACHE', '').strip()
    if cache_b64:
        try:
            cache_b64 = ''.join(cache_b64.split())
            padding = 4 - len(cache_b64) % 4
            if padding != 4:
                cache_b64 += '=' * padding
            cache.deserialize(base64.b64decode(cache_b64).decode('utf-8'))
        except Exception:
            if CACHE_FILE.exists():
                cache.deserialize(CACHE_FILE.read_text(encoding='utf-8'))
    elif CACHE_FILE.exists():
        cache.deserialize(CACHE_FILE.read_text(encoding='utf-8'))
    return cache


def _guardar_cache(cache):
    serialized = cache.serialize()
    if CACHE_FILE.parent.exists():
        CACHE_FILE.write_text(serialized, encoding='utf-8')
    os.environ['MS_TOKEN_CACHE'] = base64.b64encode(serialized.encode('utf-8')).decode()


def _get_access_token():
    client_id     = os.getenv('AZURE_CLIENT_ID', '').strip()
    client_secret = os.getenv('AZURE_CLIENT_SECRET', '').strip()

    if not client_id or not client_secret:
        raise ValueError('Configurá AZURE_CLIENT_ID y AZURE_CLIENT_SECRET en el .env')

    cache = _cargar_cache()
    app_ms = msal.ConfidentialClientApplication(
        client_id,
        authority='https://login.microsoftonline.com/consumers',
        client_credential=client_secret,
        token_cache=cache,
    )

    accounts = app_ms.get_accounts()
    if not accounts:
        raise ValueError(
            'No hay sesión Microsoft guardada. '
            'Corré get_microsoft_token.py en tu PC y pegá MS_TOKEN_CACHE en el .env'
        )

    result = app_ms.acquire_token_silent(
        ['https://graph.microsoft.com/Mail.Send'],
        account=accounts[0],
    )

    if cache.has_state_changed:
        _guardar_cache(cache)

    if not result or 'access_token' not in result:
        raise ValueError(f'No se pudo obtener token Microsoft: {result}')

    return result['access_token']


def _enviar_microsoft(nombre, cuil, pdf_bytes, pdf_filename, destinatario):
    access_token = _get_access_token()

    mensaje = {
        'message': {
            'subject': f'Solicito sorteo de demanda {nombre}',
            'body': {
                'contentType': 'Text',
                'content': CUERPO_TEMPLATE.format(nombre=nombre, cuil=cuil),
            },
            'toRecipients': [{'emailAddress': {'address': destinatario}}],
            'attachments': [
                {
                    '@odata.type': '#microsoft.graph.fileAttachment',
                    'name': pdf_filename,
                    'contentType': 'application/pdf',
                    'contentBytes': base64.b64encode(pdf_bytes).decode(),
                }
            ],
        },
        'saveToSentItems': True,
    }

    resp = requests.post(
        'https://graph.microsoft.com/v1.0/me/sendMail',
        headers={
            'Authorization': f'Bearer {access_token}',
            'Content-Type': 'application/json',
        },
        json=mensaje,
        timeout=30,
    )

    if resp.status_code != 202:
        raise Exception(f'Error Graph API {resp.status_code}: {resp.text}')


# ── Punto de entrada ──────────────────────────────────────────────────────────

def enviar_mail_planilla(nombre, cuil, pdf_bytes, pdf_filename, destinatario):
    """Envía el mail al tribunal. Usa Gmail si GMAIL_USER está configurado,
    si no usa Microsoft Graph API (Hotmail)."""
    if not destinatario:
        raise ValueError('El destinatario es obligatorio')

    gmail_user = os.getenv('GMAIL_USER', '').strip()
    gmail_pass = os.getenv('GMAIL_APP_PASSWORD', '').strip()

    if gmail_user and gmail_pass:
        _enviar_gmail(nombre, cuil, pdf_bytes, pdf_filename, destinatario)
    else:
        _enviar_microsoft(nombre, cuil, pdf_bytes, pdf_filename, destinatario)
