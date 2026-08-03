import base64
import os
import pathlib

import msal
import requests
from dotenv import load_dotenv
from io import BytesIO
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


def _cargar_cache():
    """Carga el token cache desde env var (producción) o archivo (desarrollo)."""
    cache = msal.SerializableTokenCache()
    cache_b64 = os.getenv('MS_TOKEN_CACHE', '').strip()
    if cache_b64:
        padding = 4 - len(cache_b64) % 4
        if padding != 4:
            cache_b64 += '=' * padding
        cache.deserialize(base64.b64decode(cache_b64).decode('utf-8'))
    elif CACHE_FILE.exists():
        cache.deserialize(CACHE_FILE.read_text(encoding='utf-8'))
    return cache


def _guardar_cache(cache):
    """Persiste el cache actualizado al archivo y a la variable de entorno en memoria."""
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


def enviar_mail_planilla(nombre, cuil, pdf_bytes, pdf_filename, destinatario):
    """Envía el mail al tribunal con el PDF firmado adjunto via Microsoft Graph API."""
    if not destinatario:
        raise ValueError('El destinatario es obligatorio')

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
