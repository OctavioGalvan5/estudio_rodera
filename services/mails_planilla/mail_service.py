import os
import smtplib
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from io import BytesIO

import pathlib

from dotenv import load_dotenv
from pypdf import PdfReader

# Sube desde services/mails_planilla/ → services/ → estudio_rodera/ → app-web-prueba-main/
load_dotenv(pathlib.Path(__file__).resolve().parents[3] / '.env')

DESTINATARIO_DEFAULT = 'octaviogalvan20034@gmail.com'

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
    Devuelve (nombre, cuil). Retorna ('', '') si los campos no son legibles
    (p. ej. PDF aplanado tras firma digital)."""
    try:
        reader = PdfReader(BytesIO(pdf_bytes))
        fields = reader.get_fields() or {}
        nombre = _leer_campo(fields, 'Apellido y nombre Interviniente')
        cuil   = _leer_campo(fields, 'Doc Interviniente')
        return nombre, cuil
    except Exception:
        return '', ''


def enviar_mail_planilla(nombre, cuil, pdf_bytes, pdf_filename, destinatario=None):
    """Envía el mail al tribunal con el PDF firmado adjunto.
    Levanta ValueError si faltan credenciales, o SMTPException si falla el envío."""
    gmail_user = os.getenv('GMAIL_USER', '').strip()
    gmail_pass = os.getenv('GMAIL_APP_PASSWORD', '').strip()
    if not gmail_user or not gmail_pass:
        raise ValueError('Configurá GMAIL_USER y GMAIL_APP_PASSWORD en el archivo .env')

    to = destinatario or DESTINATARIO_DEFAULT

    msg = MIMEMultipart()
    msg['From']    = gmail_user
    msg['To']      = to
    msg['Subject'] = f'Solicito sorteo de demanda {nombre}'

    msg.attach(MIMEText(CUERPO_TEMPLATE.format(nombre=nombre, cuil=cuil), 'plain', 'utf-8'))

    adjunto = MIMEApplication(pdf_bytes, _subtype='pdf')
    adjunto.add_header('Content-Disposition', 'attachment', filename=pdf_filename)
    msg.attach(adjunto)

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(gmail_user, gmail_pass)
        smtp.send_message(msg)
