"""
Script de autenticación con Microsoft (correr UNA VEZ en tu PC con navegador).

Pasos:
  1. Completá AZURE_CLIENT_ID y AZURE_CLIENT_SECRET abajo (o dejá que te los pida).
  2. Corré: python get_microsoft_token.py
  3. Se abre el navegador → iniciás sesión con estudioroderapuntaalta@hotmail.com
  4. El script guarda el token y te muestra el valor de MS_TOKEN_CACHE para el .env del servidor.
"""

import base64
import pathlib
import sys
import webbrowser
from http.server import BaseHTTPRequestHandler, HTTPServer
from urllib.parse import parse_qs, urlparse

import msal
from dotenv import load_dotenv
import os

REDIRECT_URI = 'http://localhost:8888/callback'
SCOPES = ['https://graph.microsoft.com/Mail.Send']
CACHE_FILE = pathlib.Path(__file__).parent / 'services' / 'mails_planilla' / 'ms_token_cache.json'


def main():
    load_dotenv(pathlib.Path(__file__).parent / '.env')
    print('=== Autenticación Microsoft Graph ===\n')
    client_id     = os.getenv('AZURE_CLIENT_ID', '').strip() or input('Azure Client ID:     ').strip()
    client_secret = os.getenv('AZURE_CLIENT_SECRET', '').strip() or input('Azure Client Secret: ').strip()

    cache = msal.SerializableTokenCache()

    app_ms = msal.ConfidentialClientApplication(
        client_id,
        authority='https://login.microsoftonline.com/consumers',
        client_credential=client_secret,
        token_cache=cache,
    )

    auth_url = app_ms.get_authorization_request_url(SCOPES, redirect_uri=REDIRECT_URI)
    print(f'\nAbriendo el navegador...\nSi no se abre solo, copiá esta URL:\n{auth_url}\n')
    webbrowser.open(auth_url)

    # Servidor local para recibir el callback de Microsoft
    auth_code = None

    class Handler(BaseHTTPRequestHandler):
        def do_GET(self):
            nonlocal auth_code
            params = parse_qs(urlparse(self.path).query)
            auth_code = params.get('code', [None])[0]
            self.send_response(200)
            self.end_headers()
            self.wfile.write('Autenticación exitosa. Podés cerrar esta ventana.'.encode())

        def log_message(self, *_):
            pass

    print('Esperando que inicies sesión en el navegador...')
    HTTPServer(('localhost', 8888), Handler).handle_request()

    if not auth_code:
        print('ERROR: No se recibió código de autorización.')
        sys.exit(1)

    result = app_ms.acquire_token_by_authorization_code(auth_code, SCOPES, redirect_uri=REDIRECT_URI)

    if 'access_token' not in result:
        print(f'ERROR: {result.get("error_description", result)}')
        sys.exit(1)

    # Guardar cache localmente
    CACHE_FILE.parent.mkdir(parents=True, exist_ok=True)
    serialized = cache.serialize()
    CACHE_FILE.write_text(serialized, encoding='utf-8')
    cache_b64 = base64.b64encode(serialized.encode('utf-8')).decode()

    print('\n✓ Autenticación exitosa!\n')
    print(f'Token guardado en: {CACHE_FILE}\n')
    print('=' * 60)
    print('Copiá estas líneas al .env del SERVIDOR (Dokploy):')
    print('=' * 60)
    print(f'AZURE_CLIENT_ID={client_id}')
    print(f'AZURE_CLIENT_SECRET={client_secret}')
    print(f'MS_TOKEN_CACHE={cache_b64}')
    print('=' * 60)


if __name__ == '__main__':
    main()
