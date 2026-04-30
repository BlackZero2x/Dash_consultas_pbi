# ══════════════════════════════════════════════════════════════
# ACTUALIZAR_PBI.PY
# Autentica con Power BI Service usando un Service Principal
# (app registrada en Azure AD) y dispara el refresh del dataset.
#
# Flujo de autenticación — Client Credentials (Service Principal):
#   No requiere intervención manual ni MFA. El script obtiene
#   un access_token directamente con client_id + client_secret.
#   El token dura 1 hora; se cachea en memoria durante la sesión.
#   Al no haber usuario involucrado, no hay refresh_token —
#   simplemente se solicita un token nuevo cuando expira.
#
# Requisito previo (configurado por el admin):
#   - App registrada en Azure AD con client secret
#   - "Permitir a las entidades de servicio usar las API de Power BI"
#     habilitado en el Portal de administración de Power BI
#   - La app agregada como Miembro/Colaborador en el workspace
#
# Uso:
#   python actualizar_pbi.py          → hace refresh directamente
#   from actualizar_pbi import triggerear_refresh  → desde otro script
# ══════════════════════════════════════════════════════════════

import os
import time
import logging
import requests
from pathlib import Path
from urllib.parse import quote_plus
from dotenv import load_dotenv

load_dotenv(Path(__file__).parent / '.env')

logger = logging.getLogger(__name__)

# ══════════════════════════════════════════════════════════════
# 1. CONFIGURACIÓN
# ══════════════════════════════════════════════════════════════

WORKSPACE_ID  = os.environ['PBI_WORKSPACE_ID']
DATASET_ID    = os.environ['PBI_DATASET_ID']
CLIENT_ID     = os.environ['PBI_CLIENT_ID']
CLIENT_SECRET = os.environ['PBI_CLIENT_SECRET']
TENANT_ID     = os.environ['PBI_TENANT_ID']

TOKEN_URL     = f'https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token'
SCOPE         = 'https://analysis.windows.net/powerbi/api/.default'

_token_cache = {'access_token': None, 'expires_at': 0}


def _proxies():
    proxy = os.environ.get('HTTPS_PROXY') or os.environ.get('HTTP_PROXY')
    if not proxy:
        return None
    user = os.environ.get('PROXY_USER', '')
    pwd  = os.environ.get('PROXY_PASS', '')
    if user and pwd:
        proxy = proxy.replace('http://', f'http://{quote_plus(user)}:{quote_plus(pwd)}@')
    return {'http': proxy, 'https': proxy}


# ══════════════════════════════════════════════════════════════
# 2. AUTENTICACIÓN — Client Credentials (Service Principal)
# ══════════════════════════════════════════════════════════════
#
# Client Credentials es el flujo OAuth2 para procesos desatendidos:
#   - No hay usuario → no hay MFA ni consentimiento interactivo
#   - El token se obtiene directamente con client_id + client_secret
#   - Expira en 1 hora → se renueva automáticamente al próximo uso

def obtener_access_token():
    if time.time() < _token_cache['expires_at']:
        return _token_cache['access_token']

    resp = requests.post(TOKEN_URL, data={
        'grant_type':    'client_credentials',
        'client_id':     CLIENT_ID,
        'client_secret': CLIENT_SECRET,
        'scope':         SCOPE,
    }, proxies=_proxies())
    resp.raise_for_status()
    data = resp.json()

    if 'access_token' not in data:
        raise RuntimeError(f'Error al obtener token: {data}')

    _token_cache['access_token'] = data['access_token']
    _token_cache['expires_at']   = time.time() + data.get('expires_in', 3600) - 60
    return _token_cache['access_token']


# ══════════════════════════════════════════════════════════════
# 3. REFRESH DEL DATASET
# ══════════════════════════════════════════════════════════════
#
# Llama a la Power BI REST API para disparar la actualización
# del dataset. Power BI Service leerá los archivos CSV desde
# OneDrive y actualizará el reporte publicado.
#
# Endpoint: POST /v1.0/myorg/groups/{workspace}/datasets/{dataset}/refreshes
# Respuesta esperada: HTTP 202 Accepted

def triggerear_refresh():
    access_token = obtener_access_token()
    url = (
        f'https://api.powerbi.com/v1.0/myorg/groups/{WORKSPACE_ID}'
        f'/datasets/{DATASET_ID}/refreshes'
    )
    headers = {'Authorization': f'Bearer {access_token}'}
    resp = requests.post(url, headers=headers, proxies=_proxies())

    if resp.status_code == 202:
        logger.info('Refresh de Power BI iniciado correctamente.')
        return True
    else:
        logger.error(f'Error al iniciar refresh: {resp.status_code} — {resp.text}')
        return False


# ══════════════════════════════════════════════════════════════
# 4. ESTADO DEL ÚLTIMO REFRESH
# ══════════════════════════════════════════════════════════════
#
# Consulta el historial de refreshes y devuelve el estado del
# más reciente: 'Completed', 'Failed', 'Unknown', o None si
# todavía no hay entrada (refresh en curso sin estado aún).

def obtener_estado_refresh():
    access_token = obtener_access_token()
    url = (
        f'https://api.powerbi.com/v1.0/myorg/groups/{WORKSPACE_ID}'
        f'/datasets/{DATASET_ID}/refreshes?$top=1'
    )
    headers = {'Authorization': f'Bearer {access_token}'}
    resp = requests.get(url, headers=headers, proxies=_proxies())

    if resp.status_code != 200:
        return None

    data = resp.json().get('value', [])
    if not data:
        return None

    # status puede ser: 'Completed', 'Failed', 'Disabled', 'Unknown'
    # endTime ausente significa que aún está en curso
    entry = data[0]
    if 'endTime' not in entry:
        return 'InProgress'
    return entry.get('status', 'Unknown')


# ══════════════════════════════════════════════════════════════
# 4. ENTRY POINT
# ══════════════════════════════════════════════════════════════

if __name__ == '__main__':
    triggerear_refresh()
