# ══════════════════════════════════════════════════════════════
# PUSH_DASH_PBI_CONSULTAS.PY
# Ejecución única: revisa Gmail UNA VEZ, actúa si hay correo trigger
# y termina. El Programador de Tareas lo lanza cada 10 minutos de
# lunes a viernes entre las 10:00 y las 18:00.
#
# Flujo completo al detectar el correo trigger:
#   1. Extrae datos de SQL Server + RH → BASE_CON.csv y RH.csv en OneDrive
#   2. Dispara el refresh del dataset en Power BI Service (Service Principal)
#   3. Espera a que el refresh complete
#   4. Captura la página ZONAL_SUPERVISOR del dashboard con Playwright
#   5. Envía la captura al grupo de WhatsApp con un texto
#   6. Marca el correo como leído
#
# Si no hay correo que coincida, el script termina sin hacer nada.
#
# Requisitos:
#   pip install google-auth google-auth-httplib2 google-api-python-client
#              requests playwright
#   playwright install chromium
#
# Uso:
#   python push_dash_pbi_consultas.py
# ══════════════════════════════════════════════════════════════

import os
import sys
import time
import logging
from datetime import datetime
from pathlib import Path

from dotenv import load_dotenv
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

from extraer_datos import ejecutar as extraer
from actualizar_pbi import triggerear_refresh, obtener_estado_refresh
from capturar_pbi import capturar_pagina

sys.path.insert(0, r'C:\proyectos\AVANCE_MOVISTAR\whatsapp_server')
from wa_client import WhatsAppClient

load_dotenv(Path(__file__).parent / '.env')

# ── Logging centralizado ───────────────────────────────────────
_log_dir = Path(__file__).parent / 'logs'
_log_dir.mkdir(exist_ok=True)
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(_log_dir / f"pipeline_{datetime.now().strftime('%Y%m%d')}.log", encoding='utf-8'),
        logging.StreamHandler(sys.stdout),
    ],
)
logger = logging.getLogger(__name__)
logging.getLogger('googleapiclient.discovery_cache').setLevel(logging.ERROR)


# ══════════════════════════════════════════════════════════════
# 1. CONFIGURACIÓN
# ══════════════════════════════════════════════════════════════

# ── Filtro del correo trigger ──────────────────────────────────
REMITENTE = 'e@auren.com.pe'
ASUNTO    = 'dito_consultas_hoy_csv - Actualización disponible'

# ── Autenticación Gmail ────────────────────────────────────────
SCOPES           = ['https://www.googleapis.com/auth/gmail.modify']
CREDENTIALS_FILE = 'credentials.json'
TOKEN_FILE       = 'token.json'

# ── WhatsApp ───────────────────────────────────────────────────
WA_GRUPO = 'Canal Fija 2026 Gestión AUREN'

# ── Espera máxima para que el refresh de PBI complete (seg) ───
REFRESH_POLL_INTERVAL = 30
REFRESH_TIMEOUT       = 600   # 10 minutos máximo


# ══════════════════════════════════════════════════════════════
# 2. AUTENTICACIÓN GMAIL
# ══════════════════════════════════════════════════════════════

def autenticar():
    creds = None
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_FILE, 'w') as token:
            token.write(creds.to_json())
    return creds


# ══════════════════════════════════════════════════════════════
# 3. ESPERAR QUE EL REFRESH DE PBI COMPLETE
# ══════════════════════════════════════════════════════════════
#
# Tras disparar el refresh (HTTP 202), Power BI procesa en segundo
# plano. Consultamos el estado cada REFRESH_POLL_INTERVAL segundos
# hasta que el último refresh figure como "Completed" o "Failed".

def esperar_refresh():
    logger.info('Esperando que el refresh de Power BI complete...')
    inicio = time.time()
    while time.time() - inicio < REFRESH_TIMEOUT:
        estado = obtener_estado_refresh()
        if estado == 'Completed':
            logger.info('Refresh completado correctamente.')
            return True
        if estado == 'Failed':
            logger.error('El refresh terminó con error.')
            return False
        if estado is not None:
            logger.info(f'Estado actual: {estado} — esperando...')
        time.sleep(REFRESH_POLL_INTERVAL)
    logger.error(f'Timeout: el refresh no completó en {REFRESH_TIMEOUT}s.')
    return False


# ══════════════════════════════════════════════════════════════
# 4. PIPELINE PRINCIPAL
# ══════════════════════════════════════════════════════════════

def ejecutar_pipeline(gmail_service, msg_id):
    ahora = datetime.now().strftime('%d/%m/%Y %H:%M')
    logger.info(f'Correo trigger detectado. Iniciando pipeline... [{ahora}]')

    logger.info('-- Paso 1: Extracción de datos --')
    extraer()

    logger.info('-- Paso 2: Refresh Power BI --')
    ok = triggerear_refresh()
    if not ok:
        logger.error('No se pudo iniciar el refresh. Pipeline cancelado.')
        return False

    logger.info('-- Paso 3: Esperando refresh --')
    refresh_ok = esperar_refresh()

    logger.info('-- Paso 4: Captura ZONAL_SUPERVISOR --')
    imagen, ultimo_corte = None, None
    try:
        imagen, ultimo_corte = capturar_pagina()
    except Exception as e:
        logger.error(f'Captura fallida: {e}')

    logger.info('-- Paso 5: Envío WhatsApp --')
    wa = WhatsAppClient(
        host='localhost',
        port=8002,
        config_path=r'C:\proyectos\AVANCE_MOVISTAR\whatsapp_server\config.json',
    )
    if not wa.is_ready():
        logger.warning('Servidor WhatsApp no disponible — omitiendo envío.')
    elif imagen:
        corte_txt = ultimo_corte if ultimo_corte else ahora
        estado_txt = 'actualizado correctamente' if refresh_ok else 'actualizado (verificar errores)'
        texto = (
            f'Dashboard Consultas {estado_txt}\n'
            f'Ultimo Corte: {corte_txt}\n'
            f'Vista: ZONAL / SUPERVISOR'
        )
        wa.send_image(WA_GRUPO, imagen, caption=texto)
    else:
        wa.send_text(
            WA_GRUPO,
            f'Dashboard Consultas actualizado ({ahora}) - captura no disponible.'
        )

    gmail_service.users().messages().modify(
        userId='me',
        id=msg_id,
        body={'removeLabelIds': ['UNREAD']}
    ).execute()
    logger.info('Correo marcado como leído.')
    return True


# ══════════════════════════════════════════════════════════════
# 5. ENTRY POINT — ejecución única
# ══════════════════════════════════════════════════════════════

def main():
    creds = autenticar()
    service = build('gmail', 'v1', credentials=creds)

    query = f'from:{REMITENTE} subject:{ASUNTO} is:unread'
    results = service.users().messages().list(userId='me', q=query).execute()

    if not results.get('messages'):
        # Sin correo trigger → salir silenciosamente
        sys.exit(0)

    msg_id = results['messages'][0]['id']
    exito  = ejecutar_pipeline(service, msg_id)
    sys.exit(0 if exito else 1)


if __name__ == '__main__':
    main()
