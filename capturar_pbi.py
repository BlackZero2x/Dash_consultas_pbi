# ══════════════════════════════════════════════════════════════
# CAPTURAR_PBI.PY
# Captura la región del dashboard Power BI via Playwright CDP.
#
# La captura se delega a _captura_worker.py ejecutado en un
# subproceso aislado con timeout, lo que evita que un cuelgue
# de Playwright o Chrome bloquee el pipeline principal.
#
# Requiere Chrome abierto con:
#   chrome.exe --remote-debugging-port=9222 --user-data-dir="C:\chrome_pbi_session"
#   (el bat arranque_chrome.bat lo hace automáticamente a las 9:50am)
# ══════════════════════════════════════════════════════════════

import os
import sys
import time
import json
import logging
import subprocess
import socket
from datetime import datetime, date
from pathlib import Path

from dotenv import load_dotenv

load_dotenv(Path(__file__).parent / '.env')

logger = logging.getLogger(__name__)

# ══════════════════════════════════════════════════════════════
# 1. CONFIGURACIÓN
# ══════════════════════════════════════════════════════════════

WORKSPACE_ID = os.environ['PBI_WORKSPACE_ID']
REPORT_ID    = os.environ['PBI_REPORT_ID']
PAGE_ID      = os.environ['PBI_PAGE_ID']
PAGE_NAME    = os.environ.get('PBI_PAGE_NAME', 'ZONAL_SUPERVISOR')

SCREENSHOT_DIR = os.environ.get('SCREENSHOT_DIR', str(Path(__file__).parent / 'capturas'))
_chrome_user_dir = os.environ.get('CHROME_USER_DIR', r'C:\chrome_pbi_session')
CHROME_ARGS    = ['--remote-debugging-port=9222', f'--user-data-dir={_chrome_user_dir}']

RENDER_WAIT        = 10   # segundos esperando renderizado de visuals
DIAS_RETENSION     = 7    # días que se conservan las capturas
FILTRO_HORA_INICIO = 12   # hora a partir de la cual se aplica filtro fecha=hoy
WORKER_TIMEOUT     = 90   # segundos máximos para que el worker complete

# Recorte del viewport para eliminar barras de Chrome y paneles de PBI Service.
# Valores calibrados para 1366x768 (viewport por defecto de Chrome con CDP).
CLIP_X = 265   # margen izquierdo (panel nav PBI + barra lateral)
CLIP_Y = 100   # margen superior (barra Chrome + barra PBI Service)
CLIP_W = 1350  # 1440 - 255 (izq) - 40 (derecha)
CLIP_H = 640   # 860 - 175 (sup) - 30 (inferior)

_WORKER = str(Path(__file__).parent / '_captura_worker.py')


def _detectar_chrome() -> str:
    candidatos = [
        r'C:\Program Files\Google\Chrome\Application\chrome.exe',
        r'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe',
        os.path.join(os.environ.get('LOCALAPPDATA', ''), r'Google\Chrome\Application\chrome.exe'),
    ]
    for c in candidatos:
        if os.path.exists(c):
            return c
    raise EnvironmentError('No se encontró chrome.exe. Define CHROME_EXE en el archivo .env.')


CHROME_EXE = os.environ.get('CHROME_EXE') or _detectar_chrome()


# ══════════════════════════════════════════════════════════════
# 2. UTILIDADES
# ══════════════════════════════════════════════════════════════

def _limpiar_capturas_antiguas():
    hoy = date.today()
    try:
        for f in os.listdir(SCREENSHOT_DIR):
            if not f.endswith('.png') and not f.endswith('.txt'):
                continue
            ruta = os.path.join(SCREENSHOT_DIR, f)
            dias = (hoy - date.fromtimestamp(os.path.getmtime(ruta))).days
            if dias > DIAS_RETENSION:
                os.remove(ruta)
    except Exception as e:
        logger.warning(f'Limpieza de capturas: {e}')


def _chrome_esta_abierto() -> bool:
    try:
        with socket.create_connection(('127.0.0.1', 9222), timeout=1):
            return True
    except OSError:
        return False


def _url_reporte() -> str:
    base = (
        f'https://app.powerbi.com/groups/{WORKSPACE_ID}'
        f'/reports/{REPORT_ID}/{PAGE_ID}'
    )
    if datetime.now().hour >= FILTRO_HORA_INICIO:
        hoy = date.today().isoformat()
        return base + f"?filter=BASE_CON/fecha_registro eq '{hoy}'"
    return base


def _abrir_chrome():
    report_url = _url_reporte()
    logger.info('Chrome no detectado — lanzando...')
    subprocess.Popen([CHROME_EXE] + CHROME_ARGS + [report_url])
    for _ in range(30):
        time.sleep(2)
        if _chrome_esta_abierto():
            logger.info('Chrome abierto. Esperando carga del reporte (20s)...')
            time.sleep(20)
            return True
    logger.error('Chrome no arrancó en 60s.')
    return False


# ══════════════════════════════════════════════════════════════
# 3. CAPTURA
# ══════════════════════════════════════════════════════════════

def capturar_pagina(output_path: str = None) -> tuple:
    """
    Lanza _captura_worker.py en un subproceso aislado con timeout.
    Retorna: (ruta_png: str, ultimo_corte: str | None)
    Lanza:   RuntimeError si Chrome no está disponible o el worker falla.
    """
    os.makedirs(SCREENSHOT_DIR, exist_ok=True)
    _limpiar_capturas_antiguas()

    dest = output_path or os.path.join(
        SCREENSHOT_DIR,
        f'zonal_supervisor_{datetime.now().strftime("%Y%m%d_%H%M%S")}.png'
    )

    if not _chrome_esta_abierto():
        ok = _abrir_chrome()
        if not ok:
            raise RuntimeError('No se pudo abrir Chrome con depuración remota.')

    report_url = _url_reporte()

    cmd = [
        sys.executable, _WORKER,
        dest, report_url, str(RENDER_WAIT),
        str(CLIP_X), str(CLIP_Y), str(CLIP_W), str(CLIP_H),
    ]

    env = {k: v for k, v in os.environ.items()
           if k not in ('HTTP_PROXY', 'HTTPS_PROXY', 'http_proxy', 'https_proxy')}
    env['NODE_NO_WARNINGS'] = '1'

    logger.info(f'Lanzando worker (timeout={WORKER_TIMEOUT}s)...')
    try:
        proc = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=WORKER_TIMEOUT,
            env=env,
        )
    except subprocess.TimeoutExpired:
        raise RuntimeError(f'Worker de captura superó el timeout de {WORKER_TIMEOUT}s — Chrome puede estar colgado.')

    if proc.returncode != 0 or not proc.stdout.strip():
        stderr = proc.stderr.strip()
        raise RuntimeError(f'Worker de captura falló (rc={proc.returncode}): {stderr or "sin salida"}')

    try:
        data = json.loads(proc.stdout.strip().splitlines()[-1])
    except json.JSONDecodeError:
        raise RuntimeError(f'Worker devolvió salida inesperada: {proc.stdout.strip()[:200]}')

    if not data.get('ok'):
        raise RuntimeError(f'Worker de captura reportó error: {data.get("error")}')

    ultimo_corte = data.get('ultimo_corte')
    if ultimo_corte:
        logger.info(f'Ultimo Corte leido: {ultimo_corte}')
    else:
        logger.warning('Ultimo Corte no detectado en DOM.')

    logger.info(f'Captura guardada: {dest}')
    return dest, ultimo_corte


# ══════════════════════════════════════════════════════════════
# 4. ENTRY POINT
# ══════════════════════════════════════════════════════════════

if __name__ == '__main__':
    ruta, corte = capturar_pagina()
    print(f'PNG:          {ruta}')
    print(f'Ultimo Corte: {corte}')
