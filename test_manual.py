# ══════════════════════════════════════════════════════════════
# TEST_MANUAL.PY
# Ejecuta el pipeline completo sin esperar el correo trigger.
#
# Pasos:
#   1. Extrae datos SQL + RH → BASE_CON.csv y RH.csv en OneDrive
#   2. Dispara el refresh del dataset en Power BI Service
#   3. Espera a que el refresh complete
#   4. Captura la página ZONAL_SUPERVISOR del dashboard
#   5. Envía la captura al grupo de WhatsApp
#
# Uso:
#   python test_manual.py
#   python test_manual.py --solo-extraccion
#   python test_manual.py --solo-refresh
#   python test_manual.py --solo-captura
#   python test_manual.py --solo-whatsapp <ruta_imagen>
# ══════════════════════════════════════════════════════════════

import sys
import time
import os
import socket
import subprocess
from datetime import datetime

import os
from dotenv import load_dotenv
from pathlib import Path
load_dotenv(Path(__file__).parent / '.env')

_WA_SERVER_PATH = os.environ.get('WA_SERVER_PATH', '')
if _WA_SERVER_PATH:
    sys.path.insert(0, _WA_SERVER_PATH)

from extraer_datos import ejecutar as extraer
from actualizar_pbi import triggerear_refresh, obtener_estado_refresh
from capturar_pbi import capturar_pagina, CHROME_EXE, CHROME_ARGS
from wa_client import WhatsAppClient

WA_GRUPO              = os.environ['WA_GRUPO']
REFRESH_POLL_INTERVAL = 30
REFRESH_TIMEOUT       = 600


def _chrome_esta_abierto() -> bool:
    try:
        with socket.create_connection(('127.0.0.1', 9222), timeout=1):
            return True
    except OSError:
        return False


def paso_verificar_chrome():
    print('=== VERIFICANDO CHROME ===')
    if _chrome_esta_abierto():
        print('Chrome ya está activo en puerto 9222.')
        return
    print('Chrome no detectado — arrancando...')
    bat = str(__file__).replace('test_manual.py', 'arranque_chrome.bat')
    subprocess.Popen([bat], shell=True)
    for _ in range(30):
        time.sleep(2)
        if _chrome_esta_abierto():
            print('Chrome listo. Esperando carga del reporte (20s)...')
            time.sleep(20)
            return
    print('ERROR: Chrome no arrancó en 60s.')
    sys.exit(1)


def paso_extraccion():
    print('=== EXTRACCIÓN DE DATOS ===')
    extraer()
    print('Extracción completada.')


def paso_refresh():
    print('=== REFRESH POWER BI ===')
    ok = triggerear_refresh()
    if not ok:
        sys.exit(1)


def paso_esperar_refresh():
    print('=== ESPERANDO REFRESH ===')
    inicio = time.time()
    while time.time() - inicio < REFRESH_TIMEOUT:
        estado = obtener_estado_refresh()
        if estado == 'Completed':
            print('Refresh completado correctamente.')
            return True
        if estado == 'Failed':
            print('El refresh terminó con error.')
            return False
        print(f'  Estado: {estado} — esperando...')
        time.sleep(REFRESH_POLL_INTERVAL)
    print('Timeout esperando refresh.')
    return False


def paso_captura():
    print('=== CAPTURA DASHBOARD ===')
    imagen, corte = capturar_pagina()
    print(f'Captura guardada: {imagen}')
    print(f'Ultimo Corte:     {corte}')
    return imagen, corte


def paso_whatsapp(imagen, corte=None):
    print('=== ENVÍO WHATSAPP ===')
    wa = WhatsAppClient(
        host='localhost',
        port=int(os.environ.get('WA_PORT', '8002')),
        config_path=os.environ.get('WA_CONFIG_PATH', ''),
    )
    corte_txt = corte if corte else datetime.now().strftime('%d/%m/%Y %H:%M')
    texto = (
        f'📊 *Dashboard Consultas* actualizado\n'
        f'🕐 Ultimo Corte: {corte_txt}\n'
        f'Vista: ZONAL / SUPERVISOR'
    )
    wa.send_image(WA_GRUPO, imagen, caption=texto)


def pipeline_completo():
    print('=== PIPELINE COMPLETO ===\n')

    print('-- Paso 0: Chrome --')
    paso_verificar_chrome()

    print('\n-- Paso 1: Extracción --')
    extraer()

    print('\n-- Paso 2: Refresh Power BI --')
    ok = triggerear_refresh()
    if not ok:
        print('Pipeline cancelado.')
        sys.exit(1)

    print('\n-- Paso 3: Esperando refresh --')
    paso_esperar_refresh()

    print('\n-- Paso 4: Captura dashboard --')
    imagen, corte = paso_captura()

    print('\n-- Paso 5: Envío WhatsApp --')
    paso_whatsapp(imagen, corte)

    print('\nPipeline completado exitosamente.')


if __name__ == '__main__':
    args = sys.argv[1:]

    if '--solo-extraccion' in args:
        paso_extraccion()
    elif '--solo-refresh' in args:
        paso_refresh()
        paso_esperar_refresh()
    elif '--solo-captura' in args:
        paso_verificar_chrome()
        paso_captura()
    elif '--solo-whatsapp' in args:
        idx = args.index('--solo-whatsapp')
        if idx + 1 >= len(args):
            print('Uso: python test_manual.py --solo-whatsapp <ruta_imagen>')
            sys.exit(1)
        imagen_arg = args[idx + 1]
        txt_path   = os.path.splitext(imagen_arg)[0] + '.txt'
        corte_arg  = open(txt_path, encoding='utf-8').read().strip() if os.path.exists(txt_path) else None
        paso_whatsapp(imagen_arg, corte_arg)
    else:
        pipeline_completo()
