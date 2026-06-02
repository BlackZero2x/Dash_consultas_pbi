# ══════════════════════════════════════════════════════════════
# VENDEDORES_SIN_CONSULTAS.PY
# Genera y envía el reporte de vendedores PLANILLA sin consultas
# al corte actual.
#
# Flujo de mensajes:
#   1. Mensaje introductorio al grupo
#   2. Un único mensaje con todos los supervisores mencionados
#      (@número) y sus vendedores enlistados debajo
#
# Fuentes (deben existir en disco antes de llamar):
#   BASE_CON.csv               — consultas del día (extraer_datos.py)
#   RH.csv                     — RH activos en campo (extraer_datos.py)
#   directorio_supervisores.xlsx — SUPERVISOR / TELEFONO para menciones
# ══════════════════════════════════════════════════════════════

import os
import time
import logging
import re
from datetime import date, datetime
from pathlib import Path

import pandas as pd
from dotenv import load_dotenv

load_dotenv(Path(__file__).parent / '.env')

logger = logging.getLogger(__name__)

ONEDRIVE_PATH   = os.environ.get('ONEDRIVE_PATH', '')
BASE_CON_FILE   = os.path.join(ONEDRIVE_PATH, 'BASE_CON.csv')
RH_FILE         = os.path.join(ONEDRIVE_PATH, 'RH.csv')
DIRECTORIO_FILE = str(Path(__file__).parent / 'directorio_supervisores.xlsx')


# ══════════════════════════════════════════════════════════════
# UTILIDADES
# ══════════════════════════════════════════════════════════════

def _normalizar(texto: str) -> str:
    reemplazos = str.maketrans(
        'áéíóúÁÉÍÓÚàèìòùÀÈÌÒÙñÑüÜ',
        'aeiouAEIOUaeiouAEIOUNNuU'
    )
    return re.sub(r'\s+', ' ', str(texto).upper().translate(reemplazos).strip())


def _cargar_directorio() -> dict:
    """Retorna {nombre_normalizado: wa_id} leyendo directorio_supervisores.xlsx."""
    if not os.path.exists(DIRECTORIO_FILE):
        logger.warning(f'directorio_supervisores.xlsx no encontrado en: {DIRECTORIO_FILE}')
        return {}

    df = pd.read_excel(DIRECTORIO_FILE, dtype={'TELEFONO': str})
    resultado = {}
    for _, row in df.iterrows():
        sup = str(row['SUPERVISOR']).strip()
        tel = str(row['TELEFONO']).strip().replace('-', '').replace(' ', '')
        if tel.startswith('51') and len(tel) == 11:
            tel = tel[2:]
        if len(tel) == 9:
            resultado[_normalizar(sup)] = f'51{tel}@c.us'
        else:
            logger.warning(f'Teléfono inválido para {sup}: {tel!r} — se omitirá de menciones')
    return resultado


def _wa_id(nombre_sup: str, directorio: dict) -> str | None:
    """Match exacto primero; luego containment para tolerar diferencias de acentos/espacios."""
    norm = _normalizar(nombre_sup)
    if norm in directorio:
        return directorio[norm]
    for key, wa in directorio.items():
        if key in norm or norm in key:
            return wa
    return None


# ══════════════════════════════════════════════════════════════
# LÓGICA PRINCIPAL
# ══════════════════════════════════════════════════════════════

def obtener_ausentes(fecha: date = None) -> pd.DataFrame:
    """
    Retorna DataFrame [ZONAL, SUPERVISOR, VENDEDOR] de vendedores
    PLANILLA sin consultas en `fecha` (default: hoy).
    """
    if fecha is None:
        fecha = date.today()

    for ruta in (BASE_CON_FILE, RH_FILE):
        if not os.path.exists(ruta):
            raise FileNotFoundError(f'No se encontró: {ruta}')

    base = pd.read_csv(BASE_CON_FILE, encoding='utf-8-sig', low_memory=False)
    rh   = pd.read_csv(RH_FILE,      encoding='utf-8-sig', low_memory=False)

    base['fecha_registro'] = pd.to_datetime(base['fecha_registro'], errors='coerce').dt.date
    base['documento_v']    = pd.to_numeric(base['documento_v'], errors='coerce')
    rh['DNI']              = pd.to_numeric(rh['DNI'],           errors='coerce')

    consultaron = set(
        base.loc[base['fecha_registro'] == fecha, 'documento_v'].dropna().unique()
    )
    logger.info(f'Fecha: {fecha}  |  Con consulta hoy: {len(consultaron)}')

    cols_rh = {'DNI', 'VENDEDOR', 'SUPERVISOR', 'ZONAL', 'ESQ'}
    faltantes = cols_rh - set(rh.columns)
    if faltantes:
        raise ValueError(f'RH.csv no tiene columnas: {faltantes}')

    rh = rh[rh['ESQ'] == 'PLA']
    logger.info(f'Vendedores PLANILLA en RH: {len(rh)}')

    ausentes = (
        rh[~rh['DNI'].isin(consultaron)][['ZONAL', 'SUPERVISOR', 'VENDEDOR']]
        .dropna()
        .sort_values(['ZONAL', 'SUPERVISOR', 'VENDEDOR'])
    )
    logger.info(f'Sin consulta hoy: {len(ausentes)}')
    return ausentes


def construir_mensajes(fecha: date = None, ultimo_corte: str = None) -> list[dict]:
    """
    Retorna dos mensajes listos para enviar:

      [0] Introducción  — texto simple
      [1] Detalle       — un bloque por supervisor con mención @número
                          y la lista de sus vendedores debajo.
                          Incluye 'mentions' con todos los wa_id.

    Si no hay ausentes, retorna solo [0] con mensaje de "sin ausencias".

    ultimo_corte: valor leído del visual DAX, ej. "28/04/2026 10:00 AM".
                  Si se pasa, se usa como hora del corte en el encabezado.
                  Si es None, se usa la hora actual.
    """
    if fecha is None:
        fecha = date.today()

    ahora      = datetime.now()
    dias_es    = {0: 'Lunes', 1: 'Martes', 2: 'Miércoles', 3: 'Jueves',
                  4: 'Viernes', 5: 'Sábado', 6: 'Domingo'}
    nombre_dia = dias_es[ahora.weekday()]
    fecha_fmt  = fecha.strftime('%d/%m/%Y')

    # Hora del corte: la del visual PBI si está disponible, si no la actual
    if ultimo_corte:
        # Formato esperado: "28/04/2026 10:00 AM" → extraer solo "10:00 AM"
        partes = ultimo_corte.strip().split()
        hora_fmt = ' '.join(partes[1:]) if len(partes) >= 2 else ultimo_corte
    else:
        hora_fmt = ahora.strftime('%H:%M')

    ausentes   = obtener_ausentes(fecha)
    directorio = _cargar_directorio()

    if ausentes.empty:
        return [{'texto': (
            f'✅ *Vendedores sin consultas — PLANILLA*\n'
            f'Todos los vendedores Planilla registraron consulta el {fecha_fmt}.'
        ), 'mentions': []}]

    total = len(ausentes)

    # ── Mensaje 1: introducción ───────────────────────────────
    intro = (
        f'📋 *Vendedores sin consultas — PLANILLA*\n'
        f'📅 {nombre_dia} {fecha_fmt} — Corte {hora_fmt}\n'
        f'━━━━━━━━━━━━━━━━━━━━━\n'
        f'A continuación detallo a los *{total} vendedor{"es" if total != 1 else ""}* '
        f'Planilla que no tienen consultas hasta este corte:'
    )

    # ── Mensaje 2: detalle con menciones ─────────────────────
    lineas   = []
    mentions = []

    for zonal in ausentes['ZONAL'].unique():
        lineas.append(f'📍 *{zonal}*')
        df_z = ausentes[ausentes['ZONAL'] == zonal]

        for sup in df_z['SUPERVISOR'].unique():
            wa = _wa_id(sup, directorio)

            if wa:
                # Nombre en el texto + wa_id en mentions para la notificación push
                numero = wa.replace('@c.us', '')
                lineas.append(f'@{numero} *{sup}*')
                mentions.append(wa)
            else:
                logger.warning(f'Sin teléfono para supervisor: {sup!r}')
                lineas.append(f'👤 *{sup}*')

            vendedores = df_z[df_z['SUPERVISOR'] == sup]['VENDEDOR'].tolist()
            for vdd in vendedores:
                lineas.append(f'   • {vdd}')
            lineas.append('')

        lineas.append('━━━━━━━━━━━━━━━━━━━━━')

    # Quitar separador final duplicado si quedó línea en blanco antes
    while lineas and lineas[-1] == '':
        lineas.pop()
    if not lineas or lineas[-1] != '━━━━━━━━━━━━━━━━━━━━━':
        lineas.append('━━━━━━━━━━━━━━━━━━━━━')

    return [
        {'texto': intro,               'mentions': []},
        {'texto': '\n'.join(lineas),   'mentions': mentions},
    ]


def enviar_mensajes(wa, destino: str, fecha: date = None, ultimo_corte: str = None):
    """
    Genera y envía los mensajes al destino (grupo o número).
    `wa` debe ser una instancia de WhatsAppClient.
    ultimo_corte: valor leído del visual DAX para mostrar la hora exacta del corte.
    """
    mensajes = construir_mensajes(fecha, ultimo_corte=ultimo_corte)
    for msg in mensajes:
        if msg['mentions']:
            wa.send_mention(destino, msg['texto'], mentions=msg['mentions'])
        else:
            wa.send_text(destino, msg['texto'])
        time.sleep(3)
    logger.info(f'Reporte de ausencias enviado a {destino} ({len(mensajes)} mensajes)')


# ══════════════════════════════════════════════════════════════
# ENTRY POINT (imprime en consola sin enviar)
# ══════════════════════════════════════════════════════════════

if __name__ == '__main__':
    logging.basicConfig(level=logging.INFO, format='%(levelname)s - %(message)s')
    mensajes = construir_mensajes()
    for i, msg in enumerate(mensajes, 1):
        print(f'\n{"=" * 50}')
        print(f'MENSAJE {i}')
        if msg['mentions']:
            print(f'  mentions: {msg["mentions"]}')
        print(msg['texto'])
    print(f'\n{"=" * 50}')
    print(f'Total: {len(mensajes)} mensajes')
