# ══════════════════════════════════════════════════════════════
# EXTRAER_DATOS.PY
# Extrae datos de SQL Server y Google Sheets, aplica las mismas
# transformaciones que tenía Power Query, y guarda los archivos
# resultantes en la carpeta de OneDrive para que Power BI los lea.
#
# Archivos generados:
#   BASE_CON.csv  —  tabla principal con consultas del mes actual
#                    y el mes anterior, incluyendo filas sintéticas
#                    de vendedores RUS sin intenciones en el mismo día
#   RH.csv        —  tabla de recursos humanos activos en campo,
#                    con esquema normalizado (COM / PLA) y antigüedad
#
# Uso:
#   python extraer_datos.py
# ══════════════════════════════════════════════════════════════

import os
import io
import logging
from datetime import date
from pathlib import Path

import pandas as pd
import pyodbc
import requests
from dotenv import load_dotenv

load_dotenv(Path(__file__).parent / '.env')

logger = logging.getLogger(__name__)


def _detectar_onedrive() -> str:
    """Intenta localizar la carpeta OneDrive automáticamente si no está en .env."""
    _org = os.environ.get('ONEDRIVE_ORG_FOLDER', 'datalake_processed_auren')
    candidatos = [
        Path(os.environ.get('USERPROFILE', '')) / _org,
        Path(os.environ.get('ONEDRIVE', '')) / _org,
    ]
    for c in candidatos:
        if c.exists():
            logger.info(f'OneDrive detectado automáticamente: {c}')
            return str(c)
    raise EnvironmentError(
        'No se encontró la carpeta de OneDrive. '
        'Define ONEDRIVE_PATH en el archivo .env.'
    )


# ══════════════════════════════════════════════════════════════
# 1. CONFIGURACIÓN
# ══════════════════════════════════════════════════════════════

# ── SQL Server ─────────────────────────────────────────────────
SQL_SERVER   = os.environ.get('SQL_SERVER',   '')
SQL_DATABASE = os.environ.get('SQL_DATABASE', '')
SQL_USER     = os.environ.get('SQL_USER',     '')
SQL_PASSWORD = os.environ.get('SQL_PASSWORD', '')

# ── Destino OneDrive ───────────────────────────────────────────
ONEDRIVE_PATH = os.environ.get('ONEDRIVE_PATH') or _detectar_onedrive()
OUTPUT_FILE   = os.path.join(ONEDRIVE_PATH, 'BASE_CON.csv')
RH_FILE       = os.path.join(ONEDRIVE_PATH, 'RH.csv')

# ── Google Sheets (tabla RH publicada como CSV) ────────────────
RH_URL = os.environ['RH_SHEETS_URL']


# ══════════════════════════════════════════════════════════════
# 2. FUENTE SQL SERVER — tabla fija_base_dito_consultas_hoy
# ══════════════════════════════════════════════════════════════
#
# El query trae dos bloques unidos con UNION ALL:
#
#   Bloque 1 — registros normales del mes actual y el anterior.
#              Incluye todos los tipos (INTENCIONES, RUS, etc.).
#
#   Bloque 2 — filas sintéticas: vendedores con venta RUS que NO
#              tienen ninguna intención registrada en el mismo día.
#              Se agregan como tipo "INTENCIONES" para que el
#              dashboard no los excluya del conteo de intenciones.

QUERY = """
SELECT
       [periodo]
      ,[tipo]
      ,[fecha_registro]
      ,[fecha_completa]
      ,FORMAT([hora], '00') AS hora
      ,[cms_codsrv]
      ,[documento_v]
      ,[zonal_consulta] as ZONAL
      ,[consulta_unica]
      ,[nombrecompleto_vendedor]
      ,[nombre_supervisor]
      ,[nombre_jefe]
      ,CASE
      WHEN [tipo_contrato] IS NULL THEN 'PLA'
      WHEN [nombre_supervisor] = 'ND' THEN 'PLA' ELSE SUBSTRING(SUBSTRING([tipo_contrato],12,12),1,3)
      END AS ESQ
      ,Q = 1
    FROM [eAuren].[dbo].[fija_base_dito_consultas_hoy]
    WHERE periodo IN (FORMAT(GETDATE(), 'yyyy-MM'),FORMAT(DATEADD(MONTH, -1, GETDATE()), 'yyyy-MM'))

UNION ALL

-- Filas sintéticas: ventas (RUS) sin intenciones en el mismo día
SELECT DISTINCT
       FORMAT(GETDATE(), 'yyyy-MM') AS periodo
      ,'INTENCIONES' AS tipo
      ,[fecha_registro]
      ,[fecha_completa]
      ,FORMAT([hora], '00') AS hora
      ,[cms_codsrv]
      ,[documento_v]
      ,[zonal_consulta] as ZONAL
      ,1 AS consulta_unica
      ,[nombrecompleto_vendedor]
      ,[nombre_supervisor]
      ,[nombre_jefe]
      ,CASE
      WHEN [tipo_contrato] IS NULL THEN 'PLA'
      WHEN [nombre_supervisor] = 'ND' THEN 'PLA' ELSE SUBSTRING(SUBSTRING([tipo_contrato],12,12),1,3)
      END AS ESQ
      ,Q = 1
FROM [eAuren].[dbo].[fija_base_dito_consultas_hoy] vr
WHERE vr.[tipo] = 'RUS'
  AND vr.[fecha_registro] NOT IN (
      SELECT DISTINCT [fecha_registro]
      FROM [eAuren].[dbo].[fija_base_dito_consultas_hoy]
      WHERE [documento_v] = vr.[documento_v]
        AND [tipo] = 'INTENCIONES'
        AND [fecha_registro] = vr.[fecha_registro]
  )
  AND periodo IN (FORMAT(GETDATE(), 'yyyy-MM'),FORMAT(DATEADD(MONTH, -1, GETDATE()), 'yyyy-MM'))
"""


def _conn_str():
    if SQL_USER:
        return (
            f'DRIVER={{ODBC Driver 17 for SQL Server}};'
            f'SERVER={SQL_SERVER};DATABASE={SQL_DATABASE};'
            f'UID={SQL_USER};PWD={SQL_PASSWORD}'
        )
    return (
        f'DRIVER={{ODBC Driver 17 for SQL Server}};'
        f'SERVER={SQL_SERVER};DATABASE={SQL_DATABASE};'
        f'Trusted_Connection=yes'
    )


def cargar_sql():
    logger.info('Conectando a SQL Server...')
    conn = pyodbc.connect(_conn_str())
    df = pd.read_sql(QUERY, conn)
    conn.close()

    df['fecha_registro'] = pd.to_datetime(df['fecha_registro']).dt.date
    df['documento_v']    = pd.to_numeric(df['documento_v'], errors='coerce').astype('Int64')
    df['consulta_unica'] = pd.to_numeric(df['consulta_unica'], errors='coerce').astype('Int64')
    df['fecha_completa'] = pd.to_datetime(df['fecha_completa'])

    logger.info(f'  {len(df)} filas obtenidas de SQL.')
    return df


# ══════════════════════════════════════════════════════════════
# 3. FUENTE RH — Google Sheets publicado como CSV
# ══════════════════════════════════════════════════════════════
#
# Transformaciones aplicadas (equivalentes a Power Query):
#   - Filtrar solo ESTADO = "ACTIVO" y feedback_rh = "EN CAMPO"
#   - Reemplazar "PART-TIME" por "PLANILLA" en esquema
#   - Tomar los primeros 3 caracteres → solo "COM" o "PLA"
#   - Renombrar columna ZONA → ZONAL
#   - Reemplazar "MOQUEGUA" por "ILO" en ZONAL
#   - Excluir supervisores con valor "0"
#   - Calcular antigüedad en tramos: <15d, >15d, >30d, >60d, >90d
#
# Se guarda como RH.csv en OneDrive (en vez de leer desde la web
# en Power BI) para evitar problemas de proxy en el refresh.

def _calcular_antig(f_ingreso):
    """Devuelve el tramo de antigüedad según días desde el ingreso."""
    if pd.isna(f_ingreso):
        return None
    dias = (date.today() - f_ingreso).days
    if dias <= 15:
        return '<15d'
    if dias > 90:
        return '>90d'
    if dias > 60:
        return '>60d'
    if dias > 30:
        return '>30d'
    if dias > 15:
        return '>15d'
    return None


def cargar_rh():
    logger.info('Descargando tabla RH desde Google Sheets...')
    proxy = os.environ.get('HTTPS_PROXY') or os.environ.get('HTTP_PROXY')
    if proxy:
        user = os.environ.get('PROXY_USER', '')
        pwd  = os.environ.get('PROXY_PASS', '')
        if user and pwd:
            from urllib.parse import quote_plus
            proxy = proxy.replace('http://', f'http://{quote_plus(user)}:{quote_plus(pwd)}@')
    proxies = {'http': proxy, 'https': proxy} if proxy else None
    try:
        resp = requests.get(RH_URL, timeout=30, proxies=proxies)
        resp.raise_for_status()
    except requests.RequestException as e:
        logger.error(f'Error descargando tabla RH: {e}')
        raise
    df = pd.read_csv(io.BytesIO(resp.content), encoding='utf-8')

    # ── Tipos ──────────────────────────────────────────────────
    df['DNI']        = pd.to_numeric(df['DNI'], errors='coerce').astype('Int64')
    df['F_INGRESO']  = pd.to_datetime(df['F_INGRESO'], dayfirst=True, errors='coerce').dt.date
    df['ZONA']       = df['ZONA'].astype(str)
    df['VENDEDOR']   = df['VENDEDOR'].astype(str)
    df['SUPERVISOR'] = df['SUPERVISOR'].astype(str)
    df['ESQUEMA']    = df['ESQUEMA'].astype(str)
    df['ESTADO']     = df['ESTADO'].astype(str)

    # ── Filtros ────────────────────────────────────────────────
    df = df[df['ESTADO'] == 'ACTIVO']
    df = df[(df['SUPERVISOR'] != '0') & (df['feedback_rh'] == 'EN CAMPO')]

    # ── Normalización de esquema ───────────────────────────────
    df['ESQUEMA'] = df['ESQUEMA'].str.replace('PART-TIME', 'PLANILLA', regex=False)
    df['ESQ'] = df['ESQUEMA'].str[:3]   # resulta en "COM" o "PLA"
    df = df.drop(columns=['ESQUEMA'])

    # ── Normalización de zonal ─────────────────────────────────
    df = df.rename(columns={'ZONA': 'ZONAL'})
    df['ZONAL'] = df['ZONAL'].str.replace('MOQUEGUA', 'ILO', regex=False)

    # ── Ordenar y calcular antigüedad ──────────────────────────
    df = df.sort_values('SUPERVISOR')
    df['ANTIG'] = df['F_INGRESO'].apply(_calcular_antig).astype(str)

    # ── Guardar CSV completo en OneDrive ───────────────────────
    os.makedirs(ONEDRIVE_PATH, exist_ok=True)
    df.to_csv(RH_FILE, index=False, encoding='utf-8-sig')
    logger.info(f'  RH guardado en: {RH_FILE}')
    logger.info(f'  {len(df)} filas en RH.')

    # Solo se exponen DNI, VENDEDOR y SUPERVISOR al join
    return df[['DNI', 'VENDEDOR', 'SUPERVISOR']].copy()


# ══════════════════════════════════════════════════════════════
# 4. JOIN Y COLUMNAS DERIVADAS
# ══════════════════════════════════════════════════════════════
#
# Se hace un LEFT JOIN de BASE_CON con RH por documento_v = DNI.
# Si el vendedor/supervisor aparece en RH, se usa ese nombre
# (más actualizado); si no, se cae al nombre que viene del SQL.
#
# Columnas resultantes:
#   VDD — nombre del vendedor (RH tiene prioridad sobre SQL)
#   SUP — nombre del supervisor (RH tiene prioridad sobre SQL)

def transformar(base: pd.DataFrame, rh: pd.DataFrame) -> pd.DataFrame:
    df = base.merge(rh, left_on='documento_v', right_on='DNI', how='left')

    df['VDD'] = df['VENDEDOR'].where(df['VENDEDOR'].notna(), df['nombrecompleto_vendedor'])
    df['SUP'] = df['SUPERVISOR'].where(df['SUPERVISOR'].notna(), df['nombre_supervisor'])

    df = df.drop(columns=['VENDEDOR', 'SUPERVISOR', 'DNI'], errors='ignore')
    df['VDD'] = df['VDD'].astype(str)
    return df


# ══════════════════════════════════════════════════════════════
# 5. GUARDAR EN ONEDRIVE
# ══════════════════════════════════════════════════════════════

def guardar(df: pd.DataFrame):
    os.makedirs(ONEDRIVE_PATH, exist_ok=True)
    df.to_csv(OUTPUT_FILE, index=False, encoding='utf-8-sig')
    logger.info(f'Archivo guardado en: {OUTPUT_FILE}')


# ══════════════════════════════════════════════════════════════
# 6. ENTRY POINT
# ══════════════════════════════════════════════════════════════

def ejecutar():
    base = cargar_sql()
    rh   = cargar_rh()
    df   = transformar(base, rh)
    guardar(df)
    logger.info('Extracción completada.')
    return df


if __name__ == '__main__':
    ejecutar()
