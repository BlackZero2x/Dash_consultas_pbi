# Dashboard Consultas Power BI — Publicación Automática

Sistema de automatización end-to-end que detecta la disponibilidad de nuevos datos de consultas Fija Movistar, actualiza un dataset de Power BI Service, captura el dashboard publicado y lo distribuye automáticamente al grupo de WhatsApp del equipo comercial, sin intervención manual.

---

## Tabla de contenidos

1. [Para qué sirve este sistema](#para-qué-sirve-este-sistema)
2. [Cómo funciona en el día a día](#cómo-funciona-en-el-día-a-día)
3. [Arquitectura general](#arquitectura-general)
4. [Módulos y responsabilidades](#módulos-y-responsabilidades)
5. [Estructura del proyecto](#estructura-del-proyecto)
6. [Instalación desde cero](#instalación-desde-cero)
7. [Configuración](#configuración)
8. [Ejecución manual](#ejecución-manual)
9. [Medida DAX: Ultimo Corte](#medida-dax-ultimo-corte)
10. [Configuración de Chrome para captura](#configuración-de-chrome-para-captura)
11. [Stack tecnológico](#stack-tecnológico)
12. [Glosario del dominio](#glosario-del-dominio)
13. [Preguntas frecuentes y solución de problemas](#preguntas-frecuentes-y-solución-de-problemas)

---

## Para qué sirve este sistema

Varias veces al día, a medida que llegan nuevos datos de consultas de la fuerza comercial de línea fija, el sistema:

- **Descarga los datos actualizados** desde SQL Server y Google Sheets (tabla RH).
- **Actualiza el dataset de Power BI Service** para que el reporte refleje los datos más recientes.
- **Captura automáticamente la página ZONAL_SUPERVISOR** del dashboard publicado.
- **Envía la imagen al grupo de WhatsApp** del equipo de gestión, con el texto del "Ultimo Corte" leído directamente del visual del reporte.

Esto elimina la necesidad de que alguien entre a Power BI, tome una captura de pantalla y la publique manualmente.

---

## Cómo funciona en el día a día

El proceso se dispara automáticamente con un correo electrónico:

```
10:00–18:00 (lun–vie)
    Programador de Tareas Windows lanza push_dash_pbi_consultas.py
    cada 10 minutos
          │
          ▼
    Gmail API verifica si hay un correo NO LEÍDO de:
       De:     e@auren.com.pe
       Asunto: dito_consultas_hoy_csv - Actualización disponible
          │
          ├── Sin correo → el script termina sin hacer nada
          │
          └── Correo encontrado →
                  Paso 1: extraer_datos.py
                      └─ SQL Server (eAuren) → BASE_CON.csv en OneDrive
                      └─ Google Sheets (RH publicado como CSV) → RH.csv en OneDrive
                  │
                  Paso 2: actualizar_pbi.py
                      └─ Service Principal → Power BI REST API → refresh del dataset
                  │
                  Paso 3: esperar que el refresh complete (polling cada 30s, timeout 10min)
                  │
                  Paso 4: capturar_pbi.py
                      └─ Chrome con debugging remoto (puerto 9222)
                      └─ Playwright CDP → recarga la página → captura PNG recortada
                      └─ Lee "Ultimo Corte" del DOM del visual
                  │
                  Paso 5: wa_client.py → wa_server.js (puerto 8002)
                      └─ Envía imagen PNG + texto al grupo de WhatsApp
                  │
                  Paso 6: Gmail API marca el correo como LEÍDO
```

---

## Arquitectura general

```
Fuentes de datos
    │
    ├── SQL Server eAuren (tabla fija_base_dito_consultas_hoy)
    └── Google Sheets (tabla RH publicada como CSV)
    │
    ▼
extraer_datos.py ── ETL de extracción y transformación
    │
    ▼
OneDrive / datalake_processed_auren/
    ├── BASE_CON.csv   ← dataset principal (consultas del mes actual + anterior)
    └── RH.csv         ← recursos humanos activos en campo
    │
    ▼
actualizar_pbi.py ── Power BI REST API (Service Principal)
    │
    ▼
Power BI Service ── Dataset actualizado
    │
    ▼
capturar_pbi.py ── Chrome con debugging remoto (puerto 9222)
    │   └── _captura_worker.py (subproceso aislado con timeout de 90s)
    │         └── Playwright CDP → recarga → screenshot → lee DOM
    │
    ▼
capturas/
    ├── zonal_supervisor_{YYYYMMDD_HHMMSS}.png   ← imagen capturada
    └── zonal_supervisor_{YYYYMMDD_HHMMSS}.txt   ← texto "Ultimo Corte"
    │
    ▼
wa_client.py (Python) → HTTP POST localhost:8002
    │
    ▼
wa_server.js (Node.js + open-wa) → WhatsApp Web
    │
    ▼
Grupo "Canal Fija 2026 Gestión AUREN"
```

---

## Módulos y responsabilidades

### `push_dash_pbi_consultas.py` — Orquestador principal

Es el punto de entrada del sistema. Corre una sola vez por ejecución (el Programador de Tareas lo invoca cada 10 minutos).

**Flujo:**
1. Autentica con Gmail API (OAuth2 con token persistido en `token.json`)
2. Busca correos no leídos del remitente/asunto configurados
3. Si no hay correo: termina con `sys.exit(0)` sin hacer nada
4. Si hay correo: ejecuta los 5 pasos del pipeline y marca el correo como leído

**Constantes configurables en el archivo:**

| Constante | Valor por defecto | Descripción |
|-----------|-------------------|-------------|
| `REMITENTE` | `e@auren.com.pe` | Remitente del correo trigger |
| `ASUNTO` | `dito_consultas_hoy_csv - Actualización disponible` | Asunto exacto del correo |
| `WA_GRUPO` | `Canal Fija 2026 Gestión AUREN` | Grupo de WhatsApp destino |
| `REFRESH_POLL_INTERVAL` | `30` | Segundos entre consultas de estado del refresh |
| `REFRESH_TIMEOUT` | `600` | Segundos máximos esperando el refresh (10 min) |

---

### `extraer_datos.py` — ETL de extracción

Extrae datos de dos fuentes, aplica las mismas transformaciones que tenía Power Query, y guarda los archivos CSV en OneDrive para que Power BI los lea.

**Secciones del script:**

| # | Sección | Descripción |
|---|---------|-------------|
| 1 | **Configuración** | Define conexiones SQL, rutas OneDrive y URL de Google Sheets |
| 2 | **Fuente SQL Server** | Ejecuta query con `UNION ALL`: registros normales del mes actual y anterior + filas sintéticas (ventas RUS sin intención en el mismo día, forzadas como tipo `INTENCIONES`) |
| 3 | **Fuente RH** | Descarga tabla RH desde Google Sheets publicado como CSV. Filtra `ACTIVO` + `EN CAMPO`, normaliza esquema (`COM`/`PLA`), renombra `ZONA → ZONAL`, reemplaza `MOQUEGUA → ILO`, calcula antigüedad en tramos |
| 4 | **Join y columnas derivadas** | LEFT JOIN de `BASE_CON` con `RH` por `documento_v = DNI`. Los nombres de RH tienen prioridad sobre los del SQL |
| 5 | **Guardar en OneDrive** | Escribe `BASE_CON.csv` y `RH.csv` con encoding `utf-8-sig` |

**Tramos de antigüedad calculados:**

| Tramo | Condición |
|-------|-----------|
| `<15d` | Hasta 15 días desde fecha de ingreso |
| `>15d` | Más de 15 días |
| `>30d` | Más de 30 días |
| `>60d` | Más de 60 días |
| `>90d` | Más de 90 días |

---

### `actualizar_pbi.py` — Refresh de Power BI

Autentica con Azure AD como Service Principal y dispara la actualización del dataset. No requiere intervención manual ni MFA.

**Flujo de autenticación (Client Credentials):**
```
cliente Python → POST https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token
                       con client_id + client_secret
              ← access_token (válido 1 hora, cacheado en memoria)
cliente Python → POST https://api.powerbi.com/.../datasets/{id}/refreshes
              ← HTTP 202 Accepted
```

**Funciones exportadas:**

| Función | Descripción |
|---------|-------------|
| `triggerear_refresh()` | Dispara el refresh, retorna `True` si HTTP 202 |
| `obtener_estado_refresh()` | Consulta el último refresh; retorna `'Completed'`, `'Failed'`, `'InProgress'` o `None` |
| `obtener_access_token()` | Obtiene/renueva el token OAuth2, usa caché en memoria |

**Requisito previo (configurado por el admin de Azure):**
- App registrada en Azure AD con client secret
- "Permitir a las entidades de servicio usar las API de Power BI" habilitado en el Portal de administración de Power BI
- La app agregada como Miembro o Colaborador en el workspace

---

### `capturar_pbi.py` — Captura del dashboard

Coordina la captura de pantalla del dashboard. **No importa Playwright directamente** — toda la lógica de browser está aislada en `_captura_worker.py` que se lanza como subproceso con un timeout de 90 segundos. Esto evita que un cuelgue de Playwright o Chrome bloquee el pipeline indefinidamente.

**Constantes de configuración:**

| Constante | Valor | Descripción |
|-----------|-------|-------------|
| `RENDER_WAIT` | `10` | Segundos adicionales esperando renderizado de visuals tras la carga |
| `WORKER_TIMEOUT` | `90` | Segundos máximos para que el worker complete antes de ser matado |
| `DIAS_RETENSION` | `7` | Días que se conservan las capturas (PNG + TXT) |
| `FILTRO_HORA_INICIO` | `12` | Hora a partir de la cual se aplica filtro de fecha=hoy en la URL |
| `CLIP_X` | `265` | Margen izquierdo del recorte (panel nav PBI + barra lateral) |
| `CLIP_Y` | `100` | Margen superior del recorte (barra Chrome + barra PBI Service) |
| `CLIP_W` | `1350` | Ancho de la región capturada |
| `CLIP_H` | `640` | Alto de la región capturada |

**Lógica del filtro de fecha:**
- Antes del mediodía: URL sin filtro (datos históricos disponibles)
- Desde el mediodía: URL con `?filter=BASE_CON/fecha_registro eq 'YYYY-MM-DD'` (solo datos de hoy)

---

### `_captura_worker.py` — Worker de Playwright (subproceso aislado)

Recibe 7 argumentos por línea de comando y comunica el resultado a través de stdout como JSON.

**Interfaz:**
```
python _captura_worker.py <dest_png> <report_url> <render_wait>
                          <clip_x> <clip_y> <clip_w> <clip_h>

Salida OK:    {"ok": true,  "ultimo_corte": "28/04/2026 10:00 AM"}
Salida error: {"ok": false, "error": "mensaje de error"}
```

**Lógica de navegación:**

| Situación | Acción |
|-----------|--------|
| No hay tab con `app.powerbi.com` abierto | Abre nueva pestaña con `page.goto()` + sleep 15s |
| Tab existe pero está en una página diferente | Navega con `page.goto()` + sleep 10s |
| Tab ya está en la página correcta | Recarga con `page.reload()` + sleep 20s |

Tras la navegación: `time.sleep(render_wait)` → lee `ultimo_corte` del DOM → captura con `Page.captureScreenshot` vía CDP → guarda PNG → guarda TXT sidecar.

**Lectura de Ultimo Corte:**
Busca en el DOM textos con formato `DD/MM/YYYY` que tengan al menos 16 caracteres y contengan `AM` o `PM`. Este patrón coincide con el formato de la medida DAX `Ultimo Corte`.

---

### `test_manual.py` — Ejecución manual del pipeline

Permite probar cualquier paso del pipeline de forma aislada o ejecutar el flujo completo sin esperar el correo trigger.

**Modos de uso:**

```bash
python test_manual.py                              # pipeline completo
python test_manual.py --solo-extraccion            # solo extracción SQL + RH
python test_manual.py --solo-refresh               # solo refresh PBI + esperar
python test_manual.py --solo-captura               # verifica Chrome + captura
python test_manual.py --solo-whatsapp <ruta.png>   # envía imagen existente por WA
```

El modo `--solo-whatsapp` busca automáticamente el archivo `.txt` sidecar (mismo nombre que el PNG) para obtener el "Ultimo Corte" real. Si no existe el `.txt`, envía la imagen sin ese dato.

---

### `run_pipeline.bat` — Lanzador para el Programador de Tareas

Script `.bat` invocado por Windows Task Scheduler. Establece el directorio de trabajo, fuerza encoding UTF-8 y redirige toda la salida al log diario.

```bat
cd /d "C:\proyectos\Dash_consultas_pbi"
set PYTHONIOENCODING=utf-8
"C:\proyectos\.venv\Scripts\python.exe" -u push_dash_pbi_consultas.py >> logs\pipeline.log 2>&1
```

---

### `arranque_chrome.bat` — Inicio de Chrome con debugging remoto

Mata cualquier instancia de Chrome abierta, espera 3 segundos y arranca una nueva instancia con debugging remoto habilitado en el puerto 9222. Abre directamente el reporte de Power BI.

```bat
taskkill /F /IM chrome.exe
timeout /t 3
chrome.exe --remote-debugging-port=9222
           --user-data-dir="C:\chrome_pbi_session"
           --start-maximized
           "https://app.powerbi.com/..."
```

Este `.bat` se configura en el Programador de Tareas para ejecutarse a las **9:50 AM** cada día hábil, antes de que llegue el primer correo trigger.

---

### `cerrar_chrome.bat` — Cierre de Chrome al final del día

Mata el proceso de Chrome al final de la jornada para liberar memoria y garantizar una sesión limpia al día siguiente. Configurar en el Programador de Tareas a las **19:00** cada día hábil.

```bat
taskkill /F /IM chrome.exe
```

---

## Estructura del proyecto

```
Dash_consultas_pbi/
│
├── push_dash_pbi_consultas.py   # Orquestador: trigger email → pipeline completo
├── extraer_datos.py             # ETL: SQL Server + Google Sheets → CSV en OneDrive
├── actualizar_pbi.py            # Power BI REST API: autenticación + refresh dataset
├── capturar_pbi.py              # Coordinador de captura (lanza worker como subproceso)
├── _captura_worker.py           # Worker Playwright/CDP: recarga, screenshot, DOM
├── test_manual.py               # Ejecución manual de pasos individuales o pipeline completo
├── run_pipeline.bat             # Lanzador .bat para el Programador de Tareas
├── arranque_chrome.bat          # Arranca Chrome con debugging remoto (9:50 AM)
├── cerrar_chrome.bat            # Mata Chrome al cierre del día (19:00)
│
├── requirements.txt             # Dependencias Python
├── .env                         # Variables sensibles — NO incluido en git
├── .env.example                 # Plantilla de configuración sin datos reales
├── credentials.json             # Credenciales OAuth2 Gmail — NO incluido en git
├── token.json                   # Token Gmail persistido — NO incluido en git
│
├── capturas/                    # Imágenes y archivos de texto generados (NO en git)
│   ├── zonal_supervisor_{YYYYMMDD_HHMMSS}.png
│   └── zonal_supervisor_{YYYYMMDD_HHMMSS}.txt   ← "Ultimo Corte" del visual
│
└── logs/                        # Logs de ejecución diarios (NO en git)
    ├── pipeline_{YYYYMMDD}.log
    └── pipeline.log             ← log acumulativo del .bat
```

---

## Instalación desde cero

### Prerrequisitos del sistema

- **Windows 10/11**
- **Python 3.10+** — se recomienda usar el entorno virtual compartido `C:\proyectos\.venv\`
- **Node.js 18+** — para el servidor WhatsApp (compartido con AVANCE_MOVISTAR)
- **ODBC Driver 17 for SQL Server** — para la conexión a la base de datos
- **Google Chrome** instalado (se detecta automáticamente en rutas estándar)
- Acceso de red al servidor SQL corporativo `AUREN22\AUREN`
- Cuenta de Google con las APIs de Gmail habilitadas
- WhatsApp activo en el teléfono (sesión del servidor en `AVANCE_MOVISTAR\whatsapp_server\`)

### Paso 1 — Clonar el repositorio

```bash
git clone https://github.com/BlackZero2x/Dash_consultas_pbi.git
cd Dash_consultas_pbi
```

### Paso 2 — Activar entorno virtual Python

Este proyecto usa el entorno virtual compartido del directorio padre:

```bash
C:\proyectos\.venv\Scripts\activate
pip install -r requirements.txt
```

Para instalar Playwright y descargar el navegador Chromium:

```bash
pip install playwright
playwright install chromium
```

### Paso 3 — Crear archivo `.env`

```bash
copy .env.example .env
```

Editar `.env` con los valores reales (ver sección [Configuración](#configuración)).

### Paso 4 — Autenticar Gmail (primera vez)

En la primera ejecución, el script abrirá una ventana del navegador para autorizar el acceso a Gmail. Una vez autorizado, el token se guarda en `token.json` y las ejecuciones siguientes son silenciosas.

```bash
python push_dash_pbi_consultas.py
```

### Paso 5 — Verificar conexión a Power BI

```bash
python actualizar_pbi.py
```

Debe mostrar `Refresh de Power BI iniciado correctamente.` si las credenciales en `.env` son válidas.

### Paso 6 — Verificar captura del dashboard

Asegurarse de que Chrome está abierto con debugging remoto (ejecutar `arranque_chrome.bat` manualmente), luego:

```bash
python test_manual.py --solo-captura
```

Debe generar un archivo PNG en `capturas/` y un `.txt` con el "Ultimo Corte".

### Paso 7 — Verificar envío WhatsApp

```bash
python test_manual.py --solo-whatsapp capturas\zonal_supervisor_XXXXXX_XXXXXX.png
```

### Paso 8 — Configurar el Programador de Tareas de Windows

Crear tres tareas programadas:

**Tarea 1: Arranque de Chrome**
- Disparador: cada día hábil a las **9:50 AM**
- Acción: ejecutar `arranque_chrome.bat`
- Propósito: garantizar que Chrome esté activo con debugging remoto antes del primer trigger

**Tarea 2: Pipeline principal**
- Disparador: cada día hábil, cada 10 minutos entre las **10:00** y las **18:00**
- Acción: ejecutar `run_pipeline.bat`
- Marcar "Ejecutar tanto si el usuario inició sesión como si no"
- Marcar "Ejecutar con los privilegios más altos"

**Tarea 3: Cierre de Chrome**
- Disparador: cada día hábil a las **19:00**
- Acción: ejecutar `cerrar_chrome.bat`
- Propósito: liberar memoria y evitar caché acumulada entre jornadas

---

## Configuración

### Variables de entorno (`.env`)

| Variable | Descripción | Ejemplo |
|----------|-------------|---------|
| `PBI_CLIENT_ID` | App registration de Azure AD | `xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx` |
| `PBI_CLIENT_SECRET` | Secreto de la app de Azure AD | `xxxxxx~xxxxxx` |
| `PBI_TENANT_ID` | Tenant ID de Azure AD | `xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx` |
| `PBI_WORKSPACE_ID` | ID del workspace de Power BI | `4431c026-df58-4f9c-9630-bb40072f829a` |
| `PBI_DATASET_ID` | ID del dataset a refrescar | `xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx` |
| `SQL_SERVER` | Servidor SQL (default: `AUREN22\AUREN`) | `AUREN22\AUREN` |
| `SQL_DATABASE` | Base de datos (default: `eAuren`) | `eAuren` |
| `SQL_USER` | Usuario SQL (opcional, usa Windows auth si no se define) | `eauren` |
| `SQL_PASSWORD` | Contraseña SQL | `****` |
| `ONEDRIVE_PATH` | Ruta al datalake en OneDrive (auto-detectada si no está) | `C:\Users\...\OneDrive - grupoauren.pe\datalake_processed_auren` |
| `CHROME_EXE` | Ruta a Chrome (auto-detectada si no está) | `C:\Program Files\Google\Chrome\Application\chrome.exe` |
| `SCREENSHOT_DIR` | Directorio para capturas (default: `capturas/`) | `C:\proyectos\Dash_consultas_pbi\capturas` |
| `HTTPS_PROXY` | Proxy corporativo (opcional) | `http://proxy.empresa.com:8080` |
| `PROXY_USER` | Usuario del proxy (opcional) | `usuario` |
| `PROXY_PASS` | Contraseña del proxy (opcional) | `****` |

### Constantes del reporte Power BI

Definidas directamente en `capturar_pbi.py` (no en `.env`, ya que identifican el reporte específico):

| Constante | Valor actual | Descripción |
|-----------|-------------|-------------|
| `WORKSPACE_ID` | `4431c026-...` | ID del workspace donde está el reporte |
| `REPORT_ID` | `2ab7921e-...` | ID del reporte de Power BI |
| `PAGE_ID` | `aa060e7c3b795463e58a` | ID de la página ZONAL_SUPERVISOR |
| `PAGE_NAME` | `ZONAL_SUPERVISOR` | Nombre de la página capturada |

---

## Ejecución manual

### Pipeline completo sin esperar correo

```bash
python test_manual.py
```

### Pasos individuales

```bash
# Solo extracción de datos (SQL + RH → CSV)
python test_manual.py --solo-extraccion

# Solo refresh del dataset y esperar que complete
python test_manual.py --solo-refresh

# Solo verificar Chrome + captura del dashboard
python test_manual.py --solo-captura

# Reenviar una captura existente por WhatsApp
python test_manual.py --solo-whatsapp "C:\proyectos\Dash_consultas_pbi\capturas\zonal_supervisor_20260430_120411.png"
```

### Ejecutar scripts individuales

```bash
# Extraer datos únicamente
python extraer_datos.py

# Disparar refresh manualmente
python actualizar_pbi.py

# Capturar el dashboard
python capturar_pbi.py
```

---

## Medida DAX: Ultimo Corte

El dashboard de Power BI expone la medida `Ultimo Corte` que calcula la fecha/hora del último registro redondeada al corte de hora más cercano (10am, 12pm, 2pm, 4pm). Esta medida es la que el worker lee del DOM para incluirla en el mensaje de WhatsApp.

```dax
Ultimo Corte = 
VAR _MaxFecha = CALCULATE(MAX(BASE_CON[fecha_completa]))
VAR _Hora = HOUR(_MaxFecha)
VAR _Minuto = MINUTE(_MaxFecha)
VAR _HoraDecimal = _Hora + _Minuto / 60

VAR _HoraCorte = 
    SWITCH(
        TRUE(),
        _HoraDecimal >= 9  && _HoraDecimal < 11, 10,
        _HoraDecimal >= 11 && _HoraDecimal < 13, 12,
        _HoraDecimal >= 13 && _HoraDecimal < 15, 14,
        _HoraDecimal >= 15 && _HoraDecimal < 17, 16,
        _Hora
    )

VAR _FechaCorte = 
    DATE(YEAR(_MaxFecha), MONTH(_MaxFecha), DAY(_MaxFecha)) + 
    TIME(_HoraCorte, 0, 0)

RETURN
    FORMAT(_FechaCorte, "DD/MM/YYYY HH:MM:SS AM/PM")
```

El worker busca en el DOM textos con el patrón `DD/MM/YYYY` que tengan al menos 16 caracteres y contengan `AM` o `PM`, que es exactamente el formato que produce esta medida.

---

## Configuración de Chrome para captura

### Por qué se necesita Chrome con debugging remoto

Playwright no abre un nuevo navegador, sino que se **conecta** a la instancia de Chrome que ya tiene la sesión iniciada de Power BI Service. Esto evita tener que gestionar autenticación de Microsoft/Power BI dentro del script.

### Arranque manual

```bat
"C:\Program Files\Google\Chrome\Application\chrome.exe"
    --remote-debugging-port=9222
    --user-data-dir="C:\chrome_pbi_session"
    --start-maximized
    "https://app.powerbi.com/..."
```

- `--remote-debugging-port=9222`: habilita la API de Chrome DevTools Protocol (CDP)
- `--user-data-dir`: directorio de perfil separado para mantener la sesión de Power BI aislada del Chrome personal

### Verificar que Chrome está activo

El script detecta Chrome intentando conectar al puerto 9222 via socket:

```python
import socket
with socket.create_connection(('127.0.0.1', 9222), timeout=1):
    print('Chrome activo')
```

Si no está activo, `capturar_pbi.py` lo arranca automáticamente. También puede verificarse con:

```bash
python test_manual.py --solo-captura
```

### Viewport y recorte de captura

La captura se hace con CDP `Page.captureScreenshot` con un recorte calibrado para eliminar la barra de Chrome, la barra de navegación de Power BI Service y los paneles laterales:

```
Viewport total Chrome: 1440 x 860 px (con --start-maximized en 1440x860)

Recorte aplicado:
    CLIP_X =  265 px  (margen izq: panel nav PBI + barra lateral)
    CLIP_Y =  100 px  (margen sup: barra Chrome + barra PBI Service)
    CLIP_W = 1350 px  (ancho región capturada)
    CLIP_H =  640 px  (alto región capturada)
```

Si se cambia la resolución del monitor o el zoom de Chrome, estos valores deben recalibrarse en `capturar_pbi.py`.

---

## Stack tecnológico

| Capa | Tecnología | Por qué se usa |
|------|-----------|----------------|
| Orquestación y ETL | Python 3.10, pandas 2.x | Manipulación de DataFrames, joins, transformaciones de datos |
| Acceso a base de datos | SQL Server + pyodbc | Fuente principal de registros de consultas Fija |
| Autenticación Gmail | google-auth-oauthlib, google-api-python-client | OAuth2 para leer y marcar correos en Gmail |
| Autenticación Power BI | requests + Client Credentials OAuth2 | Service Principal para refresh desatendido sin MFA |
| Captura de dashboard | Playwright (sync API) + CDP | Controla Chrome ya autenticado con la sesión de Power BI Service |
| Aislamiento del worker | subprocess + timeout | Evita que un cuelgue de Playwright/Chrome bloquee el pipeline |
| Servidor WhatsApp | Node.js 18 + @open-wa/wa-automate + Express | API REST local que encapsula la automatización de WhatsApp Web |
| Cliente WhatsApp | Python + requests | Llamadas HTTP al servidor Node.js compartido con AVANCE_MOVISTAR |
| Variables sensibles | python-dotenv | Separar credenciales del código fuente |
| Programación de tareas | Windows Task Scheduler | Ejecución automática cada 10 minutos en horario hábil |
| Logging | Python logging | Registro diario con timestamps para auditoría y debugging |

---

## Glosario del dominio

| Término | Significado |
|---------|-------------|
| **BASE_CON** | Dataset principal en Power BI; tabla de consultas de vendedores Fija Movistar |
| **RH** | Recursos Humanos — tabla que mapea DNI → vendedor → supervisor → zona |
| **Ultimo Corte** | Fecha/hora del dato más reciente redondeada al corte horario más cercano (10am, 12pm, 2pm, 4pm) |
| **ZONAL_SUPERVISOR** | Página del dashboard de Power BI que muestra el resumen por zonal y supervisor |
| **eAuren** | Base de datos SQL Server donde se almacenan las consultas de la fuerza de ventas |
| **fija_base_dito_consultas_hoy** | Vista/tabla SQL que contiene los registros del mes actual y anterior |
| **INTENCIONES** | Tipo de registro que representa una consulta/intención de venta |
| **RUS** | Tipo de registro que representa una venta cerrada |
| **Filas sintéticas** | Filas tipo `INTENCIONES` generadas artificialmente para vendedores con `RUS` pero sin intención registrada ese día |
| **Service Principal** | App registrada en Azure AD que actúa como identidad de aplicación para Power BI REST API |
| **CDP** | Chrome DevTools Protocol — API de bajo nivel para controlar Chrome remotamente |
| **Sidecar TXT** | Archivo `.txt` con el mismo nombre que el PNG de captura, que contiene el valor de "Ultimo Corte" leído del DOM |
| **datalake_processed_auren** | Carpeta en OneDrive donde se depositan los CSV procesados para que Power BI los consuma |

---

## Preguntas frecuentes y solución de problemas

### El pipeline no se ejecuta aunque llegó el correo

1. Verificar que el correo tiene exactamente el remitente y asunto configurados en `push_dash_pbi_consultas.py`.
2. Revisar el log del día: `logs\pipeline_{YYYYMMDD}.log`.
3. Verificar que el Programador de Tareas ejecutó `run_pipeline.bat` sin errores.
4. Probar manualmente: `python push_dash_pbi_consultas.py`.

### `ModuleNotFoundError` al ejecutar el script

El entorno virtual no está activado, o los paquetes no están instalados:

```bash
C:\proyectos\.venv\Scripts\activate
pip install -r requirements.txt
```

### Error de autenticación Gmail ("invalid_grant" o "Token has been expired")

El `token.json` expiró o fue revocado. Borrar el archivo y volver a autenticar:

```bash
del token.json
python push_dash_pbi_consultas.py
```

Se abrirá una ventana del navegador para volver a autorizar.

### Error de Power BI: "Unauthorized" o "Forbidden"

1. Verificar que `PBI_CLIENT_ID`, `PBI_CLIENT_SECRET` y `PBI_TENANT_ID` en `.env` son correctos.
2. Verificar en el Portal de administración de Power BI que "Permitir a las entidades de servicio usar las API de Power BI" está habilitado.
3. Verificar que la app está agregada como Miembro en el workspace en Power BI Service.

### La captura muestra el dashboard en blanco o incompleto

1. `RENDER_WAIT` puede ser insuficiente. Incrementar a 15–20 en `capturar_pbi.py`.
2. Chrome puede no haber cargado completamente el reporte. Verificar ejecutando `arranque_chrome.bat` manualmente y esperando que el reporte cargue.
3. La sesión de Power BI puede haber expirado. Abrir Chrome manualmente, navegar al reporte e iniciar sesión.

### La captura se congela / "Worker de captura superó el timeout de 90s"

Puede ocurrir si Chrome está bloqueado o el tab de Power BI tiene un error de carga.

1. Ejecutar `cerrar_chrome.bat` para matar Chrome.
2. Ejecutar `arranque_chrome.bat` para un arranque limpio.
3. Reintentar con `python test_manual.py --solo-captura`.

Si ocurre frecuentemente a primera hora del día, es señal de caché acumulada. Asegurarse de que `cerrar_chrome.bat` está configurado en el Programador de Tareas a las 19:00.

### `--solo-whatsapp` envía la hora actual en lugar del "Ultimo Corte" real

El archivo `.txt` sidecar no existe para esa imagen (fue capturada antes de que se implementara la función, o fue eliminado).

Opciones:
- Usar `python test_manual.py --solo-captura` para generar una captura nueva con su `.txt`.
- Crear el `.txt` manualmente con el contenido correcto (ej: `28/04/2026 10:00 AM`).

### El servidor WhatsApp no está disponible

1. Verificar que el servidor de AVANCE_MOVISTAR está activo: `python C:\proyectos\AVANCE_MOVISTAR\whatsapp_server\wa_client.py --health`
2. Si no responde, ejecutar `C:\proyectos\AVANCE_MOVISTAR\start_wa_server.bat`.
3. Revisar si la sesión de WhatsApp expiró (puede requerir re-escanear el QR).

---

## Licencia

Uso interno — Grupo Auren. El código es de autoría propia. Los archivos de credenciales, tokens y capturas generadas no están incluidos en este repositorio.
