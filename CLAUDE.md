# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

  Antes de escribir o modificar cualquier código de este proyecto, aplica
  siempre estas reglas de confidencialidad:

  1. CREDENCIALES: Nunca escribas valores reales (contraseñas, usuarios,
     servidores, tokens, API keys) directamente en el código. Toda credencial
     debe leerse desde variables de entorno o un archivo .env, sin valores
     por defecto que revelen datos reales. Si falta una variable requerida,
     el script debe abortar con un mensaje de error claro.

  2. DATOS PERSONALES: No incluyas nombres reales de personas, correos,
     teléfonos ni códigos de clientes en el código fuente. Usa variables
     de entorno o archivos de configuración externos.

  3. INFRAESTRUCTURA: Evita hardcodear nombres de servidores, bases de datos,
     DSNs, rutas de red internas o convenciones de codificación internas
     (como tipos de venta o estados) en lugares visibles del código. Si son
     necesarios para el funcionamiento, centralízalos en un bloque de
     configuración claramente marcado como "ajustar en cada entorno".

  4. ANTES DE VERSIONAR: Cuando vayas a preparar código para Git, revisa
     activamente si hay credenciales, datos de clientes o información de
     infraestructura interna que deba moverse a .env o eliminarse.

## Propósito

Este repositorio contiene la automatización para actualizar y publicar un dashboard de Power BI. El flujo es:

1. `push_dash_pbi_consultas.py` monitorea una bandeja de Gmail en busca de un correo específico (remitente + asunto configurados como constantes).
2. Al detectar el correo, ejecuta un script de actualización de datos y luego un script de publicación en Power BI.
3. Marca el correo como leído y repite el ciclo cada 5 minutos.

## Dependencias

Usa el entorno virtual compartido en `C:\proyectos\.venv\`. Paquetes en `C:\proyectos\requirements.txt`.

```bash
# Activar entorno
C:\proyectos\.venv\Scripts\activate
```

> Nota: `google.api_python_client` no es un paquete real — el paquete correcto es `google-api-python-client` y se importa como `googleapiclient.discovery`.

## Configuración

Las variables sensibles están en `.env` (nunca en el código):

| Variable | Descripción |
|---|---|
| `PBI_CLIENT_ID` | App registration de Azure AD |
| `PBI_CLIENT_SECRET` | Secreto de la app de Azure AD |
| `PBI_TENANT_ID` | Tenant ID de Azure AD |
| `PBI_WORKSPACE_ID` | ID del workspace de Power BI |
| `PBI_DATASET_ID` | ID del dataset a refrescar |
| `SQL_SERVER` | Servidor SQL (default: `AUREN22\AUREN`) |
| `SQL_DATABASE` | Base de datos (default: `eAuren`) |
| `ONEDRIVE_PATH` | Ruta al datalake en OneDrive (auto-detectada si no está) |
| `CHROME_EXE` | Ruta a Chrome (auto-detectada si no está) |
| `SCREENSHOT_DIR` | Directorio para capturas (default: `capturas/` relativo al proyecto) |

Las constantes en `push_dash_pbi_consultas.py` que aún deben ajustarse:
- `REMITENTE`: correo del remitente del que se espera el trigger
- `ASUNTO`: asunto del correo que activa la actualización
- `CHECK_INTERVAL`: intervalo de polling en segundos (default: 300)

La autenticación con Gmail requiere un archivo de credenciales en `C:\proyectos\shared\credentials\credentials.json` con el scope de Gmail habilitado.

## Ejecución

```bash
python push_dash_pbi_consultas.py
```

El script corre en loop infinito; detenerlo con `Ctrl+C`.

## Medidas DAX relevantes (Power BI)

**Ultimo Corte** — devuelve la fecha/hora del último registro redondeada al corte de hora más cercano (10am, 12pm, 2pm, 4pm):

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
