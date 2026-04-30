@echo off
rem Lee la URL del reporte desde la variable de entorno PBI_REPORT_URL
rem Configurar PBI_REPORT_URL en las variables de entorno del sistema o en el .env

taskkill /F /IM chrome.exe >NUL 2>&1
timeout /t 3 /nobreak >NUL

if "%PBI_REPORT_URL%"=="" (
    echo ERROR: La variable de entorno PBI_REPORT_URL no esta definida.
    echo Definela en el .env o en las variables de entorno del sistema.
    exit /b 1
)

if "%CHROME_EXE%"=="" (
    set CHROME_EXE=C:\Program Files\Google\Chrome\Application\chrome.exe
)

if "%CHROME_USER_DIR%"=="" (
    set CHROME_USER_DIR=C:\chrome_pbi_session
)

"%CHROME_EXE%" --remote-debugging-port=9222 --user-data-dir="%CHROME_USER_DIR%" --start-maximized "%PBI_REPORT_URL%"
