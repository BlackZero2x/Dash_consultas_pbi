@echo off
taskkill /F /IM chrome.exe >NUL 2>&1
timeout /t 3 /nobreak >NUL

"C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --user-data-dir="C:\chrome_pbi_session" --start-maximized "https://app.powerbi.com/groups/4431c026-df58-4f9c-9630-bb40072f829a/reports/2ab7921e-a592-472d-960b-1949fc6dec53/aa060e7c3b795463e58a"
