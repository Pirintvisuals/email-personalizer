@echo off
echo Closing Chrome...
taskkill /F /T /IM chrome.exe >nul 2>&1
wmic process where "name='chrome.exe'" delete >nul 2>&1

:waitloop
tasklist | find /i "chrome.exe" >nul 2>&1
if %errorlevel%==0 (
    taskkill /F /T /IM chrome.exe >nul 2>&1
    ping -n 2 127.0.0.1 >nul
    goto waitloop
)

echo Opening Chrome with debug port...
start "" "C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --user-data-dir="C:\ChromeDMApp" --no-first-run --no-default-browser-check
ping -n 8 127.0.0.1 >nul

curl -s http://localhost:9222/json/version
echo.

echo Starting DM App...
cd /d "%~dp0"
py dm_app.py
pause
