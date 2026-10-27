@echo off
title DM Sender — Starting...
cls
echo Starting...

REM ── Locate Chrome (Program Files / x86 / LocalAppData) ───────────────────
set "CHROME="
for %%P in (
  "%ProgramFiles%\Google\Chrome\Application\chrome.exe"
  "%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe"
  "%LocalAppData%\Google\Chrome\Application\chrome.exe"
) do if exist "%%~P" set "CHROME=%%~P"

if not defined CHROME (
  echo ERROR: Could not find chrome.exe. Install Google Chrome, or edit this
  echo file and set CHROME to the full path of chrome.exe.
  pause
  exit /b 1
)

REM ── Kill ONLY the DM Sender Chrome (the instance on debug port 9222) ──────
REM Get-Process has no CommandLine in Windows PowerShell 5.1, so use CIM.
REM Matching the debug port leaves your normal Chrome and the HU app (9223) alone.
powershell -NoProfile -Command "Get-CimInstance Win32_Process | Where-Object { $_.Name -eq 'chrome.exe' -and $_.CommandLine -like '*remote-debugging-port=9222*' } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }" >nul 2>&1
timeout /t 2 /nobreak >nul

REM ── Start Flask silently ──────────────────────────────────────────────────
start /min "DM App" cmd /c "cd /d "%~dp0" && python -u dm_app.py >> server.log 2>&1"
echo Waiting for app to start...
timeout /t 5 /nobreak >nul

REM ── Open a SEPARATE Chrome just for the DM app (won't conflict) ──────────
start "" "%CHROME%" ^
  --remote-debugging-port=9222 ^
  --user-data-dir="C:\Users\pirin\AppData\Local\Google\Chrome\DMSender" ^
  --no-first-run ^
  --no-default-browser-check ^
  "http://localhost:5001"

echo.
echo Chrome is opening. Log into Facebook if prompted (only needed once).
timeout /t 3 /nobreak >nul
exit
