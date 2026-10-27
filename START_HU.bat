@echo off
title DM Sender HU — Starting...
cls
echo Starting Hungarian DM App...

REM ── Kill only the HU Chrome (runs on port 9223) ──────────────────────────
powershell -command "Get-Process chrome -ErrorAction SilentlyContinue | Where-Object {$_.CommandLine -like '*9223*'} | Stop-Process -Force" >nul 2>&1
timeout /t 2 /nobreak >nul

REM ── Start Flask on port 5002 ─────────────────────────────────────────────
start /min "DM App HU" cmd /c "cd /d "%~dp0" && python dm_app_hu.py >> server_hu.log 2>&1"
echo Waiting for app to start...
timeout /t 6 /nobreak >nul

REM ── Open a SEPARATE Chrome for the HU app (own profile, own port) ─────────
start "" "C:\Program Files\Google\Chrome\Application\chrome.exe" ^
  --remote-debugging-port=9223 ^
  --user-data-dir="C:\Users\pirin\AppData\Local\Google\Chrome\DMSenderHU" ^
  --no-first-run ^
  --no-default-browser-check ^
  "http://localhost:5002"

echo.
echo Chrome is opening. Log into Facebook if prompted (only needed once).
timeout /t 8 /nobreak >nul
exit
