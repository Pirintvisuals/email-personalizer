@echo off
title Email Sender HU — Starting...
cls
echo Starting Hungarian Email Sender...

REM ── Start Flask on port 5003 ─────────────────────────────────────────────
start /min "Email App HU" cmd /c "cd /d "%~dp0" && python hu_email_app.py >> server_email_hu.log 2>&1"
echo Waiting for app to start...
timeout /t 4 /nobreak >nul

REM ── Open the app in your normal default browser ────────────────────────
REM No dedicated automation Chrome needed here (unlike the DM sender): this
REM app only scrapes/screenshots lead WEBSITES with its own invisible headless
REM browser, and hands finished emails to Gmail via your regular Chrome so
REM Streak tracking still works there.
start "" "http://localhost:5003"

echo.
echo Browser tab is opening. Paste your leads Excel path and click Betoltes.
timeout /t 5 /nobreak >nul
exit
