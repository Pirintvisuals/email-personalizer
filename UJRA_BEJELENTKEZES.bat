@echo off
title Gmail ujra-bejelentkezes (egyszeri)
cls
cd /d "%~dp0"
REM "python", not "py" — on this machine py.exe resolves to the Windows Store
REM stub, which exits silently and looks like nothing happened.
python reauth_email.py
if errorlevel 1 (
  echo.
  echo Nem indult el. Probald meg kezzel:
  echo    cd /d "%~dp0"
  echo    python reauth_email.py
  echo.
  pause
)
