@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"
title FB Csoport Lead Gyujto

echo ================================================
echo   FB CSOPORT LEAD GYUJTO
echo   Cegtulajdonosokat gyujt a csoport posztjaibol
echo ================================================
echo.

set "GROUPURL="
set /p "GROUPURL=Illeszd be a Facebook csoport linkjet, majd Enter: "

if "!GROUPURL!"=="" (
    echo.
    echo Nem adtal meg linket. Kilepes.
    echo.
    pause
    exit /b 1
)

echo.
echo Csoport: !GROUPURL!
echo.
echo A tovabbit innentol a program vegzi - a Chrome inditasat is.
echo A rendes Chrome-odat NEM zarja be, sajat ablakot nyit.
echo.
echo ELSO INDITASKOR: az uj Chrome ablakban be kell jelentkezned
echo a Facebookra EGYSZER. A program megvarja, es utana emlekszik ra.
echo.
echo Ne zard be ezt az ablakot.
echo.

python fb_group_scraper.py --group "!GROUPURL!" --out "leads_fb.xlsx"
set PYEXIT=!errorlevel!

echo.
if not "!PYEXIT!"=="0" (
    echo ================================================
    echo A gyujtes leallt. Olvasd el a fenti sorokat - ott
    echo van megirva, hogy pontosan mi volt a gond.
    echo ================================================
    echo.
    pause
    exit /b !PYEXIT!
)

echo ================================================
echo Kesz. Nyitom a leads_fb.xlsx fajlt...
echo ================================================
if exist "leads_fb.xlsx" start "" "leads_fb.xlsx"
echo.
echo Nyomj egy gombot a bezarashoz.
pause >nul
