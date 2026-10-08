@echo off
title ERUK exchange rates
cd /d "%~dp0"

where py >nul 2>nul && (set "PY=py") || (set "PY=python")

echo Pulling exchange rates. A browser window will open, please leave it alone.
echo If it shows "Verify you are human", tick the box.
echo Make sure eruk_fx_rates.xlsx is closed before this runs.
echo.

%PY% eruk_fix.py
if errorlevel 1 goto failed

echo.
echo Done. Opening the Excel file...
start "" "%~dp0eruk_fx_rates.xlsx"
echo.
pause
exit /b 0

:failed
echo.
echo Something went wrong. Take a screenshot of this window and send it to Pratham.
echo.
pause
exit /b 1
