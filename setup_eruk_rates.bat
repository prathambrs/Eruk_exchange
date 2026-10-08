@echo off
title ERUK exchange rates - one-time setup
cd /d "%~dp0"

where py >nul 2>nul && (set "PY=py") || (set "PY=python")

%PY% --version >nul 2>nul
if errorlevel 1 goto nopython

echo Installing the tools the rates script needs. This takes a few minutes.
echo.
%PY% -m pip install --upgrade playwright pandas openpyxl
if errorlevel 1 goto failed
%PY% -m playwright install chromium
if errorlevel 1 goto failed

echo.
echo Setup finished. From now on just double-click run_eruk_rates.bat
echo.
pause
exit /b 0

:nopython
echo Python is not installed on this computer.
echo Install it from https://www.python.org/downloads/ and tick "Add python.exe to PATH" on the first screen.
echo Then double-click this setup file again.
echo.
pause
exit /b 1

:failed
echo.
echo Setup failed. Take a screenshot of this window and send it to Pratham.
echo.
pause
exit /b 1
