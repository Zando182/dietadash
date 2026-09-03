@echo off
chcp 65001 >nul
cd /d "%~dp0"
title DietaDash

echo.
echo   Sto accendendo DietaDash...
echo.

where py >nul 2>nul
if %errorlevel%==0 goto usa_py
where python >nul 2>nul
if %errorlevel%==0 goto usa_python

echo   Non trovo Python su questo computer.
echo.
echo   Vai su https://www.python.org/downloads/ scarica Python
echo   e durante l'installazione metti la spunta su
echo   "Add python.exe to PATH". Poi riprova con questo file.
echo.
pause
exit /b 1

:usa_py
set PY=py
goto controlla

:usa_python
set PY=python

:controlla
%PY% -c "import openpyxl" >nul 2>nul
if %errorlevel%==0 goto avvia
echo   Mi manca un pezzo (openpyxl). Lo scarico, ci vuole un minuto...
%PY% -m pip install openpyxl
if not %errorlevel%==0 (
  echo.
  echo   Non sono riuscito a installarlo. Prova a scrivere a mano:
  echo   %PY% -m pip install openpyxl
  echo.
  pause
  exit /b 1
)

:avvia
%PY% scripts\server.py %1
echo.
echo   DietaDash e' stato chiuso.
pause
