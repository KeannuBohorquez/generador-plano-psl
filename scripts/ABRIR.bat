@echo off
chcp 65001 >nul
title Generador Archivo Plano PSL

set SCRIPT=%~dp0..\main.py
set PYTHON=

echo Buscando Python...

if exist "%USERPROFILE%\AppData\Local\Programs\Python\Python313\python.exe" set PYTHON=%USERPROFILE%\AppData\Local\Programs\Python\Python313\python.exe
if exist "%USERPROFILE%\AppData\Local\Programs\Python\Python312\python.exe" set PYTHON=%USERPROFILE%\AppData\Local\Programs\Python\Python312\python.exe
if exist "%USERPROFILE%\AppData\Local\Programs\Python\Python311\python.exe" set PYTHON=%USERPROFILE%\AppData\Local\Programs\Python\Python311\python.exe
if exist "%USERPROFILE%\AppData\Local\Programs\Python\Python310\python.exe" set PYTHON=%USERPROFILE%\AppData\Local\Programs\Python\Python310\python.exe
if exist "%USERPROFILE%\anaconda3\python.exe" set PYTHON=%USERPROFILE%\anaconda3\python.exe
if exist "%USERPROFILE%\Anaconda3\python.exe" set PYTHON=%USERPROFILE%\Anaconda3\python.exe
if exist "%LOCALAPPDATA%\anaconda3\python.exe" set PYTHON=%LOCALAPPDATA%\anaconda3\python.exe
if exist "C:\Python313\python.exe" set PYTHON=C:\Python313\python.exe
if exist "C:\Python312\python.exe" set PYTHON=C:\Python312\python.exe
if exist "C:\Python311\python.exe" set PYTHON=C:\Python311\python.exe
if exist "C:\Python310\python.exe" set PYTHON=C:\Python310\python.exe

if "%PYTHON%"=="" (
    where python >nul 2>&1
    if not errorlevel 1 set PYTHON=python
)

if "%PYTHON%"=="" (
    echo.
    echo [ERROR] No se encontro Python instalado.
    echo.
    echo Instala Python desde: https://www.python.org/downloads
    echo Durante la instalacion marca: Add Python to PATH
    echo.
    pause
    exit /b 1
)

echo [OK] Python: %PYTHON%
echo.
echo Verificando dependencias...
"%PYTHON%" -m pip install -r "%~dp0..\requirements.txt" --quiet --exists-action i
echo.
echo Abriendo interfaz...
"%PYTHON%" "%SCRIPT%"

if errorlevel 1 (
    echo.
    echo [ERROR] El script cerro con error.
    echo.
    pause
)
