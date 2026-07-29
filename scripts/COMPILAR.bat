@echo off
chcp 65001 >nul
title Compilador - Generador Archivo Plano PSL

echo ============================================================
echo   COMPILADOR - Generador Archivo Plano PSL (Toledana)
echo ============================================================
echo.

set PYTHON=
set SCRIPT=%~dp0..\main.py
set OUTDIR=%~dp0..\dist
set BUILDDIR=%~dp0..\build_tmp
set VENVDIR=%~dp0..\build_venv

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
    echo [ERROR] No se encontro Python.
    echo Instala desde: https://www.python.org/downloads
    pause & exit /b 1
)

echo [OK] Python: %PYTHON%
echo.

echo ============================================================
echo   IMPORTANTE - DESACTIVA EL ANTIVIRUS ANTES DE CONTINUAR
echo ============================================================
echo.
echo   Avira u otro antivirus puede bloquear la compilacion.
echo.
echo   PASOS PARA AVIRA:
echo   1. Clic derecho en el icono de Avira (barra de tareas)
echo   2. Selecciona: Desactivar proteccion en tiempo real
echo   3. Elige: 15 minutos
echo   4. Presiona cualquier tecla aqui para continuar
echo   5. Cuando termine, vuelve a activar Avira
echo.
pause
echo.

echo [1/4] Preparando entorno virtual limpio (build_venv)...
echo.
echo   Se usa un venv dedicado SOLO con las librerias de este
echo   proyecto, para que el .exe no incluya paquetes de otros
echo   proyectos que tengas instalados en tu Python global
echo   (matplotlib, jupyter, opencv, etc. si los tienes).
echo.

if not exist "%VENVDIR%\Scripts\python.exe" (
    echo   Creando entorno virtual nuevo...
    "%PYTHON%" -m venv "%VENVDIR%"
    if errorlevel 1 (
        echo [ERROR] No se pudo crear el entorno virtual.
        pause & exit /b 1
    )
) else (
    echo   Reutilizando entorno virtual existente.
)

set VENVPY=%VENVDIR%\Scripts\python.exe

echo.
echo [2/4] Instalando dependencias dentro del venv...
"%VENVPY%" -m pip install --upgrade pip --quiet
"%VENVPY%" -m pip install -r "%~dp0..\requirements.txt" --quiet
"%VENVPY%" -m pip install pyinstaller --quiet
echo [OK] Dependencias listas (aisladas del Python global).
echo.

echo [3/4] Compilando .exe (tarda 2-4 minutos)...
echo.

"%VENVPY%" -m PyInstaller ^
    --onefile --windowed ^
    --name "GeneradorArchivoPlano_PSL" ^
    --icon "%~dp0..\assets\favicon.ico" ^
    --add-data "%~dp0..\assets\favicon.ico;assets" ^
    --paths "%~dp0..\src" ^
    --collect-all pdfplumber ^
    --collect-all pdfminer ^
    --collect-all msoffcrypto ^
    --collect-all openpyxl ^
    --collect-all xlrd ^
    --collect-all xlwt ^
    --hidden-import generador_plano ^
    --hidden-import generador_plano.config ^
    --hidden-import generador_plano.extractor ^
    --hidden-import generador_plano.lector ^
    --hidden-import generador_plano.constructor ^
    --hidden-import generador_plano.exportador ^
    --hidden-import generador_plano.conciliacion ^
    --hidden-import generador_plano.conciliacion.procesador ^
    --hidden-import generador_plano.ui.app ^
    --hidden-import generador_plano.ui.widgets ^
    --hidden-import generador_plano.ui.tab_toledana ^
    --hidden-import generador_plano.ui.tab_conciliacion ^
    --hidden-import pandas ^
    --hidden-import pandas._libs.tslibs.np_datetime ^
    --hidden-import pandas._libs.tslibs.nattype ^
    --hidden-import pandas._libs.tslibs.timedeltas ^
    --hidden-import cffi ^
    --hidden-import cryptography ^
    --hidden-import charset_normalizer ^
    --exclude-module matplotlib ^
    --exclude-module scipy ^
    --exclude-module sympy ^
    --exclude-module IPython ^
    --exclude-module jupyter ^
    --exclude-module notebook ^
    --exclude-module PyQt5 ^
    --exclude-module PySide2 ^
    --exclude-module PySide6 ^
    --exclude-module cv2 ^
    --exclude-module tensorflow ^
    --exclude-module torch ^
    --exclude-module pytest ^
    --exclude-module numpy.f2py ^
    --exclude-module tkinter.test ^
    --exclude-module unittest ^
    --distpath "%OUTDIR%" ^
    --workpath "%BUILDDIR%" ^
    --specpath "%BUILDDIR%" ^
    --noconfirm ^
    "%SCRIPT%"

if errorlevel 1 (
    echo.
    echo [ERROR] La compilacion fallo.
    echo Si el antivirus elimino archivos, desactivalo y borra build_tmp.
    echo.
    pause & exit /b 1
)

echo.
echo [4/4] Limpiando archivos temporales...
rmdir /s /q "%BUILDDIR%" 2>nul

echo.
echo ============================================================
echo   COMPILACION EXITOSA
echo ============================================================
echo.
echo   Ejecutable: %OUTDIR%\GeneradorArchivoPlano_PSL.exe
echo.
for %%A in ("%OUTDIR%\GeneradorArchivoPlano_PSL.exe") do echo   Tamano: %%~zA bytes
echo.
echo   TIP: Si sigue pesando mucho, instala UPX (upx.github.io),
echo   descomprimelo y agrega la carpeta a tu PATH. PyInstaller lo
echo   detecta automaticamente y comprime el .exe bastante mas.
echo.
echo   RECUERDA: Vuelve a activar Avira si lo desactivaste.
echo.
pause
