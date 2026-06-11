@echo off
setlocal EnableExtensions
cd /d "%~dp0"

REM Comparador Excel vs Supabase — arranque Windows (doble clic o desde CMD)

where python >nul 2>&1
if errorlevel 1 (
    echo.
    echo [ERROR] Python no esta en el PATH.
    echo Instala Python 3.11+ desde https://www.python.org/downloads/
    echo Marca la opcion "Add python.exe to PATH" durante la instalacion.
    echo.
    pause
    exit /b 1
)

if not exist ".venv\Scripts\python.exe" (
    echo Creando entorno virtual...
    python -m venv .venv
    if errorlevel 1 (
        echo [ERROR] No se pudo crear el entorno virtual.
        pause
        exit /b 1
    )
    echo Instalando dependencias...
    call .venv\Scripts\python.exe -m pip install --upgrade pip
    call .venv\Scripts\pip.exe install -r requirements.txt
    if errorlevel 1 (
        echo [ERROR] Fallo la instalacion de dependencias.
        pause
        exit /b 1
    )
)

if not exist ".env" if not exist ".env.local" (
    echo.
    echo [AVISO] No se encontro .env ni .env.local
    echo Crea un archivo .env con VITE_SUPABASE_URL y SUPABASE_SERVICE_KEY
    echo.
)

call .venv\Scripts\python.exe src\main.py
if errorlevel 1 (
    echo.
    echo La aplicacion termino con error.
    pause
)
