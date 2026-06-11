@echo off
setlocal EnableExtensions
cd /d "%~dp0"

echo ============================================
echo  Instalacion - Comparador Excel vs Supabase
echo ============================================
echo.

where python >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python no encontrado.
    echo Descarga e instala Python 3.11 o 3.12 desde:
    echo   https://www.python.org/downloads/
    echo Activa "Add python.exe to PATH".
    pause
    exit /b 1
)

for /f "tokens=*" %%v in ('python -c "import sys; print(f'{sys.version_info.major}.{sys.version_info.minor}')"') do set PYVER=%%v
echo Python detectado: %PYVER%

python -c "import tkinter" 2>nul
if errorlevel 1 (
    echo.
    echo [ERROR] tkinter no esta disponible.
    echo Reinstala Python desde python.org y marca "tcl/tk and IDLE".
    pause
    exit /b 1
)
echo tkinter: OK

if exist ".venv" (
    echo Eliminando entorno virtual anterior...
    rmdir /s /q ".venv"
)

echo Creando entorno virtual...
python -m venv .venv
if errorlevel 1 (
    echo [ERROR] No se pudo crear .venv
    pause
    exit /b 1
)

call .venv\Scripts\python.exe -m pip install --upgrade pip
call .venv\Scripts\pip.exe install -r requirements.txt
if errorlevel 1 (
    echo [ERROR] pip install fallo
    pause
    exit /b 1
)

if not exist ".env" if not exist ".env.local" (
    echo.
    echo [AVISO] No hay archivo .env
    echo Copia .env.example a .env y completa las credenciales de Supabase.
    echo Ver README.md - Paso 3
)

echo.
echo ============================================
echo  Instalacion completada.
echo  Para abrir la app: doble clic en run.bat
echo ============================================
pause
