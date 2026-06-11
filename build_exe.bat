@echo off
setlocal EnableExtensions
cd /d "%~dp0"

echo ============================================
echo  Generar ComparadorCotizaciones.exe
echo  (credenciales del .env van DENTRO del exe)
echo ============================================
echo.

where python >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Instala Python 3.12 desde python.org
    pause
    exit /b 1
)

if not exist ".env" if not exist ".env.local" (
    echo [ERROR] Necesitas .env o .env.local con las credenciales Supabase
    echo         antes de generar el ejecutable.
    pause
    exit /b 1
)

if not exist ".venv\Scripts\python.exe" (
    python -m venv .venv
    call .venv\Scripts\python.exe -m pip install --upgrade pip
    call .venv\Scripts\pip.exe install -r requirements.txt -r requirements-build.txt
) else (
    call .venv\Scripts\pip.exe install -r requirements.txt -r requirements-build.txt
)

echo.
echo Empaquetando credenciales...
call .venv\Scripts\python.exe prepare_build.py
if errorlevel 1 pause & exit /b 1

echo Compilando ejecutable (2-5 minutos)...
call .venv\Scripts\pyinstaller.exe comparador.spec --noconfirm --clean
if errorlevel 1 (
    echo [ERROR] Fallo PyInstaller
    pause
    exit /b 1
)

echo.
echo ============================================
echo  LISTO: dist\ComparadorCotizaciones.exe
echo.
echo  Comparte SOLO ese archivo por Drive.
echo  El usuario: descarga y doble clic. Nada mas.
echo ============================================
pause
