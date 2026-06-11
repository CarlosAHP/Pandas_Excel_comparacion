@echo off
setlocal EnableExtensions
cd /d "%~dp0"

echo ============================================
echo  Generar ComparadorCotizaciones.exe
echo ============================================
echo.

where python >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Instala Python 3.12 desde python.org
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
echo Compilando ejecutable (puede tardar 2-5 minutos)...
call .venv\Scripts\pyinstaller.exe comparador.spec --noconfirm --clean
if errorlevel 1 (
    echo [ERROR] Fallo PyInstaller
    pause
    exit /b 1
)

copy /Y LEEME_USUARIO.txt dist\ >nul 2>&1
copy /Y .env.example dist\.env.example >nul 2>&1

echo.
echo ============================================
echo  LISTO: dist\ComparadorCotizaciones.exe
echo.
echo  Para compartir, copia a una carpeta:
echo    - ComparadorCotizaciones.exe
echo    - .env  (credenciales Supabase)
echo    - LEEME_USUARIO.txt
echo ============================================
pause
