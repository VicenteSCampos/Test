@echo off
echo ====================================
echo    Clase ^> Obsidian - Setup
echo ====================================
echo.

echo [1/2] Instalando dependencias...
py -3.11 -m pip install faster-whisper anthropic customtkinter python-dotenv
if %errorlevel% neq 0 (
    echo ERROR: Fallo la instalacion de dependencias.
    pause
    exit /b 1
)

echo.
echo [2/2] Creando archivo de configuracion...
if not exist .env (
    copy .env.example .env
    echo Archivo .env creado.
) else (
    echo Archivo .env ya existe, no se sobreescribe.
)

echo.
echo ====================================
echo  Instalacion completa!
echo ====================================
echo.
echo IMPORTANTE: Abre el archivo .env con
echo el Bloc de notas y pega tu API key
echo de Anthropic donde dice:
echo   ANTHROPIC_API_KEY=sk-ant-...
echo.
pause
