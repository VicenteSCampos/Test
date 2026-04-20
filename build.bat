@echo off
echo ====================================
echo   Compilando ClaseObsidian...
echo ====================================
echo.

py -3.11 -m PyInstaller ^
    --noconfirm ^
    --windowed ^
    --name "ClaseObsidian" ^
    --collect-all customtkinter ^
    --hidden-import="anthropic" ^
    --hidden-import="fpdf" ^
    --hidden-import="pptx" ^
    --hidden-import="faster_whisper" ^
    app.py

echo.
echo ====================================
echo  Compilacion completa!
echo ====================================
echo.
echo El .exe esta en: dist\ClaseObsidian\ClaseObsidian.exe
echo.
echo IMPORTANTE: Copia tu archivo .env
echo a la carpeta dist\ClaseObsidian\
echo antes de distribuir o usar el .exe
echo.
pause
