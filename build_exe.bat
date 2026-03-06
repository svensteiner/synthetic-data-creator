@echo off
echo ============================================
echo  Synthetische Daten Generator - EXE Build
echo ============================================
echo.

where pyinstaller >nul 2>&1
if %errorlevel% neq 0 (
    echo PyInstaller nicht gefunden. Installiere...
    pip install pyinstaller
)

echo Erstelle EXE...
python -m PyInstaller ^
    --onefile ^
    --windowed ^
    --name "SynthetischeDatenGenerator" ^
    --add-data "synthesizer.py;." ^
    --add-data "audit.py;." ^
    --exclude-module pandas ^
    --exclude-module numpy ^
    --exclude-module matplotlib ^
    --exclude-module scipy ^
    --exclude-module PIL ^
    --exclude-module IPython ^
    --exclude-module pytest ^
    --exclude-module _pytest ^
    --exclude-module notebook ^
    --exclude-module jinja2 ^
    app.py

echo.
if exist "dist\SynthetischeDatenGenerator.exe" (
    echo FERTIG!
    echo   dist\SynthetischeDatenGenerator.exe
    for %%A in ("dist\SynthetischeDatenGenerator.exe") do echo   Groesse: %%~zA Bytes
) else (
    echo FEHLER: EXE wurde nicht erstellt.
)
pause
