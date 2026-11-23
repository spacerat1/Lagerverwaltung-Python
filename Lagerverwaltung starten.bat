@echo off
rem Startet die Lagerverwaltung Comline mit dem installierten Python.
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
    echo Python wurde nicht gefunden.
    echo Bitte Python 3 installieren ^(python.org, Haken bei "Add python.exe to PATH"^).
    pause
    exit /b 1
)

python -c "import PyQt6, pandas, openpyxl, win32com.client" >nul 2>nul
if errorlevel 1 (
    echo Es fehlen Python-Pakete. Sie werden jetzt installiert ...
    python -m pip install --user -r "%~dp0requirements.txt"
    if errorlevel 1 (
        echo Installation fehlgeschlagen.
        pause
        exit /b 1
    )
)

rem pythonw: ohne schwarzes Konsolenfenster
start "" pythonw "%~dp0Lagerverwaltung.py"
