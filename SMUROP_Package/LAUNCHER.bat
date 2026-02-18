@echo off
setlocal enabledelayedexpansion

REM SMUROP Document Generator - Smart Launcher
REM Two modes: A) gen_config.json exists -> auto-run  B) open GUI

REM -- Check Python --
python --version >nul 2>&1
if errorlevel 1 (
    echo.
    echo   ERROR: Python is not installed or not in PATH
    echo   Install from: https://www.python.org/downloads/
    echo   Check "Add Python to PATH" during install!
    echo.
    pause
    exit /b 1
)

REM -- Find Python exe --
if exist ".venv\Scripts\python.exe" (
    set "PY=.venv\Scripts\python.exe"
) else (
    set "PY=python"
)

REM -- MODE A: Config-based generation --
if exist "gen_config.json" (
    echo.
    echo   ==========================================================
    echo   gen_config.json found - running document generation...
    echo   ==========================================================
    echo.
    "%PY%" run_config.py
    echo.
    pause
    exit /b 0
)

REM -- MODE B: GUI launcher --
if exist "launcher.py" (
    echo   Starting GUI...
    start "" "%PY%" launcher.py
    exit /b 0
)

echo.
echo   No gen_config.json and no launcher.py found.
echo   Use the Generator tab in S-Dnevnik_v7.html to create gen_config.json
echo   then place it in this folder and double-click LAUNCHER.bat again.
echo.
pause
exit /b 0
