@echo off
setlocal enabledelayedexpansion

REM SMUROP - Automated Setup
REM Creates .venv and installs python-docx

echo.
echo   ==========================================================
echo   SMUROP - Setup
echo   ==========================================================
echo.

REM -- Check Python --
python --version >nul 2>&1
if errorlevel 1 (
    echo   ERROR: Python is not installed or not in PATH
    echo   Install from: https://www.python.org/downloads/
    echo   Check "Add Python to PATH" during install!
    echo.
    pause
    exit /b 1
)

for /f "tokens=2" %%i in ('python --version 2^>^&1') do set PYVER=%%i
echo   Python found: %PYVER%
echo   Running setup.py ...
echo.

python setup.py --no-prompt

if errorlevel 1 (
    echo.
    echo   ERROR: Setup failed!
    pause
    exit /b 1
)

echo.
echo   ==========================================================
echo   Setup completed successfully!
echo   ==========================================================
echo.
echo   Next steps:
echo     1. Open S-Dnevnik_v7.html in Chrome
echo     2. Enter student data or import JSON
echo     3. Go to Generator tab
echo     4. Click "Save Config" button
echo     5. Double-click LAUNCHER.bat
echo.
pause
exit /b 0
