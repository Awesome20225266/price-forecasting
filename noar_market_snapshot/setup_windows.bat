@echo off
REM One-time setup: virtualenv + Playwright Chromium.
cd /d "%~dp0"

set "PYBOOT="
if exist "%LocalAppData%\Python\bin\python.exe" set "PYBOOT=%LocalAppData%\Python\bin\python.exe"
if not defined PYBOOT (
    where py >nul 2>&1 && py -3 --version >nul 2>&1 && set "PYBOOT=py -3"
)
if not defined PYBOOT (
    where python >nul 2>&1 && python --version >nul 2>&1 && set "PYBOOT=python"
)
if not defined PYBOOT (
    echo Python 3.9+ is required: https://www.python.org/downloads/
    echo If Python is installed, disable the "python.exe" app execution alias in Settings.
    exit /b 1
)

if not exist ".venv" %PYBOOT% -m venv .venv
".venv\Scripts\python.exe" -m pip install --upgrade pip
".venv\Scripts\python.exe" -m pip install -r requirements.txt || exit /b 1
".venv\Scripts\python.exe" -m playwright install chromium || exit /b 1

echo.
echo Setup complete. Run a full history pull with:
echo   run_noar.bat
echo Or yesterday only:
echo   run_noar.bat --daily
