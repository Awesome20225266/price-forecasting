@echo off
REM One-time setup: virtualenv + dependencies + daily scheduled task.
REM Does not start a download. The first pull needs you at the PC for the OTP login.
cd /d "%~dp0"

set "PYBOOT="
where python >nul 2>&1 && python --version >nul 2>&1 && set "PYBOOT=python"
if not defined PYBOOT (
    where py >nul 2>&1 && py -3 --version >nul 2>&1 && set "PYBOOT=py -3"
)
if not defined PYBOOT (
    echo Python 3.9+ is required: https://www.python.org/downloads/
    echo If Python is installed, disable the "python.exe" app execution alias in Settings.
    exit /b 1
)

if not exist ".venv" %PYBOOT% -m venv .venv
".venv\Scripts\python.exe" -m pip install --upgrade pip
".venv\Scripts\python.exe" -m pip install -r requirements.txt || exit /b 1

echo.
echo Registering daily scheduled task ...
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0setup_task.ps1"
echo.
echo Setup finished. Run run_pipeline.bat once while you are at the PC so you can complete the OTP login.
