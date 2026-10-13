@echo off
REM One-time setup: virtualenv + dependencies + backfill from 2023-04-01 + daily scheduled task.
cd /d "%~dp0"

set "PYBOOT="
where py >nul 2>&1 && py -3 --version >nul 2>&1 && set "PYBOOT=py -3"
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

echo.
echo Running PXIL TAM/GTAM backfill from 2023-04-01. Leave the window open.
".venv\Scripts\python.exe" pxil_tam_gtam.py

echo.
echo Registering daily scheduled task ...
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0setup_task.ps1"
