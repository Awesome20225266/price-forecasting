@echo off
REM One-time setup: virtualenv + dependencies + full backfill from 2020-04-01 + daily scheduled task.
cd /d "%~dp0"

REM "python" is often the Microsoft Store stub. Prefer a real interpreter, then the py launcher.
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
".venv\Scripts\python.exe" -m pip install --upgrade pip >nul
".venv\Scripts\python.exe" -m pip install -r requirements.txt || exit /b 1

echo.
echo Running full backfill since 2020-04-01: DAM, RTM, GDAM, HP-DAM, Intra-Day/DAC/TAM
echo (conventional, green, high-price), plus TAM Any-Day and TAM Monthly auction lists.
echo Auction result PDFs are not downloaded. Leave the window open.
".venv\Scripts\python.exe" iex_pipeline.py update
".venv\Scripts\python.exe" iex_pipeline.py check
".venv\Scripts\python.exe" iex_pipeline.py export

echo.
echo Registering daily scheduled task ...
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0setup_task.ps1"
