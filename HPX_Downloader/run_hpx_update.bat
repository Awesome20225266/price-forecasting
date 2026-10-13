@echo off
REM Double-click to update HPX market data (keep this file next to hpx_downloader.py)
cd /d "%~dp0"
if exist ".venv\Scripts\python.exe" (
    set "PY=.venv\Scripts\python.exe"
) else (
    where py >nul 2>&1 && py -3 --version >nul 2>&1 && set "PY=py -3"
    if not defined PY set "PY=python"
)
"%PY%" -c "import requests, pandas, openpyxl" 2>nul || (
    echo Installing required packages...
    "%PY%" -m pip install --quiet requests pandas openpyxl
)
"%PY%" hpx_downloader.py %*
echo.
pause
