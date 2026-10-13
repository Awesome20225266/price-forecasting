@echo off
REM Daily HPX market-data pull. Updates the Excel files in HPX_Data.
REM Called by Windows Task Scheduler (see setup_task.ps1). Safe to run by hand any time.
cd /d "%~dp0"
if not exist "logs" mkdir "logs"

if exist ".venv\Scripts\python.exe" (
    set "PY=.venv\Scripts\python.exe"
) else (
    where py >nul 2>&1 && py -3 --version >nul 2>&1 && set "PY=py -3"
    if not defined PY set "PY=python"
)

echo.>> "%~dp0logs\hpx_update.log"
echo ===== %DATE% %TIME% =====>> "%~dp0logs\hpx_update.log"
"%PY%" -u hpx_downloader.py %* >> "%~dp0logs\hpx_update.log" 2>&1
exit /b %ERRORLEVEL%
