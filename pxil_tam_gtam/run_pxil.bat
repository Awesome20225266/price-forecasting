@echo off
REM Daily PXIL TAM + GTAM pull. Rebuilds PXIL_TAM_GTAM.xlsx and fills any missing dates.
REM Called by Windows Task Scheduler (see setup_task.ps1). Safe to run by hand any time.
cd /d "%~dp0"
if exist ".venv\Scripts\python.exe" (
    set "PY=.venv\Scripts\python.exe"
) else (
    where py >nul 2>&1 && py -3 --version >nul 2>&1 && set "PY=py -3"
    if not defined PY set "PY=python"
)

"%PY%" pxil_tam_gtam.py %*
exit /b %ERRORLEVEL%
