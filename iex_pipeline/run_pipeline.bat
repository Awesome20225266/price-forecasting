@echo off
REM Daily IEX pull: DAM, RTM, GDAM, HP-DAM, Intra-Day/DAC/TAM, auction lists + CSV export.
REM Auction result PDFs are not downloaded.
REM Called by Windows Task Scheduler (see setup_task.ps1). Safe to run by hand any time.
cd /d "%~dp0"
if exist ".venv\Scripts\python.exe" (
    set "PY=.venv\Scripts\python.exe"
) else (
    where python >nul 2>&1 && python --version >nul 2>&1 && set "PY=python"
    if not defined PY set "PY=py -3"
)

"%PY%" iex_pipeline.py update
set RC=%ERRORLEVEL%
if not "%RC%"=="2" goto export

REM exit code 2 = portal refused plain HTTP requests -> retry through headless browser (needs playwright)
echo Plain requests blocked, retrying with --browser ...
"%PY%" iex_pipeline.py --browser update
set RC=%ERRORLEVEL%

:export
"%PY%" iex_pipeline.py export
exit /b %RC%
