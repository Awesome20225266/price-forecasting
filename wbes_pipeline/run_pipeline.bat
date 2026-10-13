@echo off
REM One WBES pass: PX (yesterday/today/tomorrow) + Net Schedule catch-up to yesterday.
REM Called by Windows Task Scheduler (see setup_task.ps1). Safe to run by hand.
REM A Chrome window opens only when the saved login has expired and OTP is required.
cd /d "%~dp0"
if not exist "logs" mkdir "logs"

if exist ".venv\Scripts\python.exe" (
    set "PY=.venv\Scripts\python.exe"
) else (
    where python >nul 2>&1 && python --version >nul 2>&1 && set "PY=python"
    if not defined PY set "PY=py -3"
)

echo WBES pull started. Chrome opens only if a fresh OTP login is needed.
echo Log: "%~dp0logs\wbes_pull.log"
echo.>> "%~dp0logs\wbes_pull.log"
echo ===== %DATE% %TIME% =====>> "%~dp0logs\wbes_pull.log"
"%PY%" -u wbes_pull.py --once >> "%~dp0logs\wbes_pull.log" 2>&1
exit /b %ERRORLEVEL%
