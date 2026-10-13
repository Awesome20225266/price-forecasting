# Registers a daily Windows Task Scheduler job for the NOAR market snapshot.
# Run from this folder:  powershell -ExecutionPolicy Bypass -File setup_task.ps1
# Default time 10:00 local. The PC must be logged in so the browser window can open.

param(
    [string]$At = "10:00",
    [string]$TaskName = "NOAR market snapshot update"
)

$here     = Split-Path -Parent $MyInvocation.MyCommand.Path
$action   = New-ScheduledTaskAction -Execute "$here\run_noar.bat" -Argument "--daily" -WorkingDirectory $here
$trigger  = New-ScheduledTaskTrigger -Daily -At $At
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable `
              -ExecutionTimeLimit (New-TimeSpan -Hours 2) `
              -RestartCount 2 -RestartInterval (New-TimeSpan -Minutes 30)

Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Settings $settings `
    -Description "Pulls yesterday's NOAR TRAS market snapshot into noar_market_snapshot.xlsx" -Force

Write-Host "Registered '$TaskName' daily at $At. Test it now with:  Start-ScheduledTask -TaskName '$TaskName'"
