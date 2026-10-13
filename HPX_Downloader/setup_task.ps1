# Registers a daily Windows Task Scheduler job for the HPX downloader.
# Run from this folder:  powershell -ExecutionPolicy Bypass -File setup_task.ps1
# Default time 10:15 local (IST), after PXIL at 09:15 and before IEX at 14:15.
# Each run fetches only new dates plus the last 3 days (in case HPX revised them).
# Missed runs (PC off) start as soon as the PC is back on.

param(
    [string]$At = "10:15",
    [string]$TaskName = "HPX Market Data update"
)

$here     = Split-Path -Parent $MyInvocation.MyCommand.Path
$action   = New-ScheduledTaskAction -Execute "$here\run_hpx.bat" -WorkingDirectory $here
$trigger  = New-ScheduledTaskTrigger -Daily -At $At
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable `
              -ExecutionTimeLimit (New-TimeSpan -Hours 2) `
              -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 30)

Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Settings $settings `
    -Description "Pulls HPX Contingency TAM, TAM, GTAM, High Price TAM, AS DAM and Green TAM into HPX_Data Excel files" -Force

Write-Host "Registered '$TaskName' daily at $At. Test it now with:  Start-ScheduledTask -TaskName '$TaskName'"
