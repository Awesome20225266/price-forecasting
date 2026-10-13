# Registers a daily Windows Task Scheduler job for the PXIL TAM/GTAM pipeline.
# Run from this folder:  powershell -ExecutionPolicy Bypass -File setup_task.ps1
# Default time 09:15 local (IST). Missed runs (PC off) start as soon as the PC is back on.
# Each run re-downloads the latest 3 days and fills any dates that were missed.

param(
    [string]$At = "09:15",
    [string]$TaskName = "PXIL TAM GTAM update"
)

$here     = Split-Path -Parent $MyInvocation.MyCommand.Path
$action   = New-ScheduledTaskAction -Execute "$here\run_pxil.bat" -WorkingDirectory $here
$trigger  = New-ScheduledTaskTrigger -Daily -At $At
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable `
              -ExecutionTimeLimit (New-TimeSpan -Hours 2) `
              -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 30)

Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Settings $settings `
    -Description "Pulls PXIL TAM and GTAM market data into PXIL_TAM_GTAM.xlsx and the local SQLite cache" -Force

Write-Host "Registered '$TaskName' daily at $At. Test it now with:  Start-ScheduledTask -TaskName '$TaskName'"
