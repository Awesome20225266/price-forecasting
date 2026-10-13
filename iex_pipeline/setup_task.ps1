# Registers a daily Windows Task Scheduler job for the IEX pipeline.
# Run from this folder:  powershell -ExecutionPolicy Bypass -File setup_task.ps1
# Default time 14:15 local (IST): DAM / GDAM / HP-DAM results for tomorrow are out by ~13:00,
# and yesterday's RTM is complete. Missed runs (PC off) start as soon as the PC is back on.

param(
    [string]$At = "14:15",
    [string]$TaskName = "IEX Market Data Pipeline"
)

$here     = Split-Path -Parent $MyInvocation.MyCommand.Path
$action   = New-ScheduledTaskAction -Execute "$here\run_pipeline.bat" -WorkingDirectory $here
$trigger  = New-ScheduledTaskTrigger -Daily -At $At
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable `
              -ExecutionTimeLimit (New-TimeSpan -Hours 2) `
              -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 30)

Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Settings $settings `
    -Description "Pulls IEX DAM/RTM/GDAM/HP-DAM, Intra-Day/DAC/TAM (incl. Green, HP) and TAM auction lists into data\iex_market_data.db" -Force

Write-Host "Registered '$TaskName' daily at $At. Test it now with:  Start-ScheduledTask -TaskName '$TaskName'"
