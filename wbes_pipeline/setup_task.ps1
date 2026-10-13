# Registers a daily Windows Task Scheduler job for the WBES pipeline.
# Run from this folder:  powershell -ExecutionPolicy Bypass -File setup_task.ps1
# 18:00 local, after the IEX 14:15 job. One pass (--once), not the continuous loop,
# so a new day does not start a second copy beside one that is still running.
# Interactive logon is required: if the token has died, Chrome must be visible for OTP.

param(
    [string]$At = "18:00",
    [string]$TaskName = "WBES Market Data Pipeline"
)

$here      = Split-Path -Parent $MyInvocation.MyCommand.Path
$action    = New-ScheduledTaskAction -Execute "$here\run_pipeline.bat" -WorkingDirectory $here
$trigger   = New-ScheduledTaskTrigger -Daily -At $At
$settings  = New-ScheduledTaskSettingsSet -StartWhenAvailable `
               -MultipleInstances IgnoreNew `
               -ExecutionTimeLimit (New-TimeSpan -Hours 6) `
               -RestartCount 1 -RestartInterval (New-TimeSpan -Minutes 15)
$principal = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" `
               -LogonType Interactive -RunLevel Limited

Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger `
    -Settings $settings -Principal $principal `
    -Description "One daily pass of WBES PX (yesterday/today/tomorrow) and Net Schedule catch-up into data\wbes.db. Opens Chrome if OTP login is needed." -Force

Write-Host "Registered '$TaskName' daily at $At (only while you are logged on)."
Write-Host "First run opens Chrome for OTP. Test with:  Start-ScheduledTask -TaskName '$TaskName'"
