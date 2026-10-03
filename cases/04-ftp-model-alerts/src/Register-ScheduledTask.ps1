# Name: Register-ScheduledTask.ps1
# Version: 1.0
# What it does: Register Watch-FtpModels.ps1 as a Windows scheduled task (every 5 minutes).
# Inputs: Watch-FtpModels.ps1 path.
# Outputs: Scheduled task in Task Scheduler.
# How to run: powershell -File Register-ScheduledTask.ps1
# Notes: Run once on the watcher PC.
$ErrorActionPreference = "Stop"

$scriptPath = Join-Path $PSScriptRoot "Watch-FtpModels.ps1"
if (-not (Test-Path $scriptPath)) {
    throw "Watch-FtpModels.ps1 not found next to this script"
}

$action = New-ScheduledTaskAction `
    -Execute "powershell.exe" `
    -Argument "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$scriptPath`""

# Once + repeat every 5 min for 10 years (MaxValue breaks Task Scheduler XML)
$trigger = New-ScheduledTaskTrigger `
    -Once `
    -At ((Get-Date).AddMinutes(1)) `
    -RepetitionInterval (New-TimeSpan -Minutes 5) `
    -RepetitionDuration (New-TimeSpan -Days 3650)

$settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -StartWhenAvailable `
    -MultipleInstances IgnoreNew `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 30)

Register-ScheduledTask `
    -TaskName "FTP Models Phone Alerts" `
    -Action $action `
    -Trigger $trigger `
    -Settings $settings `
    -Description "Poll FTP model folders and push to ntfy/telegram" `
    -Force | Out-Null

Write-Host "OK: Scheduled Task 'FTP Models Phone Alerts' registered (every 5 min)."
Write-Host "Check: Task Scheduler -> Task Scheduler Library -> FTP Models Phone Alerts"
