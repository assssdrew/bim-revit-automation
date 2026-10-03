# Name: Test-Ntfy.ps1
# Version: 1.0
# What it does: Send a test message via ntfy.sh to verify outbound HTTPS.
# Inputs: config.json ntfy settings.
# Outputs: Test notification.
# How to run: powershell -File Test-Ntfy.ps1
# Notes: Quick connectivity check.
$ErrorActionPreference = "Stop"

try {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls
} catch {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
}

$config = Get-Content (Join-Path $PSScriptRoot "config.json") -Raw -Encoding UTF8 | ConvertFrom-Json
$topic = ([string]$config.NtfyTopic).Trim()
$topic = $topic -replace '^https?://ntfy\.sh/', ''
$topic = $topic -replace '^ntfy\.sh/', ''
$topic = $topic.Trim('/')

Write-Host "1) DNS ntfy.sh..."
try {
    $dns = [System.Net.Dns]::GetHostAddresses("ntfy.sh")
    Write-Host "   OK:" ($dns | ForEach-Object { $_.IPAddressToString }) -join ", "
} catch {
    Write-Host "   FAIL DNS: $($_.Exception.Message)"
    throw
}

Write-Host "2) TCP 443..."
$tnc = Test-NetConnection ntfy.sh -Port 443 -WarningAction SilentlyContinue
Write-Host "   TcpTestSucceeded=$($tnc.TcpTestSucceeded)"
if (-not $tnc.TcpTestSucceeded) {
    throw "Cannot reach ntfy.sh:443"
}

Write-Host "3) Send test to topic '$topic'..."
$payload = @{
    topic    = $topic
    title    = "FTP alerts test"
    message  = "ntfy OK from home PC"
    priority = 3
} | ConvertTo-Json -Compress

$utf8 = New-Object System.Text.UTF8Encoding $false
Invoke-RestMethod -Method Post -Uri "https://ntfy.sh" -ContentType "application/json; charset=utf-8" -Body ($utf8.GetBytes($payload)) | Out-Null
Write-Host "   OK. Check ntfy app on iPhone (topic: $topic)"
