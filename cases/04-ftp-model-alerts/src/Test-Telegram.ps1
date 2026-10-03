# Name: Test-Telegram.ps1
# Version: 1.0
# What it does: Send a test Telegram message to verify API access.
# Inputs: config.json bot settings.
# Outputs: Test chat message.
# How to run: powershell -File Test-Telegram.ps1
# Notes: Quick connectivity check.
$ErrorActionPreference = "Stop"

try {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls
} catch {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
}

$config = Get-Content (Join-Path $PSScriptRoot "config.json") -Raw -Encoding UTF8 | ConvertFrom-Json
$token = [string]$config.TelegramBotToken
$chatId = [string]$config.TelegramChatId

Write-Host "1) DNS..."
try {
    $dns = [System.Net.Dns]::GetHostAddresses("api.telegram.org")
    Write-Host "   OK:" ($dns | ForEach-Object { $_.IPAddressToString }) -join ", "
} catch {
    Write-Host "   FAIL DNS: $($_.Exception.Message)"
}

Write-Host "2) TCP 443..."
$tnc = Test-NetConnection api.telegram.org -Port 443 -WarningAction SilentlyContinue
Write-Host "   TcpTestSucceeded=$($tnc.TcpTestSucceeded)"

Write-Host "3) getMe..."
$me = Invoke-RestMethod -Uri "https://api.telegram.org/bot$token/getMe"
Write-Host "   OK bot=@$($me.result.username)"

Write-Host "4) sendMessage to chat $chatId ..."
$payload = @{
    chat_id = $chatId
    text    = "FTP alerts test OK"
} | ConvertTo-Json -Compress
$utf8 = New-Object System.Text.UTF8Encoding $false
Invoke-RestMethod -Method Post -Uri "https://api.telegram.org/bot$token/sendMessage" -ContentType "application/json; charset=utf-8" -Body ($utf8.GetBytes($payload)) | Out-Null
Write-Host "   OK message sent. Check Telegram."
