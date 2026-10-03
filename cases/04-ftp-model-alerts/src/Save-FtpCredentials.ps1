# Name: Save-FtpCredentials.ps1
# Version: 1.0
# What it does: Store FTP password in encrypted credentials.xml for the current Windows user.
# Inputs: Get-Credential prompt.
# Outputs: credentials.xml in src\.
# How to run: powershell -File Save-FtpCredentials.ps1
# Notes: Clixml; readable only under the same Windows account.
param(
    [string]$Username = "ftpuser"
)

$cred = Get-Credential -UserName $Username -Message "Пароль FTP (как в FileZilla)"
$path = Join-Path $PSScriptRoot "credentials.xml"
$cred | Export-Clixml -Path $path
Write-Host "Сохранено: $path"
Write-Host "Файл читается только под этой учётной записью Windows."
