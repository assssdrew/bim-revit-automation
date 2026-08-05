<#
.SYNOPSIS
  Один раз сохраняет логин/пароль FTP в зашифрованном credentials.xml (под текущим Windows-пользователем).
#>

param(
    [string]$Username = "ftpuser"
)

$cred = Get-Credential -UserName $Username -Message "Пароль FTP (как в FileZilla)"
$path = Join-Path $PSScriptRoot "credentials.xml"
$cred | Export-Clixml -Path $path
Write-Host "Сохранено: $path"
Write-Host "Файл читается только под этой учётной записью Windows."
