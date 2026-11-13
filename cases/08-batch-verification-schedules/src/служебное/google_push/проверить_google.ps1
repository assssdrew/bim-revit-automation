# UTF-8 WITH BOM required (PowerShell 5.1).
# Проверка google_push.cfg и связи с веб-приложением (без Revit).

$ErrorActionPreference = "Stop"

$tool = Join-Path (Split-Path -Parent $PSScriptRoot) "скрипты"          # служебное\google_push -> служебное\скрипты
. (Join-Path $tool "google_push_cfg.ps1")

$cfg = Read-GooglePushConfig $tool
Write-Host ("Конфиг: " + $cfg.ConfigPath)

$issues = @(Test-BvsGooglePushConfigShape $cfg)
if ($issues.Count -gt 0) {
    foreach ($i in $issues) { Write-Host ("ОШИБКА: " + $i) }
    exit 1
}

$pingUrl = $cfg.WebAppUrl.TrimEnd('/') + "?ping=1"
Write-Host ("GET ping: " + $pingUrl)
try {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $ping = (Invoke-WebRequest -Uri $pingUrl -UseBasicParsing -TimeoutSec 60).Content.Trim()
    if ($ping -match '(?i)<\s*(!doctype|html)') { $ping = "html" }
    if ($ping -eq "ok") {
        Write-Host "OK ping"
    }
    elseif ($ping -match "doGet") {
        Write-Host "ВНИМАНИЕ: в развернутом скрипте нет doGet — вставьте Code.gs из репозитория и «новая версия» развертывания."
    }
    else {
        Write-Host "ВНИМАНИЕ: ping не ok (это не блокирует POST, если секрет верный)."
    }
}
catch {
    Write-Host ("GET ping не удался (не критично): " + $_.Exception.Message)
}

$body = @{
    secret = $cfg.Secret
    hide   = @()
    sheets = @()
} | ConvertTo-Json -Compress
$bytes = (New-Object System.Text.UTF8Encoding $false).GetBytes($body)
Write-Host "POST test (пустые листы)..."
$text = (Invoke-GoogleAppsScriptPost -Url $cfg.WebAppUrl -Body $bytes -Attempts 1)
$res = Get-GooglePushResult $text
if ($res.Ok) {
    if ($res.Legacy) { Write-Host "OK POST (старая версия Code.gs: вставьте новый Code.gs и сделайте «новая версия» развертывания)" }
    else { Write-Host "OK POST (секрет принят, веб-приложение отвечает, версия Code.gs актуальная)" }
    exit 0
}
if ($res.Message -match 'forbidden') {
    Write-Host "ОШИБКА: forbidden — secret в cfg не совпадает с PUSH_SECRET (свойство скрипта или Code.gs; или старая версия развертывания)."
    exit 2
}
Write-Host ("ОШИБКА: " + $res.Message)
exit 3
