# UTF-8 WITH BOM required (PowerShell 5.1).
# Отправка листов «Сводка» и разделов в Google-таблицу через Apps Script Web App.

param(
    [string]$XlsxPath = "",
    [string]$ToolDir = "",
    [string]$PayloadPath = "",
    [switch]$ExportPayloadOnly,
    [switch]$PayloadOnly,
    [string]$ProjectConfigPath = "",      # распределение_<проект>.txt: адрес Google берётся из google_<проект>.cfg рядом с ним
    [switch]$AllowShared                   # пользователь подтвердил отправку в общую таблицу из google_push.cfg
)

$ErrorActionPreference = "Stop"

try {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
}
catch { }

$__pushDir = Split-Path -Parent $MyInvocation.MyCommand.Path
. (Join-Path $__pushDir "google_push_cfg.ps1")
. (Join-Path $__pushDir "google_payload_lib.ps1")

$PushSheetNames = $script:BvsGooglePushSheetNames
$HideSheetNames = $script:BvsGoogleHideSheetNames

function Get-ToolDir {
    if (-not [string]::IsNullOrWhiteSpace($ToolDir) -and (Test-Path -LiteralPath $ToolDir)) {
        return $ToolDir
    }
    return $__pushDir
}

if (-not $PayloadOnly -and -not $ExportPayloadOnly -and [string]::IsNullOrWhiteSpace($XlsxPath)) {
    throw "Укажите -XlsxPath."
}

$script:BvsPushDummyXlsx = $null
if ($PayloadOnly) {
    if ([string]::IsNullOrWhiteSpace($PayloadPath)) {
        throw "Для -PayloadOnly укажите -PayloadPath (google_push_payload.json)."
    }
    if (-not (Test-Path -LiteralPath $PayloadPath)) {
        throw "Нет payload: $PayloadPath"
    }
    $XlsxPath = Join-Path ([System.IO.Path]::GetTempPath()) ("bvs_push_dummy_{0}.xlsx" -f [Guid]::NewGuid().ToString("N"))
    [IO.File]::WriteAllBytes($XlsxPath, [byte[]](0x50, 0x4B, 0x05, 0x06, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0))
    $script:BvsPushDummyXlsx = $XlsxPath
}
elseif (-not (Test-Path -LiteralPath $XlsxPath)) {
    throw "Нет файла задания: $XlsxPath"
}

$root = Get-ToolDir
$cfg = @{ WebAppUrl = ""; Secret = "" }
if (-not $ExportPayloadOnly) {
    if ($ProjectConfigPath) {
        # адрес на проект; общий google_push.cfg — только после подтверждения (-AllowShared)
        $target = Resolve-GooglePushTarget $ProjectConfigPath $root $AllowShared.IsPresent
        if ($target.Kind -eq "none") { throw "У проекта нет своего адреса Google, общий google_push.cfg не настроен." }
        if ($target.Kind -eq "needs-confirm") { throw "У проекта нет своего адреса Google. Общая таблица не используется без подтверждения." }
        $cfg = $target.Cfg
        if ($target.Warning) { Write-Output $target.Warning }          # общий запасной google_push.cfg (только с -AllowShared)
        Write-Output ("Google: " + $target.Label)        # куда уходит отправка (проект + конец адреса), без секрета
    }
    else {
        $cfg = Read-GooglePushConfig $root
    }
    $cfgIssues = @(Test-BvsGooglePushConfigShape $cfg)
    if ($cfgIssues.Count -gt 0) {
        $where = $cfg.ConfigPath
        if ([string]::IsNullOrWhiteSpace($where)) { $where = (Join-Path (Split-Path -Parent $root) "google_push\google_push.cfg") }
        throw ("Настройка Google: " + ($cfgIssues -join " ") + " Файл: " + $where)
    }
}

$sheetsPayload = New-Object System.Collections.Generic.List[object]
# выборочная выгрузка: payload сам говорит, какие листы очистить (только затронутые разделы без замечаний)
$partialClear = $null
$payloadFile = $PayloadPath
if ([string]::IsNullOrWhiteSpace($payloadFile)) {
    $payloadFile = Join-Path (Split-Path -Parent $XlsxPath) "google_push_payload.json"
}
if (-not $ExportPayloadOnly -and (Test-Path -LiteralPath $payloadFile)) {
    $raw = [System.IO.File]::ReadAllText($payloadFile, (New-Object System.Text.UTF8Encoding $false))
    $parsed = $raw | ConvertFrom-Json
    if ($parsed.partial) { $partialClear = @($parsed.clear | Where-Object { $_ }) }
    if ($parsed.sheets) {
        $byName = @{}
        foreach ($sh in @($parsed.sheets)) { $byName[[string]$sh.name] = $sh }
        foreach ($name in $PushSheetNames) {
            if ($byName.ContainsKey($name)) {
                [void]$sheetsPayload.Add((Normalize-PayloadSheet $byName[$name]))
            }
        }
    }
}
if ($sheetsPayload.Count -lt 1) {
    . (Join-Path $__pushDir "bvs_assignment_guard.ps1")
    if (-not (Test-BvsAssignmentXlsx $XlsxPath)) {
        throw "xlsx не открывается для Google (повреждён или устаревшая сборка). Пересоберите задание."
    }
    . (Join-Path $__pushDir "excel_isolated.ps1")
    $excelHandle = New-IsolatedExcelApplication
    $excel = $excelHandle.Application
    try {
        $book = $excel.Workbooks.Open($XlsxPath, 0, $true)
        try {
            $built = Build-GoogleSheetsPayloadFromWorkbook $book
            foreach ($sh in @($built)) { [void]$sheetsPayload.Add($sh) }
        }
        finally {
            try { $book.Close($false) } catch { }
        }
    }
    finally {
        Stop-IsolatedExcelApplication $excelHandle
    }
}

if ($sheetsPayload.Count -lt 1) {
    throw "В файле нет листов для Google."
}

if ($ExportPayloadOnly) {
    $outPath = $payloadFile
    if ([string]::IsNullOrWhiteSpace($outPath)) {
        $outPath = Join-Path (Split-Path -Parent $XlsxPath) "google_push_payload.json"
    }
    Write-GooglePushPayloadFile $sheetsPayload $outPath
    Write-Output ("OK payload " + $outPath)
    return
}

$sheetArray = @($sheetsPayload.ToArray())
# Листы разделов, которых нет в этой выгрузке (замечаний не осталось): Code.gs очистит их,
# иначе в Google останутся данные прошлой отправки и таблица разойдётся с xlsx.
$present = @{}
foreach ($sh in $sheetArray) { $present[[string]$sh.name] = $true }
$clearNames = @($PushSheetNames | Where-Object { $_ -ne "Сводка" -and -not $present.ContainsKey($_) })
if ($null -ne $partialClear) { $clearNames = @($partialClear | Where-Object { $PushSheetNames -contains $_ -and $_ -ne "Сводка" -and -not $present.ContainsKey($_) }) }
$body = @{
    secret = $cfg.Secret
    hide   = @($HideSheetNames)
    clear  = $clearNames
    sheets = $sheetArray
}
$json = $body | ConvertTo-Json -Depth 30 -Compress
$utf8 = New-Object System.Text.UTF8Encoding $false
$payload = $utf8.GetBytes($json)
try {
    $bodyText = (Invoke-GoogleAppsScriptPost -Url $cfg.WebAppUrl -Body $payload)
    $result = Get-GooglePushResult $bodyText
    if (-not $result.Ok) {
        throw ("Google отклонил запрос: " + $result.Message)
    }
    $note = ""
    if ($result.Legacy) { $note = " (старая версия Code.gs: обновите развёртывание, иначе очистка листов и защита от переполнения групп не работают)" }
    Write-Output ("OK " + $result.Message + $note)
}
catch {
    throw ("Google: " + $_.Exception.Message)
}
finally {
    if ($script:BvsPushDummyXlsx -and (Test-Path -LiteralPath $script:BvsPushDummyXlsx)) {
        Remove-Item -LiteralPath $script:BvsPushDummyXlsx -Force -ErrorAction SilentlyContinue
    }
}
