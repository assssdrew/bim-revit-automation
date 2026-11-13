# UTF-8 WITH BOM required (PowerShell 5.1).
# Вариант 1 из выгрузки сводной. Имя спецификации берётся целиком из export.txt [manifest] (прежде manifest.txt).

param(
    [string]$ExportDir = '',
    [Parameter(Mandatory = $true)][string]$OutXlsx,
    [string]$HistoryPath = '',
    [string]$DynamicsFrom = '',
    [string]$DynamicsTo = '',
    [switch]$DynamicsOnly,
    # схлопывать сборки за день: остаётся только последняя за сегодня (по умолчанию каждая сборка — отдельный срез)
    [switch]$CollapseDay,
    # сколько дней исправленные замечания остаются на листах разделов
    [int]$KeepFixedDays = 30,
    # выборочная выгрузка (assign_scope=partial): проверенные пары обновляются, остальные берутся из буфера (assignment_partial.ps1)
    [switch]$Partial,
    # куда класть служебные файлы прогона (JSON для Google, копия фильтров); по умолчанию — рядом с xlsx (прежнее поведение)
    [string]$AuxDir = '',
    # префикс проекта в именах связей/моделей (PROJ в PROJ_IOS_OV_…); из настройки проекта. Пусто — любой префикс
    [string]$ProjectPrefix = ''
)

$ErrorActionPreference = "Stop"
try {
    $utf8Out = New-Object System.Text.UTF8Encoding $false
    [Console]::OutputEncoding = $utf8Out
    $OutputEncoding = $utf8Out
}
catch { }

# Порядок листов — разделы.txt. Нет файла — встроенный список (тот же) и предупреждение.
try {
    if (-not (Get-Command Get-BvsReportCodes -ErrorAction SilentlyContinue)) {
        $__bvsSections = if ($PSScriptRoot) { Join-Path $PSScriptRoot "bvs_sections.ps1" } else { Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "bvs_sections.ps1" }
        if (Test-Path -LiteralPath $__bvsSections) { . $__bvsSections }
    }
}
catch { Write-Warning ("разделы: " + $_.Exception.Message) }
if (Get-Command Get-BvsReportCodes -ErrorAction SilentlyContinue) {
    $SectionOrder = @(Get-BvsReportCodes)
    $SectionTitle = Get-BvsTitleMap
}
else {
    Write-Warning "встроенный порядок листов задания (bvs_sections.ps1 не подключён)"
    $SectionOrder = @("EOM", "SS", "OV", "PD", "PT", "VK", "OT", "HS", "ITP", "AVT", "HC")
    $SectionTitle = @{
        EOM = "ЭОМ"; SS = "СС"; OV = "ОВ"; PD = "ПД"; PT = "ПТ"; VK = "ВК"; OT = "ОТ"; HS = "ХС"; ITP = "ИТП"; AVT = "АВТ"; HC = "ХЦ"
    }
}
$Sep = [char]0x1F

function To-Text($Value) {
    if ($null -eq $Value) { return "" }
    return ([string]$Value).Trim()
}

function Get-ModelFileName([string]$Link) {
    $text = $Link.Trim()
    if ($text -match '[^\\/]+$') { $text = $Matches[0] }
    return ($text -replace '(?i)\.rvt$', '')
}

function Get-ModelSort([string]$FileName) {
    $body = $FileName
    if ($FileName -match 'ORG_(.+)_R\d+$') { $body = $Matches[1] }
    if ($body -match '^K(\d+)$') { return [int]$Matches[1] }
    if ($body -match '^K(\d+)_') { return 1000 + [int]$Matches[1] }
    return 8000
}

function Count-Text([int]$Now, $Prev) {
    if ($null -eq $Prev) { return [string]$Now }
    $was = [int]$Prev
    if ($was -ne $Now) { return ("{0} ({1})" -f $Now, $was) }
    return [string]$Now
}

function Get-SectionFromModel([string]$Model) {
    if ($Model -match '(?i)_IOS_([A-Z]+)_') { return $Matches[1].ToUpper() }
    $s = Get-BvsLinkSection $Model          # имена без «_IOS_» (SampleSite): _AVT_, _HC_ …
    if ($s) { return ([string]$s).ToUpper() }
    return ""
}

function Read-History([string]$Path) {
    $list = New-Object System.Collections.Generic.List[object]
    if (-not (Test-Path -LiteralPath $Path)) { return $list }
    foreach ($line in @(Get-Content -LiteralPath $Path -Encoding UTF8)) {
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        # 7-я колонка (необязательная) — имя параметра для показа; ключ (Param, 5-я) остаётся прежним
        # 8-я колонка (необязательная) — идентификатор сборки «yyyy-MM-dd HH:mm:ss» (срез); у старых строк — дата + 00:00:00
        $parts = @($line -split "`t", 8)
        if ($parts.Count -ge 6) {
            $shown = ""
            if ($parts.Count -ge 7) { $shown = $parts[6].Trim() }
            $bid = $parts[0].Trim() + " 00:00:00"
            if ($parts.Count -ge 8 -and $parts[7].Trim() -match '^\d{4}-\d\d-\d\d \d\d:\d\d:\d\d$') { $bid = $parts[7].Trim() }
            $n = 0
            [void][int]::TryParse($parts[5].Trim(), [ref]$n)
            $list.Add([pscustomobject]@{
                Date    = $parts[0].Trim()
                Section = $parts[1].Trim()
                Model   = $parts[2].Trim()
                Check   = (ConvertTo-CanonicalSpecName $parts[3].Trim())
                Param   = $parts[4].Trim()
                ParamName = $shown
                Count   = $n
                BuildId = $bid
            })
            continue
        }
        $parts = @($line -split "`t", 4)
        if ($parts.Count -lt 4) { continue }
        $n = 0
        [void][int]::TryParse($parts[3].Trim(), [ref]$n)
        $model = $parts[1].Trim()
        $list.Add([pscustomobject]@{
            Date    = $parts[0].Trim()
            Section = (Get-SectionFromModel $model)
            Model   = $model
            Check   = (ConvertTo-CanonicalSpecName $parts[2].Trim())
            Param   = ""
            Count   = $n
            BuildId = ($parts[0].Trim() + " 00:00:00")
        })
    }
    return $list
}

function Get-HistoryKey($Row) {
    return ($Row.Model + [char]0x1F + $Row.Check + [char]0x1F + $Row.Param)
}

function Get-DynamicsCorrections($HistoryRows, [string]$From, [string]$To) {
    $out = New-Object System.Collections.Generic.List[object]
    if ([string]::IsNullOrWhiteSpace($From) -or [string]::IsNullOrWhiteSpace($To)) { return $out }
    $byKey = @{}
    foreach ($h in @($HistoryRows)) {
        $k = Get-HistoryKey $h
        if (-not $byKey.ContainsKey($k)) { $byKey[$k] = New-Object System.Collections.Generic.List[object] }
        $byKey[$k].Add($h)
    }
    foreach ($k in @($byKey.Keys)) {
        $rows = @($byKey[$k] | Sort-Object BuildId)
        $baseline = $null
        foreach ($r in $rows) {
            if ($r.Date -lt $From) { $baseline = $r }
        }
        $inRange = @($rows | Where-Object { $_.Date -ge $From -and $_.Date -le $To })
        if ($inRange.Count -lt 1) { continue }
        $end = $inRange[$inRange.Count - 1]
        $startCount = 0
        if ($null -ne $baseline) { $startCount = [int]$baseline.Count }
        else { $startCount = [int]$inRange[0].Count }
        $endCount = [int]$end.Count
        $done = $startCount - $endCount
        if ($done -lt 1) { continue }
        $sample = $end
        $out.Add([pscustomobject]@{
            Section     = $sample.Section
            Model       = $sample.Model
            Check       = $sample.Check
            Param       = $sample.Param
            ParamName   = $sample.ParamName
            Corrections = $done
            Family      = ""
            Item        = ""
            Value       = ""
        })
    }
    return $out
}


# Колонки: Модель, Спецификация, Параметр, Семейство и типоразмер, Наименование, Значение (+ Статус, Найдено, Исправлено, Дней).
# Замечание — строка спецификации; счётчика «Замечаний» нет: строки модели и спецификации — только заголовки групп.
function Build-SectionOutlineBlock($Items, [int]$StartExcelRow, [switch]$Wide) {
    $cols = 6
    if ($Wide) { $cols = 10 }
    $lines = New-Object System.Collections.Generic.List[object]
    $inners = New-Object System.Collections.Generic.List[object]
    $outers = New-Object System.Collections.Generic.List[object]
    $modelRows = New-Object System.Collections.Generic.List[int]
    $checkRows = New-Object System.Collections.Generic.List[int]
    $fixedRows = New-Object System.Collections.Generic.List[int]
    $newRows = New-Object System.Collections.Generic.List[int]
    $byModel = @{}
    $modelOrder = New-Object System.Collections.Generic.List[string]
    foreach ($it in @($Items)) {
        if (-not $byModel.ContainsKey($it.Model)) {
            $byModel[$it.Model] = New-Object System.Collections.Generic.List[object]
            $modelOrder.Add($it.Model)
        }
        $byModel[$it.Model].Add($it)
    }
    $models = @($modelOrder | Sort-Object { Get-ModelSort $_ }, { $_ })
    $excelRow = $StartExcelRow
    foreach ($model in $models) {
        if (-not $byModel.ContainsKey($model)) { continue }
        $modelItems = @($byModel[$model] | ForEach-Object { $_ })
        if ($modelItems.Count -lt 1) { continue }
        $fixedCount = @($modelItems | Where-Object { $_.Status -eq "fixed" }).Count
        $line = New-Object object[] $cols
        for ($i = 0; $i -lt $cols; $i++) { $line[$i] = "" }
        $line[0] = $model
        if ($Wide -and $fixedCount -gt 0) { $line[8] = $fixedCount }
        $lines.Add($line)
        $modelRows.Add($excelRow)
        $excelRow++
        $outerFirst = $excelRow
        $checks = @(Sort-CheckNames @($modelItems | ForEach-Object { $_.Check } | Select-Object -Unique))
        foreach ($check in $checks) {
            $chunk = @($modelItems | Where-Object { $_.Check -eq $check })
            if ($chunk.Count -lt 1) { continue }
            $param = $chunk[0].ParamName
            if (-not $param) { $param = $chunk[0].Param }
            $chunkFixed = @($chunk | Where-Object { $_.Status -eq "fixed" }).Count
            $line = New-Object object[] $cols
            for ($i = 0; $i -lt $cols; $i++) { $line[$i] = "" }
            $line[1] = $check; $line[2] = $param
            if ($Wide -and $chunkFixed -gt 0) { $line[8] = $chunkFixed }
            $lines.Add($line)
            $checkRows.Add($excelRow)
            $excelRow++
            $itemFirst = $excelRow
            foreach ($it in $chunk) {
                $line = New-Object object[] $cols
                for ($i = 0; $i -lt $cols; $i++) { $line[$i] = "" }
                $line[3] = $it.Family; $line[4] = $it.Item; $line[5] = $it.Value
                if ($Wide) {
                    if ($it.Status -eq "fixed") {
                        $line[6] = "Исправлено"
                        $line[7] = Format-DayRu $it.FirstSeen
                        $line[8] = Format-DayRu $it.FixedAt
                        $line[9] = Get-DaysBetween $it.FirstSeen $it.FixedAt
                        $fixedRows.Add($excelRow)
                    }
                    else {
                        if ($it.IsNew) { $line[6] = "Новое"; $newRows.Add($excelRow) } else { $line[6] = "Открыто" }
                        $line[7] = Format-DayRu $it.FirstSeen
                        $line[9] = Get-DaysBetween $it.FirstSeen $script:BuildNow
                    }
                }
                $lines.Add($line)
                $excelRow++
            }
            $itemLast = $excelRow - 1
            if ($itemLast -ge $itemFirst) {
                $inners.Add([pscustomobject]@{ First = $itemFirst; Last = $itemLast })
            }
        }
        $outerLast = $excelRow - 1
        if ($outerLast -ge $outerFirst) {
            $outers.Add([pscustomobject]@{ First = $outerFirst; Last = $outerLast })
        }
    }
    $grid = New-Object 'object[,]' $lines.Count, $cols
    for ($r = 0; $r -lt $lines.Count; $r++) {
        $line = $lines[$r]
        for ($c = 0; $c -lt $cols; $c++) { $grid[$r, $c] = $line[$c] }
    }
    return @{
        Grid      = $grid
        Inners    = $inners
        Outers    = $outers
        ModelRows = $modelRows
        CheckRows = $checkRows
        FixedRows = $fixedRows
        NewRows   = $newRows
        Cols      = $cols
    }
}


function Build-DynamicsOutlineBlock($Items, [int]$StartExcelRow) {
    $lines = New-Object System.Collections.Generic.List[object]
    $inners = New-Object System.Collections.Generic.List[object]
    $outers = New-Object System.Collections.Generic.List[object]
    $byModel = @{}
    $modelOrder = New-Object System.Collections.Generic.List[string]
    foreach ($it in @($Items)) {
        if (-not $byModel.ContainsKey($it.Model)) {
            $byModel[$it.Model] = New-Object System.Collections.Generic.List[object]
            $modelOrder.Add($it.Model)
        }
        $byModel[$it.Model].Add($it)
    }
    $models = @($modelOrder | Sort-Object { Get-ModelSort $_ }, { $_ })
    $excelRow = $StartExcelRow
    foreach ($model in $models) {
        $modelItems = @($byModel[$model] | ForEach-Object { $_ })
        $modelSum = 0
        foreach ($mi in $modelItems) { $modelSum += [int]$mi.Corrections }
        $lines.Add(@($model, $modelSum, "", "", "", "", "", "", "", ""))
        $excelRow++
        $outerFirst = $excelRow
        $checks = @(Sort-CheckNames @($modelItems | ForEach-Object { $_.Check } | Select-Object -Unique))
        foreach ($check in $checks) {
            $chunk = @($modelItems | Where-Object { $_.Check -eq $check })
            if ($chunk.Count -lt 1) { continue }
            $checkSum = 0
            foreach ($ch in $chunk) { $checkSum += [int]$ch.Corrections }
            $lines.Add(@("", $checkSum, $check, "", "", "", "", "", "", ""))
            $excelRow++
            $itemFirst = $excelRow
            foreach ($it in $chunk) {
                $shownParam = $it.ParamName
                if (-not $shownParam) { $shownParam = $it.Param }
                $found = ""; $fixedOn = ""; $days = ""
                if ($it.FirstSeen) { $found = Format-DayRu $it.FirstSeen }
                if ($it.FixedAt) { $fixedOn = Format-DayRu $it.FixedAt; $days = Get-DaysBetween $it.FirstSeen $it.FixedAt }
                $lines.Add(@("", [int]$it.Corrections, "", $shownParam, $it.Family, $it.Item, $it.Value, $found, $fixedOn, $days))
                $excelRow++
            }
            $itemLast = $excelRow - 1
            if ($itemLast -ge $itemFirst) {
                $inners.Add([pscustomobject]@{ First = $itemFirst; Last = $itemLast })
            }
        }
        $outerLast = $excelRow - 1
        if ($outerLast -ge $outerFirst) {
            $outers.Add([pscustomobject]@{ First = $outerFirst; Last = $outerLast })
        }
    }
    $cols = 10
    $grid = New-Object 'object[,]' $lines.Count, $cols
    for ($r = 0; $r -lt $lines.Count; $r++) {
        $line = $lines[$r]
        for ($c = 0; $c -lt $cols; $c++) { $grid[$r, $c] = $line[$c] }
    }
    return @{ Grid = $grid; Inners = $inners; Outers = $outers; Cols = $cols }
}

if ($PSScriptRoot) {
    $__buildDir = $PSScriptRoot
}
else {
    $__buildDir = Split-Path -Parent $MyInvocation.MyCommand.Path
}
$script:BvsAssignmentDir = $__buildDir
$script:BvsExcelModuleDir = $__buildDir
. (Join-Path $__buildDir "assignment_export_stream.ps1")
$script:BvsProjectPrefix = $ProjectPrefix          # после подключения модуля (он сбрасывает префикс)
. (Join-Path $__buildDir "assignment_ledger.ps1")
# буфер выборочной выгрузки: обязателен только для -Partial; в режиме «Всё» его отсутствие ничего не ломает
$script:PartialModuleLoaded = $false
$__partialMod = Join-Path $__buildDir "assignment_partial.ps1"
if (Test-Path -LiteralPath $__partialMod) { . $__partialMod; $script:PartialModuleLoaded = $true }
elseif ($Partial) { throw "Нет assignment_partial.ps1 в папке инструмента. Скопируйте папку инструмента целиком." }

function Parse-DateYmd([string]$Text) {
    if ([string]::IsNullOrWhiteSpace($Text)) { return $null }
    try { return [datetime]::Parse($Text).Date } catch { return $null }
}

# Время сборки. Переменная окружения BVS_FAKE_NOW="yyyy-MM-dd HH:mm:ss" нужна только тестам (детерминированные срезы);
# без неё — текущее время, поведение прежнее.
function Get-BuildNow {
    if ($env:BVS_FAKE_NOW) {
        try { return [datetime]::ParseExact($env:BVS_FAKE_NOW, "yyyy-MM-dd HH:mm:ss", [System.Globalization.CultureInfo]::InvariantCulture) } catch { }
    }
    return (Get-Date)
}

function Resolve-DynamicsPeriod([string]$From, [string]$To) {
    $today = Get-BuildNow
    $toDate = $today.Date
    $fromDate = $toDate.AddDays(-6)
    $parsedTo = Parse-DateYmd $To
    if ($null -ne $parsedTo) { $toDate = $parsedTo }
    $parsedFrom = Parse-DateYmd $From
    if ($null -ne $parsedFrom) { $fromDate = $parsedFrom }
    if ($fromDate -gt $toDate) { $swap = $fromDate; $fromDate = $toDate; $toDate = $swap }
    $fromS = $fromDate.ToString("yyyy-MM-dd")
    $toS = $toDate.ToString("yyyy-MM-dd")
    $label = ("Период: {0} — {1}." -f $fromDate.ToString("dd.MM.yyyy"), $toDate.ToString("dd.MM.yyyy"))
    return @{ From = $fromS; To = $toS; Label = $label }
}

# ---------- история по срезам (каждая сборка — отдельный срез) ----------
function New-CurrentSlice($Stream, $HistoryRows, [string]$PrevBuild, [string]$BuildId, $CheckedPairs = $null) {
    $rows = New-Object System.Collections.Generic.List[object]
    $day = $BuildId.Substring(0, 10)
    $seen = @{}
    foreach ($b in @($Stream.HistoryBuckets.Values)) {
        if ([int]$b.Count -lt 1) { continue }
        $seen[($b.Model + "`t" + $b.Check + "`t" + $b.Param)] = $true
        $rows.Add([pscustomobject]@{
            Date = $day; Section = $b.Section; Model = $b.Model; Check = $b.Check; Param = $b.Param
            ParamName = $b.ParamName; Count = [int]$b.Count; BuildId = $BuildId
        })
    }
    # ноль вместо пропажи: ключ был в прошлом срезе, спецификация проверена сейчас, строк нет -> исправлено
    # (без списка проверенных спецификаций — например, если Revit не нашёл ни одной — нулей не пишем)
    # выборочная выгрузка: ноль только для проверенной пары «модель + спецификация», остальные пары не трогаем
    if ($PrevBuild -and $Stream.CheckedSpecs.Count -gt 0) {
        foreach ($h in @($HistoryRows)) {
            if ($h.BuildId -ne $PrevBuild -or [int]$h.Count -lt 1) { continue }
            if ($null -ne $CheckedPairs) {
                if (-not $CheckedPairs.ContainsKey(([string]$h.Model + [char]0x1F + [string]$h.Check))) { continue }
            }
            elseif (-not $Stream.CheckedSpecs.ContainsKey($h.Check)) { continue }
            $k = $h.Model + "`t" + $h.Check + "`t" + $h.Param
            if ($seen.ContainsKey($k)) { continue }
            $seen[$k] = $true
            $rows.Add([pscustomobject]@{
                Date = $day; Section = $h.Section; Model = $h.Model; Check = $h.Check; Param = $h.Param
                ParamName = $h.ParamName; Count = 0; BuildId = $BuildId
            })
        }
    }
    return $rows
}

function Write-HistoryRows([string]$Path, $Rows) {
    $lines = New-Object System.Collections.Generic.List[string]
    foreach ($h in @($Rows)) {
        $lines.Add(("{0}`t{1}`t{2}`t{3}`t{4}`t{5}`t{6}`t{7}" -f $h.Date, $h.Section, $h.Model, $h.Check, $h.Param, $h.Count, $h.ParamName, $h.BuildId))
    }
    $dir = Split-Path -Parent $Path
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $bom = New-Object System.Text.UTF8Encoding $true
    [System.IO.File]::WriteAllLines($Path, $lines.ToArray(), $bom)
}

# старше N дней оставляем только последний срез каждого дня, чтобы файл не рос бесконечно
function Compress-OldHistory($Rows, [string]$Today, [int]$KeepDays = 60) {
    $limit = ([datetime]::Parse($Today)).AddDays(-$KeepDays).ToString("yyyy-MM-dd")
    $lastOfDay = @{}
    foreach ($h in @($Rows)) {
        if ($h.Date -ge $limit) { continue }
        if (-not $lastOfDay.ContainsKey($h.Date) -or $h.BuildId -gt $lastOfDay[$h.Date]) { $lastOfDay[$h.Date] = $h.BuildId }
    }
    return @($Rows | Where-Object { $_.Date -ge $limit -or $_.BuildId -eq $lastOfDay[$_.Date] })
}

# «Динамика»: каждое исправленное замечание отдельной строкой (из журнала) + остаток из истории по счётчикам,
# которого журнал не знает (исправлено до ведения журнала): такие строки без детализации.
function Get-DynamicsItems($HistoryRows, [string]$From, [string]$To, $Ledger) {
    $out = New-Object System.Collections.Generic.List[object]
    $fixed = @(Get-LedgerFixedInPeriod $Ledger $From $To)
    $perSpec = @{}
    foreach ($rec in $fixed) {
        $out.Add([pscustomobject]@{
            Section = $rec.Section; Model = $rec.Model; Check = $rec.Check; Param = $rec.ParamName; ParamName = $rec.ParamName
            Corrections = 1; Family = $rec.Family; Item = $rec.Item; Value = $rec.Value
            FirstSeen = $rec.FirstSeen; FixedAt = $rec.FixedAt
        })
        $k = $rec.Model + $Sep + $rec.Check
        if (-not $perSpec.ContainsKey($k)) { $perSpec[$k] = 0 }
        $perSpec[$k] = [int]$perSpec[$k] + 1
    }
    $histSum = @{}
    $sample = @{}
    foreach ($h in @(Get-DynamicsCorrections $HistoryRows $From $To)) {
        $k = $h.Model + $Sep + $h.Check
        if (-not $histSum.ContainsKey($k)) { $histSum[$k] = 0 }
        $histSum[$k] = [int]$histSum[$k] + [int]$h.Corrections
        $sample[$k] = $h
    }
    foreach ($k in @($histSum.Keys)) {
        $have = 0
        if ($perSpec.ContainsKey($k)) { $have = [int]$perSpec[$k] }
        $rest = [int]$histSum[$k] - $have
        if ($rest -lt 1) { continue }
        $h = $sample[$k]
        $out.Add([pscustomobject]@{
            Section = $h.Section; Model = $h.Model; Check = $h.Check; Param = $h.Param; ParamName = $h.ParamName
            Corrections = $rest; Family = "(без детализации: исправлено до ведения журнала)"; Item = ""; Value = ""
            FirstSeen = ""; FixedAt = ""
        })
    }
    return $out.ToArray()
}

$buildId = (Get-BuildNow).ToString("yyyy-MM-dd HH:mm:ss")
$script:BuildNow = $buildId
$today = $buildId.Substring(0, 10)
if ([string]::IsNullOrWhiteSpace($HistoryPath)) {
    $historyPath = Join-Path (Split-Path -Parent $OutXlsx) "задание_история.txt"
}
else {
    $historyPath = $HistoryPath
}
$history = @(Read-History $historyPath)
if ($CollapseDay -and -not $DynamicsOnly) {
    # режим «одна сборка в день»: сегодняшние прежние срезы заменяются новым
    $history = @($history | Where-Object { $_.Date -ne $today })
}
$period = Resolve-DynamicsPeriod $DynamicsFrom $DynamicsTo
# журнал — на проект: история_<код>.txt -> задание_журнал_<код>.txt рядом; прочие имена истории (задание_история.txt и т. п.) — прежний общий задание_журнал.txt
$ledgerName = "задание_журнал.txt"
if ((Split-Path -Leaf $historyPath) -match '^история_(.+)\.txt$') { $ledgerName = "задание_журнал_" + $Matches[1] + ".txt" }
$ledgerPath = Join-Path (Split-Path -Parent $historyPath) $ledgerName

$stream = $null
$partialPaths = $null
$partialChecks = $null
if (-not $DynamicsOnly) {
    if ([string]::IsNullOrWhiteSpace($ExportDir)) {
        throw "Не указана папка выгрузки Revit."
    }
    if ($Partial) {
        # буфер «последнее состояние по паре модель + спецификация»; нет файла — засев из журнала и последнего среза истории
        $partialPaths = Get-PartialStatePaths $historyPath
        $catalog = Read-ExportCatalogLines (Get-BvsExportResult $ExportDir).Catalog          # export.txt [catalog] (прежде catalog.txt)
        $partialState = Read-PartialState $partialPaths          # <история>_состояние.txt (прежние _буфер/_проверки/_каталог.txt — если ещё не перенесены)
        $partialChecks = $partialState.Checks
        $bufferRows = $partialState.Buffer
        if ($null -eq $bufferRows) {
            $seed = New-SeedBuffer (Read-IssueLedger $ledgerPath) $history
            if ($null -eq $seed) {
                throw "Для режима «Выборочно» нужна база: нет ни буфера, ни журнала, ни истории. Сначала сделайте полную выгрузку (режим «Всё»)."
            }
            $bufferRows = $seed.Rows
            foreach ($k in @($seed.Checks.Keys)) { if (-not $partialChecks.ContainsKey($k)) { $partialChecks[$k] = $seed.Checks[$k] } }
        }
        $bufferRows = @(Select-CatalogRows $bufferRows $catalog)
        $partialChecks = Select-CatalogChecks $partialChecks $catalog
        $stream = Import-ExportDirStreaming $ExportDir -Partial -BufferRows $bufferRows
    }
    else {
        $stream = Import-ExportDirStreaming $ExportDir
    }
}

# прошлый срез — последняя предыдущая сборка (не обязательно вчерашняя)
$prevBuild = ""
foreach ($b in @($history | ForEach-Object { $_.BuildId } | Select-Object -Unique | Sort-Object)) {
    if ($b -lt $buildId) { $prevBuild = $b }
}
$prev = @{}
if ($prevBuild) {
    foreach ($h in @($history)) {
        if ($h.BuildId -ne $prevBuild) { continue }
        $key = $h.Model + $Sep + $h.Check
        if (-not $prev.ContainsKey($key)) { $prev[$key] = 0 }
        $prev[$key] = [int]$prev[$key] + [int]$h.Count
    }
}

if ($stream -and $stream.Partial -and $catalog -and $catalog.Has) {
    # модель или спецификация, которых больше нет в сводной, уходят из отчёта сразу (без синего нуля «были, сейчас нет»)
    foreach ($k in @($prev.Keys)) {
        $pk = @($k.Split($Sep))
        if (($catalog.Specs.Count -gt 0 -and -not $catalog.Specs.ContainsKey($pk[1])) -or
            ($catalog.Models.Count -gt 0 -and -not $catalog.Models.ContainsKey((Get-ModelNorm $pk[0])))) { $prev.Remove($k) }
    }
}

$allHistory = @($history)
$newSlice = @()
if ($stream) {
    if ($stream.Partial) { $newSlice = @(New-CurrentSlice $stream $history $prevBuild $buildId $stream.CheckedPairs) }
    else { $newSlice = @(New-CurrentSlice $stream $history $prevBuild $buildId) }
    $allHistory = @($history) + $newSlice
}
# журнал замечаний: обновляется по выгрузке этой сборки (в режиме «только динамика» — только читается)
$ledger = Read-IssueLedger $ledgerPath
$deltas = @{}
$ledgerStats = @{ New = 0; Reopened = 0; Fixed = 0 }
if ($stream) {
    $allIssues = New-Object System.Collections.Generic.List[object]
    foreach ($sec in @(Get-StreamSections $stream $SectionOrder)) {
        foreach ($it in @(Read-SectionIssuesFromStream $stream.SectionDir $sec)) { $allIssues.Add($it) }
    }
    if ($stream.Partial) {
        # журнал видит только строки проверенных пар; пары из буфера не «замечены» и не закрываются
        $freshIssues = @($allIssues.ToArray() | Where-Object { -not $_.Buffered })
        $ledgerStats = Update-IssueLedger $ledger $freshIssues $stream.CheckedSpecs $buildId ($ledger.Count -gt 0) $stream.CheckedPairs
    }
    else {
        $ledgerStats = Update-IssueLedger $ledger $allIssues.ToArray() $stream.CheckedSpecs $buildId ($ledger.Count -gt 0)
    }
    $deltas = Get-LedgerDeltas $ledger
}
# «Динамика» считается уже с текущей сборкой
$dynItems = @(Get-DynamicsItems $allHistory $period.From $period.To $ledger)

# сумма замечаний по модели в каждом срезе (колонки Сводки)
$script:ByDateModelSum = @{}
foreach ($h in $allHistory) {
    $dm = $h.BuildId + $Sep + $h.Model
    if (-not $script:ByDateModelSum.ContainsKey($dm)) { $script:ByDateModelSum[$dm] = 0 }
    $script:ByDateModelSum[$dm] = [int]$script:ByDateModelSum[$dm] + [int]$h.Count
}

$dateCols = New-Object System.Collections.Generic.List[string]
$pool = New-Object System.Collections.Generic.List[string]
foreach ($d in @($allHistory | ForEach-Object { $_.BuildId } | Select-Object -Unique | Sort-Object)) { $pool.Add($d) }
if ($stream -and -not $pool.Contains($buildId)) { $pool.Add($buildId) }
$take = @($pool | Select-Object -Last 3)
foreach ($d in $take) { $dateCols.Add($d) }

$dir = Split-Path -Parent $OutXlsx
if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
if (Test-Path -LiteralPath $OutXlsx) { Remove-Item -LiteralPath $OutXlsx -Force }
$auxOut = $dir
if (-not [string]::IsNullOrWhiteSpace($AuxDir)) { $auxOut = $AuxDir }
if (-not (Test-Path -LiteralPath $auxOut)) { New-Item -ItemType Directory -Path $auxOut -Force | Out-Null }
# Одна копия фильтров. Выгрузка пишет её сразу как фильтры_спецификаций.txt.
# filters_debug.txt — старое имя: копируем, только если папка назначения другая (без -AuxDir файл ложился рядом с xlsx).
if ($ExportDir) {
    $named = Join-Path $ExportDir "фильтры_спецификаций.txt"
    $legacy = Join-Path $ExportDir "filters_debug.txt"
    $srcFilters = ""
    if (Test-Path -LiteralPath $named) { $srcFilters = $named }
    elseif (Test-Path -LiteralPath $legacy) { $srcFilters = $legacy }
    $destFilters = Join-Path $auxOut "фильтры_спецификаций.txt"
    if ($srcFilters) {
        $sameFilters = ([System.IO.Path]::GetFullPath($srcFilters) -eq [System.IO.Path]::GetFullPath($destFilters))
        if (-not $sameFilters) {
            try { Copy-Item -LiteralPath $srcFilters -Destination $destFilters -Force -ErrorAction Stop }
            catch { Write-Warning ("фильтры_спецификаций.txt не обновлён: " + $_.Exception.Message) }
        }
    }
}

$payloadPath = Join-Path $auxOut "google_push_payload.json"
if (Test-Path -LiteralPath $payloadPath) { Remove-Item -LiteralPath $payloadPath -Force }

$issueCount = 0
if ($stream) { $issueCount = [int]$stream.IssueCount }

. (Join-Path $__buildDir "bvs_assignment_guard.ps1")
. (Join-Path $__buildDir "assignment_xlsx.ps1")
# выборочная выгрузка: свежесть пар и разделы, которые уходят в Google (Сводка + разделы проверенных моделей)
$partialInfo = $null
$checksNow = $null
if ($stream -and $stream.Partial) {
    $checksNow = $partialChecks
    foreach ($k in @($stream.CheckedPairs.Keys)) { $checksNow[$k] = $buildId }
    $checksNow = Select-CatalogChecks $checksNow $catalog
    $touched = New-Object System.Collections.Generic.List[string]
    foreach ($k in @($stream.CheckedPairs.Keys)) {
        $model = $k.Split([char]0x1F)[0]
        $sec = ""
        $modelSec = Get-BvsLinkSection $model
        if ($modelSec) { $sec = $modelSec.ToUpper() }
        if ($sec -and -not $touched.Contains($sec)) { $touched.Add($sec) }
    }
    $partialInfo = @{ Checks = $checksNow; CheckedPairs = $stream.CheckedPairs; BuildId = $buildId; StaleDays = 7; TouchedSections = $touched.ToArray() }
}
try {
    Export-AssignmentWorkbookOpenXml -OutXlsx $OutXlsx -ExportDir $ExportDir -PayloadPath $payloadPath -Stream $stream -IssueCount $issueCount `
        -Prev $prev -DateCols @($dateCols) -Today $today -DynItems $dynItems -Period $period `
        -SectionOrder $SectionOrder -SectionTitle $SectionTitle -Sep $Sep -ByDateModelSum $script:ByDateModelSum -DynamicsOnly:$DynamicsOnly `
        -Ledger $ledger -Deltas $deltas -BuildId $buildId -KeepFixedDays $KeepFixedDays -PartialInfo $partialInfo
    Assert-BvsAssignmentXlsx $OutXlsx
}
catch {
    if (Test-Path -LiteralPath $OutXlsx) { Remove-Item -LiteralPath $OutXlsx -Force -ErrorAction SilentlyContinue }
    if (Test-Path -LiteralPath $payloadPath) { Remove-Item -LiteralPath $payloadPath -Force -ErrorAction SilentlyContinue }
    throw
}

if (-not $DynamicsOnly -and $stream) {
    Write-HistoryRows $historyPath (Compress-OldHistory $allHistory $today)
    Write-IssueLedger $ledgerPath $ledger $today
    # состояние для выборочной выгрузки (буфер, даты проверок, каталог): только добавляет файлы рядом с историей и не может сорвать сборку
    if ($script:PartialModuleLoaded) {
        try {
            $paths = Get-PartialStatePaths $historyPath
            if ($stream.Partial) {
                $checks = $checksNow
            }
            else {
                $checks = @{}
                foreach ($sp in @($stream.CheckedSpecs.Keys)) { $checks[("*" + [char]0x1F + $sp)] = $buildId }
            }
            Save-PartialState $paths $allIssues.ToArray() $checks $ExportDir
        }
        catch { Write-Warning ("Не удалось записать состояние выборочной выгрузки: " + $_.Exception.Message) }
    }
    if ($stream.SectionDir -and (Test-Path -LiteralPath $stream.SectionDir)) {
        Remove-Item -LiteralPath $stream.SectionDir -Recurse -Force -ErrorAction SilentlyContinue
    }
}

if (-not $DynamicsOnly) {
    if (-not (Test-Path -LiteralPath $payloadPath)) {
        throw "Не создан google_push_payload.json рядом с заданием."
    }
}

Write-Output ("ISSUES=" + $issueCount + " FIXED=" + $ledgerStats.Fixed + " NEW=" + $ledgerStats.New + " BUILDER=OPENXML DYNAMICS=" + $dynItems.Count + " FILE=" + $OutXlsx + " PAYLOAD=" + $payloadPath)
if ($stream -and $stream.Partial) {
    Write-Output ("PARTIAL=1 PAIRS=" + $stream.CheckedPairs.Count + " BUFFERED=" + @($allIssues | Where-Object { $_.Buffered }).Count)
}
