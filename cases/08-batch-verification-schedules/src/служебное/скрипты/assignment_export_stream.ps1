# UTF-8 WITH BOM required (PowerShell 5.1).
# Потоковый разбор выгрузки Revit (export.txt: [manifest] + tsv) без списка $issues на все разделы.

# Каноническое имя спецификации: старая опечатка «Без значния» → «Без значения». Применяется ко всем именам, которые читаются
# из файлов прошлых сборок (история, журнал, буфер, проверки, каталог) и из manifest, чтобы ключи (модель + спецификация + параметр)
# после переименования спецификаций в моделях не менялись и история не терялась. При следующей записи файлы получают новое имя.
# Префикс проекта в именах связей и моделей: «<префикс>_IOS_<РАЗДЕЛ>_…» (например PROJ_IOS_OV_ORG_K01_R24). Задаётся из настройки проекта
# (build_assignment.ps1 -ProjectPrefix); не задан — подходит любой префикс из букв и цифр. Для PROJ результат тот же, что при зашитом «PROJ».
$script:BvsProjectPrefix = ""
function Get-BvsPrefixPattern {
    if ([string]::IsNullOrWhiteSpace($script:BvsProjectPrefix)) { return '[A-Za-z0-9]+' }
    return [regex]::Escape($script:BvsProjectPrefix.Trim())
}
# Код раздела из имени связи/модели («…_IOS_OV_…» -> «OV»); нет — $null
# Запасной вариант без «_IOS_» (SampleSite: SITE_ORG_AVT_K1_R22, SITE_ORG_HC_K1_R22): известный код раздела в виде «_КОД_», ближайший к началу имени.
# Коды — латинские из разделы.txt (порядок столбцов; на выбор ближайшего не влияет).
try {
    if (-not (Get-Command Get-BvsColumnCodes -ErrorAction SilentlyContinue)) {
        $__bvsSections = if ($PSScriptRoot) { Join-Path $PSScriptRoot "bvs_sections.ps1" } else { Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "bvs_sections.ps1" }
        if (Test-Path -LiteralPath $__bvsSections) { . $__bvsSections }
    }
}
catch { Write-Warning ("разделы: " + $_.Exception.Message) }
if (Get-Command Get-BvsColumnCodes -ErrorAction SilentlyContinue) {
    $script:BvsKnownSectionCodes = @(Get-BvsColumnCodes)
}
else {
    Write-Warning "встроенные коды разделов связи (bvs_sections.ps1 не подключён)"
    $script:BvsKnownSectionCodes = @("OV", "PD", "VK", "PT", "OT", "HS", "ITP", "EOM", "SS", "AVT", "HC")
}
function Get-BvsLinkSection([string]$Text) {
    if ($Text -match ((Get-BvsPrefixPattern) + '_IOS_([A-Z]+)_')) { return $Matches[1] }
    # префикс проекта задан — запасной вариант только для моделей этого проекта (чужой префикс разделов не даёт)
    if (-not [string]::IsNullOrWhiteSpace($script:BvsProjectPrefix) -and ([string]$Text -notmatch ('(?i)(^|[\\/])' + (Get-BvsPrefixPattern) + '_'))) { return $null }
    $up = ([string]$Text).ToUpperInvariant()
    $best = $null; $bestPos = [int]::MaxValue
    foreach ($c in $script:BvsKnownSectionCodes) {
        $pos = $up.IndexOf("_" + $c + "_", [System.StringComparison]::Ordinal)
        if ($pos -ge 0 -and $pos -lt $bestPos) { $best = $c; $bestPos = $pos }
    }
    return $best
}
# Имя начинается с «<префикс>_IOS_» (так называются семейства/типы, подставленные из имени файла связи) — такое значение не показываем
function Test-BvsIosPrefixedName([string]$Text) {
    return ([string]$Text -match ('^' + (Get-BvsPrefixPattern) + '_IOS_'))
}

try {
    $__bvsCommon = if ($PSScriptRoot) { Join-Path $PSScriptRoot "bvs_common.ps1" } else { Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "bvs_common.ps1" }
    if ((Test-Path -LiteralPath $__bvsCommon) -and -not (Get-Command ConvertTo-CanonicalSpecName -ErrorAction SilentlyContinue)) { . $__bvsCommon }
}
catch { Write-Warning ("bvs_common.ps1: " + $_.Exception.Message) }
try {
    # итог выгрузки: export.txt (или прежние manifest.txt / export_result.txt / catalog.txt)
    $__bvsExport = if ($PSScriptRoot) { Join-Path $PSScriptRoot "bvs_export_result.ps1" } else { Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "bvs_export_result.ps1" }
    if (-not (Get-Command Get-BvsExportResult -ErrorAction SilentlyContinue)) { . $__bvsExport }
}
catch { Write-Warning ("bvs_export_result.ps1: " + $_.Exception.Message) }
if (-not (Get-Command ConvertTo-CanonicalSpecName -ErrorAction SilentlyContinue)) {
    function ConvertTo-CanonicalSpecName([string]$Name) {
        return ([string]$Name).Replace("значния", "значения").Replace(([string][char]0x43 + "антехприборы"), ([string][char]0x421 + "антехприборы")).Replace(([string][char]0x63 + "антехприборы"), ([string][char]0x441 + "антехприборы"))
    }
}

function New-ExportStreamState {
    return [pscustomobject]@{
        IssueCount     = 0
        TodayCount     = @{}
        HistoryBuckets = @{}
        Models         = @{}
        CheckOrder     = New-Object System.Collections.Generic.List[string]
        CheckSeen      = @{}
        SectionWriters = @{}
        SectionDir     = ""
        # все проверенные спецификации (в т.ч. пустые): без них нельзя отличить «исправлено» от «не проверялось»
        CheckedSpecs   = @{}
        # выборочная выгрузка (manifest: *scope* partial): проверенные пары «Модель<0x1F>Спецификация»; в режиме «Всё» не заполняется
        Partial        = $false
        CheckedPairs   = @{}
    }
}

function Close-ExportStreamState($State) {
    foreach ($k in @($State.SectionWriters.Keys)) {
        $w = $State.SectionWriters[$k]
        if ($w) { $w.Flush(); $w.Close() }
    }
    $State.SectionWriters = @{}
}

function Get-SectionWriter($State, [string]$Section) {
    if ($State.SectionWriters.ContainsKey($Section)) {
        return $State.SectionWriters[$Section]
    }
    if (-not (Test-Path -LiteralPath $State.SectionDir)) {
        New-Item -ItemType Directory -Path $State.SectionDir -Force | Out-Null
    }
    $path = Join-Path $State.SectionDir ("{0}.jsonl" -f $Section)
    $enc = New-Object System.Text.UTF8Encoding $false
    $w = New-Object System.IO.StreamWriter($path, $false, $enc)
    $State.SectionWriters[$Section] = $w
    return $w
}

function Register-StreamIssue($State, $Issue, [switch]$Buffered) {
    $State.IssueCount++
    $State.Models[$Issue.Model] = $true
    $ck = $Issue.Model + [char]0x1F + $Issue.Check
    if (-not $State.TodayCount.ContainsKey($ck)) { $State.TodayCount[$ck] = 0 }
    $State.TodayCount[$ck] = [int]$State.TodayCount[$ck] + 1
    if (-not $State.CheckSeen.ContainsKey($Issue.Check)) {
        $State.CheckSeen[$Issue.Check] = $true
        $State.CheckOrder.Add($Issue.Check)
    }
    $hk = $Issue.Model + "`t" + $Issue.Check + "`t" + (To-Text $Issue.Param)
    if (-not $State.HistoryBuckets.ContainsKey($hk)) {
        $State.HistoryBuckets[$hk] = [pscustomobject]@{
            Section = $Issue.Section
            Model   = $Issue.Model
            Check   = $Issue.Check
            Param   = (To-Text $Issue.Param)
            ParamName = (To-Text $Issue.ParamName)
            Count   = 0
        }
    }
    $State.HistoryBuckets[$hk].Count = [int]$State.HistoryBuckets[$hk].Count + 1
    $json = @{
        Section = $Issue.Section
        Model   = $Issue.Model
        Check   = $Issue.Check
        Param   = (To-Text $Issue.Param)
        ParamName = (To-Text $Issue.ParamName)
        Family  = (To-Text $Issue.Family)
        Item    = (To-Text $Issue.Item)
        Value   = (To-Text $Issue.Value)
    }
    # строка взята из буфера (пара не проверялась в этом прогоне): помечаем только в выборочном режиме, прежний JSON не меняется
    if ($Buffered) { $json.Buf = 1 }
    $json = $json | ConvertTo-Json -Compress
    $sw = Get-SectionWriter $State $Issue.Section
    $sw.WriteLine($json)
}

# $fieldNames — имена полей по колонкам (manifest, 3-й столбец), $filterNames — параметры из фильтра спецификации (4-й).
# Колонки определяем и по заголовку, и по имени поля: заголовок в Revit можно переименовать
# («Проверить, чтобы…» вместо «Семейство и типоразмер»), имя поля при этом остаётся.
function Parse-ExportTsvIssue($headers, $cells, [string]$scheduleName, $fieldNames = $null, $filterNames = $null) {
    $fields = @()
    if ($null -ne $fieldNames) { $fields = @($fieldNames) }
    # «0_Ошибка типоразмера»: проверяемый параметр ровно ADSK_Наименование. Колонка с этим именем иначе
    # съедается правилом «Наименование» (подстрока) и не попадает в «Значение». «ADSK_Материал наименование» — другое имя, его не трогаем.
    $literalValue = @{}
    if ($null -ne $filterNames) {
        foreach ($fn in @($filterNames)) {
            $t = ([string]$fn).Trim()
            if ($t -eq "ADSK_Наименование") { $literalValue[$t] = $true }
        }
    }
    $linkCol = -1
    $familyCol = -1
    $nameCol = -1
    $oldParamCols = New-Object System.Collections.Generic.List[int]
    $paramCols = New-Object System.Collections.Generic.List[int]
    for ($c = 0; $c -lt $headers.Count; $c++) {
        $h = $headers[$c].Trim()
        $f = ""
        if ($c -lt $fields.Count) { $f = ([string]$fields[$c]).Trim() }
        $isLiteral = (($f -and $literalValue.ContainsKey($f)) -or ($h -and $literalValue.ContainsKey($h)))
        # --- прежние правила (только по заголовку): нужны для ключа истории Param ---
        if ($h -match 'Связанный файл') { $linkCol = $c }
        elseif ($h -eq "Семейство и типоразмер") { if ($familyCol -lt 0) { $familyCol = $c } }
        elseif (-not $isLiteral -and ($h -match 'Наименование')) { if ($nameCol -lt 0) { $nameCol = $c } }
        elseif (-not ($h -match 'Количество' -or $h -eq "Категория" -or $h -eq "")) { $oldParamCols.Add($c) }
        # --- новые правила (заголовок ИЛИ имя поля): что показываем ---
        if ($h -match 'Связанный файл') { continue }
        if ($h -eq "Семейство и типоразмер" -or $f -eq "Семейство и типоразмер") { if ($familyCol -lt 0 -or $h -eq "Семейство и типоразмер") { $familyCol = $c }; continue }
        if (-not $isLiteral -and ($h -match 'Наименование' -or $f -match 'Наименование')) { if ($nameCol -lt 0 -or $h -match 'Наименование') { $nameCol = $c }; continue }
        if (-not $isLiteral -and ($h -match 'Количество' -or $f -match 'Количество' -or $h -eq "Категория" -or $h -eq "")) { continue }
        $paramCols.Add($c)
    }
    if ($linkCol -lt 0) { $linkCol = $headers.Count - 1 }
    $link = ""
    if ($linkCol -lt $cells.Count) { $link = To-Text $cells[$linkCol] }
    $section = Get-BvsLinkSection $link
    if (-not $section) { return $null }
    $valueParts = New-Object System.Collections.Generic.List[string]
    foreach ($pc in @($paramCols)) {
        if ($pc -lt $cells.Count) {
            $pv = To-Text $cells[$pc]
            if (-not [string]::IsNullOrWhiteSpace($pv)) { $valueParts.Add($pv) }
        }
    }
    $family = ""
    $item = ""
    if ($familyCol -ge 0 -and $familyCol -lt $cells.Count) { $family = To-Text $cells[$familyCol] }
    if ($nameCol -ge 0 -and $nameCol -lt $cells.Count) { $item = To-Text $cells[$nameCol] }
    if (Test-BvsIosPrefixedName $family) { $family = "" }
    if (Test-BvsIosPrefixedName $item) { $item = "" }
    # ключ истории — как раньше: заголовки колонок по прежним правилам
    $paramName = (@($oldParamCols | ForEach-Object { $headers[$_].Trim() }) -join ", ")
    # показываем: параметр из фильтра спецификации; иначе имена полей/заголовки колонок-параметров
    $shownName = ""
    if ($null -ne $filterNames) { $shownName = (@($filterNames | Where-Object { $_ -and ([string]$_).Trim() } | ForEach-Object { ([string]$_).Trim() }) -join ", ") }
    if (-not $shownName) {
        $shownName = (@($paramCols | ForEach-Object {
            $fn = ""
            if ($_ -lt $fields.Count) { $fn = ([string]$fields[$_]).Trim() }
            if ($fn) { $fn } else { $headers[$_].Trim() }
        }) -join ", ")
    }
    if (-not $shownName) { $shownName = $paramName }
    return [pscustomobject]@{
        ParamName = $shownName
        Section = $section
        Model   = (Get-ModelFileName $link)
        Check   = $scheduleName
        Param   = $paramName
        Family  = $family
        Item    = $item
        Value   = ($valueParts -join " · ")
    }
}

function Import-ExportDirStreaming {
    param([string]$Dir, [switch]$Partial, $BufferRows = $null)
    $export = Get-BvsExportResult $Dir
    if (-not $export.Has) {
        throw "Нет итога выгрузки (export.txt). Выгрузка из сводной не записала спецификации."
    }
    $state = New-ExportStreamState
    $state.SectionDir = Join-Path $Dir "_sections"
    if (Test-Path -LiteralPath $state.SectionDir) {
        Remove-Item -LiteralPath $state.SectionDir -Recurse -Force -ErrorAction SilentlyContinue
    }
    foreach ($line in $export.Manifest) {
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        $tab = $line.IndexOf("`t")
        if ($tab -lt 1) { continue }
        $fileName = $line.Substring(0, $tab).Trim()
        if ($fileName -eq "*scope*") { $state.Partial = $true; continue }
        if ($fileName -eq "*pair*") {
            $pp = @($line.Split([char]9))
            if ($pp.Count -ge 3) { $state.CheckedPairs[($pp[1].Trim() + [char]0x1F + (ConvertTo-CanonicalSpecName $pp[2].Trim()))] = $true }
            continue
        }
        $scheduleName = $line.Substring($tab + 1).Trim()
        $fieldNames = $null
        $filterNames = $null
        $mparts = @($line.Split([char]9))
        if ($mparts.Count -ge 3) {
            $scheduleName = $mparts[1].Trim()
            $fieldNames = @($mparts[2].Split([char]'|'))
        }
        if ($mparts.Count -ge 4 -and $mparts[3].Trim()) {
            $filterNames = @($mparts[3].Trim().Split([char]'|'))
        }
        $scheduleName = ConvertTo-CanonicalSpecName $scheduleName
        $state.CheckedSpecs[$scheduleName] = $true
        if ($fileName -eq "*empty*") { continue }
        $path = Join-Path $Dir $fileName
        if (-not (Test-Path -LiteralPath $path)) { continue }
        $reader = New-Object System.IO.StreamReader($path, [System.Text.Encoding]::UTF8, $true)
        $headerLine = $reader.ReadLine()
        if ([string]::IsNullOrWhiteSpace($headerLine)) {
            $reader.Close()
            continue
        }
        $headers = @($headerLine.Split([char]9))
        while ($null -ne ($dataLine = $reader.ReadLine())) {
            if ([string]::IsNullOrWhiteSpace($dataLine)) { continue }
            $cells = @($dataLine.Split([char]9))
            $issue = Parse-ExportTsvIssue $headers $cells $scheduleName $fieldNames $filterNames
            if ($null -eq $issue) { continue }
            Register-StreamIssue $state $issue
        }
        $reader.Close()
    }
    if ($state.Partial -and -not $Partial) {
        Close-ExportStreamState $state
        throw "Выгрузка выборочная (assign_scope=partial), а сборка запущена в режиме «Всё». Сделайте полную выгрузку или выберите режим «Выборочно»."
    }
    if ($Partial -and -not $state.Partial) {
        Close-ExportStreamState $state
        throw "Режим «Выборочно», но выгрузка полная (в итоге выгрузки нет *scope*). Повторите выгрузку."
    }
    if ($state.Partial) {
        # виртуально полная выгрузка: к проверенным в этом прогоне парам добавляем прежнее состояние остальных пар из буфера
        foreach ($r in @($BufferRows)) {
            if ($state.CheckedPairs.ContainsKey(([string]$r.Model + [char]0x1F + [string]$r.Check))) { continue }
            Register-StreamIssue $state $r -Buffered
            $state.CheckedSpecs[[string]$r.Check] = $true
        }
    }
    Close-ExportStreamState $state
    return $state
}

function Read-SectionIssuesFromStream {
    param([string]$SectionDir, [string]$SectionCode)
    $path = Join-Path $SectionDir ("{0}.jsonl" -f $SectionCode)
    $list = New-Object System.Collections.Generic.List[object]
    if (-not (Test-Path -LiteralPath $path)) { return @() }
    foreach ($line in [System.IO.File]::ReadLines($path, [System.Text.UTF8Encoding]::new($false))) {
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        $o = $line | ConvertFrom-Json
        $list.Add([pscustomobject]@{
            Section = [string]$o.Section
            Model   = [string]$o.Model
            Check   = [string]$o.Check
            Param   = [string]$o.Param
            ParamName = $(if ($o.ParamName) { [string]$o.ParamName } else { [string]$o.Param })
            Family  = [string]$o.Family
            Item    = [string]$o.Item
            Value   = [string]$o.Value
            Buffered = [bool]$o.Buf
        })
    }
    return $list.ToArray()
}

function Get-StreamSections($State, $SectionOrder) {
    $out = New-Object System.Collections.Generic.List[string]
    foreach ($code in $SectionOrder) {
        $path = Join-Path $State.SectionDir ("{0}.jsonl" -f $code)
        if (Test-Path -LiteralPath $path) { $out.Add($code) }
    }
    return @($out)
}

function Write-HistorySnapshotFromBuckets {
    param([string]$Path, $HistoryRows, [string]$Today, $HistoryBuckets)
    $past = @($HistoryRows | Where-Object { $_.Date -ne $Today })
    $keep = New-Object System.Collections.Generic.List[string]
    foreach ($h in $past) {
        $keep.Add(("{0}`t{1}`t{2}`t{3}`t{4}`t{5}`t{6}" -f $h.Date, $h.Section, $h.Model, $h.Check, $h.Param, $h.Count, $h.ParamName))
    }
    foreach ($b in @($HistoryBuckets.Values)) {
        if ([int]$b.Count -lt 1) { continue }
        $keep.Add(("{0}`t{1}`t{2}`t{3}`t{4}`t{5}`t{6}" -f $Today, $b.Section, $b.Model, $b.Check, $b.Param, $b.Count, $b.ParamName))
    }
    $bom = New-Object System.Text.UTF8Encoding $true
    $dir = Split-Path -Parent $Path
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    [System.IO.File]::WriteAllLines($Path, @($keep), $bom)
}

# ---------- порядок спецификаций ----------
# Номер в начале имени («П   14.7_Длина_…», «16.0» раньше «16.1») сравниваем как версию: 0 < 1 < 2.1 < 4.4 < 5.7 < 13 < 14.7 < 16.0 < 16.1 < 99.
# При равном номере спецификации одной проверки (одинаковое имя без суффикса) стоят рядом:
# «Пустые», затем «999» (1_Классификатор_999 сразу после 1_Классификатор_Пустые), затем «Без значения».
# Суффикс «Без значения» в именах встречается и с опечаткой («Без значния»), поэтому ищем «Без знач…».
function Get-CheckNameParts([string]$Name) {
    $t = ([string]$Name).Trim() -replace '^(?i)(РД|ПП|П)(?=\s|\d)\s*', ''
    $kind = ""
    $base = $t
    if ($t -match '^(?<b>.*?)[_\s]+(?<s>Пустые|Без\s+знач\w*|999)\s*$') {
        $base = $Matches['b']
        if ($Matches['s'] -match '^(?i)Пустые') { $kind = "empty" }
        elseif ($Matches['s'] -match '^999') { $kind = "code999" }
        else { $kind = "novalue" }
    }
    $base = (($base -replace '\s+', ' ').Trim()).ToLowerInvariant()
    return @{ Base = $base; Kind = $kind }
}

function Get-CheckSortKey([string]$Name) {
    $t = ([string]$Name).Trim() -replace '^(?i)(РД|ПП|П)(?=\s|\d)\s*', ''
    $prefix = "~"
    if ($t -match '^(\d+(?:\.\d+)*)(?=[_\s]|$)') {
        $prefix = (@($Matches[1].Split('.') | ForEach-Object { $_.PadLeft(6, '0') }) -join '.')
    }
    $parts = Get-CheckNameParts $Name
    $rank = "0"
    if ($parts.Kind -eq "empty") { $rank = "1" }
    elseif ($parts.Kind -eq "code999") { $rank = "2" }
    elseif ($parts.Kind -eq "novalue") { $rank = "3" }
    return ($prefix + "|" + $parts.Base + "|" + $rank + "|" + ([string]$Name).ToLowerInvariant())
}

function Sort-CheckNames($Names) {
    $list = New-Object 'System.Collections.Generic.List[string]'
    foreach ($n in @($Names)) { if (-not [string]::IsNullOrWhiteSpace([string]$n)) { $list.Add([string]$n) } }
    $keys = @{}
    foreach ($n in $list) { $keys[$n] = Get-CheckSortKey $n }
    $cmp = [System.Comparison[string]] { param($a, $b) [string]::CompareOrdinal($keys[$a], $keys[$b]) }
    $list.Sort($cmp)
    return $list.ToArray()
}

function Get-ChecksWithRemarksFromStream($State, $Prev, $Sep) {
    $active = @{}
    foreach ($ck in @($State.TodayCount.Keys)) {
        $pair = @($ck -split $Sep, 2)
        if ($pair.Count -ge 2 -and -not [string]::IsNullOrWhiteSpace($pair[1])) {
            $active[$pair[1]] = $true
        }
    }
    $checked = $State.CheckedSpecs
    foreach ($k in @($Prev.Keys)) {
        if ([int]$Prev[$k] -lt 1) { continue }
        $pair = @($k -split $Sep, 2)
        if ($pair.Count -ge 2 -and -not [string]::IsNullOrWhiteSpace($pair[1])) {
            # спецификация, которой не было в этой выгрузке (убрана из списка), «не проверялась»: колонку не показываем
            if ($checked -and $checked.Count -gt 0 -and -not $checked.ContainsKey($pair[1])) { continue }
            $active[$pair[1]] = $true
        }
    }
    $order = New-Object System.Collections.Generic.List[string]
    $seen = @{}
    foreach ($name in @($State.CheckOrder)) {
        if ([string]::IsNullOrWhiteSpace($name)) { continue }
        if (-not $active.ContainsKey($name)) { continue }
        if (-not $seen.ContainsKey($name)) {
            $seen[$name] = $true
            $order.Add($name)
        }
    }
    foreach ($name in @($active.Keys | Sort-Object)) {
        if (-not $seen.ContainsKey($name)) { $order.Add($name) }
    }
    return @(Sort-CheckNames $order)
}
