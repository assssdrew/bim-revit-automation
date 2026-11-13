# UTF-8 WITH BOM required (PowerShell 5.1).
# Выборочная выгрузка: буфер «последнее состояние по паре модель + спецификация».
# Подключается сборкой только при -Partial; режим «Всё» эти функции не вызывает (кроме записи файлов состояния после сборки).
#
# Состояние лежит рядом с историей одним файлом <история>_состояние.txt (<история> — имя файла истории без расширения), UTF-8 с BOM:
#   [buffer]  — замечания по всем парам, как их видел последний прогон (TAB): Section, Model, Check, Param, ParamName, Family, Item, Value
#   [checks]  — когда пара проверялась в последний раз (TAB): Model, Check, BuildId; Model = * — проверена спецификация целиком (режим «Всё»)
#   [catalog] — каталог последней выгрузки (export.txt [catalog]): spec<TAB>имя / model<TAB>имя<TAB>loaded|unloaded (из них UI строит списки)
# Строки разделов — ровно строки прежних файлов <история>_буфер.txt, _проверки.txt, _каталог.txt (до фазы 7). Прежние файлы читаются,
# если их ещё не перенесли (Invoke-PartialStateMigration при открытии окна: резервная копия, затем переименование в .migrated_<дата>).

# Имя спецификации (см. assignment_export_stream.ps1); запасное определение, если файл подключён отдельно
if (-not (Get-Command ConvertTo-CanonicalSpecName -ErrorAction SilentlyContinue)) {
    function ConvertTo-CanonicalSpecName([string]$Name) { return ([string]$Name).Replace("значния", "значения").Replace(([string][char]0x43 + "антехприборы"), ([string][char]0x421 + "антехприборы")).Replace(([string][char]0x63 + "антехприборы"), ([string][char]0x441 + "антехприборы")) }
}

$script:PartialSep = [char]0x1F

function Get-PartialStatePaths([string]$HistoryPath) {
    $dir = Split-Path -Parent $HistoryPath
    $stem = [System.IO.Path]::GetFileNameWithoutExtension($HistoryPath)
    return @{
        State   = (Join-Path $dir ($stem + "_состояние.txt"))
        Buffer  = (Join-Path $dir ($stem + "_буфер.txt"))          # прежние файлы (до фазы 7): только чтение и перенос
        Checks  = (Join-Path $dir ($stem + "_проверки.txt"))
        Catalog = (Join-Path $dir ($stem + "_каталог.txt"))
    }
}

$script:PartialStateParts = @("buffer", "checks", "catalog")
function Get-PartialOldStatePath($Paths, [string]$Part) {
    switch ($Part) { "buffer" { return $Paths.Buffer } "checks" { return $Paths.Checks } "catalog" { return $Paths.Catalog } }
    return ""
}

# Строки состояния по частям: @{ buffer|checks|catalog = string[] | $null (части нет) }.
# Часть берётся из <история>_состояние.txt; если её там нет или прежний файл новее (после отката на прежнюю версию) — из прежнего файла.
function Read-PartialStateLines($Paths) {
    $res = @{ buffer = $null; checks = $null; catalog = $null }
    $enc = New-Object System.Text.UTF8Encoding $true
    $stateTime = $null
    if ($Paths.State -and (Test-Path -LiteralPath $Paths.State)) {
        $stateTime = (Get-Item -LiteralPath $Paths.State).LastWriteTimeUtc
        $cur = ""
        foreach ($line in [System.IO.File]::ReadAllLines($Paths.State, $enc)) {
            $t = $line.Trim()
            if ($t -eq "[buffer]" -or $t -eq "[checks]" -or $t -eq "[catalog]") {
                $cur = $t.Substring(1, $t.Length - 2)
                if ($null -eq $res[$cur]) { $res[$cur] = New-Object System.Collections.Generic.List[string] }
                continue
            }
            if (-not $cur -or [string]::IsNullOrWhiteSpace($line)) { continue }          # до первого раздела — заголовок-комментарий
            $res[$cur].Add($line)
        }
    }
    foreach ($part in $script:PartialStateParts) {
        $old = Get-PartialOldStatePath $Paths $part
        if (-not $old -or -not (Test-Path -LiteralPath $old)) { continue }
        if ($null -eq $res[$part] -or ($null -ne $stateTime -and (Get-Item -LiteralPath $old).LastWriteTimeUtc -gt $stateTime)) {
            $res[$part] = New-Object System.Collections.Generic.List[string]
            foreach ($line in [System.IO.File]::ReadAllLines($old, $enc)) { if (-not [string]::IsNullOrWhiteSpace($line)) { $res[$part].Add($line) } }
        }
    }
    return $res
}

# Запись <история>_состояние.txt: сначала во временный файл рядом, затем подмена — недописанный файл не остаётся на месте прежнего
function Write-PartialStateFile([string]$Path, $Parts) {
    $L = New-Object System.Collections.Generic.List[string]
    $L.Add("# Состояние «Выборочно»: буфер замечаний, даты проверок, каталог последней выгрузки. Служебный файл сборки задания, руками не править.")
    foreach ($part in $script:PartialStateParts) {
        if ($null -eq $Parts[$part]) { continue }
        $L.Add("[" + $part + "]")
        foreach ($line in @($Parts[$part])) { if (-not [string]::IsNullOrWhiteSpace([string]$line)) { $L.Add([string]$line) } }
    }
    $tmp = $Path + ".tmp"
    [System.IO.File]::WriteAllText($tmp, (($L -join "`r`n") + "`r`n"), (New-Object System.Text.UTF8Encoding $true))
    if (Test-Path -LiteralPath $Path) {
        try { [System.IO.File]::Replace($tmp, $Path, $null) }
        catch { [System.IO.File]::Copy($tmp, $Path, $true); [System.IO.File]::Delete($tmp) }          # папка, где подмена не поддерживается
    }
    else { [System.IO.File]::Move($tmp, $Path) }
}

# Всё состояние разом: Buffer (строки | $null — нет буфера или он пуст, как прежде), Checks (карта), Catalog (как Read-ExportCatalog), HasBuffer
function Read-PartialState($Paths) {
    $lines = Read-PartialStateLines $Paths
    $buffer = $null
    if ($null -ne $lines.buffer) {
        $rows = @(ConvertFrom-PartialBufferLines $lines.buffer)
        if ($rows.Count -gt 0) { $buffer = $rows }
    }
    return [pscustomobject]@{
        Buffer    = $buffer
        HasBuffer = ($null -ne $lines.buffer)
        Checks    = (ConvertFrom-PartialChecksLines $lines.checks)
        Catalog   = (Read-ExportCatalogLines @($lines.catalog))
        Lines     = $lines
    }
}

# Имя модели для сравнения: без пути, обрезано по «.rvt», без регистра (как в export_schedules.py)
function Get-ModelNorm([string]$Name) {
    $t = ([string]$Name).Trim()
    if ($t -match '[^\\/]+$') { $t = $Matches[0] }
    $i = $t.ToLowerInvariant().IndexOf(".rvt")
    if ($i -ge 0) { $t = $t.Substring(0, $i) }
    return $t.Trim().ToLowerInvariant()
}

# Прежний файл буфера (до фазы 7); сборка и окно читают через Read-PartialState
function Read-PartialBuffer([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path)) { return $null }
    return (ConvertFrom-PartialBufferLines @(Get-Content -LiteralPath $Path -Encoding UTF8))
}

function ConvertFrom-PartialBufferLines($Lines) {
    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($line in @($Lines)) {
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        $p = @($line.Split([char]9))
        if ($p.Count -lt 8) { continue }
        $rows.Add([pscustomobject]@{
            Section = $p[0]; Model = $p[1]; Check = (ConvertTo-CanonicalSpecName $p[2]); Param = $p[3]; ParamName = $p[4]
            Family = $p[5]; Item = $p[6]; Value = $p[7]
        })
    }
    return $rows.ToArray()
}

function ConvertTo-PartialBufferLines($Rows) {
    $lines = New-Object System.Collections.Generic.List[string]
    foreach ($r in @($Rows)) {
        if ($null -eq $r) { continue }
        $lines.Add(("{0}`t{1}`t{2}`t{3}`t{4}`t{5}`t{6}`t{7}" -f $r.Section, $r.Model, $r.Check, $r.Param, $r.ParamName, $r.Family, $r.Item, $r.Value))
    }
    return , $lines.ToArray()
}

# {ключ "Модель<0x1F>Спецификация" -> BuildId}; Model = "*" — вся спецификация
# Прежний файл дат проверок (до фазы 7); сборка и окно читают через Read-PartialState
function Read-PartialChecks([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path)) { return @{} }
    return (ConvertFrom-PartialChecksLines @(Get-Content -LiteralPath $Path -Encoding UTF8))
}

function ConvertFrom-PartialChecksLines($Lines) {
    $map = @{}
    foreach ($line in @($Lines)) {
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        $p = @($line.Split([char]9))
        if ($p.Count -lt 3) { continue }
        $map[($p[0] + $script:PartialSep + (ConvertTo-CanonicalSpecName $p[1]))] = $p[2]
    }
    return $map
}

function ConvertTo-PartialChecksLines($Map) {
    $lines = New-Object System.Collections.Generic.List[string]
    foreach ($k in @($Map.Keys | Sort-Object)) {
        $pair = @($k.Split($script:PartialSep))
        $lines.Add(("{0}`t{1}`t{2}" -f $pair[0], $pair[1], $Map[$k]))
    }
    return , $lines.ToArray()
}

# Свежесть пары: самая поздняя из записи пары и записи «вся спецификация»; "" — не проверялась
function Get-PairChecked($Checks, [string]$Model, [string]$Check) {
    $best = ""
    foreach ($k in @(($Model + $script:PartialSep + $Check), ("*" + $script:PartialSep + $Check))) {
        if ($Checks.ContainsKey($k) -and [string]$Checks[$k] -gt $best) { $best = [string]$Checks[$k] }
    }
    return $best
}

# файл каталога (<история>_каталог.txt, прежний catalog.txt) -> @{ Specs = hashtable; Models = hashtable(norm -> loaded|unloaded); Has = есть ли данные }
function Read-ExportCatalog([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path)) { return (Read-ExportCatalogLines @()) }
    return (Read-ExportCatalogLines @(Get-Content -LiteralPath $Path -Encoding UTF8))
}

# Строки каталога (catalog.txt / <история>_каталог.txt / export.txt [catalog]) -> то же, что Read-ExportCatalog
function Read-ExportCatalogLines($Lines) {
    $cat = @{ Specs = @{}; Models = @{}; ModelNames = @{}; Has = $false }
    foreach ($line in @($Lines)) {
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        $p = @($line.Split([char]9))
        if ($p.Count -ge 2 -and $p[0] -eq "spec") { $cat.Specs[(ConvertTo-CanonicalSpecName $p[1].Trim())] = $true; $cat.Has = $true }
        elseif ($p.Count -ge 2 -and $p[0] -eq "model") {
            $st = "loaded"
            if ($p.Count -ge 3) { $st = $p[2].Trim() }
            $n = Get-ModelNorm $p[1]
            $cat.Models[$n] = $st
            $cat.ModelNames[$n] = $p[1].Trim()
            $cat.Has = $true
        }
    }
    return $cat
}

# Убирает из буфера строки спецификаций и моделей, которых больше нет в сводной (переименованы/удалены).
# Без каталога (старый выгрузчик) или с пустой частью каталога — эту часть не трогаем.
function Select-CatalogRows($Rows, $Catalog) {
    if (-not $Catalog -or -not $Catalog.Has) { return @($Rows) }
    $out = New-Object System.Collections.Generic.List[object]
    foreach ($r in @($Rows)) {
        if ($Catalog.Specs.Count -gt 0 -and -not $Catalog.Specs.ContainsKey($r.Check)) { continue }
        if ($Catalog.Models.Count -gt 0 -and -not $Catalog.Models.ContainsKey((Get-ModelNorm $r.Model))) { continue }
        $out.Add($r)
    }
    return $out.ToArray()
}

function Select-CatalogChecks($Checks, $Catalog) {
    if (-not $Catalog -or -not $Catalog.Has) { return $Checks }
    $out = @{}
    foreach ($k in @($Checks.Keys)) {
        $pair = @($k.Split($script:PartialSep))
        if ($Catalog.Specs.Count -gt 0 -and -not $Catalog.Specs.ContainsKey($pair[1])) { continue }
        if ($pair[0] -ne "*" -and $Catalog.Models.Count -gt 0 -and -not $Catalog.Models.ContainsKey((Get-ModelNorm $pair[0]))) { continue }
        $out[$k] = $Checks[$k]
    }
    return $out
}

# Начальный буфер, если его ещё нет (первый «Выборочно» после полных прогонов): открытые замечания из журнала +
# параметры и даты проверки из последнего среза истории. Без журнала и истории базы нет — нужна полная выгрузка.
function New-SeedBuffer($Ledger, $HistoryRows) {
    $last = ""
    foreach ($h in @($HistoryRows)) { if ([string]$h.BuildId -gt $last) { $last = [string]$h.BuildId } }
    $open = @($Ledger.Values | Where-Object { $_.Status -eq "open" })
    if (-not $last -and $open.Count -eq 0) { return $null }
    $paramOf = @{}
    $checks = @{}
    foreach ($h in @($HistoryRows)) {
        if ([string]$h.BuildId -ne $last) { continue }
        $pk = $h.Model + $script:PartialSep + $h.Check
        if (-not $paramOf.ContainsKey($pk)) { $paramOf[$pk] = $h.Param }
        $checks[$pk] = $last
    }
    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($rec in @($open | Sort-Object Model, Check, Family, Item)) {
        $pk = $rec.Model + $script:PartialSep + $rec.Check
        $param = $rec.ParamName
        if ($paramOf.ContainsKey($pk)) { $param = $paramOf[$pk] }
        $rows.Add([pscustomobject]@{
            Section = $rec.Section; Model = $rec.Model; Check = $rec.Check; Param = $param; ParamName = $rec.ParamName
            Family = $rec.Family; Item = $rec.Item; Value = $rec.Value
        })
        if (-not $checks.ContainsKey($pk)) { $checks[$pk] = [string]$rec.LastSeen }
    }
    return @{ Rows = $rows.ToArray(); Checks = $checks }
}

# Состояние после сборки. Режим «Всё»: буфер = вся выгрузка, проверена каждая спецификация целиком.
# Выборочно: буфер = виртуально полная выгрузка (проверенные пары + прежнее состояние остальных), проверены пары прогона.
# Один файл <история>_состояние.txt. Каталог — из этой выгрузки; в выгрузке его нет — остаётся прежний (как раньше оставался прежний _каталог.txt).
function Save-PartialState {
    param($Paths, $Rows, $Checks, [string]$ExportDir)
    $parts = @{ buffer = (ConvertTo-PartialBufferLines $Rows); checks = (ConvertTo-PartialChecksLines $Checks); catalog = $null }
    $export = $null
    if ($ExportDir) { $export = Get-BvsExportResult $ExportDir }
    if ($null -ne $export -and $export.HasCatalog) { $parts.catalog = @($export.Catalog | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) }
    else { $parts.catalog = (Read-PartialStateLines $Paths).catalog }
    Write-PartialStateFile $Paths.State $parts
}

# ---------- перенос прежних файлов состояния (один раз, при открытии окна) ----------
# Новое имя прежнего файла: <имя>.migrated_<yyyy-MM-dd>.txt (занято — _2, _3 …). Ничего не затирается и не удаляется.
function Get-PartialMigratedPath([string]$Path, $Now = $null) {
    if ($null -eq $Now) { $Now = Get-Date }
    $dir = Split-Path -Parent $Path
    $stem = [System.IO.Path]::GetFileNameWithoutExtension($Path)
    $day = ([datetime]$Now).ToString("yyyy-MM-dd", [System.Globalization.CultureInfo]::InvariantCulture)
    $to = Join-Path $dir ("{0}.migrated_{1}.txt" -f $stem, $day)
    $n = 2
    while (Test-Path -LiteralPath $to) { $to = Join-Path $dir ("{0}.migrated_{1}_{2}.txt" -f $stem, $day, $n); $n++ }
    return $to
}

# Одна история: <история>_буфер.txt + _проверки.txt + _каталог.txt -> <история>_состояние.txt (части, которых в прежних файлах нет, остаются
# из уже имеющегося файла состояния). До записи — резервная копия ($EnsureBackup); запись проверяется перечитыванием; затем прежние файлы
# переименовываются. Нет прежних файлов — ничего (повторный запуск ничего не делает). Возвращает @{ Action = none|merged; Renamed; Log }
function Invoke-PartialStateMigration([string]$HistoryPath, [scriptblock]$EnsureBackup, $Now = $null) {
    $res = @{ Action = "none"; Renamed = (New-Object System.Collections.Generic.List[string]); Log = (New-Object System.Collections.Generic.List[string]) }
    $paths = Get-PartialStatePaths $HistoryPath
    $olds = @($script:PartialStateParts | ForEach-Object { Get-PartialOldStatePath $paths $_ } | Where-Object { $_ -and (Test-Path -LiteralPath $_ -PathType Leaf) })
    if ($olds.Count -eq 0) { return $res }
    $backup = & $EnsureBackup
    if (-not $backup) { throw "нет резервной копии — перенос не выполнен" }
    $parts = Read-PartialStateLines $paths
    Write-PartialStateFile $paths.State $parts
    $check = Read-PartialStateLines @{ State = $paths.State }
    foreach ($part in $script:PartialStateParts) {
        $a = @($parts[$part]) -join "`n"; $b = @($check[$part]) -join "`n"
        if (($null -eq $parts[$part]) -ne ($null -eq $check[$part]) -or $a -ne $b) { throw ("{0}: после записи не совпал раздел [{1}] — прежние файлы оставлены" -f (Split-Path -Leaf $paths.State), $part) }
    }
    foreach ($old in $olds) {
        $to = Get-PartialMigratedPath $old $Now
        [System.IO.File]::Move($old, $to)
        $res.Renamed.Add($to)
        $res.Log.Add(("{0} -> {1}" -f (Split-Path -Leaf $old), (Split-Path -Leaf $to)))
    }
    $cnt = { param($x) if ($null -eq $x) { "нет" } else { [string]@($x).Count } }
    $res.Log.Insert(0, ("{0}: {1} файла(ов) объединены в {2} (строк: буфер {3}, проверки {4}, каталог {5})" -f [System.IO.Path]::GetFileNameWithoutExtension($HistoryPath), $olds.Count, (Split-Path -Leaf $paths.State),
        (& $cnt $parts.buffer), (& $cnt $parts.checks), (& $cnt $parts.catalog)))
    $res.Action = "merged"
    return $res
}

# Все истории в папке «история»: по каждому <история>_буфер|_проверки|_каталог.txt (кроме .migrated_*) — Invoke-PartialStateMigration
function Invoke-PartialStateMigrationAll([string]$HistoryDir, [scriptblock]$EnsureBackup, $Now = $null) {
    $out = @{ Results = (New-Object System.Collections.Generic.List[object]); Errors = (New-Object System.Collections.Generic.List[string]); Log = (New-Object System.Collections.Generic.List[string]) }
    if (-not $HistoryDir -or -not (Test-Path -LiteralPath $HistoryDir)) { return $out }
    $stems = New-Object System.Collections.Generic.List[string]
    foreach ($f in @(Get-ChildItem -LiteralPath $HistoryDir -File -ErrorAction SilentlyContinue | Sort-Object Name)) {
        if ($f.Name -match '\.migrated_') { continue }
        if ($f.Name -match '^(.+)_(буфер|проверки|каталог)\.txt$' -and -not $stems.Contains($Matches[1])) { $stems.Add($Matches[1]) }
    }
    foreach ($stem in $stems) {
        try {
            $r = Invoke-PartialStateMigration (Join-Path $HistoryDir ($stem + ".txt")) $EnsureBackup $Now
            $out.Results.Add($r)
            foreach ($l in $r.Log) { $out.Log.Add($l) }
        }
        catch { $out.Errors.Add(("{0}: {1}" -f $stem, $_.Exception.Message)); $out.Log.Add(("{0}: ОШИБКА {1} — прежние файлы оставлены как есть" -f $stem, $_.Exception.Message)) }
    }
    return $out
}
