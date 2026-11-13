# UTF-8 WITH BOM required (PowerShell 5.1).
# Раскладка папки «отчеты» (путь к ней не меняется):
#   отчеты\                    — итоговые .xlsx
#   отчеты\история\            — постоянные файлы между прогонами (история, журнал, буфер/проверки/каталог «Выборочно»). Автоматически не удаляются
#   отчеты\служебные\<время>\  — промежуточные файлы одного прогона (выгрузка из Revit, manifest, диагностика фильтров, JSON для Google).
#                                Подпапки старше AuxKeepDays (по умолчанию 30, ui_settings.txt: auxkeepdays=) удаляются автоматически
#   отчеты\операции\           — дампы и CSV служебных .bat. Папки и файлы старше 30 дней удаляются (как служебные).
#   отчеты\логи\               — журналы ЭТОГО инструмента: логи RBP текущего запуска BVS, google_push.log, partial_sets.log, ошибки .py.
#                                Папка передаётся в BatchRvt только как --log_folder в командной строке запуска (без смены настроек RBP).
#                                Запуски ModelOps, других инструментов и ручного окна RBP пишут логи по-своему / в папку по умолчанию.
#                                BatchRvt_*.log старше AuxKeepDays удаляются. Чужие логи (%LOCALAPPDATA%\RevitBatchProcessor) не трогаются и не переносятся.
#   *.bad.xlsx                 — битый xlsx задания; старше 30 дней удаляется. Итоговые xlsx, динамика, история, отчеты\проект\ не трогаются.
#   отчеты\проект\_backup_<дата>\ — резервные копии перед переносом файлов состояния (project_config.ps1: New-ProjectStateBackup). Автоочисткой не удаляются никогда.
# Миграция старой раскладки (файлы прямо в «отчеты\») — Invoke-ReportMigration: файлы переносятся, ничего не удаляется.

# Раскладка папки: <корень>\Проверочные спецификации.vbs, ИНСТРУКЦИЯ.txt, servers.cfg, отчеты\, служебное\ (в нём скрипты\ — эти .ps1/.py, lib\, google_push, разработка\инструменты\ — служебные .bat …).
# Все пути — только от расположения скриптов ($ToolDir = служебное\скрипты), папку можно копировать куда угодно (в т. ч. на UNC).
function Get-BvsServiceDir([string]$ToolDir) {
    $t = ([string]$ToolDir).TrimEnd([char[]]@([char]92, [char]47))
    if ((Split-Path -Leaf $t) -eq "скрипты") { return (Split-Path -Parent $t) }
    return (Join-Path $t "служебное")
}
function Get-BvsRootDir([string]$ToolDir) {
    $t = ([string]$ToolDir).TrimEnd([char[]]@([char]92, [char]47))
    if ((Split-Path -Leaf $t) -eq "скрипты") { return (Split-Path -Parent (Split-Path -Parent $t)) }
    return $t
}
function Get-BvsLibDir([string]$ToolDir) { return (Join-Path (Get-BvsServiceDir $ToolDir) "lib") }
function Get-ReportRootDir([string]$ToolDir) { return (Join-Path (Get-BvsRootDir $ToolDir) "отчеты") }

$script:RlHistoryName = "история"
$script:RlAuxName = "служебные"
$script:RlOpsName = "операции"
$script:RlLogsName = "логи"
$script:RlDefaultKeepDays = 30

function Get-ReportLayout([string]$ReportDir) {
    return [pscustomobject]@{
        Root    = $ReportDir
        History = (Join-Path $ReportDir $script:RlHistoryName)
        Aux     = (Join-Path $ReportDir $script:RlAuxName)
        Ops     = (Join-Path $ReportDir $script:RlOpsName)
        Logs    = (Join-Path $ReportDir $script:RlLogsName)
    }
}

# PS 5.1 Start-Process -ArgumentList с массивом не заключает значения с пробелами в кавычки.
# Кавычки добавляем только если они нужны: пути без пробелов остаются как были.
function Quote-BvsProcessArg([string]$Value) {
    if ($null -eq $Value) { return "" }
    if ($Value -notmatch '\s' -and $Value -notmatch '"') { return $Value }
    return ('"{0}"' -f ($Value -replace '"', '\"'))
}

# Куда писать журналы окна: отчеты\логи, если раскладка уже есть; иначе запасной AppDir.
function Get-BvsToolLogDir([string]$AppDir = "") {
    if ($script:Layout -and $script:Layout.Logs) { return [string]$script:Layout.Logs }
    if ($AppDir) { return $AppDir }
    return ""
}

# Срок хранения служебных папок: auxkeepdays= в ui_settings.txt, иначе 30 (допустимо 1..3650)
function Get-AuxKeepDays([string]$SettingsPath = "") {
    if (-not $SettingsPath -and $env:LOCALAPPDATA) { $SettingsPath = Join-Path $env:LOCALAPPDATA "BatchRvt\batch_verification_schedules\ui_settings.txt" }
    if ($SettingsPath -and (Test-Path -LiteralPath $SettingsPath)) {
        foreach ($line in @(Get-Content -LiteralPath $SettingsPath -Encoding UTF8)) {
            if ($line -match '^(?i)auxkeepdays=\s*(\d+)\s*$') {
                $n = [int]$Matches[1]
                if ($n -ge 1 -and $n -le 3650) { return $n }
            }
        }
    }
    return $script:RlDefaultKeepDays
}

function Test-RlSameFile([string]$A, [string]$B) {
    try {
        $fa = Get-Item -LiteralPath $A -ErrorAction Stop
        $fb = Get-Item -LiteralPath $B -ErrorAction Stop
        if ($fa.Length -ne $fb.Length) { return $false }
        $ha = (Get-FileHash -LiteralPath $A -Algorithm SHA256).Hash
        $hb = (Get-FileHash -LiteralPath $B -Algorithm SHA256).Hash
        return ($ha -eq $hb)
    }
    catch { return $false }
}

function Move-RlItem([string]$From, [string]$To, [string]$ConflictDir, $Moves, [string]$Kind) {
    $dir = Split-Path -Parent $To
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $final = $To
    $note = ""
    if (Test-Path -LiteralPath $To) {
        # назначение занято: ничего не затираем и не удаляем — исходный файл уходит в папку конфликтов
        if (-not (Test-Path -LiteralPath $ConflictDir)) { New-Item -ItemType Directory -Path $ConflictDir -Force | Out-Null }
        $leaf = Split-Path -Leaf $From
        $same = (Test-Path -LiteralPath $From -PathType Leaf) -and (Test-RlSameFile $From $To)
        $final = Join-Path $ConflictDir $leaf
        $n = 2
        while (Test-Path -LiteralPath $final) {
            $final = Join-Path $ConflictDir ("{0}_{1}{2}" -f [System.IO.Path]::GetFileNameWithoutExtension($leaf), $n, [System.IO.Path]::GetExtension($leaf))
            $n++
        }
        if ($same) { $note = "дубликат: такой же файл уже на месте" } else { $note = "назначение занято другим файлом" }
    }
    Move-Item -LiteralPath $From -Destination $final -Force -ErrorAction Stop
    $Moves.Add([pscustomobject]@{ Kind = $Kind; From = $From; To = $final; Note = $note })
}

# Что лежит в «отчеты\» по старой раскладке -> куда его перенести. Итоговые .xlsx не трогаются.
function Invoke-ReportMigration([string]$ReportDir) {
    $moves = New-Object System.Collections.Generic.List[object]
    $warnings = New-Object System.Collections.Generic.List[string]
    if (-not (Test-Path -LiteralPath $ReportDir)) { return [pscustomobject]@{ Moves = @(); Warnings = @() } }
    $layout = Get-ReportLayout $ReportDir
    $legacyAux = Join-Path $layout.Aux "_до_переноса"
    foreach ($item in @(Get-ChildItem -LiteralPath $ReportDir -Force -ErrorAction SilentlyContinue)) {
        $name = $item.Name
        $target = $null
        $kind = ""
        $conflicts = ""
        if ($item.PSIsContainer) {
            if ($name -match '^дамп_фильтров_') { $target = Join-Path $layout.Ops $name; $kind = "операции"; $conflicts = Join-Path $layout.Ops "_конфликты_переноса" }
        }
        else {
            if ($name -match '^(история_.+|задание_история(_.+)?|задание_журнал(_.+)?)\.txt$') {
                $target = Join-Path $layout.History $name; $kind = "история"; $conflicts = Join-Path $layout.History "_конфликты_переноса"
            }
            elseif ($name -match '^(исправление_фильтров|замена_фильтра|приведение_к_эталону|сравнение_фильтров)_.*\.csv$') {
                $target = Join-Path $layout.Ops $name; $kind = "операции"; $conflicts = Join-Path $layout.Ops "_конфликты_переноса"
            }
            elseif ($name -eq "google_push_payload.json" -or $name -eq "фильтры_спецификаций.txt") {
                $target = Join-Path $legacyAux $name; $kind = "служебные"; $conflicts = Join-Path $legacyAux "_конфликты_переноса"
            }
        }
        if (-not $target) { continue }
        try { Move-RlItem $item.FullName $target $conflicts $moves $kind }
        catch { $warnings.Add(("не перенесён {0}: {1}" -f $name, $_.Exception.Message)) }
    }
    return [pscustomobject]@{ Moves = $moves.ToArray(); Warnings = $warnings.ToArray() }
}

# Подпапки «служебные\<время>» старше KeepDays (время — из имени yyyyMMdd_HHmmss, иначе по дате изменения). История, операции и .xlsx не затрагиваются.
function Remove-OldAuxDirs([string]$AuxRoot, [int]$KeepDays = 30, $Now = $null) {
    $removed = New-Object System.Collections.Generic.List[string]
    if (-not (Test-Path -LiteralPath $AuxRoot)) { return @() }
    if ($KeepDays -lt 1) { $KeepDays = $script:RlDefaultKeepDays }
    if ($null -eq $Now) { $Now = Get-Date }
    $limit = ([datetime]$Now).AddDays(-$KeepDays)
    $rootFull = [System.IO.Path]::GetFullPath($AuxRoot).TrimEnd('\', '/')
    foreach ($d in @(Get-ChildItem -LiteralPath $AuxRoot -Directory -Force -ErrorAction SilentlyContinue)) {
        $stamp = $d.LastWriteTime
        if ($d.Name -match '^(\d{8})_(\d{6})') {
            $parsed = [datetime]::MinValue
            if ([datetime]::TryParseExact(($Matches[1] + $Matches[2]), "yyyyMMddHHmmss", [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$parsed)) { $stamp = $parsed }
        }
        if ($stamp -ge $limit) { continue }
        $full = [System.IO.Path]::GetFullPath($d.FullName)
        if (-not $full.StartsWith($rootFull + [System.IO.Path]::DirectorySeparatorChar) -and -not $full.StartsWith($rootFull + "/")) { continue }
        try { Remove-Item -LiteralPath $d.FullName -Recurse -Force -ErrorAction Stop; $removed.Add($d.Name) }
        catch { }
    }
    return $removed.ToArray()
}

# Дата из имени (yyyyMMdd_HHmmss или yyyy-MM-dd_HH-mm-ss). Нет штампа — $null, вызывающий берёт дату изменения.
function Get-BvsNamedStamp([string]$Name) {
    $parsed = [datetime]::MinValue
    $culture = [System.Globalization.CultureInfo]::InvariantCulture
    $style = [System.Globalization.DateTimeStyles]::None
    if ($Name -match '(\d{8})_(\d{6})') {
        if ([datetime]::TryParseExact(($Matches[1] + $Matches[2]), "yyyyMMddHHmmss", $culture, $style, [ref]$parsed)) { return $parsed }
    }
    if ($Name -match '(\d{4})-(\d{2})-(\d{2})_(\d{2})-(\d{2})-(\d{2})') {
        $compact = $Matches[1] + $Matches[2] + $Matches[3] + $Matches[4] + $Matches[5] + $Matches[6]
        if ([datetime]::TryParseExact($compact, "yyyyMMddHHmmss", $culture, $style, [ref]$parsed)) { return $parsed }
    }
    return $null
}

function Test-BvsOlderThan($Item, [datetime]$Limit) {
    $stamp = Get-BvsNamedStamp $Item.Name
    if ($null -eq $stamp) { $stamp = $Item.LastWriteTime }
    return ($stamp -lt $Limit)
}

# Подпапки и файлы «операции\» старше KeepDays. Саму папку операции не удаляет. История и xlsx здесь не лежат.
function Remove-OldOpsItems([string]$OpsRoot, [int]$KeepDays = 30, $Now = $null) {
    $removed = New-Object System.Collections.Generic.List[string]
    if (-not (Test-Path -LiteralPath $OpsRoot)) { return @() }
    if ($KeepDays -lt 1) { $KeepDays = $script:RlDefaultKeepDays }
    if ($null -eq $Now) { $Now = Get-Date }
    $limit = ([datetime]$Now).AddDays(-$KeepDays)
    foreach ($item in @(Get-ChildItem -LiteralPath $OpsRoot -Force -ErrorAction SilentlyContinue)) {
        if (-not (Test-BvsOlderThan $item $limit)) { continue }
        try {
            if ($item.PSIsContainer) { Remove-Item -LiteralPath $item.FullName -Recurse -Force -ErrorAction Stop }
            else { Remove-Item -LiteralPath $item.FullName -Force -ErrorAction Stop }
            $removed.Add($item.Name)
        }
        catch { }
    }
    return $removed.ToArray()
}

# Только файлы *.bad.xlsx старше KeepDays. Папка отчеты\проект\ и отчеты\история\ не просматриваются.
# Обычные xlsx (задание, вставка, динамика) не подходят под имя *.bad.xlsx и не удаляются.
function Remove-OldBadXlsx([string]$ReportDir, [int]$KeepDays = 30, $Now = $null) {
    $removed = New-Object System.Collections.Generic.List[string]
    if (-not (Test-Path -LiteralPath $ReportDir)) { return @() }
    if ($KeepDays -lt 1) { $KeepDays = $script:RlDefaultKeepDays }
    if ($null -eq $Now) { $Now = Get-Date }
    $limit = ([datetime]$Now).AddDays(-$KeepDays)
    $root = [System.IO.Path]::GetFullPath($ReportDir).TrimEnd('\', '/')
    foreach ($f in @(Get-ChildItem -LiteralPath $ReportDir -Recurse -File -Force -ErrorAction SilentlyContinue)) {
        if (-not $f.Name.EndsWith(".bad.xlsx")) { continue }
        $full = [System.IO.Path]::GetFullPath($f.FullName)
        if (-not $full.StartsWith($root)) { continue }
        $rel = $full.Substring($root.Length).TrimStart('\', '/')
        $parts = @($rel -split '[\\/]')
        $top = $parts[0]
        if ($top -eq "проект" -or $top -eq "история") { continue }
        if (@($parts | Where-Object { $_ -like "_backup_*" }).Count -gt 0) { continue }          # резервные копии — никогда
        if (-not (Test-BvsOlderThan $f $limit)) { continue }
        try { Remove-Item -LiteralPath $f.FullName -Force -ErrorAction Stop; $removed.Add($rel) }
        catch { }
    }
    return $removed.ToArray()
}

# Оставляет хвост файла, если он разросся. Журналы дописываются; ошибка обрезки не должна ронять окно.
function Trim-BvsTextFile([string]$Path, [int]$LimitBytes = 1048576, [int]$KeepBytes = 524288) {
    if (-not $Path -or -not (Test-Path -LiteralPath $Path -PathType Leaf)) { return }
    $len = (Get-Item -LiteralPath $Path).Length
    if ($len -le $LimitBytes) { return }
    if ($KeepBytes -ge $len) { $KeepBytes = [Math]::Max(1, [int]($len / 2)) }
    $fs = $null
    try {
        $fs = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
        [void]$fs.Seek($len - $KeepBytes, [System.IO.SeekOrigin]::Begin)
        $buf = New-Object byte[] $KeepBytes
        $n = $fs.Read($buf, 0, $KeepBytes)
    }
    catch { return }
    finally { if ($fs) { $fs.Dispose() } }
    $i = 0
    while ($i -lt $n -and $buf[$i] -ne 10) { $i++ }
    if ($i -lt $n) { $i++ }
    $tail = New-Object byte[] ($n - $i)
    if ($tail.Length -gt 0) { [Array]::Copy($buf, $i, $tail, 0, $tail.Length) }
    try { [System.IO.File]::WriteAllBytes($Path, $tail) } catch { }
}

# Журналы окна и файлы ошибок этого инструмента (в отчеты\логи; обрезка хвоста, если разрослись).
# Остатки в прежних местах (AppDir, %TEMP%) тоже обрезаются; чужие *_error.txt в %TEMP% не трогаются.
function Trim-BvsToolLogs([string]$LogDir = "", [string]$AppDir = "") {
    $dirs = New-Object System.Collections.Generic.List[string]
    if ($LogDir) { $dirs.Add($LogDir) }
    if ($AppDir -and $AppDir -ne $LogDir) { $dirs.Add($AppDir) }
    foreach ($d in $dirs) {
        Trim-BvsTextFile (Join-Path $d "google_push.log")
        Trim-BvsTextFile (Join-Path $d "partial_sets.log")
        foreach ($name in @("read_template_specs_error.txt", "dump_schedule_filters_error.txt")) {
            Trim-BvsTextFile (Join-Path $d $name)
        }
    }
    $temp = [System.IO.Path]::GetTempPath()
    foreach ($name in @("read_template_specs_error.txt", "dump_schedule_filters_error.txt")) {
        Trim-BvsTextFile (Join-Path $temp $name)
    }
}

# Переносит журналы окна из AppDir и файлы ошибок из %TEMP% в отчеты\логи. Содержимое не теряется:
# если файл уже есть на новом месте, старый дописывается в конец, исходник удаляется.
# Чужие логи RBP (%LOCALAPPDATA%\RevitBatchProcessor и других инструментов) не читает и не переносит.
function Move-BvsAppDirLogs([string]$AppDir = "", [string]$LogDir = "") {
    $moved = New-Object System.Collections.Generic.List[string]
    if (-not $LogDir) { return @() }
    if (-not (Test-Path -LiteralPath $LogDir)) { New-Item -ItemType Directory -Path $LogDir -Force | Out-Null }
    $pairs = New-Object System.Collections.Generic.List[object]
    if ($AppDir) {
        foreach ($name in @("google_push.log", "partial_sets.log")) {
            $pairs.Add([pscustomobject]@{ From = (Join-Path $AppDir $name); To = (Join-Path $LogDir $name) })
        }
    }
    $temp = [System.IO.Path]::GetTempPath()
    foreach ($name in @("read_template_specs_error.txt", "dump_schedule_filters_error.txt")) {
        $pairs.Add([pscustomobject]@{ From = (Join-Path $temp $name); To = (Join-Path $LogDir $name) })
    }
    foreach ($p in $pairs) {
        if (-not (Test-Path -LiteralPath $p.From -PathType Leaf)) { continue }
        $fromFull = [System.IO.Path]::GetFullPath($p.From)
        $toFull = [System.IO.Path]::GetFullPath($p.To)
        if ($fromFull -eq $toFull) { continue }
        try {
            if (-not (Test-Path -LiteralPath $p.To -PathType Leaf)) {
                Move-Item -LiteralPath $p.From -Destination $p.To -Force -ErrorAction Stop
            }
            else {
                $extra = [System.IO.File]::ReadAllText($p.From)
                if ($extra) { [System.IO.File]::AppendAllText($p.To, $extra, (New-Object System.Text.UTF8Encoding $true)) }
                Remove-Item -LiteralPath $p.From -Force -ErrorAction Stop
            }
            $moved.Add((Split-Path -Leaf $p.From))
        }
        catch { }
    }
    return $moved.ToArray()
}

# Только файлы BatchRvt_*.log в отчеты\логи старше KeepDays. Журналы окна и файлы ошибок не удаляет.
# Папку %LOCALAPPDATA%\RevitBatchProcessor не просматривает и ничего оттуда не удаляет.
function Remove-OldLogItems([string]$LogRoot, [int]$KeepDays = 30, $Now = $null) {
    $removed = New-Object System.Collections.Generic.List[string]
    if (-not (Test-Path -LiteralPath $LogRoot)) { return @() }
    if ($KeepDays -lt 1) { $KeepDays = $script:RlDefaultKeepDays }
    if ($null -eq $Now) { $Now = Get-Date }
    $limit = ([datetime]$Now).AddDays(-$KeepDays)
    $rootFull = [System.IO.Path]::GetFullPath($LogRoot).TrimEnd('\', '/')
    foreach ($f in @(Get-ChildItem -LiteralPath $LogRoot -File -Force -ErrorAction SilentlyContinue)) {
        if ($f.Name -notlike "BatchRvt_*.log") { continue }
        $full = [System.IO.Path]::GetFullPath($f.FullName)
        if (-not $full.StartsWith($rootFull + [System.IO.Path]::DirectorySeparatorChar) -and -not $full.StartsWith($rootFull + "/")) { continue }
        if (-not (Test-BvsOlderThan $f $limit)) { continue }
        try { Remove-Item -LiteralPath $f.FullName -Force -ErrorAction Stop; $removed.Add($f.Name) }
        catch { }
    }
    return $removed.ToArray()
}

# 30 дней: служебные, операции, BatchRvt_*.log в отчеты\логи, *.bad.xlsx; плюс обрезка журналов окна.
# Не трогает проект (и его _backup_*), историю, google_*.cfg, итоговые xlsx и чужие логи RBP вне отчеты\логи.
function Invoke-BvsHousekeeping {
    param([string]$ReportDir, [int]$KeepDays = 30, [string]$AppDir = "", $Now = $null)
    if ($KeepDays -lt 1) { $KeepDays = $script:RlDefaultKeepDays }
    $layout = Get-ReportLayout $ReportDir
    if ($layout.Logs -and -not (Test-Path -LiteralPath $layout.Logs)) { New-Item -ItemType Directory -Path $layout.Logs -Force | Out-Null }
    $moved = @()
    if ($layout.Logs) { $moved = @(Move-BvsAppDirLogs $AppDir $layout.Logs) }
    $aux = @(Remove-OldAuxDirs $layout.Aux $KeepDays $Now)
    $ops = @(Remove-OldOpsItems $layout.Ops $KeepDays $Now)
    $bad = @(Remove-OldBadXlsx $ReportDir $KeepDays $Now)
    $logs = @(Remove-OldLogItems $layout.Logs $KeepDays $Now)
    Trim-BvsToolLogs -LogDir $layout.Logs -AppDir $AppDir
    return [pscustomobject]@{ Aux = $aux; Ops = $ops; Bad = $bad; Logs = $logs; Moved = $moved }
}

# Создаёт папки, переносит старые файлы (с журналом переноса), чистит старые служебные папки.
# Возвращает layout + Migration (что перенесено) + Removed (какие служебные папки удалены).
function Initialize-ReportLayout([string]$ReportDir, [int]$KeepDays = 0, [string]$SettingsPath = "", $Now = $null, [string]$AppDir = "") {
    if (-not (Test-Path -LiteralPath $ReportDir)) { New-Item -ItemType Directory -Path $ReportDir -Force | Out-Null }
    $layout = Get-ReportLayout $ReportDir
    foreach ($d in @($layout.History, $layout.Aux, $layout.Ops, $layout.Logs)) {
        if (-not (Test-Path -LiteralPath $d)) { New-Item -ItemType Directory -Path $d -Force | Out-Null }
    }
    $mig = $null
    try { $mig = Invoke-ReportMigration $ReportDir } catch { $mig = [pscustomobject]@{ Moves = @(); Warnings = @($_.Exception.Message) } }
    if (@($mig.Moves).Count -gt 0) {
        try {
            $logDir = Join-Path $layout.Aux ("_перенос_" + (Get-Date -Format "yyyyMMdd_HHmmss"))
            New-Item -ItemType Directory -Path $logDir -Force | Out-Null
            $lines = @($mig.Moves | ForEach-Object { "{0}`t{1}`t{2}`t{3}" -f $_.Kind, $_.From, $_.To, $_.Note })
            [System.IO.File]::WriteAllLines((Join-Path $logDir "перенос.txt"), [string[]]$lines, (New-Object System.Text.UTF8Encoding $true))
        }
        catch { }
    }
    if ($KeepDays -lt 1) { $KeepDays = Get-AuxKeepDays $SettingsPath }
    $removed = @(Remove-OldAuxDirs $layout.Aux $KeepDays $Now)
    [void](Remove-OldOpsItems $layout.Ops $KeepDays $Now)
    [void](Remove-OldBadXlsx $ReportDir $KeepDays $Now)
    if ($AppDir) { [void](Move-BvsAppDirLogs $AppDir $layout.Logs) }
    [void](Remove-OldLogItems $layout.Logs $KeepDays $Now)
    [void](Trim-BvsToolLogs -LogDir $layout.Logs -AppDir $AppDir)
    return [pscustomobject]@{ Layout = $layout; Migration = $mig; Removed = $removed; KeepDays = $KeepDays }
}

# Папка служебных файлов прогона: служебные\<runId> (runId = yyyyMMdd_HHmmss)
function New-RunAuxDir([string]$AuxRoot, [string]$RunId) {
    if (-not (Test-Path -LiteralPath $AuxRoot)) { New-Item -ItemType Directory -Path $AuxRoot -Force | Out-Null }
    $dir = Join-Path $AuxRoot $RunId
    $n = 2
    while (Test-Path -LiteralPath $dir) { $dir = Join-Path $AuxRoot ("{0}_{1}" -f $RunId, $n); $n++ }
    New-Item -ItemType Directory -Path $dir -Force | Out-Null
    return $dir
}

# После успешной сборки крупные промежуточные файлы выгрузки (NNNN.tsv, _sections) не нужны; мелкая диагностика остаётся до автоочистки
function Remove-RunBulk([string]$Dir) {
    if (-not $Dir -or -not (Test-Path -LiteralPath $Dir)) { return }
    foreach ($f in @(Get-ChildItem -LiteralPath $Dir -Filter "*.tsv" -File -ErrorAction SilentlyContinue)) { Remove-Item -LiteralPath $f.FullName -Force -ErrorAction SilentlyContinue }
    $sections = Join-Path $Dir "_sections"
    if (Test-Path -LiteralPath $sections) { Remove-Item -LiteralPath $sections -Recurse -Force -ErrorAction SilentlyContinue }
}

# Журнал замечаний раньше был один на все проекты (история_<код>.txt — на проект, задание_журнал.txt — общий). Теперь задание_журнал_<код>.txt на проект;
# старый общий журнал переносится в проект PROJ (если там журнала ещё нет; иначе — в папку конфликтов, ничего не затирается и не удаляется).
$script:RlLegacyLedgerProject = "PROJ"
function Invoke-LedgerPerProjectMigration([string]$ReportDir, [string]$Project = "") {
    if (-not $Project) { $Project = $script:RlLegacyLedgerProject }
    $moves = New-Object System.Collections.Generic.List[object]
    $layout = Get-ReportLayout $ReportDir
    $old = Join-Path $layout.History "задание_журнал.txt"
    if (-not (Test-Path -LiteralPath $old -PathType Leaf)) { return @() }
    $target = Join-Path $layout.History ("задание_журнал_{0}.txt" -f $Project)
    try { Move-RlItem $old $target (Join-Path $layout.History "_конфликты_переноса") $moves "журнал" } catch { }
    return $moves.ToArray()
}

# Итог прогона вставки (шаги «Сводная»/«Разделы») по <id>_сводка.csv: Модель;Раздел;Удалено;Вставлено;Пропущено;Ошибка.
# Возвращает @{ Total = строк (моделей); Errors = моделей с непустой «Ошибка»; Names = их имена }
function Get-RunSummaryStats([string]$SummaryCsv) {
    $total = 0; $errors = 0
    $names = New-Object System.Collections.Generic.List[string]
    if ($SummaryCsv -and (Test-Path -LiteralPath $SummaryCsv)) {
        $text = [System.IO.File]::ReadAllText($SummaryCsv, (New-Object System.Text.UTF8Encoding $true))
        $first = $true
        foreach ($line in @($text -split "`r?`n")) {
            if (-not $line.Trim()) { continue }
            if ($first) { $first = $false; continue }          # заголовок
            $p = $line.Split(";")
            $total++
            if ($p.Count -ge 6 -and $p[5].Trim()) { $errors++; $names.Add($p[0]) }
        }
    }
    return @{ Total = $total; Errors = $errors; Names = $names.ToArray() }
}

# Текст итога прогона: пусто, если всё хорошо (код RBP 0, модели есть, ошибок нет); иначе короткое сообщение с числом моделей с ошибкой
function Get-RunFailureText([int]$ExitCode, $Stats) {
    if ($Stats.Total -eq 0) { return ("Не готово: ни одна модель не обработана (код RBP {0}) — Revit не запустился или не открыл модели." -f $ExitCode) }
    if ($Stats.Errors -gt 0) { return ("Не готово: с ошибкой {0} из {1} моделей (код RBP {2}). Подробности — в отчёте, лист «Сводка»." -f $Stats.Errors, $Stats.Total, $ExitCode) }
    if ($ExitCode -ne 0) { return ("Не готово: Revit Batch Processor завершился с кодом {0}; обработано моделей: {1}." -f $ExitCode, $Stats.Total) }
    return ""
}
