# UTF-8 WITH BOM required (PowerShell 5.1).
# Выгрузка фильтров всех спецификаций выбранных rvt-файлов в текстовые файлы (только чтение: модели не меняются и не
# сохраняются, в Google ничего не пишется). Формат — как у фильтры_спецификаций.txt: Спецификация, Поле фильтра, Условие, Значение.
# Что именно пишется — см. dump_schedule_filters.py. Сравнение двух выгрузок — compare_filter_dumps.ps1.
#
# Запуск: перетащите на «Выгрузить фильтры спек.bat» файлы .rvt, папку (поиск по вложенным папкам),
#   файл-список .txt (путь на строку) либо
#   powershell -File dump_schedule_filters.ps1 "путь\файл1.rvt" "путь\файл2.rvt"
# Несколько путей в одном аргументе — через «;». Без аргументов скрипт спросит путь.
# Из папки не берутся резервные копии Revit (*.NNNN.rvt), сводные (SVOD/СВОД) и каталоги
# Резерв*, Архив*, Archive*, Backup*, _backup*, Old, Стар/Старые/Старое/Старый (не «Старт»), OLD_*, а также папки-даты (2026-10-19…).
# Одинаковое имя файла в нескольких местах — один раз, с предупреждением.
# Файлы: отчеты\операции\дамп_фильтров_<время>\фильтры_<имя модели>.txt и спеки_<имя модели>.txt (только имена спецификаций).

# PositionalBinding=$false: при перетаскивании двух файлов второй путь раньше попадал в -OutDir (позиционный параметр),
# и «папкой вывода» становился rvt-файл. Теперь все позиционные аргументы идут только в $Models.
[CmdletBinding(PositionalBinding = $false)]
param(
    [Parameter(Position = 0, ValueFromRemainingArguments = $true)][string[]]$Models = @(),
    [string]$OutDir = "",
    [string]$NameContains = "",
    [string]$RbpExe = "",
    [string]$AppDir = "",
    [switch]$NoPause
)

$ErrorActionPreference = "Stop"
$ServiceDir = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot "..\.."))  # служебное (скрипт лежит в служебное\разработка\инструменты)
$ToolDir = Join-Path $ServiceDir "скрипты"          # служебное\скрипты: общие скрипты окна
$TaskScript = Join-Path $PSScriptRoot "dump_schedule_filters.py"
. (Join-Path $ToolDir "report_layout.ps1")
if (-not $AppDir) { $AppDir = Join-Path $env:LOCALAPPDATA "BatchRvt\batch_verification_schedules" }

function Stop-Here([string]$Message, [int]$Code = 1) {
    Write-Host $Message -ForegroundColor Red
    if (-not $NoPause) { [void](Read-Host "Нажмите Enter") }
    exit $Code
}

function Test-ModelReady([string]$Path) {
    if ([string]::IsNullOrWhiteSpace($Path)) { return $false }
    if ($Path.Trim() -match '^(?i)RSN://.+\.rvt$') { return $true }
    return (Test-Path -LiteralPath $Path)
}

function Test-DumpExcludedSegment([string]$Name) {
    if ([string]::IsNullOrWhiteSpace($Name)) { return $false }
    # Резерв*, Архив*, Archive*, Backup*, _backup*, Old, Стар/Старые/Старое/Старый, OLD_*, папка-дата 2026-10-19…
    if ($Name -match '^(?i)(Резерв|Архив|Archive|Backup|_backup)') { return $true }
    if ($Name -match '^(?i)Old$') { return $true }
    if ($Name -match '^(?i)Стар(ые|ое|ый)?$') { return $true }
    if ($Name -match '^(?i)OLD_') { return $true }
    if ($Name -match '^\d{4}-\d{2}-\d{2}') { return $true }
    return $false
}

function Test-DumpExcludedPath([string]$Path) {
    foreach ($seg in @(([string]$Path) -split '[\\/]')) {
        if (Test-DumpExcludedSegment $seg) { return $true }
    }
    return $false
}

function Test-DumpBackupPath([string]$Path) {
    $name = [System.IO.Path]::GetFileName($Path)
    if ($name -match '(?i)\.\d{4}\.rvt$') { return $true }
    return (Test-DumpExcludedPath $Path)
}

function Test-DumpSummaryModel([string]$Path) {
    $name = [System.IO.Path]::GetFileNameWithoutExtension($Path)
    return ($name -match '(?i)(_SVOD_|_СВОД_|СВОДН)')
}

# Папка -> *.rvt во вложенных папках. В каталоги резервных копий, архива и дат не заходим. Явный файл .rvt остаётся как есть.
function Expand-DumpFolder([string]$Dir, $Skipped) {
    $found = New-Object System.Collections.Generic.List[string]
    $queue = New-Object System.Collections.Generic.Queue[string]
    $queue.Enqueue($Dir)
    while ($queue.Count -gt 0) {
        $cur = $queue.Dequeue()
        foreach ($d in @(Get-ChildItem -LiteralPath $cur -Directory -ErrorAction SilentlyContinue)) {
            if (Test-DumpExcludedSegment $d.Name) { $Skipped.BackupFolders++ ; continue }
            $queue.Enqueue($d.FullName)
        }
        foreach ($f in @(Get-ChildItem -LiteralPath $cur -File -Filter "*.rvt" -ErrorAction SilentlyContinue)) {
            if ($f.Name -match '(?i)\.\d{4}\.rvt$') { $Skipped.Backup++ ; continue }
            if (Test-DumpSummaryModel $f.FullName) { $Skipped.Summary++ ; continue }
            $found.Add($f.FullName)
        }
    }
    return $found
}

# Одно и то же имя файла в нескольких местах: берём копию вне папок резерва/архива/дат (если таких несколько — с самым коротким путём) и предупреждаем.
function Select-DumpUniqueModels($Paths, $Warnings) {
    $byName = @{}
    foreach ($p in @($Paths)) {
        $key = [System.IO.Path]::GetFileName([string]$p).ToLowerInvariant()
        if (-not $key) { continue }
        if (-not $byName.ContainsKey($key)) { $byName[$key] = New-Object System.Collections.Generic.List[string] }
        [void]$byName[$key].Add([string]$p)
    }
    $out = New-Object System.Collections.Generic.List[string]
    foreach ($key in @($byName.Keys | Sort-Object)) {
        $group = @($byName[$key])
        $clean = @($group | Where-Object { -not (Test-DumpExcludedPath $_) })
        $pool = $(if ($clean.Count -gt 0) { $clean } else { @() })
        if ($pool.Count -eq 0) { continue }
        $best = @($pool | Sort-Object { @($_ -split '[\\/]').Count }, { $_.Length }, { $_ } | Select-Object -First 1)
        $out.Add([string]$best)
        if ($group.Count -gt 1) {
            $drop = @($group | Where-Object { $_ -ne [string]$best })
            $Warnings.Add(("{0}: оставлено {1}; пропущены дубли: {2}" -f [System.IO.Path]::GetFileName([string]$best), [string]$best, ($drop -join "; ")))
        }
    }
    return $out
}

function Add-DumpToken([string]$Token, $Into, $Skipped) {
    $t = $Token.Trim().Trim('"').Trim("'")
    if (-not $t -or $t.StartsWith("#")) { return }
    if ($t -match '(?i)\.txt$' -and (Test-Path -LiteralPath $t -PathType Leaf)) {
        $Skipped.Expanded = $true
        foreach ($line in @([System.IO.File]::ReadAllLines($t))) {
            foreach ($piece in @($line.Split(';'))) { Add-DumpToken $piece $Into $Skipped }
        }
        return
    }
    if ((Test-Path -LiteralPath $t -PathType Container)) {
        $Skipped.Expanded = $true
        foreach ($p in @(Expand-DumpFolder $t $Skipped)) { if (-not $Into.Contains($p)) { $Into.Add($p) } }
        return
    }
    if (Test-DumpBackupPath $t) { $Skipped.Backup++ ; return }
    if (-not $Into.Contains($t)) { $Into.Add($t) }
}

function Find-BatchRvt {
    $names = @("BatchRvt.exe", "BatchRvtGUI.exe")
    $dirs = New-Object System.Collections.Generic.List[string]
    Get-Process -ErrorAction SilentlyContinue |
        Where-Object { $_.ProcessName -match '^(?i)BatchRvt(GUI)?$|RevitBatchProcessor$' } |
        ForEach-Object {
            try { $d = Split-Path -Parent $_.MainModule.FileName; if ($d) { $dirs.Add($d) } } catch { }
        }
    foreach ($root in @($env:LOCALAPPDATA, $env:ProgramFiles, ${env:ProgramFiles(x86)}, $env:ProgramW6432)) {
        if ([string]::IsNullOrWhiteSpace($root)) { continue }
        foreach ($rel in @("RevitBatchProcessor", "BatchRvt", "Revit Batch Processor")) { $dirs.Add((Join-Path $root $rel)) }
    }
    $seen = @{}
    foreach ($dir in $dirs) {
        if ([string]::IsNullOrWhiteSpace($dir)) { continue }
        try { $full = [System.IO.Path]::GetFullPath($dir) } catch { continue }
        if ($seen.ContainsKey($full)) { continue }
        $seen[$full] = $true
        foreach ($name in $names) {
            $exe = Join-Path $full $name
            if (Test-Path -LiteralPath $exe) { return $exe }
        }
    }
    return $null
}

function Enable-RsnForRbp([string]$Exe) {
    $patch = Join-Path $ServiceDir "rbp_rsn_patch\apply_rsn_support.ps1"
    if (-not (Test-Path -LiteralPath $patch)) { return $false }
    $scriptsDir = Join-Path (Split-Path -Parent $Exe) "Scripts"
    if (Test-Path -LiteralPath (Join-Path $scriptsDir "revit_file_list.py")) {
        & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $patch -ScriptsPath $scriptsDir
    }
    else {
        & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $patch
    }
    return ($LASTEXITCODE -eq 0)
}


# ---------- список моделей ----------
$tokens = New-Object System.Collections.Generic.List[string]
foreach ($m in @($Models)) { foreach ($piece in @(([string]$m).Split(';'))) { $t = $piece.Trim().Trim('"'); if ($t) { $tokens.Add($t) } } }
if ($tokens.Count -eq 0) {
    $typed = Read-Host "Путь к rvt, папке или txt-списку (несколько — через «;»)"
    foreach ($piece in @(([string]$typed).Split(';'))) { $t = $piece.Trim().Trim('"'); if ($t) { $tokens.Add($t) } }
}
$skipped = [pscustomobject]@{ Backup = 0; BackupFolders = 0; Summary = 0; Expanded = $false }
$raw = New-Object System.Collections.Generic.List[string]
foreach ($token in $tokens) { Add-DumpToken $token $raw $skipped }
$dupes = New-Object System.Collections.Generic.List[string]
$list = Select-DumpUniqueModels $raw $dupes
$ready = New-Object System.Collections.Generic.List[string]
Write-Host ""
Write-Host "Выгрузка фильтров спецификаций (модели не меняются)" -ForegroundColor Cyan
foreach ($p in $list) {
    if (Test-ModelReady $p) { $ready.Add($p); Write-Host ("  {0}" -f $p) }
    else { Write-Host ("  ПРОПУЩЕН, файл не найден: {0}" -f $p) -ForegroundColor Yellow }
}
if ($skipped.Backup -gt 0) { Write-Host ("  пропущены резервные копии Revit: {0}" -f $skipped.Backup) -ForegroundColor DarkGray }
if ($skipped.Summary -gt 0) { Write-Host ("  пропущены сводные (SVOD): {0}" -f $skipped.Summary) -ForegroundColor DarkGray }
foreach ($w in $dupes) { Write-Host ("  дубль: {0}" -f $w) -ForegroundColor Yellow }
Write-Host ("Найдено моделей: {0}" -f $ready.Count) -ForegroundColor Cyan
Write-Host ("Пропущено папок резервных копий: {0}" -f $skipped.BackupFolders) -ForegroundColor Cyan
if ($ready.Count -eq 0) { Stop-Here "Нет ни одного доступного файла." }
if ($skipped.Expanded -and -not $NoPause) {
    $answer = Read-Host "Выгрузить фильтры этих моделей? (д/н)"
    if ($answer -notmatch '^(?i)\s*(д|да|y|yes)\s*$') { Write-Host "Отменено."; exit 0 }
}

# ---------- Revit Batch Processor ----------
$exe = $RbpExe
if (-not $exe) { $exe = Find-BatchRvt }
if (-not $exe) { Stop-Here "Не найден Revit Batch Processor (BatchRvt.exe). Установите RBP и откройте его один раз." }
$batch = Join-Path (Split-Path -Parent $exe) "BatchRvt.exe"
if (-not $RbpExe -and (Test-Path -LiteralPath $batch)) { $exe = $batch }
if (@($ready | Where-Object { $_ -match '^(?i)RSN://' }).Count -gt 0) {
    Write-Host "Есть пути RSN:// — подключаю поддержку в Revit Batch Processor."
    if (-not (Enable-RsnForRbp $exe)) { Stop-Here "Поддержка RSN:// не подключилась (rbp_rsn_patch). Запуск остановлен." }
}
$runId = Get-Date -Format "yyyyMMdd_HHmmss"
# Папка вывода — всегда отчеты\операции\дамп_фильтров_<время> в папке инструмента; путь файла/модели вместо папки не принимается.
if ($OutDir -and ((Test-Path -LiteralPath $OutDir -PathType Leaf) -or $OutDir -match '(?i)\.(rvt|rte|rfa|rft)$')) {
    Write-Host ("  -OutDir указывает на файл, игнорирую: {0}" -f $OutDir) -ForegroundColor Yellow
    $OutDir = ""
}
$layout = (Initialize-ReportLayout (Get-ReportRootDir $ToolDir) -AppDir $AppDir).Layout
if (-not $OutDir) { $OutDir = Join-Path $layout.Ops ("дамп_фильтров_" + $runId) }
if (-not (Test-Path -LiteralPath $OutDir)) { New-Item -ItemType Directory -Path $OutDir -Force | Out-Null }
if (-not (Test-Path -LiteralPath $AppDir)) { New-Item -ItemType Directory -Path $AppDir -Force | Out-Null }
$logDir = $layout.Logs
if ($logDir -and -not (Test-Path -LiteralPath $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
$bom = New-Object System.Text.UTF8Encoding $true
$cfgLines = @(("report_dir={0}" -f $OutDir), ("name_contains={0}" -f $NameContains))
if ($logDir) { $cfgLines += @("log_folder={0}" -f $logDir) }
$cfg = $cfgLines -join "`r`n"
[System.IO.File]::WriteAllText((Join-Path $AppDir "dump_filters.cfg"), $cfg + "`r`n", $bom)

# модели разных версий Revit — отдельными проходами (RBP запускает одну версию Revit за раз)
. (Join-Path $ToolDir "revit_versions.ps1")
$serverYears = Get-BvsServerYears (Join-Path (Get-BvsRootDir $ToolDir) "servers.cfg")
$groups = @{}
foreach ($m in $ready) {
    $yr = Get-ModelRevitYear $m $serverYears
    if ($null -eq $yr.Year) { Stop-Here ("Не удалось определить версию Revit: {0} ({1})." -f (Get-RevitLeaf $m), $yr.Why) }
    $y = [string]$yr.Year
    if (-not $groups.ContainsKey($y)) { $groups[$y] = New-Object System.Collections.Generic.List[string] }
    $groups[$y].Add($m)
}
$exitCode = 0
$pass = 0
foreach ($year in @($groups.Keys | Sort-Object)) {
    $pass++
    $listFile = Join-Path $AppDir ("rvt_list_filter_dump_{0}.txt" -f $pass)
    [System.IO.File]::WriteAllLines($listFile, [string[]]$groups[$year].ToArray(), (New-Object System.Text.UTF8Encoding $false))
    $argText = '--task_script "{0}" --file_list "{1}" --create_new_local --worksets close_all' -f $TaskScript, $listFile
    $argText += (' --revit_version {0}' -f $year)          # всегда
    if ($logDir) { $argText += (' --log_folder "{0}"' -f $logDir) }   # только этот запуск BVS; настройки RBP не меняем
    Write-Host ""
    Write-Host ("Revit {0}: моделей {1}. Revit откроется поверх, окно не закрывайте..." -f $(if ($year) { $year } else { "(по умолчанию)" }), $groups[$year].Count) -ForegroundColor Cyan
    $env:BVS_DUMP_OUT = $OutDir
    $startArgs = @{ FilePath = $exe; ArgumentList = $argText; WorkingDirectory = (Split-Path -Parent $exe); PassThru = $true }
    if ($env:OS -eq "Windows_NT") { $startArgs.WindowStyle = "Hidden" }
    $proc = Start-Process @startArgs
    while (-not $proc.HasExited) {
        Start-Sleep -Seconds 3
        $done = @(Get-ChildItem -LiteralPath $OutDir -File -Filter "фильтры_*.txt" -ErrorAction SilentlyContinue).Count
        Write-Host ("`r  готово файлов: {0}   " -f $done) -NoNewline
    }
    Write-Host ""
    if ($proc.ExitCode -ne 0) { Write-Host ("Revit Batch Processor завершился с кодом {0}." -f $proc.ExitCode) -ForegroundColor Yellow; $exitCode = 2 }
}

# ---------- итог ----------
Write-Host ""
$files = @(Get-ChildItem -LiteralPath $OutDir -File -Filter "фильтры_*.txt" -ErrorAction SilentlyContinue | Sort-Object Name)
foreach ($f in $files) {
    $rows = @(Get-Content -LiteralPath $f.FullName -Encoding UTF8 | Select-Object -Skip 1 | Where-Object { $_.Trim() })
    $specNames = @($rows | ForEach-Object { ($_ -split "`t")[0].Trim() } | Where-Object { $_ -and $_ -ne "(нет фильтров)" } | Select-Object -Unique)
    $specFile = Join-Path $OutDir ("спеки_" + $f.Name.Substring(("фильтры_").Length))
    [System.IO.File]::WriteAllLines($specFile, [string[]]$specNames, $bom)
    Write-Host ("  {0}: спецификаций {1}, строк фильтров {2}; список имён: {3}" -f $f.Name, $specNames.Count, $rows.Count, (Split-Path -Leaf $specFile)) -ForegroundColor Green
}
if ($files.Count -lt $ready.Count) {
    Write-Host ("Файлов {0}, а моделей {1}: не все модели открылись, смотрите журнал BatchRvt." -f $files.Count, $ready.Count) -ForegroundColor Yellow
    $exitCode = 2
}
Write-Host ("Папка: {0}" -f $OutDir)
Write-Host "Сравнить две выгрузки: служебное\разработка\инструменты\Сравнить фильтры спек.bat (перетащите на него два файла: «было» и «стало»)."
if (-not $NoPause) { [void](Read-Host "Нажмите Enter") }
exit $exitCode
