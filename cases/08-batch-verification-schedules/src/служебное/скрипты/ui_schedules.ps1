# UTF-8 WITH BOM required (PowerShell 5.1).
# Окно «Проверочные спецификации». Запуск: Проверочные спецификации.vbs

param([switch]$SelfCheck)     # -SelfCheck: проверка без окна и без Revit — находит ли скрипт все свои файлы от своего расположения

$ErrorActionPreference = "Continue"
$ToolDir = Split-Path -Parent $MyInvocation.MyCommand.Path
if ($SelfCheck) {
    # Папка с кодом и корень отчётов определяются только от расположения этого файла (<папка с кодом>\отчеты)
    $required = @(
        "excel_isolated.ps1", "report_xlsx.ps1", "bvs_assignment_guard.ps1", "assignment_export_stream.ps1", "assignment_partial.ps1", "assignment_partial_ui.ps1",
        "assignment_xlsx.ps1", "assignment_ledger.ps1", "build_assignment.ps1", "report_layout.ps1", "folder_dialog.ps1",
        "project_config.ps1", "bvs_export_result.ps1", "project_dist_ui.ps1", "push_google.ps1", "google_push_cfg.ps1", "google_payload_lib.ps1",
        "insert_schedules.py", "export_schedules.py", "set_assembly_code.py", "read_template_specs.py", "..\lib\DocumentFormat.OpenXml.dll",
        "revit_versions.ps1", "выбрать_с_сервера.ps1", "..\..\servers.cfg", "..\choose_models.ps1", "..\rbp_rsn_patch\apply_rsn_support.ps1",
        "..\google_push\google_push.cfg.example"
    )
    $missing = @($required | Where-Object { -not (Test-Path -LiteralPath (Join-Path $ToolDir $_)) })
    $reportRoot = "?"
    $rl = Join-Path $ToolDir "report_layout.ps1"
    if (Test-Path -LiteralPath $rl) { . $rl; $reportRoot = Get-ReportRootDir $ToolDir }
    Write-Output ("TOOLDIR=" + $ToolDir)
    Write-Output ("REPORTDIR=" + $reportRoot)
    Write-Output ("REQUIRED=" + $required.Count)
    foreach ($m in $missing) { Write-Output ("MISSING=" + $m) }
    if ($missing.Count -gt 0) { exit 3 }
    Write-Output "SELFCHECK OK"
    exit 0
}
Add-Type -AssemblyName System.Windows.Forms
# Пояснения к полям — всплывающими подсказками (ToolTip), а не абзацами текста в окне
$script:UiTips = $null
function Set-UiTip($Control, [string]$Text) {
    if ($null -eq $script:UiTips) { $script:UiTips = New-Object System.Windows.Forms.ToolTip; $script:UiTips.AutoPopDelay = 15000 }
    $script:UiTips.SetToolTip($Control, $Text)
}
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()
. (Join-Path $ToolDir "excel_isolated.ps1")          # запасное чтение xlsx для Google (push_google.ps1), если нет payload; отчёт прогона Excel не использует
. (Join-Path $ToolDir "report_xlsx.ps1")             # отчёт прогона в xlsx без Excel (zip + XML)
. (Join-Path $ToolDir "bvs_assignment_guard.ps1")
# режим «Выборочно»: списки моделей/спецификаций и буфер. Нет файлов — режим просто недоступен, «Всё» работает как раньше.
$script:PartialUiReady = $false
$__pFiles = @("assignment_export_stream.ps1", "assignment_partial.ps1", "assignment_partial_ui.ps1")
if (@($__pFiles | Where-Object { -not (Test-Path -LiteralPath (Join-Path $ToolDir $_)) }).Count -eq 0) {
    foreach ($__f in $__pFiles) { . (Join-Path $ToolDir $__f) }
    $script:PartialUiReady = $true
}
$TaskScript = Join-Path $ToolDir "insert_schedules.py"
# «Распределение»: единственный источник наборов спецификаций (вставка, «Сводная», «Разделы», «Выборочно», проверка по эталону)
. (Join-Path $ToolDir "project_config.ps1")
. (Join-Path $ToolDir "bvs_export_result.ps1")      # итог выгрузки из сводной: export.txt (прежние выгрузки — manifest.txt / export_result.txt / catalog.txt)
. (Join-Path $ToolDir "google_push_cfg.ps1")         # имя проекта из пути настройки, Google-адрес проекта (без него проект не выбирается в списке)
$script:Project = $null
$script:ProjectPath = ""
# раскладка: отчеты\ — итоговые xlsx, отчеты\история\ — постоянные файлы, отчеты\служебные\<время>\ — файлы прогона (автоочистка), отчеты\операции\ — отчёты служебных .bat
. (Join-Path $ToolDir "report_layout.ps1")
. (Join-Path $ToolDir "revit_versions.ps1")          # версия Revit по моделям (BasicFileInfo / servers.cfg), не по имени
$ServersCfgPath = Join-Path (Get-BvsRootDir $ToolDir) "servers.cfg"
. (Join-Path $ToolDir "folder_dialog.ps1")          # Show-FolderPathDialog: единственное окно выбора папки
$ReportDir = Get-ReportRootDir $ToolDir
if (-not (Test-Path -LiteralPath $ReportDir)) {
    New-Item -ItemType Directory -Path $ReportDir -Force | Out-Null
}
$script:Layout = Get-ReportLayout $ReportDir
$script:AuxKeepDays = 30
$script:LastOpenPath = ""

$AppDir = Join-Path $env:LOCALAPPDATA "BatchRvt\batch_verification_schedules"
if (-not (Test-Path -LiteralPath $AppDir)) {
    New-Item -ItemType Directory -Path $AppDir -Force | Out-Null
}
$SettingsPath = Join-Path $AppDir "ui_settings.txt"
$ListPath = Join-Path $AppDir "rvt_list.txt"
$CfgPath = Join-Path $AppDir "run.cfg"

$script:Models = New-Object System.Collections.Generic.List[string]
$script:SummaryDone = $false        # «В сводную» выполнена (запоминается; сбрасывается при «Загрузить спецификации и модели»). Кнопку «В разделы» не включает и не выключает.
$script:Run = @{ Process = $null; Summary = ""; ReportDir = ""; Partial = $false }
$script:LayoutNote = ""
# выбор для «Выборочно»: имена отмеченных моделей/спецификаций (хэш); $null — ещё не выбирали
$script:AsgSel = @{ Models = $null; Specs = $null; Pairs = $null }   # Pairs — «модель|спека» (новый формат asgpair=)
# «Выборочно»: полный список спецификаций, отметки пар (модель -> спека -> да/нет), наборы спецификаций моделей и причины, где набор не определён
$script:AsgAllSpecs = @(); $script:AsgPicks = @{}; $script:AsgModelSets = @{}; $script:AsgModelWhy = @{}; $script:AsgAllModels = @(); $script:AsgShown = @(); $script:AsgPrevModels = @()
$script:AsgSpecBusy = $false
$script:AsgBusy = $false

function Read-Settings {
    $map = @{
        Source = ""; Pilot = "0"; ProjStage = "П"; ProjFile = ""; ProjModelsDir = ""; ProjRsn = @(); PrefixFrom = ""; PrefixTo = ""; Summary = ""; Google = ""
        InsertModeSummary = "initial"; InsertModeSections = "initial"
        DynDays = "7"; DynCustom = "0"; DynFrom = ""; DynTo = ""; CollapseDay = "0"; SummaryDone = "0"
        AsgScope = "all"; AsgModels = $null; AsgSpecs = $null; AsgPairs = $null    # $null — выбор ещё не делали (отмечено всё)
        LastProject = ""; Migrated = $false
    }
    $hasSummaryMode = $false
    $hasSectionsMode = $false
    $legacyInsertMode = ""
    if (Test-Path -LiteralPath $SettingsPath) {
        foreach ($line in @(Get-Content -LiteralPath $SettingsPath -Encoding UTF8)) {
            if ($line -match '^(?i)lastproject=(.*)$') { $map.LastProject = $Matches[1].Trim() }
            elseif ($line -match '^(?i)migrated=1\s*$') { $map.Migrated = $true }
            elseif ($line -match '^(?i)source=(.*)$') { $map.Source = $Matches[1] }
            elseif ($line -match '^(?i)pilot=(.*)$') { $map.Pilot = $Matches[1] }
            elseif ($line -match '^(?i)projstage=(.*)$') { $map.ProjStage = $Matches[1] }
            elseif ($line -match '^(?i)projfile=(.*)$') { $map.ProjFile = $Matches[1] }
            elseif ($line -match '^(?i)projmodelsdir=(.*)$') { $map.ProjModelsDir = $Matches[1] }
            elseif ($line -match '^(?i)projrsn=(RSN://.+)$') { $map.ProjRsn += $Matches[1].Trim() }
            elseif ($line -match '^(?i)prefixfrom=(.*)$') { $map.PrefixFrom = $Matches[1] }
            elseif ($line -match '^(?i)prefixto=(.*)$') { $map.PrefixTo = $Matches[1] }
            elseif ($line -match '^(?i)summary=(.*)$') { $map.Summary = $Matches[1] }
            elseif ($line -match '^(?i)googleurl=(.*)$') { $map.Google = $Matches[1] }
            elseif ($line -match '^(?i)dyndays=(.*)$') { $map.DynDays = $Matches[1] }
            elseif ($line -match '^(?i)dyncustom=(.*)$') { $map.DynCustom = $Matches[1] }
            elseif ($line -match '^(?i)collapseday=(.*)$') { $map.CollapseDay = $Matches[1] }
            elseif ($line -match '^(?i)summarydone=(.*)$') { $map.SummaryDone = $Matches[1].Trim() }
            elseif ($line -match '^(?i)auxkeepdays=\s*(\d+)\s*$') { $n = [int]$Matches[1]; if ($n -ge 1 -and $n -le 3650) { $script:AuxKeepDays = $n } }
            elseif ($line -match '^(?i)asgscope=(.*)$') { $map.AsgScope = $Matches[1] }
            elseif ($line -match '^(?i)asgmodels=(.*)$') { $map.AsgModels = $Matches[1] }
            elseif ($line -match '^(?i)asgspecs=(.*)$') { $map.AsgSpecs = $Matches[1] }
            elseif ($line -match '^(?i)asgpairsfmt=') { if ($null -eq $map.AsgPairs) { $map.AsgPairs = @{} } }
            elseif ($line -match '^(?i)asgpair=(.+)$') { if ($null -eq $map.AsgPairs) { $map.AsgPairs = @{} }; $map.AsgPairs[$Matches[1].Trim()] = $true }
            elseif ($line -match '^(?i)dynfrom=(.*)$') { $map.DynFrom = $Matches[1] }
            elseif ($line -match '^(?i)dynto=(.*)$') { $map.DynTo = $Matches[1] }
            elseif ($line -match '^(?i)insertmodesummary=(.*)$') {
                $map.InsertModeSummary = $Matches[1]
                $hasSummaryMode = $true
            }
            elseif ($line -match '^(?i)insertmodesections=(.*)$') {
                $map.InsertModeSections = $Matches[1]
                $hasSectionsMode = $true
            }
            elseif ($line -match '^(?i)insertmode=(.*)$') { $legacyInsertMode = $Matches[1] }
        }
    }
    if ($legacyInsertMode) {
        if (-not $hasSummaryMode) { $map.InsertModeSummary = $legacyInsertMode }
        if (-not $hasSectionsMode) { $map.InsertModeSections = $legacyInsertMode }
    }
    return $map
}

# ui_settings.txt: только последний открытый проект и общие параметры окна. Пути, Google, режим вставки, «Выборочно» и т. п. — в настройке проекта
# (Save-DistProject), одно значение не хранится в двух местах. migrated=1 — старые значения уже перенесены в проекты (Invoke-LegacySettingsMigration).
function Save-Settings {
    if ($script:PartialUiReady -and $ui.AsgTree) { Save-AsgSelection }
    [void](Save-DistProject)
    $text = @(
        ("lastproject={0}" -f (Get-CurrentProjectName)),
        "migrated=1",
        ("dyndays={0}" -f (Get-DynamicsDaysValue)),
        ("dyncustom={0}" -f ($(if ($ui.DynCustom.Checked) { "1" } else { "0" }))),
        ("collapseday={0}" -f ($(if ($ui.CollapseDay.Checked) { "1" } else { "0" }))),
        ("auxkeepdays={0}" -f $script:AuxKeepDays),
        ("dynfrom={0}" -f $ui.DynFrom.Value.ToString("yyyy-MM-dd")),
        ("dynto={0}" -f $ui.DynTo.Value.ToString("yyyy-MM-dd"))
    )
    [System.IO.File]::WriteAllText($SettingsPath, ($text -join "`r`n"), (New-Object System.Text.UTF8Encoding $true))
}

function Get-CurrentProjectName {
    if (-not $script:ProjectPath) { return "" }
    return (Get-ProjectNameFromConfigPath $script:ProjectPath)
}

# Прежний список моделей разделов (rvt_list.txt, шаг «Разделы»): только читается — из него берётся источник моделей, если он ещё не задан.
# Файл не перезаписывается и не удаляется.
function Read-LegacyModelList {
    $out = New-Object System.Collections.Generic.List[string]
    if (Test-Path -LiteralPath $ListPath) {
        foreach ($line in @(Get-Content -LiteralPath $ListPath -Encoding UTF8)) {
            $t = $line.Trim()
            if ($t -and -not $t.StartsWith("#")) { [void]$out.Add($t) }
        }
    }
    return $out.ToArray()
}

# «Распределить» и «Посчитать элементы» — после загрузки. «В сводную» и «В разделы» — после распределения, независимо друг от друга.
# Путь сводной и SummaryDone кнопку «В разделы» не выключают. Во время запуска всё серое.
$script:ActionsBusy = $false
function Test-DistLoaded { return ($null -ne $script:Project -and $script:Project.Specs.Count -gt 0 -and $script:Project.Models.Count -gt 0) }
function Update-ActionState {
    $busy = [bool]$script:ActionsBusy
    $loaded = Test-DistLoaded
    $distributed = ($loaded -and (Test-ProjectDistributed $script:Project))
    $hasProject = [bool]$script:ProjectPath          # без проекта кнопки шагов неактивны: негде хранить пути, выбор и журнал
    $on = @{
        Tabs = (-not $busy); Run = ((-not $busy) -and $hasProject); DistLoad = ((-not $busy) -and $hasProject); DynBuild = (-not $busy)
        ProjectBox = (-not $busy); ProjectNew = (-not $busy)
        DistAssign = ((-not $busy) -and $loaded)
        DistCount = ((-not $busy) -and $loaded)
        DistFromSummary = ((-not $busy) -and $loaded)
        RunSummary = ((-not $busy) -and $distributed)
        RunSections = ((-not $busy) -and $distributed)
        RunClassifier = ((-not $busy) -and $loaded)
    }
    foreach ($key in @($on.Keys)) {
        if ($ui.ContainsKey($key) -and $ui[$key]) { $ui[$key].Enabled = [bool]$on[$key] }
    }
}
function Set-ActionsEnabled([bool]$On) {
    $script:ActionsBusy = (-not $On)
    Update-ActionState
}

function Find-BatchRvt {
    $names = @("BatchRvt.exe", "BatchRvtGUI.exe")
    $dirs = New-Object System.Collections.Generic.List[string]
    Get-Process -ErrorAction SilentlyContinue |
        Where-Object { $_.ProcessName -match '^(?i)BatchRvt(GUI)?$|RevitBatchProcessor$' } |
        ForEach-Object {
            try {
                $exeDir = Split-Path -Parent $_.MainModule.FileName
                if ($exeDir) { $dirs.Add($exeDir) }
            }
            catch { }
        }
    foreach ($root in @($env:LOCALAPPDATA, $env:ProgramFiles, ${env:ProgramFiles(x86)}, $env:ProgramW6432)) {
        if ([string]::IsNullOrWhiteSpace($root)) { continue }
        foreach ($rel in @("RevitBatchProcessor", "BatchRvt", "Revit Batch Processor")) {
            $dirs.Add((Join-Path $root $rel))
        }
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

function Test-ModelReady([string]$Path) {
    if ([string]::IsNullOrWhiteSpace($Path)) { return $false }
    if ($Path.Trim() -match '^(?i)RSN://.+\.rvt$') { return $true }
    return (Test-Path -LiteralPath $Path)
}

function Get-ModelLeaf([string]$Path) {
    if ($Path -match '^(?i)RSN://') {
        $text = $Path.Trim().TrimEnd("/")
        $i = $text.LastIndexOf("/")
        if ($i -ge 0 -and $i -lt ($text.Length - 1)) { return $text.Substring($i + 1) }
        return $text
    }
    return [System.IO.Path]::GetFileName($Path)
}

function Enable-RsnForRbp([string]$RbpExe) {
    $patch = $null
    foreach ($candidate in @(
            (Join-Path (Split-Path -Parent $ToolDir) "rbp_rsn_patch\apply_rsn_support.ps1"),
            (Join-Path (Split-Path -Parent (Get-BvsRootDir $ToolDir)) "batch_compact_save\служебное\rbp_rsn_patch\apply_rsn_support.ps1")
        )) {
        if (Test-Path -LiteralPath $candidate) { $patch = $candidate; break }
    }
    if (-not $patch) { return $false }
    $scriptsDir = Join-Path (Split-Path -Parent $RbpExe) "Scripts"
    if (Test-Path -LiteralPath (Join-Path $scriptsDir "revit_file_list.py")) {
        & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $patch -ScriptsPath $scriptsDir
    }
    else {
        & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $patch
    }
    return ($LASTEXITCODE -eq 0)
}

function Get-ServerModels {
    $picker = Join-Path $ToolDir "выбрать_с_сервера.ps1"
    if (-not (Test-Path -LiteralPath $picker)) {
        [System.Windows.Forms.MessageBox]::Show($form, "Нет файла выбрать_с_сервера.ps1.", "Проверочные спецификации") | Out-Null
        return @()
    }
    $out = Join-Path $env:TEMP ("verif_rsn_" + [guid]::NewGuid().ToString("N") + ".txt")
    $arg = "-NoProfile -WindowStyle Hidden -STA -ExecutionPolicy Bypass -File `"$picker`" -OutFile `"$out`""
    $ui.Status.Text = "Список Revit Server открыт отдельно. Выберите модель в том окне."
    Set-ActionsEnabled $false
    [System.Windows.Forms.Application]::DoEvents()
    $proc = Start-Process -FilePath "powershell.exe" -ArgumentList $arg -PassThru -WindowStyle Hidden
    try {
        while ($proc -and -not $proc.HasExited) {
            if ($form.IsDisposed) { break }
            [System.Windows.Forms.Application]::DoEvents()
            Start-Sleep -Milliseconds 100
        }
    }
    finally {
        Set-ActionsEnabled $true
    }
    $lines = @()
    if (Test-Path -LiteralPath $out) {
        $lines = @(Get-Content -LiteralPath $out -Encoding UTF8 | ForEach-Object { $_.Trim() } | Where-Object { $_ -match '^(?i)RSN://' })
        Remove-Item -LiteralPath $out -Force -ErrorAction SilentlyContinue
    }
    if ($proc.ExitCode -ne 0 -and $proc.ExitCode -ne 2 -and $lines.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show($form, "Список с Revit Server не открылся.", "Проверочные спецификации") | Out-Null
    }
    return $lines
}

function Set-OneServerModel($Box) {
    $picked = @(Get-ServerModels)
    if ($picked.Count -eq 0) { return }
    if ($picked.Count -gt 1) {
        [System.Windows.Forms.MessageBox]::Show($form, "Для сводной нужна одна модель.", "Проверочные спецификации") | Out-Null
        return
    }
    $Box.Text = [string]$picked[0]
    Save-Settings
}

# Первый запуск после появления проектов: старые общие настройки окна переносятся в проект (Invoke-LegacySettingsMigration). Исходный ui_settings.txt
# копируется в ui_settings.до_проектов.txt (один раз), прежний список моделей разделов (rvt_list.txt) даёт папку моделей, если её нет. Возвращает имя проекта или ""
function Invoke-StartupSettingsMigration($Saved) {
    if ($Saved.Migrated -or -not (Test-Path -LiteralPath $SettingsPath)) { return "" }
    $backup = Join-Path (Split-Path -Parent $SettingsPath) "ui_settings.до_проектов.txt"
    if (-not (Test-Path -LiteralPath $backup)) { try { Copy-Item -LiteralPath $SettingsPath -Destination $backup -ErrorAction Stop } catch { } }
    $legacy = @{}
    foreach ($k in @($Saved.Keys)) { $legacy[$k] = $Saved[$k] }
    $legacy.SummaryDone = ($Saved.SummaryDone -eq "1")
    $legacy.InsertMode = Get-CombinedInsertMode $Saved.InsertModeSummary $Saved.InsertModeSections
    $legacy.AsgModels = $null
    if ($null -ne $Saved.AsgModels) {
        $set = @{}
        foreach ($x in @(([string]$Saved.AsgModels).Split([char]'|'))) { if ($x.Trim()) { $set[$x.Trim()] = $true } }
        $legacy.AsgModels = $set
    }
    if (-not $legacy.ProjModelsDir) {
        $fromList = Get-LegacyModelsSource @(Read-LegacyModelList)
        $legacy.ProjModelsDir = $fromList.Source
        if (@($fromList.Rsn).Count -gt 0 -and @($legacy.ProjRsn).Count -eq 0) { $legacy.ProjRsn = @($fromList.Rsn) }
    }
    $r = Invoke-LegacySettingsMigration $legacy $ReportDir
    return [string]$r.Name
}

# При смене проекта списки «Выборочно» относятся к прежнему проекту: сбрасываются, чтобы их отметки не попали в настройку нового
function Reset-PartialLists {
    if (-not $ui.AsgTree) { return }
    $ui.AsgTree.Nodes.Clear()
    $script:AsgAllSpecs = @(); $script:AsgPicks = @{}; $script:AsgModelSets = @{}; $script:AsgModelWhy = @{}; $script:AsgAllModels = @(); $script:AsgShown = @(); $script:AsgPrevModels = @()
    $ui.AsgSpecs.Items.Clear()
}

# Режим вставки — один переключатель для «В сводную» и «В разделы». «initial» из старых настроек читается как «Добавить недостающие» (в модели без пакета это то же самое)
$script:InsertModeValues = @("add_diff", "replace", "initial")
$script:InsertModeLabels = @("Добавить недостающие", "Заменить все")

function Normalize-InsertModeValue([string]$Value) {
    $v = ($Value -as [string]).Trim().ToLower()
    if (-not $v) { return "initial" }
    if ($v -in @("first", "first_load")) { return "initial" }
    if ($v -in @("add_missing", "add_new_tails", "add_all")) { return "add_diff" }
    if ($script:InsertModeValues -contains $v) { return $v }
    return "initial"
}

# Выпадающий список режима вставки рядом с «В разделы»; значения — $script:InsertModeValues (add_diff, replace)
function New-InsertModeBox {
    $box = New-Object System.Windows.Forms.ComboBox
    $box.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $box.Width = 190
    foreach ($label in $script:InsertModeLabels) { [void]$box.Items.Add($label) }
    $box.SelectedIndex = 0
    # пояснение — всплывающей подсказкой, не строкой в окне
    Set-UiTip $box "Совпадение по номеру спецификации. «Заменить все» удаляет старые спецификации списка и вставляет заново. Посторонние спецификации не удаляются. Режим один для «В сводную» и «В разделы»."
    return $box
}

# Старые настройки хранили режим отдельно для сводной и для разделов; единый переключатель — «Заменить все» только если он был в обоих
function Get-CombinedInsertMode([string]$Summary, [string]$Sections) {
    if ((Normalize-InsertModeValue $Summary) -eq "replace" -and (Normalize-InsertModeValue $Sections) -eq "replace") { return "replace" }
    return "add_diff"
}

function Get-InsertModeValue {
    param([string]$Scope = "")          # прежний параметр (сводная/разделы) оставлен для совместимости: переключатель один
    $box = $ui.InsertModeBox
    if (-not $box) { return "initial" }
    $i = [int]$box.SelectedIndex
    if ($i -ge 0 -and $i -lt $script:InsertModeLabels.Count) { return [string]$script:InsertModeValues[$i] }
    return "initial"
}

# «initial» из старых настроек читается как «Добавить недостающие» (первый пункт списка)
function Set-InsertModeBox {
    param($Box, [string]$Value)
    if (-not $Box) { return }
    $v = Normalize-InsertModeValue $Value
    $i = [Array]::IndexOf([string[]]@($script:InsertModeValues), $v)
    if ($i -lt 0 -or $i -ge $Box.Items.Count) { $i = 0 }
    $Box.SelectedIndex = $i
}

function Get-ProjectCode([string]$Summary) {
    # открыт проект — код истории и журнала это имя проекта (у каждого проекта свои история_<код>.txt и задание_журнал_<код>.txt); иначе по имени сводной
    if ($script:ProjectPath) { return (Get-CurrentProjectName) }
    $leaf = "PROJ"
    if (-not [string]::IsNullOrWhiteSpace($Summary)) {
        $leaf = Get-ModelLeaf $Summary
        $leaf = [System.IO.Path]::GetFileNameWithoutExtension($leaf)
    }
    if ($leaf -match '^([A-Z0-9]+)_IOS_') { return $Matches[1] }
    if ($leaf -match '^([A-Z0-9]+)_') { return $Matches[1] }
    return "PROJ"
}

# Отчёты проекта лежат в отчеты\<код проекта>\ (PROJ, SITE...); проект не выбран — в корне отчетов, как раньше.
# Имя папки очищается так же, как имя файла настройки проекта; папка создаётся при первой записи.
function Get-ProjectReportDir([string]$Code) {
    if (-not $script:ProjectPath) { return $ReportDir }
    $c = ([string]$Code).Trim()
    if (-not $c) { return $ReportDir }
    $c = ($c -replace '[\\/:*?"<>|]', '_')
    $dir = Join-Path $ReportDir $c
    if (-not (Test-Path -LiteralPath $dir)) {
        try { New-Item -ItemType Directory -Path $dir -Force | Out-Null } catch { return $ReportDir }
    }
    return $dir
}

function Get-AssignmentPath([string]$Summary) {
    $stamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
    $code = Get-ProjectCode $Summary
    $outDir = Get-ProjectReportDir $code
    $path = Join-Path $outDir ("{0}_{1}.xlsx" -f $stamp, $code)
    $n = 2
    while (Test-Path -LiteralPath $path) {
        $path = Join-Path $outDir ("{0}_{1}_{2}.xlsx" -f $stamp, $code, $n)
        $n++
    }
    return $path
}

function Get-DynamicsOnlyPath([string]$Summary) {
    $stamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
    $code = Get-ProjectCode $Summary
    $outDir = Get-ProjectReportDir $code
    $path = Join-Path $outDir ("{0}_{1}_динамика.xlsx" -f $stamp, $code)
    $n = 2
    while (Test-Path -LiteralPath $path) {
        $path = Join-Path $outDir ("{0}_{1}_динамика_{2}.xlsx" -f $stamp, $code, $n)
        $n++
    }
    return $path
}

# Журнал замечаний — на проект (раньше один задание_журнал.txt на все): задание_журнал_<код>.txt рядом с историей
function Get-LedgerPath([string]$Summary) {
    return (Join-Path $script:Layout.History ("задание_журнал_" + (Get-ProjectCode $Summary) + ".txt"))
}
function Get-HistoryPath([string]$Summary) {
    $code = Get-ProjectCode $Summary
    return (Join-Path $script:Layout.History ("история_" + $code + ".txt"))
}

function Get-DynamicsDaysValue {
    if ($ui.DynDays -and $null -ne $ui.DynDays.SelectedItem) {
        $text = [string]$ui.DynDays.SelectedItem
        if ($text -match '^(\d+)') { return $Matches[1] }
    }
    return "7"
}

function Get-DynamicsPeriodFromUi {
    if ($ui.DynCustom -and $ui.DynCustom.Checked) {
        $from = $ui.DynFrom.Value.Date
        $to = $ui.DynTo.Value.Date
        if ($from -gt $to) {
            $swap = $from
            $from = $to
            $to = $swap
        }
        return @{
            From = $from.ToString("yyyy-MM-dd")
            To   = $to.ToString("yyyy-MM-dd")
        }
    }
    $days = 7
    if ($ui.DynDays -and $null -ne $ui.DynDays.SelectedItem) {
        $text = [string]$ui.DynDays.SelectedItem
        if ($text -match '^(\d+)') { $days = [int]$Matches[1] }
    }
    $to = (Get-Date).Date
    $from = $to.AddDays(-($days - 1))
    return @{
        From = $from.ToString("yyyy-MM-dd")
        To   = $to.ToString("yyyy-MM-dd")
    }
}

# Журнал отправок в Google (без секрета): когда, куда (проект + конец адреса), результат
function Add-GooglePushLog([string]$Line) {
    try {
        $dir = Get-BvsToolLogDir $AppDir
        if (-not $dir) { return }
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        $log = Join-Path $dir "google_push.log"
        [System.IO.File]::AppendAllText($log, ((Get-Date -Format "yyyy-MM-dd HH:mm:ss") + "  " + $Line + "`r`n"), (New-Object System.Text.UTF8Encoding $true))
        Trim-BvsTextFile $log
    } catch { }
}

# Куда пойдёт отправка и нужно ли спросить пользователя. Возвращает @{ Go; Target; Text }:
# у проекта свой адрес — отправляем; своего нет, есть общий google_push.cfg — вопрос (по умолчанию «Нет»); адреса нет нигде — не отправляем
# Подключение Google на вкладке «Задание»: адрес и секрет читаются молча из отчеты\проект\google_<проект>.cfg и нигде не показываются.
# На экране — строка «Google: проект X, адрес …<конец>» и кнопки «Проверить» / «Подключить Google…» («Изменить подключение…»).
# Существующие cfg не переписываются: файл пишется только после подтверждения в диалоге.
function Get-GoogleStatusText {
    if (-not $script:ProjectPath) { return "Google: нет проекта" }
    $name = Get-ProjectNameFromConfigPath $script:ProjectPath
    $cfg = Read-ProjectGoogleConfig (Get-ProjectGoogleConfigPath $script:ProjectPath)
    if ($cfg.WebAppUrl) { return ("Google: проект {0}, адрес {1}" -f $name, (Get-GoogleUrlTail $cfg.WebAppUrl)) }
    return ("Google: проект {0}, подключение не задано" -f $name)
}

function Show-ProjectGoogle {
    if (-not $ui.ContainsKey("GoogleInfo") -or -not $ui.GoogleInfo) { return }
    $cfg = $null
    if ($script:ProjectPath) { $cfg = Read-ProjectGoogleConfig (Get-ProjectGoogleConfigPath $script:ProjectPath) }
    $ui.GoogleInfo.Text = Get-GoogleStatusText
    $has = ($null -ne $cfg -and [bool]$cfg.WebAppUrl)
    $ui.GoogleConnect.Text = $(if ($has) { "Изменить подключение…" } else { "Подключить Google…" })
    foreach ($c in @($ui.GoogleCheck, $ui.GoogleConnect)) { $c.Enabled = [bool]$script:ProjectPath }
}

# Сохранение подключения из диалога: формат — как Test-BvsGooglePushConfigShape. Пустой секрет при «Изменить» — оставить прежний.
# Возвращает список замечаний (пусто — сохранено)
function Save-GoogleConnection([string]$ProjectConfigPath, [string]$Url, [string]$Secret) {
    $path = Get-ProjectGoogleConfigPath $ProjectConfigPath
    $cur = Read-ProjectGoogleConfig $path
    $u = ([string]$Url).Trim()
    $sec = ([string]$Secret).Trim()
    if (-not $sec) { $sec = $cur.Secret }
    $issues = @(Test-BvsGooglePushConfigShape @{ WebAppUrl = $u; Secret = $sec; ConfigPath = $path })
    if ($issues.Count -gt 0) { return $issues }
    Write-ProjectGoogleConfig $path $u $sec
    return @()
}

# Небольшой диалог «Подключить Google…» / «Изменить подключение…»: адрес (…/exec) и секрет (маскированное поле, прежний секрет не показывается)
function Show-GoogleConnectDialog {
    if (-not $script:ProjectPath) { return }
    $path = Get-ProjectGoogleConfigPath $script:ProjectPath
    $cur = Read-ProjectGoogleConfig $path
    $change = [bool]$cur.WebAppUrl
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = $(if ($change) { "Изменить подключение Google" } else { "Подключить Google" })
    $dlg.StartPosition = "CenterParent"; $dlg.FormBorderStyle = "FixedDialog"; $dlg.MaximizeBox = $false; $dlg.MinimizeBox = $false
    $dlg.ClientSize = New-Object System.Drawing.Size(560, 168)
    $l1 = New-Object System.Windows.Forms.Label; $l1.Text = "Адрес веб-приложения (…/exec) проекта " + (Get-ProjectNameFromConfigPath $script:ProjectPath); $l1.AutoSize = $true; $l1.Location = New-Object System.Drawing.Point(12, 12)
    $tUrl = New-Object System.Windows.Forms.TextBox; $tUrl.Location = New-Object System.Drawing.Point(12, 34); $tUrl.Size = New-Object System.Drawing.Size(536, 24); $tUrl.Text = [string]$cur.WebAppUrl
    $l2 = New-Object System.Windows.Forms.Label; $l2.Text = $(if ($change) { "Секрет (пусто — оставить прежний)" } else { "Секрет" }); $l2.AutoSize = $true; $l2.Location = New-Object System.Drawing.Point(12, 68)
    $tSec = New-Object System.Windows.Forms.TextBox; $tSec.Location = New-Object System.Drawing.Point(12, 90); $tSec.Size = New-Object System.Drawing.Size(536, 24); $tSec.UseSystemPasswordChar = $true
    $ok = New-Object System.Windows.Forms.Button; $ok.Text = "Сохранить"; $ok.Location = New-Object System.Drawing.Point(368, 128); $ok.Size = New-Object System.Drawing.Size(88, 28); $ok.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $cancel = New-Object System.Windows.Forms.Button; $cancel.Text = "Отмена"; $cancel.Location = New-Object System.Drawing.Point(462, 128); $cancel.Size = New-Object System.Drawing.Size(88, 28); $cancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $dlg.AcceptButton = $ok; $dlg.CancelButton = $cancel
    $dlg.Controls.AddRange(@($l1, $tUrl, $l2, $tSec, $ok, $cancel))
    $dlg.Add_FormClosing({
        param($sender, $e)
        if ($sender.DialogResult -ne [System.Windows.Forms.DialogResult]::OK) { return }
        $issues = @(Save-GoogleConnection $script:ProjectPath $tUrl.Text $tSec.Text)
        if ($issues.Count -gt 0) { [System.Windows.Forms.MessageBox]::Show($sender, ($issues -join [Environment]::NewLine), "Google") | Out-Null; $e.Cancel = $true }
    })
    try { [void]$dlg.ShowDialog($form) } finally { $dlg.Dispose() }
    Show-ProjectGoogle
    $ui.Status.Text = Get-GoogleStatusText
}

# «Проверить»: формат как в Test-BvsGooglePushConfigShape, затем ping и пробный POST. Секрет в статус и журнал не попадает
function Invoke-GoogleCheck {
    $cfg = Read-ProjectGoogleConfig (Get-ProjectGoogleConfigPath $script:ProjectPath)
    if (-not $cfg.WebAppUrl) { $ui.Status.Text = "Google не подключён: нажмите «Подключить Google…»."; return }
    $ui.GoogleCheck.Enabled = $false
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    $ui.Status.Text = ("Проверяю связь с Google: {0}..." -f (Get-GoogleUrlTail $cfg.WebAppUrl))
    [System.Windows.Forms.Application]::DoEvents()
    try { $r = Test-GooglePushConnection $cfg }
    finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default; $ui.GoogleCheck.Enabled = [bool]$script:ProjectPath }
    $text = (@($r.Lines) -join " ")
    Add-GooglePushLog ("проверка: проект {0}, {1}: {2}" -f (Get-ProjectNameFromConfigPath $script:ProjectPath), (Get-GoogleUrlTail $cfg.WebAppUrl), $(if ($r.Ok) { "ok" } else { "ошибка" }))
    $ui.Status.Text = ("Google: {0}" -f $text)
}

function Get-GooglePushDecision([string]$ProjectConfigPath, [string]$Root, [scriptblock]$Ask) {
    $t = Resolve-GooglePushTarget $ProjectConfigPath $Root $false
    if ($t.Kind -eq "project") { return @{ Go = $true; AllowShared = $false; Target = $t; Text = "" } }
    if ($t.Kind -eq "none") { return @{ Go = $false; AllowShared = $false; Target = $t; Text = "" } }
    if ($t.Warning) { Add-GooglePushLog $t.Warning }
    $yes = [bool](& $Ask "У проекта нет своего адреса Google. Отправить в общую таблицу из google_push.cfg?")
    if ($yes) { return @{ Go = $true; AllowShared = $true; Target = (Resolve-GooglePushTarget $ProjectConfigPath $Root $true); Text = "" } }
    return @{ Go = $false; AllowShared = $false; Target = $t; Text = "Google не обновлён: у проекта нет своего адреса, в общую таблицу не отправлено." }
}

# Раньше копировала фильтры и JSON в папку проекта рядом с xlsx (вторая копия ~1 МБ на каждый прогон).
# Единственные копии остаются в служебной папке прогона: google_push_payload.json читает push_google.ps1,
# фильтры_спецификаций.txt остаётся там же для проверки колонки «Параметр». Папка проекта хранит только xlsx.
function Copy-RunDiagnosticsToProject {
    param([string]$WorkDir, [string]$XlsxPath)
    return
}

function Push-GoogleAssignment {
    param([string]$XlsxPath, [string]$PayloadPath = "", [switch]$AllowShared)
    $pushScript = Join-Path $ToolDir "push_google.ps1"
    if (-not (Test-Path -LiteralPath $pushScript)) {
        return @{ Ok = $false; Message = "Нет push_google.ps1" }
    }
    try {
        # JSON для Google лежит в служебной папке прогона (отчеты\служебные\<время>); без явного пути — рядом с xlsx, как раньше
        $payloadPath = $PayloadPath
        if ([string]::IsNullOrWhiteSpace($payloadPath)) { $payloadPath = Join-Path (Split-Path -Parent $XlsxPath) "google_push_payload.json" }
        $pushParams = @{
            XlsxPath = $XlsxPath
            ToolDir  = $ToolDir
        }
        if (Test-Path -LiteralPath $payloadPath) { $pushParams.PayloadPath = $payloadPath }
        if ($script:ProjectPath) { $pushParams.ProjectConfigPath = $script:ProjectPath }
        if ($AllowShared) { $pushParams.AllowShared = $true }
        $out = & $pushScript @pushParams 2>&1
        $text = ($out | Out-String).Trim()
        # push_google.ps1 при успехе последней строкой печатает "OK <ответ скрипта>", при любой ошибке — бросает исключение.
        # Предупреждения/прочий вывод выше не мешают, но успех — только если последняя строка именно такая.
        $lines = @($text -split "\r?\n" | Where-Object { $_.Trim() })
        $last = ""
        if ($lines.Count -gt 0) { $last = $lines[$lines.Count - 1].Trim() }
        if ($last -match '^OK\s+(?<tail>.+)$') {
            $tail = $Matches['tail']
            if ($tail -match '(?i)^(forbidden|error)\b' -or $tail -match '(?i)^\{\s*"ok"\s*:\s*false') {
                return @{ Ok = $false; Message = $text }
            }
            return @{ Ok = $true; Message = $last }
        }
        return @{ Ok = $false; Message = $text }
    }
    catch {
        return @{ Ok = $false; Message = $_.Exception.Message }
    }
}

function Invoke-BuildAssignment {
    param(
        [string]$ExportDir,
        [string]$OutXlsx,
        [string]$HistoryPath,
        [hashtable]$Period,
        [switch]$DynamicsOnly,
        [switch]$CollapseDay,
        [switch]$Partial,
        [string]$AuxDir = ""
    )
    $build = Join-Path $ToolDir "build_assignment.ps1"
    $params = @{
        ExportDir    = $ExportDir
        OutXlsx      = $OutXlsx
        HistoryPath  = $HistoryPath
        DynamicsFrom = $Period.From
        DynamicsTo   = $Period.To
    }
    if ($DynamicsOnly) { $params.DynamicsOnly = $true }
    if ($CollapseDay) { $params.CollapseDay = $true }
    if ($Partial) { $params.Partial = $true }
    if ($AuxDir) { $params.AuxDir = $AuxDir }
    $prefix = Get-AssignProjectPrefix $script:ProjectPath $ui.Summary.Text.Trim()
    if ($prefix) { $params.ProjectPrefix = $prefix }          # префикс проекта вместо зашитого PROJ
    return & $build @params 2>&1
}

function Read-BuildLogText([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path)) { return "" }
    $bytes = [System.IO.File]::ReadAllBytes($Path)
    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
        return ([System.Text.Encoding]::UTF8.GetString($bytes, 3, $bytes.Length - 3)).Trim()
    }
    if ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) {
        return ([System.Text.Encoding]::Unicode.GetString($bytes)).Trim()
    }
    $utf8 = [System.Text.Encoding]::UTF8.GetString($bytes)
    if ($utf8 -match '[\u0400-\u04FF]') { return $utf8.Trim() }
    try {
        $cp1251 = [System.Text.Encoding]::GetEncoding(1251)
        return $cp1251.GetString($bytes).Trim()
    }
    catch {
        return $utf8.Trim()
    }
}

function Get-AssignBuildFailureSummary {
    param(
        [string]$Out,
        [string]$Err,
        [int]$BuildExitCode,
        [bool]$XlsxValid
    )
    $combined = (($Out + "`n" + $Err).Trim())
    if ($combined -match 'Split-Path|BvsExcelModuleDir|ParameterArgumentValidationErrorNullNotAllowed') {
        return ("Ошибка пути к excel_isolated.ps1 при фоновой сборке. Обновите batch_verification_schedules (excel_isolated.ps1 с PSScriptRoot) и повторите.`n`n" + $combined)
    }
    if ($combined -match 'COM к нему не подключился|не подключился по COM|не подключился\.') {
        return ("Сборка xlsx: отдельный Excel не подключился по COM. Закройте зависшие EXCEL.EXE в диспетчере задач и снова «Собрать задание». Если Excel на ПК ещё не активировали — один раз откройте и закройте вручную.`n`n" + $combined)
    }
    if ($combined -notmatch 'BUILDER=(OPENXML|COM)') {
        if ($combined -match 'DocumentFormat\.OpenXml|assignment_xlsx|Нет .*OpenXml') {
            return ("Сборка xlsx: нет или не загружается lib\DocumentFormat.OpenXml.dll. Скопируйте batch_verification_schedules с папкой lib (только dll, без openxml_pkg).`n`n" + $combined)
        }
        if ($combined -match 'excel_isolated|COM к нему') {
            return ("Сборка через Excel COM не удалась. Обновите инструмент: задание собирается через Open XML (lib\DocumentFormat.OpenXml.dll).`n`n" + $combined)
        }
        return ("Сборка не завершена (нет BUILDER=OPENXML в логе). Обновите папку инструмента целиком.`n`n" + $combined)
    }
    if (-not $XlsxValid) {
        return ("Файл задания повреждён или собран устаревшим скриптом. Обновите инструмент и соберите снова.`n`n" + $combined)
    }
    return ("Excel-задание не создано (код сборки {0}).`n`n{1}" -f $BuildExitCode, $combined)
}

function Finish-AssignBuild {
    param(
        [int]$RevitCode,
        [string]$WorkDir,
        [string]$XlsxPath,
        [string]$BuildLogPath,
        [string]$BuildErrPath,
        [int]$BuildExitCode,
        [string]$BuildStdout = ""
    )
    Set-ActionsEnabled $true
    $out = $BuildStdout
    if ([string]::IsNullOrWhiteSpace($out) -and $BuildLogPath -and (Test-Path -LiteralPath $BuildLogPath)) {
        $out = Read-BuildLogText $BuildLogPath
    }
    $err = ""
    if ($BuildErrPath -and (Test-Path -LiteralPath $BuildErrPath)) {
        $err = Read-BuildLogText $BuildErrPath
    }
    if ($BuildLogPath -and (Test-Path -LiteralPath $BuildLogPath) -and [string]::IsNullOrWhiteSpace($out)) {
        $out = Read-BuildLogText $BuildLogPath
    }
    $exportedRows = -1
    $partialNote = ""
    $resultText = (Get-BvsExportResult $WorkDir).ResultText          # export.txt [result] (прежде export_result.txt)
    if ($resultText) {
        if ($resultText -match 'exported=(\d+)') { $exportedRows = [int]$Matches[1] }
        # выборочная выгрузка: связи, которые не были загружены (их замечания не закрываются), и спецификации, которых нет в сводной
        if ($resultText -match '(?m)^unloaded_models=(.+?)\s*$') { $partialNote += (" Связь не загружена, не проверялась: " + ($Matches[1] -replace '\|', ', ') + ".") }
        if ($resultText -match '(?m)^missing_specs=(.+?)\s*$') { $partialNote += (" Нет в сводной: " + ($Matches[1] -replace '\|', ', ') + ".") }
    }
    $buildOk = ($BuildExitCode -eq 0) -and ($out -match 'BUILDER=(OPENXML|COM)')
    $xlsxValid = $false
    if (Test-Path -LiteralPath $XlsxPath) {
        $xlsxValid = Test-BvsAssignmentXlsx $XlsxPath
    }
    $built = $buildOk -and $xlsxValid
    if ($built) {
        # крупные промежуточные файлы (tsv, _sections) больше не нужны; export.txt, диагностика и JSON для Google остаются в служебной папке прогона до автоочистки
        Remove-RunBulk $WorkDir
        Set-OpenFolderTarget $XlsxPath
        Copy-RunDiagnosticsToProject -WorkDir $WorkDir -XlsxPath $XlsxPath
        if ($exportedRows -eq 0 -and -not $script:Run.Partial) {
            $ui.Status.Text = ("Файл создан, но он пустой: из сводной не вышло ни одной строки. Это не задание. Файл: {0}" -f $XlsxPath)
        }
        else {
            $ui.Status.Text = ("Задание записано (xlsx). Имена спецификаций как в Revit. Лист «Динамика» дополнен. Файл: {0}" -f $XlsxPath)
            if ($script:Run.Partial) { $ui.Status.Text = ("Выборочная выгрузка: обновлены только выбранные пары, остальное взято из прошлых проверок.{0} Файл: {1}" -f $partialNote, $XlsxPath) }
            # Адрес веб-приложения — свой у каждого проекта (google_<проект>.cfg); общий google_push.cfg — только после подтверждения.
            # Поле со ссылкой на таблицу — только для чата и отправку не включает.
            $decision = Get-GooglePushDecision $script:ProjectPath $ToolDir {
                param($Question)
                ([System.Windows.Forms.MessageBox]::Show($form, $Question, "Google", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning, [System.Windows.Forms.MessageBoxDefaultButton]::Button2) -eq [System.Windows.Forms.DialogResult]::Yes)
            }
            if ($decision.Go) {
                $ui.Status.Text = $ui.Status.Text + (" Отправляю в Google: {0}..." -f $decision.Target.Label)
                Add-GooglePushLog ("отправка: " + $decision.Target.Label)
                [System.Windows.Forms.Application]::DoEvents()
                $push = Push-GoogleAssignment -XlsxPath $XlsxPath -PayloadPath (Join-Path $WorkDir "google_push_payload.json") -AllowShared:$decision.AllowShared
                if ($push.Ok) {
                    Add-GooglePushLog ("готово: " + $decision.Target.Label)
                    $ui.Status.Text = ("Задание записано и Google обновлён ({0}).{1} Файл: {2}" -f $decision.Target.Label, $partialNote, $XlsxPath)
                    if ($push.Message -match '"sharing"\s*:\s*"error') {
                        $ui.Status.Text += " ВНИМАНИЕ: доступ к таблице по ссылке не открыт (запретил домен или нет прав владельца), см. служебное\google_push\README.md."
                    }
                }
                else {
                    Add-GooglePushLog ("ошибка: " + $decision.Target.Label)
                    $ui.Status.Text = ("Задание на диске: {0}. Google не обновлён ({1}): {2}" -f $XlsxPath, $decision.Target.Label, $push.Message)
                }
            }
            elseif ($decision.Text) {
                Add-GooglePushLog ("не отправлено: " + $decision.Target.Label)
                $ui.Status.Text = ("Задание записано. {0} Файл: {1}" -f $decision.Text, $XlsxPath)
            }
        }
        Refresh-PartialLists
        Show-Finished $ui.Status.Text
        return
    }
    if (Test-Path -LiteralPath $XlsxPath) {
        Remove-Item -LiteralPath $XlsxPath -Force -ErrorAction SilentlyContinue
    }
    $payloadSide = Join-Path $WorkDir "google_push_payload.json"
    if ($WorkDir -and (Test-Path -LiteralPath $payloadSide)) {
        Remove-Item -LiteralPath $payloadSide -Force -ErrorAction SilentlyContinue
    }
    if ($RevitCode -eq 0) {
        $ui.Status.Text = Get-AssignBuildFailureSummary -Out $out -Err $err -BuildExitCode $BuildExitCode -XlsxValid $xlsxValid
    }
    else {
        $tail = (($out + "`n" + $err).Trim())
        $ui.Status.Text = ("Задание не собрано (Revit код {0}). {1}" -f $RevitCode, $tail)
    }
    Show-Finished $ui.Status.Text
}

function Start-DynamicsOnly {
    Save-Settings
    $summary = $ui.Summary.Text.Trim()
    if (-not (Test-ModelReady $summary)) {
        [System.Windows.Forms.MessageBox]::Show($form, "Укажите сводную модель на вкладке «Спецификации».", "Проверочные спецификации") | Out-Null
        return
    }
    $historyPath = Get-HistoryPath $summary
    if (-not (Test-Path -LiteralPath $historyPath)) {
        [System.Windows.Forms.MessageBox]::Show($form, "Нет файла истории. Сначала соберите задание хотя бы один раз.", "Проверочные спецификации") | Out-Null
        return
    }
    $period = Get-DynamicsPeriodFromUi
    $xlsxPath = Get-DynamicsOnlyPath $summary
    try { [void](Invoke-BvsHousekeeping -ReportDir $ReportDir -KeepDays $script:AuxKeepDays -AppDir $AppDir) } catch { }
    $auxDir = New-RunAuxDir $script:Layout.Aux (Get-Date -Format "yyyyMMdd_HHmmss")
    $ui.Status.Text = "Собираю лист «Динамика»..."
    [System.Windows.Forms.Application]::DoEvents()
    Set-ActionsEnabled $false
    try {
        $out = Invoke-BuildAssignment -ExportDir "" -OutXlsx $xlsxPath -HistoryPath $historyPath -Period $period -DynamicsOnly -AuxDir $auxDir
        if (Test-Path -LiteralPath $xlsxPath) {
            Set-OpenFolderTarget $xlsxPath
            $ui.Status.Text = ("Лист «Динамика» записан. Период {0} — {1}. Файл: {2}" -f $period.From, $period.To, $xlsxPath)
            Show-Finished $ui.Status.Text
        }
        else {
            $ui.Status.Text = ("Динамика не собрана. {0}" -f (($out | Out-String).Trim()))
            Show-Finished $ui.Status.Text
        }
    }
    finally {
        Set-ActionsEnabled $true
    }
}

function Write-RunFiles {
    param(
        [string[]]$Models,
        [string]$Source,
        [string]$ReportDir,
        [string]$RunId,
        [string]$Mode,
        [string]$ProjectPath,
        [string]$PrefixFrom,
        [string]$PrefixTo,
        [string]$InsertMode,
        [string[]]$Extra = @(),
        # четыре CSV отчёта вставки. Их читают Export-ReportXlsx, итог прогона и «Открыть CSV».
        # Для выгрузки, классификатора и «Посчитать» они пустые и никем не читаются — не создаются.
        [switch]$InsertReport
    )
    if (-not (Test-Path -LiteralPath $ReportDir)) {
        New-Item -ItemType Directory -Path $ReportDir -Force | Out-Null
    }
    $utf8 = New-Object System.Text.UTF8Encoding $false
    [System.IO.File]::WriteAllLines((Join-Path $AppDir "rvt_list_run.txt"), $Models, $utf8)

    $cfg = @(
        ("source={0}" -f $Source),
        ("report_dir={0}" -f $ReportDir),
        ("run_id={0}" -f $RunId),
        ("mode={0}" -f $Mode),
        ("project_path={0}" -f $ProjectPath),
        ("prefix_from={0}" -f $PrefixFrom),
        ("prefix_to={0}" -f $PrefixTo),
        ("insert_mode={0}" -f $InsertMode)
    ) + @($Extra)
    if ($script:Layout -and $script:Layout.Logs) { $cfg = @($cfg) + @("log_folder={0}" -f $script:Layout.Logs) }
    $cfg = $cfg -join "`r`n"
    $bom = New-Object System.Text.UTF8Encoding $true
    [System.IO.File]::WriteAllText($CfgPath, $cfg + "`r`n", $bom)

    $summary = ""
    $params = ""
    $skips = ""
    $inserted = ""
    if ($InsertReport) {
        $summary = Join-Path $ReportDir ("{0}_сводка.csv" -f $RunId)
        $params = Join-Path $ReportDir ("{0}_параметры.csv" -f $RunId)
        $skips = Join-Path $ReportDir ("{0}_пропуски.csv" -f $RunId)
        $inserted = Join-Path $ReportDir ("{0}_вставлено.csv" -f $RunId)
        [System.IO.File]::WriteAllText($summary, "Модель;Раздел;Удалено;Вставлено;Пропущено;Ошибка;Режим`r`n", $bom)
        [System.IO.File]::WriteAllText($params, "Модель;Спецификация;Параметр;Замечание`r`n", $bom)
        [System.IO.File]::WriteAllText($skips, "Модель;Спецификация;Причина`r`n", $bom)
        [System.IO.File]::WriteAllText($inserted, "Модель;Раздел;Спецификация`r`n", $bom)
    }
    return @{
        List     = (Join-Path $AppDir "rvt_list_run.txt")
        Summary  = $summary
        Params   = $params
        Skips    = $skips
        Inserted = $inserted
    }
}

function Read-ReportCsv {
    param([string]$Path)
    $empty = @{ Headers = @(); Rows = @() }
    if (-not (Test-Path -LiteralPath $Path)) { return $empty }
    $lines = @(Get-Content -LiteralPath $Path -Encoding UTF8 | Where-Object { $_ -ne "" })
    if ($lines.Count -eq 0) { return $empty }
    $headers = @($lines[0].Split([char]';'))
    $rows = New-Object System.Collections.Generic.List[object]
    for ($i = 1; $i -lt $lines.Count; $i++) {
        $rows.Add(@($lines[$i].Split([char]';')))
    }
    return @{ Headers = $headers; Rows = $rows }
}

function Get-ModelFileName {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return "" }
    return [IO.Path]::GetFileName($Path)
}

function Convert-ReportNumber {
    param($Value)
    $n = 0
    if ([int]::TryParse([string]$Value, [ref]$n)) { return $n }
    if ([string]::IsNullOrWhiteSpace([string]$Value)) { return $null }
    return [string]$Value
}

function Write-SheetGrid {
    param($Sheet, [string[]]$Headers, $DataRows)
    $cols = $Headers.Count
    $count = 0
    if ($null -ne $DataRows) { $count = $DataRows.Count }
    $n = 1 + $count
    $grid = New-Object 'object[,]' $n, $cols
    for ($c = 0; $c -lt $cols; $c++) { $grid[0, $c] = $Headers[$c] }
    for ($r = 0; $r -lt $count; $r++) {
        $row = $DataRows[$r]
        for ($c = 0; $c -lt $cols; $c++) {
            $val = $null
            if ($null -ne $row -and $c -lt $row.Count) { $val = $row[$c] }
            if ($null -eq $val) { $val = "" }
            $grid[($r + 1), $c] = $val
        }
    }
    $range = $Sheet.Range($Sheet.Cells.Item(1, 1), $Sheet.Cells.Item($n, $cols))
    $range.Value2 = $grid
    return $n
}

function Format-ReportSheet {
    param($Excel, $Sheet, [int]$Cols, [int]$LastRow, [int[]]$ParentRows, [int[]]$CountCols, [int[]]$Widths)
    $header = $Sheet.Range($Sheet.Cells.Item(1, 1), $Sheet.Cells.Item(1, $Cols))
    $header.Font.Bold = $true
    $header.Interior.Color = 0xF2E1D9
    $header.HorizontalAlignment = -4108
    foreach ($pr in @($ParentRows)) {
        if ($pr -lt 2) { continue }
        $row = $Sheet.Range($Sheet.Cells.Item($pr, 1), $Sheet.Cells.Item($pr, $Cols))
        $row.Font.Bold = $true
        $row.Interior.Color = 0xF2F2F2
    }
    foreach ($cc in @($CountCols)) {
        if ($cc -lt 1 -or $LastRow -lt 2) { continue }
        $Sheet.Range($Sheet.Cells.Item(2, $cc), $Sheet.Cells.Item($LastRow, $cc)).NumberFormat = "0"
    }
    for ($c = 0; $c -lt $Widths.Count; $c++) {
        $Sheet.Columns.Item($c + 1).ColumnWidth = $Widths[$c]
    }
    if ($LastRow -ge 1) {
        $Sheet.Range($Sheet.Cells.Item(1, 1), $Sheet.Cells.Item($LastRow, $Cols)).AutoFilter() | Out-Null
    }
    $Sheet.Activate() | Out-Null
    $Excel.ActiveWindow.SplitRow = 1
    $Excel.ActiveWindow.FreezePanes = $false
    $Excel.ActiveWindow.FreezePanes = $true
}

function Add-ModelGroups {
    param($Sheet, $Groups)
    if ($null -eq $Groups -or $Groups.Count -lt 1) { return }
    $has = $false
    foreach ($g in $Groups) { if ($g.ChildCount -ge 1) { $has = $true; break } }
    if (-not $has) { return }
    $Sheet.Outline.SummaryRow = 0
    foreach ($g in $Groups) {
        if ($g.ChildCount -lt 1) { continue }
        $first = [int]$g.Parent + 1
        $last = [int]$g.Parent + [int]$g.ChildCount
        $Sheet.Rows.Item("${first}:${last}").Group() | Out-Null
    }
    $Sheet.Outline.ShowLevels(1) | Out-Null
}

function Build-SummaryRows {
    param($Csv)
    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($cells in $Csv.Rows) {
        $file = Get-ModelFileName $(if ($cells.Count -gt 0) { $cells[0] } else { "" })
        $sec = if ($cells.Count -gt 1) { $cells[1] } else { "" }
        $del = if ($cells.Count -gt 2) { $cells[2] } else { "" }
        $ins = if ($cells.Count -gt 3) { $cells[3] } else { "" }
        $skip = if ($cells.Count -gt 4) { $cells[4] } else { "" }
        $err = if ($cells.Count -gt 5) { $cells[5] } else { "" }
        $mode = if ($cells.Count -gt 6) { $cells[6] } else { "" }
        $rows.Add(@($file, $sec, (Convert-ReportNumber $del), (Convert-ReportNumber $ins), (Convert-ReportNumber $skip), $err, $mode))
    }
    return $rows
}

function Build-GroupedRows {
    param($Csv, [int]$GroupKeyIndex, [scriptblock]$ParentValues, [scriptblock]$ChildValues)
    $rows = New-Object System.Collections.Generic.List[object]
    $groups = New-Object System.Collections.Generic.List[object]
    $order = New-Object System.Collections.Generic.List[string]
    $map = @{}
    foreach ($cells in $Csv.Rows) {
        $key = ""
        if ($GroupKeyIndex -lt $cells.Count) { $key = [string]$cells[$GroupKeyIndex] }
        if (-not $map.ContainsKey($key)) {
            $map[$key] = New-Object System.Collections.Generic.List[object]
            $order.Add($key)
        }
        $map[$key].Add($cells)
    }
    foreach ($key in $order) {
        $items = $map[$key]
        $parent = & $ParentValues $key $items
        $rows.Add($parent)
        $parentRow = $rows.Count + 1
        foreach ($item in $items) {
            $rows.Add((& $ChildValues $item))
        }
        $groups.Add([pscustomobject]@{ Parent = $parentRow; ChildCount = $items.Count })
    }
    return @{ Rows = $rows; Groups = $groups }
}

# Отчёт прогона (Сводка, Вставлено, Параметры, Пропуски) пишется в .xlsx напрямую (report_xlsx.ps1): Excel и COM не используются, временные книги не создаются.
# Ошибка записи — $null и причина в $script:ReportXlsxError (статус и ссылка «Открыть CSV»).
# Список -> обычный массив (@() и .ToArray() на List[object] с PSCustomObject в некоторых сборках PowerShell дают «Argument types do not match»)
function ConvertTo-PlainArray($List) {
    $out = New-Object System.Collections.ArrayList
    foreach ($i in $List) { [void]$out.Add($i) }
    return ,$out.ToArray()
}

function Export-ReportXlsx {
    param([string]$Summary, [string]$Params, [string]$Skips, [string]$Inserted, [string]$XlsxPath)
    try {
        $summaryRows = Build-SummaryRows (Read-ReportCsv $Summary)
        $insertedBuilt = Build-GroupedRows -Csv (Read-ReportCsv $Inserted) -GroupKeyIndex 0 -ParentValues {
            param($key, $items)
            $sec = ""
            if ($items.Count -gt 0 -and $items[0].Count -gt 1) { $sec = $items[0][1] }
            return @((Get-ModelFileName $key), $sec, $items.Count, "")
        } -ChildValues {
            param($item)
            $name = if ($item.Count -gt 2) { $item[2] } else { "" }
            return @("", "", "", $name)
        }
        $paramsBuilt = Build-GroupedRows -Csv (Read-ReportCsv $Params) -GroupKeyIndex 0 -ParentValues {
            param($key, $items)
            return @((Get-ModelFileName $key), $items.Count, "", "", "")
        } -ChildValues {
            param($item)
            $a = if ($item.Count -gt 1) { $item[1] } else { "" }
            $b = if ($item.Count -gt 2) { $item[2] } else { "" }
            $c = if ($item.Count -gt 3) { $item[3] } else { "" }
            return @("", "", $a, $b, $c)
        }
        $skipsBuilt = Build-GroupedRows -Csv (Read-ReportCsv $Skips) -GroupKeyIndex 0 -ParentValues {
            param($key, $items)
            return @((Get-ModelFileName $key), $items.Count, "", "")
        } -ChildValues {
            param($item)
            $a = if ($item.Count -gt 1) { $item[1] } else { "" }
            $b = if ($item.Count -gt 2) { $item[2] } else { "" }
            return @("", "", $a, $b)
        }
        $sheets = @(
            @{ Name = "Сводка"; Headers = @("Модель", "Раздел", "Удалено", "Вставлено", "Пропущено", "Ошибка", "Режим"); Rows = (ConvertTo-PlainArray $summaryRows)
               Widths = @(46, 12, 12, 14, 14, 36, 24); CountCols = @(3, 4, 5); ParentRows = @(); Groups = @() },
            @{ Name = "Вставлено"; Headers = @("Модель", "Раздел", "Число", "Спецификация"); Rows = (ConvertTo-PlainArray $insertedBuilt.Rows)
               Widths = @(46, 12, 10, 72); CountCols = @(3); ParentRows = @(foreach ($g in (ConvertTo-PlainArray $insertedBuilt.Groups)) { [int]$g.Parent }); Groups = (ConvertTo-PlainArray $insertedBuilt.Groups) },
            @{ Name = "Параметры"; Headers = @("Модель", "Число", "Спецификация", "Параметр", "Замечание"); Rows = (ConvertTo-PlainArray $paramsBuilt.Rows)
               Widths = @(46, 10, 72, 28, 42); CountCols = @(2); ParentRows = @(foreach ($g in (ConvertTo-PlainArray $paramsBuilt.Groups)) { [int]$g.Parent }); Groups = (ConvertTo-PlainArray $paramsBuilt.Groups) },
            @{ Name = "Пропуски"; Headers = @("Модель", "Число", "Спецификация", "Причина"); Rows = (ConvertTo-PlainArray $skipsBuilt.Rows)
               Widths = @(46, 10, 72, 42); CountCols = @(2); ParentRows = @(foreach ($g in (ConvertTo-PlainArray $skipsBuilt.Groups)) { [int]$g.Parent }); Groups = (ConvertTo-PlainArray $skipsBuilt.Groups) }
        )
        return (Write-ReportXlsx -Path $XlsxPath -Sheets $sheets)
    }
    catch {
        $script:ReportXlsxError = $_.Exception.Message
        return $null
    }
}

function Get-NextReportPath {
    param([string]$Dir)
    $day = Get-Date -Format "yyyy-MM-dd"
    $path = Join-Path $Dir ($day + ".xlsx")
    if (-not (Test-Path -LiteralPath $path)) { return $path }
    $n = 2
    while ($n -lt 100) {
        $path = Join-Path $Dir ("{0}_{1}.xlsx" -f $day, $n)
        if (-not (Test-Path -LiteralPath $path)) { return $path }
        $n++
    }
    return (Join-Path $Dir ("{0}_{1}.xlsx" -f $day, (Get-Date -Format "HHmmss")))
}

function Read-LiveStatus {
    $path = Join-Path $AppDir "live_status.txt"
    if (-not (Test-Path -LiteralPath $path)) { return $null }
    $map = @{}
    foreach ($line in @(Get-Content -LiteralPath $path -Encoding UTF8 -ErrorAction SilentlyContinue)) {
        if ($line -match '^(index|total|name|inserted|skipped|params|result)=(.*)$') {
            $map[$Matches[1]] = $Matches[2]
        }
    }
    if (-not $map.ContainsKey("name")) { return $null }
    return $map
}

function Format-LiveStatus($Live) {
    $params = 0
    try { $params = [int]$Live.params } catch { $params = 0 }
    $holes = if ($params -gt 0) { "параметры: $params" } else { "параметры: нет" }
    return ("{0} из {1} · {2} · вставлено {3} · пропущено {4} · {5}" -f $Live.index, $Live.total, $Live.name, $Live.inserted, $Live.skipped, $holes)
}

# Архив*, Archive*, Old, Стар/Старые/Старое/Старый (не «Старт»), OLD_* и папка-дата. Резерв/Backup/_backup — отдельно, ниже.
function Test-SkippedArchiveFolderName([string]$Name) {
    if ([string]::IsNullOrWhiteSpace($Name)) { return $false }
    if ($Name -match '^(?i)(Архив|Archive)') { return $true }
    if ($Name -match '^(?i)Old$') { return $true }
    if ($Name -match '^(?i)Стар(ые|ое|ый)?$') { return $true }
    if ($Name -match '^(?i)OLD_') { return $true }
    if ($Name -match '^\d{4}-\d{2}-\d{2}') { return $true }
    return $false
}

function Test-ExcludedModelPath([string]$Path) {
    if ([string]::IsNullOrWhiteSpace($Path)) { return $true }
    $p = $Path.Replace("/", "\")
    if ($p -match '(?i)\\Backup\\') { return $true }
    if ($p -match '(?i)\\Revit_temp\\') { return $true }
    if ($p -match '(?i)_backup(\\|$)') { return $true }
    if ($p -match '\\Резерв\\') { return $true }
    if ($p -match '\\Семейства\\') { return $true }
    foreach ($seg in @($p -split '\\')) { if (Test-SkippedArchiveFolderName $seg) { return $true } }
    return $false
}

function Test-IsIosRvtName([string]$Name) {
    if ([string]::IsNullOrWhiteSpace($Name)) { return $false }
    if ($Name -notmatch '(?i)\.rvt$') { return $false }
    if ($Name -match '\.\d{4}\.rvt$') { return $false }
    if ($Name -match '(?i)_backup') { return $false }
    return $true
}

function Test-SkipWalkFolderName([string]$Name) {
    if ([string]::IsNullOrWhiteSpace($Name)) { return $true }
    if ($Name -match '^(?i)(DWG|PDF|EXCEL|WORD|XLSX|NWC|NWD|IFC|Backup|Revit_temp)$') { return $true }
    if ($Name -eq "Хлам") { return $true }
    if (Test-SkippedArchiveFolderName $Name) { return $true }
    return $false
}

function Add-IosRvtFromFolder {
    param([string]$Folder, [System.Collections.Generic.List[string]]$Found)
    $files = @()
    try { $files = @(Get-ChildItem -LiteralPath $Folder -File -Force -ErrorAction Stop) }
    catch { return }
    foreach ($f in $files) {
        if (-not (Test-IsIosRvtName $f.Name)) { continue }
        $native = Get-NativePath $f.FullName
        if ($native -and -not (Test-ExcludedModelPath $native)) { [void]$Found.Add($native) }
    }
}

function Get-IosModelsFromFolder([string]$Folder) {
    $Folder = Resolve-PickedFolderPath $Folder
    if ([string]::IsNullOrWhiteSpace($Folder)) { return @() }
    $found = New-Object System.Collections.Generic.List[string]
    $maxDepth = 8
    $queue = New-Object System.Collections.Generic.Queue[object]
    $queue.Enqueue(@{ Path = $Folder; Depth = 0 })
    $guard = 0
    while ($queue.Count -gt 0 -and $guard -lt 4000) {
        $guard++
        $cur = $queue.Dequeue()
        $path = [string]$cur.Path
        $depth = [int]$cur.Depth
        if (Test-ExcludedModelPath ($path.TrimEnd("\") + "\")) { continue }
        $leaf = ""
        try { $leaf = [System.IO.Path]::GetFileName($path.TrimEnd("\")) } catch { $leaf = "" }
        if ($leaf -ieq "RVT") {
            Add-IosRvtFromFolder -Folder $path -Found $found
            continue
        }
        if ($depth -ge $maxDepth) { continue }
        $kids = @()
        try { $kids = @(Get-ChildItem -LiteralPath $path -Directory -Force -ErrorAction Stop) }
        catch { continue }
        foreach ($k in $kids) {
            if (Test-ExcludedModelPath ($k.FullName.TrimEnd("\") + "\")) { continue }
            if (Test-SkipWalkFolderName $k.Name) { continue }
            $queue.Enqueue(@{ Path = $k.FullName; Depth = ($depth + 1) })
        }
    }
    return @($found | Sort-Object -Unique)
}

function Test-SectionModelName([string]$Name) {
    if ([string]::IsNullOrWhiteSpace($Name)) { return $false }
    return (Test-ProjectEngineeringModel $Name)          # единственное правило: код раздела «_КОД_» в имени (project_config.ps1)
}

function Get-SectionModelsNearSummary([string]$SummaryPath) {
    if ([string]::IsNullOrWhiteSpace($SummaryPath)) { return @() }
    if ($SummaryPath.Trim() -match '^(?i)RSN://') { return @() }
    if (-not (Test-Path -LiteralPath $SummaryPath)) { return @() }
    $dir = Split-Path -Parent $SummaryPath
    $leaf = ""
    try { $leaf = [System.IO.Path]::GetFileName($dir.TrimEnd("\")) } catch { $leaf = "" }
    if ($leaf -ieq "RVT") { $dir = Split-Path -Parent $dir }
    $root = Split-Path -Parent $dir
    if ([string]::IsNullOrWhiteSpace($root)) { $root = $dir }
    $summaryLeaf = [System.IO.Path]::GetFileName($SummaryPath)
    $all = @(Get-IosModelsFromFolder $root)
    return @($all | Where-Object {
        $name = [System.IO.Path]::GetFileName($_)
        ($name -ne $summaryLeaf) -and (Test-SectionModelName $name)
    })
}

# ---- «Выборочно»: слева модели (по разделам), справа спецификации ВЫДЕЛЕННОГО узла; отметки спек — для каждой модели отдельно ----
function Get-AsgCheckedModels {
    $out = New-Object System.Collections.Generic.List[string]
    foreach ($sec in @($ui.AsgTree.Nodes)) {
        foreach ($n in @($sec.Nodes)) { if ($n.Checked) { $out.Add([string]$n.Tag) } }
    }
    return $out.ToArray()
}

# модели выделенного (кликнутого) узла: модель — она сама, раздел — все его модели, ничего не выделено — пусто
function Get-AsgSelectedModels {
    $node = $ui.AsgTree.SelectedNode
    if ($null -eq $node) { return @() }
    if ($node.Nodes.Count -gt 0) { return @($node.Nodes | ForEach-Object { [string]$_.Tag }) }
    if ($node.Tag) { return @([string]$node.Tag) }
    return @()
}

function Get-AsgSelectedTitle {
    $node = $ui.AsgTree.SelectedNode
    if ($null -eq $node) { return "" }
    if ($node.Nodes.Count -gt 0) { return ([string]$node.Text -replace '\s*\(\d+\)\s*$', '') }
    return [string]$node.Tag
}

# выгружаемые пары: отмеченные модели × отмеченные у них спецификации
function Get-AsgPairs {
    return @(Get-PartialPairs $script:AsgPicks @(Get-AsgCheckedModels) $script:AsgAllSpecs $script:AsgModelSets)
}

function Update-AsgCaptions {
    $models = @(Get-AsgSelectedModels)
    if ($models.Count -eq 0) {
        $ui.AsgSpecCount.Text = "Выберите модель или раздел"
    }
    else {
        $text = Format-PartialNodeCaption (Get-AsgSelectedTitle) $script:AsgPicks $models $script:AsgAllSpecs $script:AsgModelSets
        if ($models.Count -eq 1 -and -not (@(Get-AsgCheckedModels) -contains $models[0])) { $text += " — модель не отмечена, не выгружается" }
        $ui.AsgSpecCount.Text = $text
    }
    $ui.AsgTotal.Text = Format-PartialTotalCaption (Get-AsgPairs)
}

# Правый список = спецификации выделенного узла. Галочка: у модели — её отметка; у раздела — у всех его моделей, где спека есть
# (частично — состояние Indeterminate и пометка в названии)
function Show-AsgSpecList {
    $models = @(Get-AsgSelectedModels)
    $specs = @()
    if ($models.Count -gt 0) { $specs = @(Get-PartialNodeSpecs $models $script:AsgAllSpecs $script:AsgModelSets) }
    $script:AsgShown = $specs
    $top = 0
    try { $top = $ui.AsgSpecs.TopIndex } catch { $top = 0 }
    $script:AsgSpecBusy = $true
    $ui.AsgSpecs.BeginUpdate()
    try {
        $ui.AsgSpecs.Items.Clear()
        foreach ($sp in $specs) {
            $st = Get-PartialSpecState $script:AsgPicks $models $script:AsgModelSets $sp
            $caption = $sp
            if ($st -eq "Indeterminate") { $caption += "   [частично]" }
            $i = $ui.AsgSpecs.Items.Add($caption)
            $ui.AsgSpecs.SetItemCheckState($i, [System.Windows.Forms.CheckState]::$st)
        }
        if ($top -gt 0 -and $top -lt $ui.AsgSpecs.Items.Count) { $ui.AsgSpecs.TopIndex = $top }
    }
    finally {
        $ui.AsgSpecs.EndUpdate()          # без EndUpdate список остаётся «замороженным» и пустым при любой ошибке в теле
        $script:AsgSpecBusy = $false
    }
    Update-AsgCaptions
}

# Ключ выделенного узла дерева «Модели (по разделам)»: для модели — её имя, для раздела — название без числа моделей в скобках
function Get-AsgNodeKey($Node) {
    if ($null -eq $Node) { return "" }
    if ($Node.Nodes.Count -gt 0) { return ("S:" + ([string]$Node.Text -replace '\s*\(\d+\)\s*$', '')) }
    return ("M:" + [string]$Node.Tag)
}

# Модели отметили/сняли слева: у вновь отмеченных все спецификации отмечаются, итог и список пересчитываются
function Sync-AsgSpecsWithModels {
    $checked = @(Get-AsgCheckedModels)
    Update-PartialPicksOnModelCheck $script:AsgPicks $script:AsgAllSpecs $script:AsgModelSets $script:AsgPrevModels $checked
    $script:AsgPrevModels = $checked
    Show-AsgSpecList
}

# «все»/«ничего» справа — только для спецификаций выделенного узла
function Set-AsgNodeSpecs([bool]$On) {
    $models = @(Get-AsgSelectedModels)
    if ($models.Count -eq 0) { return }
    Set-PartialNodeAll $script:AsgPicks $models $script:AsgAllSpecs $script:AsgModelSets $On
    Show-AsgSpecList
}

# Для выгрузки: уникальные модели и спецификации из пар (прежние assign_models/assign_specs) и сами пары
function Get-AsgPairSummary($Pairs) {
    $m = New-Object System.Collections.Generic.List[string]; $sp = New-Object System.Collections.Generic.List[string]
    $mSeen = @{}; $sSeen = @{}
    foreach ($p in @($Pairs)) {
        if (-not $mSeen.ContainsKey($p.Model)) { $mSeen[$p.Model] = $true; $m.Add($p.Model) }
        if (-not $sSeen.ContainsKey($p.Spec)) { $sSeen[$p.Spec] = $true; $sp.Add($p.Spec) }
    }
    return @{ Models = $m.ToArray(); Specs = $sp.ToArray() }
}

# запоминаем текущие отметки (если списки уже построены): отмеченные модели и пары модель×спека
function Save-AsgSelection {
    if ($script:DistRestoring) { return }
    if ($ui.AsgTree.Nodes.Count -gt 0 -or @($script:AsgAllSpecs).Count -gt 0) {
        $m = @{}
        foreach ($x in @(Get-AsgCheckedModels)) { $m[$x] = $true }
        $pairs = @{}
        foreach ($model in @($script:AsgPicks.Keys)) {
            $pairs[($model + "|*")] = $true      # модель уже была в выборе: при следующем запуске её отметки берутся из сохранённого
            foreach ($sp in @($script:AsgPicks[$model].Keys)) { if ($script:AsgPicks[$model][$sp]) { $pairs[($model + "|" + $sp)] = $true } }
        }
        $script:AsgSel.Models = $m
        $script:AsgSel.Pairs = $pairs
        $script:AsgSel.Specs = $null      # прежний формат asgspecs= только читается (совместимость)
    }
}

# Диагностика наборов: у каких моделей набор спецификаций не определён (и почему) — в строку статуса, подсказку и журнал
function Write-AsgSetsDiagnostics($Why, $Models) {
    $logDir = Get-BvsToolLogDir $AppDir
    if (-not $logDir) { $logDir = $AppDir }
    if ($logDir -and -not (Test-Path -LiteralPath $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
    $log = Join-Path $logDir "partial_sets.log"
    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add(("{0}  моделей: {1}; набор не определён у: {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), @($Models).Count, @($Why.Keys).Count))
    foreach ($k in @($Why.Keys | Sort-Object)) { $lines.Add(("{0}: набор не определён — {1} (считаем, что у модели все спецификации)" -f $k, $Why[$k])) }
    try { [System.IO.File]::WriteAllLines($log, [string[]]$lines.ToArray(), (New-Object System.Text.UTF8Encoding $true)) } catch { }
    return $log
}

function Refresh-PartialLists {
    if (-not $ui.AsgTree) { return }
    if (-not $script:PartialUiReady) {
        $ui.AsgHint.Text = "Режим «Выборочно» недоступен: в папке инструмента нет assignment_partial.ps1 / assignment_partial_ui.ps1."
        $ui.AsgPart.Enabled = $false
        return
    }
    Save-AsgSelection
    $selKey = Get-AsgNodeKey $ui.AsgTree.SelectedNode          # выделенный узел запоминается и после перестроения восстанавливается
    $summary = $ui.Summary.Text.Trim()
    $choices = Get-PartialChoices (Get-HistoryPath $summary) (Get-LedgerPath $summary)
    $script:AsgBusy = $true
    try {
        $ui.AsgTree.BeginUpdate()
        $ui.AsgTree.Nodes.Clear()
        $groups = [ordered]@{}
        foreach ($m in @($choices.Models)) {
            if (-not $groups.Contains($m.SectionTitle)) { $groups[$m.SectionTitle] = New-Object System.Collections.Generic.List[object] }
            $groups[$m.SectionTitle].Add($m)
        }
        foreach ($title in @($groups.Keys)) {
            $parent = New-Object System.Windows.Forms.TreeNode ("{0} ({1})" -f $title, $groups[$title].Count)
            $allOn = $true
            foreach ($m in $groups[$title]) {
                $caption = $m.Name
                if (-not $m.Loaded) { $caption += "  (связь не загружена)" }
                $child = New-Object System.Windows.Forms.TreeNode $caption
                $child.Tag = $m.Name
                $on = ($null -eq $script:AsgSel.Models) -or $script:AsgSel.Models.ContainsKey($m.Name)
                $child.Checked = $on
                if (-not $on) { $allOn = $false }
                [void]$parent.Nodes.Add($child)
            }
            $parent.Checked = $allOn
            [void]$ui.AsgTree.Nodes.Add($parent)
            $parent.Expand()
        }
        $ui.AsgTree.EndUpdate()
        # выделение: прежний узел по имени, иначе первый раздел (иначе список спецификаций справа пуст, пока не щёлкнуть по дереву)
        $pick = $null
        if ($selKey) {
            foreach ($sec in @($ui.AsgTree.Nodes)) {
                if ((Get-AsgNodeKey $sec) -eq $selKey) { $pick = $sec; break }
                foreach ($c in @($sec.Nodes)) { if ((Get-AsgNodeKey $c) -eq $selKey) { $pick = $c; break } }
                if ($pick) { break }
            }
        }
        if ($null -eq $pick -and $ui.AsgTree.Nodes.Count -gt 0) { $pick = $ui.AsgTree.Nodes[0] }
        $ui.AsgTree.SelectedNode = $pick
        $script:AsgAllSpecs = @($choices.Specs)
        $script:AsgAllModels = @($choices.Models | ForEach-Object { $_.Name })
        # наборы — из распределения проекта (вкладка «Распределение»)
        $info = Get-PartialModelSpecInfo $script:AsgAllModels $script:Project
        $script:AsgModelSets = $info.Sets
        $script:AsgModelWhy = $info.Why
        $script:AsgPicks = New-PartialPicks $script:AsgAllModels $script:AsgAllSpecs $script:AsgModelSets $script:AsgSel.Pairs $script:AsgSel.Specs
        $script:AsgPrevModels = @(Get-AsgCheckedModels)     # сохранённые отметки не пересчитываем
        Show-AsgSpecList
    }
    finally { $script:AsgBusy = $false }
    $why = $script:AsgModelWhy
    $diag = ""
    if ($why -and $why.Count -gt 0 -and @($choices.Models).Count -gt 0) {
        $log = Write-AsgSetsDiagnostics $why $script:AsgAllModels
        $first = @($why.Keys | Sort-Object)[0]
        $diag = (" ВНИМАНИЕ: набор спецификаций не определён у {0} из {1} моделей (у них показаны все спецификации). Например, {2}: {3}. Все причины: {4}" -f $why.Count, @($choices.Models).Count, $first, $why[$first], $log)
        $ui.Status.Text = ("Выборочно: набор спецификаций не определён у {0} моделей — {1}: {2}" -f $why.Count, $first, $why[$first])
    }
    if (@($choices.Models).Count -eq 0) {
        $ui.AsgHint.Text = "Списков нет — сначала соберите задание в режиме «Всё»."
    }
    else {
        # подробное пояснение режима «Выборочно» — в ИНСТРУКЦИЯ.txt; в окне только диагностика, если она есть
        $ui.AsgHint.Text = $diag.Trim()
    }
}

function Set-AsgScopeEnabled {
    $on = ($ui.AsgPart.Checked -and $script:PartialUiReady)
    foreach ($c in @($ui.AsgTree, $ui.AsgSpecs, $ui.AsgModelsAll, $ui.AsgModelsNone, $ui.AsgSpecsAll, $ui.AsgSpecsNone)) { $c.Enabled = $on }
}

# Путь классификатора — в настройке проекта (Get-ProjectClassifierPath, project_config.ps1), а не в тексте set_assembly_code.py.
function Test-ClassifierReadable([string]$Path) {
    if ([string]::IsNullOrWhiteSpace($Path)) { return $false }
    $stream = $null
    try {
        if (-not [System.IO.File]::Exists($Path)) { return $false }
        $stream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
        return $true
    }
    catch { return $false }
    finally { if ($stream) { $stream.Dispose() } }
}

function Get-ClassifierRows([string]$CsvPath) {
    $rows = New-Object System.Collections.Generic.List[object]
    if (-not $CsvPath -or -not (Test-Path -LiteralPath $CsvPath)) { return $rows }
    $text = [System.IO.File]::ReadAllText($CsvPath, (New-Object System.Text.UTF8Encoding $true))
    $first = $true
    foreach ($line in @($text -split "`r?`n")) {
        if (-not $line.Trim()) { continue }
        if ($first) { $first = $false; continue }
        $p = @($line.Split(';'))
        while ($p.Count -lt 5) { $p += "" }
        $rows.Add(@([string]$p[0], [string]$p[1], [string]$p[2], [string]$p[3], [string]$p[4]))
    }
    return $rows
}

function Get-ClassifierReportPath {
    $day = Get-Date -Format "yyyy-MM-dd"
    $dir = Get-ProjectReportDir (Get-CurrentProjectName)
    $path = Join-Path $dir ($day + "_классификатор.xlsx")
    if (-not (Test-Path -LiteralPath $path)) { return $path }
    $n = 2
    while ($n -lt 100) {
        $path = Join-Path $dir ("{0}_классификатор_{1}.xlsx" -f $day, $n)
        if (-not (Test-Path -LiteralPath $path)) { return $path }
        $n++
    }
    return (Join-Path $dir ("{0}_классификатор_{1}.xlsx" -f $day, (Get-Date -Format "HHmmss")))
}

function Finish-ClassifierRun([int]$ExitCode) {
    $rows = @(Get-ClassifierRows $script:Run.ClassifierCsv)
    $updated = 0; $same = 0; $skipped = 0; $errors = 0
    foreach ($r in $rows) {
        switch ($r[3]) {
            "updated" { $updated++ }
            "already_set" { $same++ }
            "skipped_not_section" { $skipped++ }
            default { $errors++ }
        }
    }
    $xlsxPath = Get-ClassifierReportPath
    $script:ReportXlsxError = ""
    $xlsx = $null
    try {
        $sheetRows = New-Object System.Collections.Generic.List[object]
        foreach ($r in $rows) { $sheetRows.Add($r) }
        $sheets = @(@{
            Name = "Классификатор"
            Headers = @("Модель", "Прежний путь", "Новый путь", "Результат", "Ошибка")
            Rows = (ConvertTo-PlainArray $sheetRows)
            Widths = @(42, 78, 78, 24, 36)
            CountCols = @(); ParentRows = @(); Groups = @()
        })
        $xlsx = Write-ReportXlsx -Path $xlsxPath -Sheets $sheets
    }
    catch {
        $script:ReportXlsxError = $_.Exception.Message
        $xlsx = $null
    }
    if ($xlsx) { Set-OpenFolderTarget $xlsxPath; Set-CsvLink "" } else { Set-CsvLink $script:Run.ClassifierCsv }
    $fail = ""
    if ($rows.Count -eq 0) { $fail = ("Не готово: ни одна модель не обработана (код RBP {0})." -f $ExitCode) }
    elseif ($errors -gt 0) { $fail = ("Не готово: с ошибкой {0} из {1} моделей (код RBP {2})." -f $errors, $rows.Count, $ExitCode) }
    elseif ($ExitCode -ne 0) { $fail = ("Не готово: Revit Batch Processor завершился с кодом {0}; обработано моделей: {1}." -f $ExitCode, $rows.Count) }
    if ($fail) {
        $where = if ($xlsx) { " Отчёт: " + $xlsxPath } else { " xlsx не собран." }
        $ui.Status.Text = $fail + $where
        Show-Finished $fail
        return
    }
    $ui.Status.Text = ("Классификатор готов. Обновлено {0}, уже стоял {1}, пропущено {2}. Отчёт: {3}" -f $updated, $same, $skipped, $(if ($xlsx) { $xlsxPath } else { $script:Run.ClassifierCsv }))
    Show-Finished $ui.Status.Text
}

# Что показать перед запуском: проект и таблица Google, куда пойдёт результат (адрес — только конец, без секрета)
function Get-RunTargetText {
    $name = Get-CurrentProjectName
    $g = Resolve-GooglePushTarget $script:ProjectPath $ToolDir $false
    $where = switch ($g.Kind) {
        "project" { "Google: " + (Get-GoogleUrlTail $g.Cfg.WebAppUrl) }
        "needs-confirm" { "Google: у проекта нет своего адреса (общая таблица — только по подтверждению)" }
        default { "Google: не задан" }
    }
    return ("Проект {0}. {1}" -f $name, $where)
}

function Start-Batch {
    param([string]$Mode)
    if (-not (Confirm-DistDirtyLeave)) { return }
    Save-Settings
    $script:Run.Mode = $Mode
    $script:Run.TargetText = Get-RunTargetText          # в статусе на всё время запуска: проект и таблица Google
    $source = $ui.Source.Text.Trim()
    $projectPath = ""
    $report = $ReportDir
    $task = $TaskScript
    $chosen = New-Object System.Collections.Generic.List[string]
    $extraCfg = @()
    $partialPairs = @()
    $script:Run.Partial = $false
    if ($Mode -eq "assign") {
        $summary = $ui.Summary.Text.Trim()
        if (-not (Test-ModelReady $summary)) {
            [System.Windows.Forms.MessageBox]::Show($form, "Укажите сводную модель.", "Проверочные спецификации") | Out-Null
            return
        }
        $task = Join-Path $ToolDir "export_schedules.py"
        if (-not (Test-Path -LiteralPath $task)) {
            [System.Windows.Forms.MessageBox]::Show($form, "Нет скрипта выгрузки.", "Проверочные спецификации") | Out-Null
            return
        }
        if (-not (Test-BvsAssignmentBuilderReady $ToolDir)) {
            $msg = Get-BvsAssignmentBuilderMessage $ToolDir
            [System.Windows.Forms.MessageBox]::Show($form, $msg, "Проверочные спецификации") | Out-Null
            return
        }
        [void]$chosen.Add($summary)
        # «Выборочно»: только отмеченные спецификации и модели; нужна база прошлых сборок (буфер, журнал или история)
        if ($ui.AsgPart -and $ui.AsgPart.Checked) {
            if (-not $script:PartialUiReady) {
                [System.Windows.Forms.MessageBox]::Show($form, "Для режима «Выборочно» нужны файлы assignment_partial.ps1, assignment_partial_ui.ps1 и assignment_export_stream.ps1 в папке инструмента. Скопируйте папку целиком.", "Проверочные спецификации") | Out-Null
                return
            }
            $histForPartial = Get-HistoryPath $summary
            if (-not (Test-PartialBaseReady $histForPartial (Get-LedgerPath $summary))) {
                [System.Windows.Forms.MessageBox]::Show($form, "Режим «Выборочно» дополняет прошлые проверки, а прошлых сборок ещё нет. Сначала сделайте полную выгрузку (режим «Всё»).", "Проверочные спецификации") | Out-Null
                return
            }
            # выгружаются пары «модель × спецификация»: отмеченные модели и отмеченные у них спецификации
            $partialPairs = @(Get-AsgPairs)
            if ($partialPairs.Count -eq 0) {
                [System.Windows.Forms.MessageBox]::Show($form, "В режиме «Выборочно» отметьте хотя бы одну модель и хотя бы одну её спецификацию (кликните по модели слева и отметьте спецификации справа).", "Проверочные спецификации") | Out-Null
                return
            }
            $script:Run.Partial = $true
        }
    }
    elseif ($Mode -eq "classifier") {
        # Тот же список, что у «В разделы». Общий файл спецификаций не нужен. Сводная не открывается.
        $task = Join-Path $ToolDir "set_assembly_code.py"
        if (-not (Test-Path -LiteralPath $task)) {
            [System.Windows.Forms.MessageBox]::Show($form, "Нет скрипта set_assembly_code.py.", "Проверочные спецификации") | Out-Null
            return
        }
        # путь классификатора — из настройки проекта (classifier= в распределение_<проект>.txt); пусто — для PROJ официальный путь (Get-ProjectClassifierPath)
        $wantProject = Get-CurrentProjectName
        $classifierPath = Get-ProjectClassifierPath $script:Project $wantProject
        if (-not $wantProject -or -not $classifierPath) {
            [System.Windows.Forms.MessageBox]::Show($form, ("Для проекта «{0}» не задан путь классификатора. Впишите строку classifier=<UNC-путь> в раздел [project] файла отчеты\проект\распределение_{0}.txt." -f $wantProject), "Проверочные спецификации") | Out-Null
            return
        }
        if (-not $classifierPath.StartsWith('\\')) {
            [System.Windows.Forms.MessageBox]::Show($form, ("Путь классификатора должен быть абсолютным UNC.`r`n`r`n" + $classifierPath), "Проверочные спецификации") | Out-Null
            return
        }
        if (-not (Test-ClassifierReadable $classifierPath)) {
            [System.Windows.Forms.MessageBox]::Show($form, ("Файл классификатора не найден или не читается. Модели не открываются.`r`n`r`n" + $classifierPath), "Проверочные спецификации") | Out-Null
            return
        }
        if (-not $script:ProjectPath -or $null -eq $script:Project) {
            [System.Windows.Forms.MessageBox]::Show($form, "Сначала выберите проект и нажмите «Загрузить спецификации и модели».", "Проверочные спецификации") | Out-Null
            return
        }
        $mp = Get-DistProjectModelPaths
        if (@($mp.Missing).Count -gt 0) {
            $missingNote = ("нет файла в папке: {0}." -f (@($mp.Missing) -join ", "))
            $ui.Status.Text = $missingNote
            [System.Windows.Forms.MessageBox]::Show($form, ($missingNote + " Проверьте поле «Модели»."), "Проверочные спецификации") | Out-Null
            return
        }
        $prefix = $wantProject.ToUpperInvariant() + "_"
        $extraCfg = @(("classifier_path={0}" -f $classifierPath), ("classifier_project={0}" -f $wantProject))      # set_assembly_code.py: нет строк — встроенный CLASSIFIER_PATH
        foreach ($x in @($mp.Paths)) {
            $path = [string]$x
            if (Test-ExcludedModelPath $path) { continue }
            $leaf = ""
            try { $leaf = [System.IO.Path]::GetFileNameWithoutExtension($path) } catch { $leaf = "" }
            if (-not $leaf.ToUpperInvariant().StartsWith($prefix)) { continue }
            [void]$chosen.Add($path)
        }
        if ($chosen.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show($form, "Нет моделей разделов. Нажмите «Загрузить спецификации и модели».", "Проверочные спецификации") | Out-Null
            return
        }
    }
    else {
        if (-not $source -or -not (Test-ModelReady $source)) {          # общий файл может быть RSN://
            [System.Windows.Forms.MessageBox]::Show($form, "Укажите файл со спецификациями.", "Проверочные спецификации") | Out-Null
            return
        }
        # то, что на экране, и есть то, что вставится: настройка сохраняется перед запуском; пустой путь — не запускаем
        [void](Save-DistProject)
        $blocker = Get-ProjectRunBlocker $script:ProjectPath
        if ($blocker) {
            [System.Windows.Forms.MessageBox]::Show($form, $blocker, "Проверочные спецификации") | Out-Null
            return
        }
        $projectPath = $script:ProjectPath
        if ($Mode -eq "summary") {
            $summary = $ui.Summary.Text.Trim()
            if (-not (Test-ModelReady $summary)) {
                [System.Windows.Forms.MessageBox]::Show($form, "Укажите сводную модель.", "Проверочные спецификации") | Out-Null
                return
            }
            [void]$chosen.Add($summary)
        }
        else {
            # модели разделов — из распределения («Загрузить спецификации и модели»)
            $mp = Get-DistProjectModelPaths
            foreach ($x in $mp.Paths) { [void]$chosen.Add([string]$x) }
            if ($chosen.Count -eq 0) {
                [System.Windows.Forms.MessageBox]::Show($form, "Нет моделей разделов. Нажмите «Загрузить спецификации и модели».", "Проверочные спецификации") | Out-Null
                return
            }
            if (@($mp.Missing).Count -gt 0) {
                $missingNote = ("нет файла в папке: {0}." -f (@($mp.Missing) -join ", "))
                $ui.Status.Text = $missingNote
                [System.Windows.Forms.MessageBox]::Show($form, ($missingNote + " Проверьте поле «Модели»."), "Проверочные спецификации") | Out-Null
                return
            }
        }
        if (-not (Test-Path -LiteralPath $TaskScript)) {
            [System.Windows.Forms.MessageBox]::Show($form, "Нет скрипта вставки.", "Проверочные спецификации") | Out-Null
            return
        }
    }
    if (($Mode -eq "summary" -or $Mode -eq "sections") -and (Get-InsertModeValue) -eq "replace") {
        $ask = [System.Windows.Forms.MessageBox]::Show($form, (Get-ReplaceConfirmText $Mode @($chosen)), "Заменить все", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning, [System.Windows.Forms.MessageBoxDefaultButton]::Button2)
        if ($ask -ne [System.Windows.Forms.DialogResult]::Yes) { $ui.Status.Text = "Заменить все: отменено."; return }
    }
    Set-CsvLink ""
    $exe = Find-BatchRvt
    if (-not $exe) {
        [System.Windows.Forms.MessageBox]::Show(
            "Не найден Revit Batch Processor (BatchRvt.exe). Установите RBP и откройте его один раз.",
            "Проверочные спецификации"
        ) | Out-Null
        return
    }
    $batch = Join-Path (Split-Path -Parent $exe) "BatchRvt.exe"
    if (Test-Path -LiteralPath $batch) { $exe = $batch }

    $hasRsn = @(@($chosen) + @($(if ($Mode -eq "summary" -or $Mode -eq "sections") { $source })) | Where-Object { $_ -match '^(?i)RSN://' }).Count -gt 0          # и когда RSN:// только у общего файла
    if ($hasRsn) {
        $ui.Status.Text = "Список с Revit Server. Подключаю поддержку RSN://."
        [System.Windows.Forms.Application]::DoEvents()
        if (-not (Enable-RsnForRbp $exe)) {
            [System.Windows.Forms.MessageBox]::Show(
                $form,
                "Путь RSN:// без правки Revit Batch Processor не откроется: программа считает, что файла нет. Правка из папки «служебное» сейчас не применилась.",
                "Проверочные спецификации"
            ) | Out-Null
            $ui.Status.Text = "RSN:// не подключён. Запуск остановлен."
            return
        }
    }

    # версия Revit — по самим моделям; несогласованные версии или общий файл новее моделей — остановка до запуска
    $verTemplate = ""
    if ($Mode -eq "summary" -or $Mode -eq "sections") { $verTemplate = $source }
    $ver = Resolve-RevitRun @($chosen) $verTemplate (Get-BvsServerYears $ServersCfgPath)
    if ($ver.Problem) {
        [System.Windows.Forms.MessageBox]::Show($form, $ver.Problem, "Проверочные спецификации") | Out-Null
        $ui.Status.Text = $ver.Problem
        return
    }
    $runId = Get-Date -Format "yyyyMMdd_HHmmss"
    # файлы прогона — в отчеты\служебные\<время>; логи RBP этого запуска — в отчеты\логи (--log_folder только здесь, настройки RBP не меняем)
    try { [void](Invoke-BvsHousekeeping -ReportDir $ReportDir -KeepDays $script:AuxKeepDays -AppDir $AppDir) } catch { }
    $work = New-RunAuxDir $script:Layout.Aux $runId
    if ($Mode -eq "classifier") {
        $clfCsv = Join-Path $work ("{0}_классификатор.csv" -f $runId)
        $bomEnc = New-Object System.Text.UTF8Encoding $true
        [System.IO.File]::WriteAllText($clfCsv, "Модель;Прежний путь;Новый путь;Результат;Ошибка`r`n", $bomEnc)
        $script:Run.ClassifierCsv = $clfCsv
    }
    if ($script:Run.Partial -and $partialPairs) {
        # точные пары — в файл служебной папки прогона (читает export_schedules.py через .NET); assign_models/assign_specs — их объединение
        $pairsFile = Join-Path $work "assign_pairs.txt"
        Write-PartialPairsFile $pairsFile $partialPairs
        $pairSummary = Get-AsgPairSummary $partialPairs
        $extraCfg = ConvertTo-PartialCfgLines $pairSummary.Specs $pairSummary.Models $pairsFile
    }
    if ($Mode -eq "summary" -or $Mode -eq "sections") {
        # готовый план вставки: insert_schedules.py не разбирает файл настройки сам. Не записался — без plan_path (скрипт разберёт файл настройки, как раньше)
        try {
            $planPath = Write-InsertPlan $projectPath (Join-Path $work "plan.txt")
            $extraCfg = @($extraCfg) + @(("plan_path={0}" -f $planPath))
        }
        catch { Write-Warning ("plan.txt не записан: " + $_.Exception.Message) }
    }
    $scope = if ($Mode -eq "sections") { "sections" } else { "summary" }
    $insertMode = Get-InsertModeValue
    $files = Write-RunFiles -Models @($chosen) -Source $source -ReportDir $work -RunId $runId -Mode $Mode -ProjectPath $projectPath -PrefixFrom "" -PrefixTo "" -InsertMode $insertMode -Extra $extraCfg -InsertReport:($Mode -eq "summary" -or $Mode -eq "sections")

    $worksets = "close_all"
    if ($Mode -eq "assign") { $worksets = "open_all" }
    # PS 5.1 Start-Process не кавычит элементы массива с пробелами — Quote-BvsProcessArg только для путей с пробелами
    $logDir = $script:Layout.Logs
    if ($logDir -and -not (Test-Path -LiteralPath $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
    $args = @(
        "--task_script", (Quote-BvsProcessArg $task),
        "--file_list", (Quote-BvsProcessArg $files.List),
        "--create_new_local",
        "--worksets", $worksets,
        "--revit_version", [string]$ver.Year,          # всегда: версию определили по моделям
        "--log_folder", (Quote-BvsProcessArg $logDir)  # только этот запуск BVS; глобальную папку логов RBP не меняем
    )

    $script:Run.Summary = $files.Summary
    $script:Run.Params = $files.Params
    $script:Run.Skips = $files.Skips
    $script:Run.Inserted = $files.Inserted
    $script:Run.ReportDir = $report
    $script:Run.WorkDir = $work
    $script:Run.RunId = $runId
    Set-ActionsEnabled $false
    if ($Mode -eq "summary") {
        $script:Run.WaitText = ("{0}. Идёт загрузка в сводную. Revit откроется поверх. Окно не закрывай." -f (Get-RunTargetText))
    }
    elseif ($Mode -eq "assign") {
        $script:Run.WaitText = ("{0}. Идёт выгрузка из сводной. Revit откроется поверх. Окно не закрывай." -f (Get-RunTargetText))
        if ($script:Run.Partial) { $script:Run.WaitText = ("{0}. Идёт выборочная выгрузка из сводной. Revit откроется поверх. Окно не закрывай." -f (Get-RunTargetText)) }
    }
    elseif ($Mode -eq "classifier") {
        $script:Run.WaitText = ("{0}. Задаю код по классификатору в моделях разделов. Revit откроется поверх. Окно не закрывай." -f (Get-RunTargetText))
    }
    else {
        $script:Run.WaitText = ("{0}. Идёт загрузка в разделы. Revit откроется поверх. Окно не закрывай." -f (Get-RunTargetText))
    }
    $ui.Status.Text = $script:Run.WaitText
    $livePath = Join-Path $AppDir "live_status.txt"
    if (Test-Path -LiteralPath $livePath) { Remove-Item -LiteralPath $livePath -Force -ErrorAction SilentlyContinue }
    $p = Start-Process -FilePath $exe -ArgumentList $args -WorkingDirectory (Split-Path -Parent $exe) -PassThru -WindowStyle Hidden
    $script:Run.Process = $p
    $ui.Timer.Start()
}

$ui = @{}
. (Join-Path $ToolDir "project_dist_ui.ps1")

$form = New-Object System.Windows.Forms.Form
$form.Text = "Проверочные спецификации"
$form.Size = New-Object System.Drawing.Size(1040, 700)
$form.MinimumSize = New-Object System.Drawing.Size(960, 640)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$form.BackColor = [System.Drawing.SystemColors]::Window

# ---- раскладка вкладок: TableLayoutPanel, чтобы страницы тянулись вслед за окном ----
function New-LayoutTable {
    $t = New-Object System.Windows.Forms.TableLayoutPanel
    $t.Dock = "Fill"
    $t.Padding = New-Object System.Windows.Forms.Padding(8)
    $t.ColumnCount = 0
    $t.RowCount = 0
    $t.BackColor = [System.Drawing.SystemColors]::Window
    $t.Tag = New-Object System.Collections.ArrayList     # подписи, которые переносятся по ширине (Register-WrapControl)
    $t.Add_Resize({
        param($sender, $e)
        $w = $sender.ClientSize.Width - $sender.Padding.Horizontal
        foreach ($c in @($sender.Tag)) {
            $mw = $w - $c.Margin.Horizontal - 8
            if ($mw -gt 120) { $c.MaximumSize = New-Object System.Drawing.Size($mw, 0) }
        }
    })
    return $t
}

function Add-LayoutColumn($Table, [string]$Kind, [double]$Size) {
    if ($Kind -eq "percent") { $cs = New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, [single]$Size) }
    else { $cs = New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, [single]$Size) }
    $Table.ColumnCount = $Table.ColumnCount + 1
    [void]$Table.ColumnStyles.Add($cs)
}

function Add-LayoutRow($Table, [bool]$Fill = $false) {
    $idx = $Table.RowCount
    $Table.RowCount = $idx + 1
    if ($Fill) { $rs = New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, [single]100) }
    else { $rs = New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize) }
    [void]$Table.RowStyles.Add($rs)
    return $idx
}

function Add-LayoutControl($Table, $Control, [int]$Col, [int]$Row, [int]$Span = 1) {
    $Table.Controls.Add($Control, $Col, $Row)
    if ($Span -gt 1) { $Table.SetColumnSpan($Control, $Span) }
}

# длинный текст переносится по ширине окна, а не обрезается
function Register-WrapControl($Table, $Control) {
    $Control.AutoSize = $true
    [void]$Table.Tag.Add($Control)
}

# «Открыть отчёт»: выделяет в проводнике последний записанный xlsx (или открывает папку отчетов, пока файла нет)
function Set-OpenFolderTarget([string]$Path) { $script:LastOpenPath = $Path }
function Open-ReportFolder {
    $p = $script:LastOpenPath
    try {
        if ($p -and (Test-Path -LiteralPath $p -PathType Leaf)) { Start-Process -FilePath "explorer.exe" -ArgumentList ('/select,"{0}"' -f $p) }
        else {
            $target = $ReportDir
            if ($script:ProjectPath) {
                $pd = Join-Path $ReportDir (((Get-CurrentProjectName) -replace '[\\/:*?"<>|]', '_'))
                if (Test-Path -LiteralPath $pd -PathType Container) { $target = $pd }
            }
            Start-Process -FilePath "explorer.exe" -ArgumentList ('"{0}"' -f $target)
        }
    }
    catch { $ui.Status.Text = ("Не удалось открыть папку: {0}" -f $_.Exception.Message) }
}

# «Открыть CSV»: когда xlsx отчёта не собран. Показывает CSV сводки прогона в проводнике (Excel не запускается)
$script:CsvPath = ""
function Set-CsvLink([string]$Path) {
    $script:CsvPath = $Path
    if ($ui.ContainsKey("OpenCsv") -and $ui.OpenCsv) { $ui.OpenCsv.Visible = [bool]($Path -and (Test-Path -LiteralPath $Path)) }
}
function Open-RunCsv {
    $p = $script:CsvPath
    try {
        if ($p -and (Test-Path -LiteralPath $p -PathType Leaf)) { Start-Process -FilePath "explorer.exe" -ArgumentList ('/select,"{0}"' -f $p) }
        elseif ($p) { Start-Process -FilePath "explorer.exe" -ArgumentList ('"{0}"' -f (Split-Path -Parent $p)) }
    }
    catch { $ui.Status.Text = ("Не удалось открыть CSV: {0}" -f $_.Exception.Message) }
}

# «Логи»: папка отчеты\логи (журналы этого скрипта, не системная папка RBP)
function Open-BvsLogFolder {
    $target = $script:Layout.Logs
    if (-not $target) { $target = Join-Path $ReportDir "логи" }
    try {
        if (-not (Test-Path -LiteralPath $target)) { New-Item -ItemType Directory -Path $target -Force | Out-Null }
        Start-Process -FilePath "explorer.exe" -ArgumentList ('"{0}"' -f $target)
    }
    catch { $ui.Status.Text = ("Не удалось открыть папку логов: {0}" -f $_.Exception.Message) }
}

# «Заменить все»: что будет удалено и вставлено заново. Модели без спецификаций списка получат первичную загрузку (без удаления).
# ModelNames — имена/пути моделей; для сводной — спецификации «В проекте» и «Сводная», для разделов — отмеченные у каждой модели
function Get-ReplacePlan([string]$Mode, $ModelNames) {
    $models = @($ModelNames).Count
    $specs = 0
    if ($Mode -eq "summary") { $specs = @(Get-ProjectSummarySpecs $script:Project).Count }
    else { foreach ($m in @($ModelNames)) { $specs += @(Get-ProjectEnabledSpecs $script:Project (Get-ProjectModelName ([string]$m))).Count } }
    return @{ Models = $models; Specs = $specs }
}

function Get-ReplaceConfirmText([string]$Mode, $ModelNames) {
    $p = Get-ReplacePlan $Mode $ModelNames
    return ("Режим «Заменить все».{0}Моделей: {1}. Спецификаций списка: {2} (по всем моделям).{0}{0}В моделях они будут удалены и вставлены заново; если вид стоял на листе, он пропадёт. Модели, где спецификаций списка ещё нет, получат первичную загрузку без удаления.{0}{0}Продолжить?" -f [Environment]::NewLine, $p.Models, $p.Specs)
}

function Show-Finished([string]$Text) {
    try {
        $form.WindowState = [System.Windows.Forms.FormWindowState]::Normal
        $form.TopMost = $true
        $form.Activate()
        $form.TopMost = $false
    }
    catch { }
    [System.Windows.Forms.MessageBox]::Show($form, $Text, "Проверочные спецификации") | Out-Null
}

$ui.PageAssign = New-Object System.Windows.Forms.Panel
$ui.PageAssign.Dock = "Fill"
$ui.PageAssign.BackColor = [System.Drawing.SystemColors]::Window
$lblAsg = New-Object System.Windows.Forms.Label
$lblAsg.Text = "Задание"
$lblAsg.Location = New-Object System.Drawing.Point(8, 16)
$lblAsg.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$lblAsg.AutoSize = $true
$lblDyn = New-Object System.Windows.Forms.Label
$lblDyn.Text = "Интервал для листа «Динамика»"
$lblDyn.Location = New-Object System.Drawing.Point(8, 92)
$lblDyn.AutoSize = $true
$lblDyn.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$ui.DynDays = New-Object System.Windows.Forms.ComboBox
$ui.DynDays.Location = New-Object System.Drawing.Point(8, 116)
$ui.DynDays.Size = New-Object System.Drawing.Size(120, 26)
$ui.DynDays.DropDownStyle = "DropDownList"
[void]$ui.DynDays.Items.AddRange(@("7 дней", "14 дней", "30 дней"))
$ui.DynCustom = New-Object System.Windows.Forms.CheckBox
$ui.DynCustom.Text = "Свои даты"
$ui.DynCustom.Location = New-Object System.Drawing.Point(140, 118)
$ui.DynCustom.AutoSize = $true
$ui.DynFrom = New-Object System.Windows.Forms.DateTimePicker
$ui.DynFrom.Location = New-Object System.Drawing.Point(248, 114)
$ui.DynFrom.Size = New-Object System.Drawing.Size(120, 26)
$ui.DynFrom.Format = "Short"
$ui.DynFrom.Enabled = $false
$ui.DynTo = New-Object System.Windows.Forms.DateTimePicker
$ui.DynTo.Location = New-Object System.Drawing.Point(380, 114)
$ui.DynTo.Size = New-Object System.Drawing.Size(120, 26)
$ui.DynTo.Format = "Short"
$ui.DynTo.Enabled = $false
$ui.DynBuild = New-Object System.Windows.Forms.Button
$ui.DynBuild.Text = "Собрать динамику"
$ui.DynBuild.Location = New-Object System.Drawing.Point(8, 148)
$ui.DynBuild.Size = New-Object System.Drawing.Size(200, 32)
$ui.DynCustom.Add_CheckedChanged({
    $on = $ui.DynCustom.Checked
    $ui.DynFrom.Enabled = $on
    $ui.DynTo.Enabled = $on
    $ui.DynDays.Enabled = -not $on
})
$ui.DynBuild.Add_Click({ Start-DynamicsOnly })
$lblUrl = New-Object System.Windows.Forms.Label
$lblUrl.Text = "Ссылка на Google-таблицу проекта"
$lblUrl.Location = New-Object System.Drawing.Point(8, 192)
$lblUrl.AutoSize = $true
$ui.Google = New-Object System.Windows.Forms.TextBox
$ui.Google.Location = New-Object System.Drawing.Point(8, 216)
$ui.Google.Size = New-Object System.Drawing.Size(640, 26)
# Адрес веб-приложения Google (…/exec) и секрет — для этого проекта: хранятся в google_<проект>.cfg рядом с распределение_<проект>.txt (в git не попадает)
$ui.GoogleInfo = New-Object System.Windows.Forms.Label
$ui.GoogleInfo.AutoSize = $true
$ui.GoogleInfo.Text = "Google: нет проекта"
$ui.GoogleCheck = New-Object System.Windows.Forms.Button
$ui.GoogleCheck.Text = "Проверить"
$ui.GoogleCheck.Size = New-Object System.Drawing.Size(100, 28)
$ui.GoogleConnect = New-Object System.Windows.Forms.Button
$ui.GoogleConnect.Text = "Подключить Google…"
$ui.GoogleConnect.AutoSize = $true
$ui.GoogleConnect.MinimumSize = New-Object System.Drawing.Size(150, 28)
Set-UiTip $ui.GoogleInfo "Адрес и секрет читаются из отчеты\проект\google_<проект>.cfg; секрет нигде не показывается."
Set-UiTip $ui.GoogleCheck "Проверка связи с веб-приложением: ping и пробный POST без листов (как проверить_google.ps1)."
Set-UiTip $ui.GoogleConnect "Адрес веб-приложения (…/exec) и секрет этого проекта. Сохраняются в google_<проект>.cfg."
$ui.CollapseDay = New-Object System.Windows.Forms.CheckBox
$ui.CollapseDay.Text = "Одна сборка в день: заменять сегодняшние прежние сборки (иначе каждая сборка — отдельный срез истории)"
$ui.CollapseDay.Location = New-Object System.Drawing.Point(8, 284)
$ui.CollapseDay.AutoSize = $true
Set-UiTip $ui.Google "Ссылка для чата. Автообновление таблицы — google_push.cfg в служебное\google_push (см. ИНСТРУКЦИЯ.txt)."
Set-UiTip $ui.DynBuild "Только лист «Динамика» из истории, без Revit."
# ---- что выгружать: «Всё» (как раньше) или «Выборочно» ----
$ui.PageAssign.AutoScroll = $true
$lblScope = New-Object System.Windows.Forms.Label
$lblScope.Text = "Что выгружать"
$lblScope.Location = New-Object System.Drawing.Point(8, 320)
$lblScope.AutoSize = $true
$lblScope.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$ui.AsgAll = New-Object System.Windows.Forms.RadioButton
$ui.AsgAll.Text = "Всё: все спецификации и все модели"
$ui.AsgAll.Location = New-Object System.Drawing.Point(8, 346)
$ui.AsgAll.Size = New-Object System.Drawing.Size(330, 24)
$ui.AsgAll.Checked = $true
$ui.AsgPart = New-Object System.Windows.Forms.RadioButton
$ui.AsgPart.Text = "Выборочно: отмеченные ниже"
$ui.AsgPart.Location = New-Object System.Drawing.Point(350, 346)
$ui.AsgPart.Size = New-Object System.Drawing.Size(330, 24)
$lblAsgModels = New-Object System.Windows.Forms.Label
$lblAsgModels.Text = "Модели (по разделам)"
$lblAsgModels.Location = New-Object System.Drawing.Point(8, 380)
$lblAsgModels.AutoSize = $true
$ui.AsgModelsAll = New-Object System.Windows.Forms.Button
$ui.AsgModelsAll.Text = "все"
$ui.AsgModelsAll.Location = New-Object System.Drawing.Point(190, 374)
$ui.AsgModelsAll.Size = New-Object System.Drawing.Size(70, 26)
$ui.AsgModelsNone = New-Object System.Windows.Forms.Button
$ui.AsgModelsNone.Text = "ничего"
$ui.AsgModelsNone.Location = New-Object System.Drawing.Point(266, 374)
$ui.AsgModelsNone.Size = New-Object System.Drawing.Size(70, 26)
$ui.AsgTree = New-Object System.Windows.Forms.TreeView
$ui.AsgTree.CheckBoxes = $true
$ui.AsgTree.HideSelection = $false
$ui.AsgTree.Location = New-Object System.Drawing.Point(8, 404)
$ui.AsgTree.Size = New-Object System.Drawing.Size(330, 220)
$ui.AsgTree.Dock = "Fill"
$ui.AsgTree.Margin = New-Object System.Windows.Forms.Padding(3)
$lblAsgSpecs = New-Object System.Windows.Forms.Label
$lblAsgSpecs.Text = "Спецификации (по номерам)"
$lblAsgSpecs.Location = New-Object System.Drawing.Point(350, 380)
$lblAsgSpecs.AutoSize = $true
$ui.AsgSpecsAll = New-Object System.Windows.Forms.Button
$ui.AsgSpecsAll.Text = "все"
$ui.AsgSpecsAll.Location = New-Object System.Drawing.Point(560, 374)
$ui.AsgSpecsAll.Size = New-Object System.Drawing.Size(70, 26)
$ui.AsgSpecsNone = New-Object System.Windows.Forms.Button
$ui.AsgSpecsNone.Text = "ничего"
$ui.AsgSpecsNone.Location = New-Object System.Drawing.Point(636, 374)
$ui.AsgSpecsNone.Size = New-Object System.Drawing.Size(70, 26)
$ui.AsgSpecs = New-Object System.Windows.Forms.CheckedListBox
$ui.AsgSpecs.CheckOnClick = $true
$ui.AsgSpecs.HorizontalScrollbar = $true
$ui.AsgSpecs.IntegralHeight = $false    # иначе высота урезается до целого числа строк и низ списка не совпадает с деревом моделей
$ui.AsgSpecs.Location = New-Object System.Drawing.Point(350, 404)
$ui.AsgSpecs.Size = New-Object System.Drawing.Size(356, 220)
$ui.AsgSpecs.Dock = "Fill"
$ui.AsgSpecs.Margin = New-Object System.Windows.Forms.Padding(3)
$ui.AsgTotal = New-Object System.Windows.Forms.Label
$ui.AsgTotal.Text = "Всего к выгрузке: 0 пар, 0 моделей"
$ui.AsgTotal.AutoSize = $true
$ui.AsgTotal.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$ui.AsgSpecCount = New-Object System.Windows.Forms.Label
$ui.AsgSpecCount.Text = "Выберите модель или раздел"
$ui.AsgSpecCount.AutoSize = $true
$ui.AsgSpecCount.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
$ui.AsgHint = New-Object System.Windows.Forms.Label
$ui.AsgHint.Location = New-Object System.Drawing.Point(8, 632)
$ui.AsgHint.AutoSize = $true
$ui.AsgHint.ForeColor = [System.Drawing.SystemColors]::GrayText
$ui.AsgTree.Add_AfterCheck({
    param($sender, $e)
    if ($script:AsgBusy) { return }
    if (-not $script:DistLoading -and -not $script:DistRestoring) { Enter-DistDirty }
    $script:AsgBusy = $true
    try {
        $node = $e.Node
        if ($node.Nodes.Count -gt 0) {
            # галочка на разделе — на все его модели
            foreach ($c in @($node.Nodes)) { $c.Checked = $node.Checked }
        }
        elseif ($node.Parent) {
            $all = $true
            foreach ($c in @($node.Parent.Nodes)) { if (-not $c.Checked) { $all = $false } }
            $node.Parent.Checked = $all
        }
    }
    finally { $script:AsgBusy = $false }
    Sync-AsgSpecsWithModels
})
$ui.AsgModelsAll.Add_Click({
    if (-not $script:DistLoading -and -not $script:DistRestoring) { Enter-DistDirty }
    $script:AsgBusy = $true
    try { foreach ($sec in @($ui.AsgTree.Nodes)) { $sec.Checked = $true; foreach ($c in @($sec.Nodes)) { $c.Checked = $true } } }
    finally { $script:AsgBusy = $false }
    Sync-AsgSpecsWithModels
})
$ui.AsgModelsNone.Add_Click({
    if (-not $script:DistLoading -and -not $script:DistRestoring) { Enter-DistDirty }
    $script:AsgBusy = $true
    try { foreach ($sec in @($ui.AsgTree.Nodes)) { $sec.Checked = $false; foreach ($c in @($sec.Nodes)) { $c.Checked = $false } } }
    finally { $script:AsgBusy = $false }
    Sync-AsgSpecsWithModels
})
# Выделение (клик) модели или раздела слева — справа спецификации этого узла
$ui.AsgTree.Add_AfterSelect({
    if ($script:AsgBusy) { return }
    try { Show-AsgSpecList }
    catch { $ui.Status.Text = ("Список спецификаций не показан: {0}" -f $_.Exception.Message) }
})
# щелчок по узлу (в т.ч. по галочке) выделяет его: справа сразу спецификации этого узла
$ui.AsgTree.Add_NodeMouseClick({
    param($sender, $e)
    if ($e.Node -and -not [object]::ReferenceEquals($ui.AsgTree.SelectedNode, $e.Node)) { $ui.AsgTree.SelectedNode = $e.Node }
})
# Галочка спецификации: у модели — её отметка; у раздела — у всех его моделей, где спека есть. Из «частично» щелчок отмечает у всех.
# ItemCheck приходит до смены значения: берём NewValue (для Indeterminate принудительно Checked)
$ui.AsgSpecs.Add_ItemCheck({
    param($sender, $e)
    if ($script:AsgSpecBusy) { return }
    if (-not $script:DistLoading -and -not $script:DistRestoring) { Enter-DistDirty }
    $idx = [int]$e.Index
    if ($idx -lt 0 -or $idx -ge @($script:AsgShown).Count) { return }
    if ($e.CurrentValue -eq [System.Windows.Forms.CheckState]::Indeterminate) { $e.NewValue = [System.Windows.Forms.CheckState]::Checked }
    $on = ($e.NewValue -eq [System.Windows.Forms.CheckState]::Checked)
    Set-PartialSpecPick $script:AsgPicks @(Get-AsgSelectedModels) $script:AsgModelSets ([string]$script:AsgShown[$idx]) $on
    Update-AsgCaptions
})
# «все»/«ничего» справа — только для выбранного узла
$ui.AsgSpecsAll.Add_Click({ if (-not $script:DistLoading -and -not $script:DistRestoring) { Enter-DistDirty }; Set-AsgNodeSpecs $true })
$ui.AsgSpecsNone.Add_Click({ if (-not $script:DistLoading -and -not $script:DistRestoring) { Enter-DistDirty }; Set-AsgNodeSpecs $false })
$ui.AsgPart.Add_CheckedChanged({ if (-not $script:DistLoading -and -not $script:DistRestoring) { Enter-DistDirty }; Set-AsgScopeEnabled })
# ---- раскладка вкладки «Задание»: две колонки по 50 %, списки тянутся по ширине и высоте, тексты переносятся ----
$asg = New-LayoutTable
Add-LayoutColumn $asg "percent" 50
Add-LayoutColumn $asg "percent" 50
$tm = New-Object System.Windows.Forms.Padding(3, 3, 3, 3)

$r = Add-LayoutRow $asg
$lblAsg.Margin = New-Object System.Windows.Forms.Padding(3, 6, 3, 3)
Add-LayoutControl $asg $lblAsg 0 $r 2

$r = Add-LayoutRow $asg
$lblDyn.Margin = New-Object System.Windows.Forms.Padding(3, 12, 3, 3)
Add-LayoutControl $asg $lblDyn 0 $r 2
$dynFlow = New-Object System.Windows.Forms.FlowLayoutPanel
$dynFlow.FlowDirection = "LeftToRight"
$dynFlow.WrapContents = $false
$dynFlow.AutoSize = $true
$dynFlow.Anchor = "Left"
$dynFlow.Margin = New-Object System.Windows.Forms.Padding(0)
foreach ($c in @($ui.DynDays, $ui.DynCustom, $ui.DynFrom, $ui.DynTo)) { $c.Margin = New-Object System.Windows.Forms.Padding(3, 4, 8, 3); $dynFlow.Controls.Add($c) }
$r = Add-LayoutRow $asg
Add-LayoutControl $asg $dynFlow 0 $r 2
$ui.DynBuild.Margin = New-Object System.Windows.Forms.Padding(3, 6, 3, 3)
$r = Add-LayoutRow $asg
Add-LayoutControl $asg $ui.DynBuild 0 $r 2

$r = Add-LayoutRow $asg
$lblUrl.Margin = New-Object System.Windows.Forms.Padding(3, 12, 3, 3)
Add-LayoutControl $asg $lblUrl 0 $r 2
$r = Add-LayoutRow $asg
$ui.Google.Anchor = "Left, Right"
$ui.Google.Margin = $tm
Add-LayoutControl $asg $ui.Google 0 $r 2
$r = Add-LayoutRow $asg
$gFlow = New-Object System.Windows.Forms.FlowLayoutPanel
$gFlow.AutoSize = $true; $gFlow.Dock = "Fill"; $gFlow.WrapContents = $true
$ui.GoogleInfo.Margin = New-Object System.Windows.Forms.Padding(3, 8, 12, 0)
$ui.GoogleCheck.Margin = $tm; $ui.GoogleConnect.Margin = $tm
$gFlow.Controls.AddRange(@($ui.GoogleInfo, $ui.GoogleCheck, $ui.GoogleConnect))
Add-LayoutControl $asg $gFlow 0 $r 2
$ui.Google.Add_TextChanged({ if (-not $script:DistLoading -and -not $script:DistRestoring) { Enter-DistDirty } })
$ui.GoogleCheck.Add_Click({ Invoke-GoogleCheck })
$ui.GoogleConnect.Add_Click({ Show-GoogleConnectDialog })
$r = Add-LayoutRow $asg
$ui.CollapseDay.Margin = New-Object System.Windows.Forms.Padding(3, 6, 3, 3)
Register-WrapControl $asg $ui.CollapseDay
Add-LayoutControl $asg $ui.CollapseDay 0 $r 2

# «Что выгружать»: «Всё» над списком моделей, «Выборочно» над списком спецификаций — на одной линии
$r = Add-LayoutRow $asg
$lblScope.Margin = New-Object System.Windows.Forms.Padding(3, 14, 3, 3)
Add-LayoutControl $asg $lblScope 0 $r 2
$r = Add-LayoutRow $asg
$ui.AsgAll.AutoSize = $true
$ui.AsgPart.AutoSize = $true
$ui.AsgAll.Anchor = "Left"
$ui.AsgPart.Anchor = "Left"
$ui.AsgAll.Margin = $tm
$ui.AsgPart.Margin = $tm
Add-LayoutControl $asg $ui.AsgAll 0 $r
Add-LayoutControl $asg $ui.AsgPart 1 $r

# заголовок списка: подпись слева, «все»/«ничего» прижаты к правому краю своего списка
$r = Add-LayoutRow $asg
$colIndex = 0
foreach ($hd in @(
        @{ Label = $lblAsgModels; All = $ui.AsgModelsAll; None = $ui.AsgModelsNone; Caption = $ui.AsgTotal },
        @{ Label = $lblAsgSpecs; All = $ui.AsgSpecsAll; None = $ui.AsgSpecsNone; Caption = $ui.AsgSpecCount })) {
    $head = New-Object System.Windows.Forms.TableLayoutPanel
    $head.ColumnCount = 3
    $head.RowCount = 1
    $head.AutoSize = $true
    $head.Dock = "Fill"
    $head.Margin = $tm
    [void]$head.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, [single]100)))
    [void]$head.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, [single]76)))
    [void]$head.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, [single]82)))
    [void]$head.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    $hd.Label.Anchor = "Left"
    $hd.All.Size = New-Object System.Drawing.Size(70, 26)
    $hd.None.Size = New-Object System.Drawing.Size(76, 26)
    $hd.All.Anchor = "Left, Right"
    $hd.None.Anchor = "Left, Right"
    $head.Controls.Add($hd.Label, 0, 0)
    $head.Controls.Add($hd.All, 1, 0)
    $head.Controls.Add($hd.None, 2, 0)
    if ($hd.Caption) {
        # подпись над правым списком: «Выбрано N из M (для K моделей)» — второй строкой шапки
        $head.RowCount = 2
        [void]$head.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
        $hd.Caption.Anchor = "Left"
        $hd.Caption.Margin = New-Object System.Windows.Forms.Padding(3, 4, 3, 0)
        $head.Controls.Add($hd.Caption, 0, 1)
        $head.SetColumnSpan($hd.Caption, 3)
    }
    Add-LayoutControl $asg $head $colIndex $r
    $colIndex++
}

# списки одной высоты, делят ширину пополам и занимают всё свободное место по высоте
$r = Add-LayoutRow $asg $true
Add-LayoutControl $asg $ui.AsgTree 0 $r
Add-LayoutControl $asg $ui.AsgSpecs 1 $r

$r = Add-LayoutRow $asg
$ui.AsgHint.Margin = New-Object System.Windows.Forms.Padding(3, 6, 3, 3)
Register-WrapControl $asg $ui.AsgHint
Add-LayoutControl $asg $ui.AsgHint 0 $r 2

# «Собрать задание»: сводная и настройка проекта берутся с вкладки «Спецификации», повторно не спрашиваются
$r = Add-LayoutRow $asg
$ui.Run = New-Object System.Windows.Forms.Button
$ui.Run.Text = "Собрать задание"
$ui.Run.Size = New-Object System.Drawing.Size(220, 40)
$ui.Run.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$ui.Run.Anchor = "Left"
$ui.Run.Margin = New-Object System.Windows.Forms.Padding(3, 8, 3, 3)
$ui.Run.Add_Click({ Start-Batch -Mode "assign" })
Set-UiTip $ui.Run "Выгрузка из сводной модели (путь — на вкладке «Спецификации») через Revit, затем xlsx в папке отчетов."
Add-LayoutControl $asg $ui.Run 0 $r 2

# при маленьком окне страница прокручивается, а не сжимает списки до нуля
$ui.PageAssign.AutoScrollMinSize = New-Object System.Drawing.Size(760, 680)
[void]$ui.PageAssign.Controls.Add($asg)

# две вкладки: «Спецификации» (пути, распределение, кнопки шагов) и «Задание»; внизу общая строка статуса и «Открыть отчёт»
$ui.PageDist = New-DistPage
$ui.Tabs = New-Object System.Windows.Forms.TabControl
$ui.Tabs.Dock = "Fill"
$ui.Tabs.Padding = New-Object System.Drawing.Point(18, 6)
$tabSpecs = New-Object System.Windows.Forms.TabPage
$tabSpecs.Text = "Спецификации"
$tabAsg = New-Object System.Windows.Forms.TabPage
$tabAsg.Text = "Задание"
foreach ($tp in @($tabSpecs, $tabAsg)) {
    $tp.UseVisualStyleBackColor = $false
    $tp.BackColor = [System.Drawing.SystemColors]::Window
    $tp.Padding = New-Object System.Windows.Forms.Padding(8)
}
$tabSpecs.Controls.Add($ui.PageDist)
$tabAsg.Controls.Add($ui.PageAssign)
$ui.Tabs.Controls.Add($tabSpecs)
$ui.Tabs.Controls.Add($tabAsg)
$ui.Tabs.Add_SelectedIndexChanged({ if ($ui.Tabs.SelectedIndex -eq 1) { Refresh-PartialLists } })

$bottom = New-Object System.Windows.Forms.Panel
$bottom.Height = 44
$bottom.Dock = "Bottom"
$bottom.BackColor = [System.Drawing.SystemColors]::Window
$ui.Status = New-Object System.Windows.Forms.Label
$ui.Status.Text = ""      # статус — только для реальных состояний и ошибок, без постоянной подсказки
$ui.Status.Dock = "Fill"
$ui.Status.Padding = New-Object System.Windows.Forms.Padding(12, 0, 8, 0)
$ui.Status.ForeColor = [System.Drawing.SystemColors]::GrayText
$ui.Status.TextAlign = "MiddleLeft"
$openHost = New-Object System.Windows.Forms.Panel
$openHost.Width = 150
$openHost.Dock = "Right"
$ui.OpenFolder = New-Object System.Windows.Forms.Button
$ui.OpenFolder.Text = "Открыть отчёт"
$ui.OpenFolder.Size = New-Object System.Drawing.Size(134, 30)
$ui.OpenFolder.Location = New-Object System.Drawing.Point(8, 7)
$ui.OpenFolder.Add_Click({ Open-ReportFolder })
$openHost.Controls.Add($ui.OpenFolder)
$csvHost = New-Object System.Windows.Forms.Panel
$csvHost.Width = 110
$csvHost.Dock = "Right"
$ui.OpenCsv = New-Object System.Windows.Forms.LinkLabel
$ui.OpenCsv.Text = "Открыть CSV"
$ui.OpenCsv.AutoSize = $true
$ui.OpenCsv.Location = New-Object System.Drawing.Point(8, 13)
$ui.OpenCsv.Visible = $false
$ui.OpenCsv.Add_LinkClicked({ Open-RunCsv })
Set-UiTip $ui.OpenCsv "Отчёт прогона в CSV (xlsx не собран). Открывается в проводнике, Excel не запускается."
$csvHost.Controls.Add($ui.OpenCsv)
$logsHost = New-Object System.Windows.Forms.Panel
$logsHost.Width = 60
$logsHost.Dock = "Right"
$ui.OpenLogs = New-Object System.Windows.Forms.LinkLabel
$ui.OpenLogs.Text = "Логи"
$ui.OpenLogs.AutoSize = $true
$ui.OpenLogs.Location = New-Object System.Drawing.Point(8, 13)
$ui.OpenLogs.Add_LinkClicked({ Open-BvsLogFolder })
Set-UiTip $ui.OpenLogs "Папка журналов этого скрипта (отчеты\логи): логи Revit Batch Processor этого запуска, отправки в Google, «Выборочно»."
$logsHost.Controls.Add($ui.OpenLogs)
$bottom.Controls.Add($ui.Status)
$bottom.Controls.Add($logsHost)
$bottom.Controls.Add($csvHost)
$bottom.Controls.Add($openHost)

# вверху окна: список проектов (файлы отчеты\проект\распределение_*.txt) и «Новый проект»; смена проекта подставляет всё сразу
$top = New-Object System.Windows.Forms.Panel
$top.Height = 44
$top.Dock = "Top"
$top.BackColor = [System.Drawing.SystemColors]::Window
$lblProject = New-Object System.Windows.Forms.Label
$lblProject.Text = "Проект"
$lblProject.AutoSize = $true
$lblProject.Location = New-Object System.Drawing.Point(12, 14)
$ui.ProjectBox = New-Object System.Windows.Forms.ComboBox
$ui.ProjectBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
$ui.ProjectBox.Location = New-Object System.Drawing.Point(72, 10)
$ui.ProjectBox.Size = New-Object System.Drawing.Size(280, 26)
$ui.ProjectNew = New-DistButton "Новый проект"
$ui.ProjectNew.Location = New-Object System.Drawing.Point(364, 8)
Set-UiTip $ui.ProjectBox "Проект = файл отчеты\проект\распределение_<проект>.txt. У проекта свои пути, выбор «В проекте», Google, журнал и история."
Set-UiTip $ui.ProjectNew "Создать проект (имя — обычно префикс моделей, например PROJ)."
$ui.ProjectBox.Add_SelectedIndexChanged({ if (-not $script:ProjectListBusy -and $ui.ProjectBox.SelectedItem) { Switch-Project ([string]$ui.ProjectBox.SelectedItem) } })
$ui.ProjectNew.Add_Click({ Invoke-NewProject })
$top.Controls.AddRange(@($lblProject, $ui.ProjectBox, $ui.ProjectNew))
$form.Controls.Add($ui.Tabs)
$form.Controls.Add($top)
$form.Controls.Add($bottom)

$ui.Timer = New-Object System.Windows.Forms.Timer
$ui.Timer.Interval = 1000
$ui.Timer.Add_Tick({
    $proc = $script:Run.Process
    if (-not $proc) { return }
    $live = Read-LiveStatus
    if (-not $proc.HasExited) {
        if ($live -and $script:Run.Mode -eq "assign") {
            $ui.Status.Text = ("{0}. Выгрузка {1} из {2} · {3}. Окно не закрывай." -f $script:Run.TargetText, $live.index, $live.total, $live.name)
        }
        elseif ($live -and $script:Run.Mode -eq "classifier") {
            $res = ""
            if ($live.ContainsKey("result")) { $res = [string]$live.result }
            $ui.Status.Text = ("{0}. Классификатор {1} из {2} · {3} · {4}" -f $script:Run.TargetText, $live.index, $live.total, $live.name, $res)
        }
        elseif ($live) { $ui.Status.Text = ($script:Run.TargetText + ". " + (Format-LiveStatus $live)) }
        elseif ($script:Run.WaitText) { $ui.Status.Text = $script:Run.WaitText }
        else { $ui.Status.Text = "Идёт обработка. Revit откроется поверх. Окно не закрывай." }
        return
    }
    $code = $proc.ExitCode
    Set-ActionsEnabled $true
    $script:Run.Process = $null
    if ($script:Run.Mode -eq "assign") {
        $summaryPath = $ui.Summary.Text.Trim()
        $xlsxPath = Get-AssignmentPath $summaryPath
        $historyPath = Get-HistoryPath $summaryPath
        $period = Get-DynamicsPeriodFromUi
        if (-not (Test-BvsExportDone $script:Run.WorkDir)) {
            $ui.Timer.Stop()
            $ui.Status.Text = ("Выгрузка оборвалась (код {0}). Задание не собрано." -f $code)
            Show-Finished $ui.Status.Text
            return
        }
        $ui.Timer.Stop()
        Set-ActionsEnabled $false
        $ui.Status.Text = "Собираю задание (Open XML, без Excel). Подождите..."
        [System.Windows.Forms.Application]::DoEvents()
        $buildOut = ""
        $buildCode = 1
        try {
            $buildOut = (Invoke-BuildAssignment -ExportDir $script:Run.WorkDir -OutXlsx $xlsxPath -HistoryPath $historyPath -Period $period -CollapseDay:$ui.CollapseDay.Checked -Partial:$script:Run.Partial -AuxDir $script:Run.WorkDir | Out-String)
            if ($buildOut -match 'BUILDER=OPENXML') { $buildCode = 0 }
        }
        catch {
            $buildOut = ($_.Exception.Message + "`r`n" + $_.ScriptStackTrace)
        }
        $script:Run.Mode = $null
        Finish-AssignBuild -RevitCode $code -WorkDir $script:Run.WorkDir -XlsxPath $xlsxPath `
            -BuildLogPath "" -BuildErrPath "" -BuildExitCode $buildCode -BuildStdout $buildOut
        return
    }
    if ($script:Run.Mode -eq "classifier") {
        $ui.Timer.Stop()
        Finish-ClassifierRun -ExitCode $code
        $script:Run.Mode = $null
        Update-ActionState
        return
    }
    $ui.Timer.Stop()
    $xlsxPath = Get-NextReportPath (Get-ProjectReportDir (Get-ProjectCode $script:Run.Summary))
    $script:ReportXlsxError = ""
    $script:ExcelPidUnknown = $false
    $xlsx = Export-ReportXlsx -Summary $script:Run.Summary -Params $script:Run.Params -Skips $script:Run.Skips -Inserted $script:Run.Inserted -XlsxPath $xlsxPath
    $runStats = Get-RunSummaryStats $script:Run.Summary
    $failText = Get-RunFailureText $code $runStats
    if ($xlsx) { Set-OpenFolderTarget $xlsxPath; Set-CsvLink "" } else { Set-CsvLink $script:Run.Summary }
    if ($failText) {
        # «готова» — только при коде выхода RBP 0 и без ошибок в колонке «Ошибка»
        $where = if ($xlsx) { " Отчёт: " + $xlsxPath } else { " xlsx не собран." + $(if ($script:ExcelPidUnknown) { " Excel не закрыт, PID неизвестен." } else { "" }) }
        $ui.Status.Text = $failText + $where
        Show-Finished $failText
    }
    elseif ($xlsx) {
        if ($script:Run.Mode -eq "summary") {
            $ui.Status.Text = "Сводная готова."
            Publish-SummaryDone $true
            Save-Settings
        }
        else {
            $ui.Status.Text = "Модели разделов готовы."
        }
        Show-Finished $ui.Status.Text
    }
    else {
        $why = ""
        if ($script:ReportXlsxError) { $why = " Причина: " + $script:ReportXlsxError }
        $pidNote = if ($script:ExcelPidUnknown -and $why -notmatch 'PID неизвестен') { " Excel не закрыт, PID неизвестен." } else { "" }
        $ui.Status.Text = ("Готово (код {0}). xlsx не собран, отчёт — в CSV.{1}{2}" -f $code, $why, $pidNote)
        Show-Finished $ui.Status.Text
    }
    # «В сводную» прошла без ошибок (даже если xlsx не собран) — запомнить, что сводная выполнена. Кнопку «В разделы» это не включает и не выключает.
    if ($script:Run.Mode -eq "summary" -and -not $failText -and -not $script:SummaryDone) { Publish-SummaryDone $true; Save-Settings }
    Update-ActionState
})
$form.Controls.Add($ui.Timer)

$form.Add_FormClosing({
    if (-not (Confirm-DistDirtyLeave)) { $_.Cancel = $true; return }
    if ($script:Run.Process -and -not $script:Run.Process.HasExited) {
        $ask = [System.Windows.Forms.MessageBox]::Show(
            "Обработка ещё идёт. Закрыть окно? Revit Batch Processor продолжит работу.",
            "Проверочные спецификации",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )
        if ($ask -ne [System.Windows.Forms.DialogResult]::Yes) { $_.Cancel = $true }
    }
    else {
        Flush-DistSave          # отложенное автосохранение «Распределения» не должно потеряться при закрытии
        Save-Settings
    }
})

$saved = Read-Settings
# раскладка отчетов: перенос старых файлов (ничего не удаляется) и автоочистка служебных папок старше срока хранения
$layoutInit = Initialize-ReportLayout $ReportDir $script:AuxKeepDays -AppDir $AppDir
$script:Layout = $layoutInit.Layout
if (@($layoutInit.Migration.Moves).Count -gt 0) {
    $ui.Status.Text = ("Файлы отчётов разложены по папкам: перенесено {0} (история — в отчеты\история, служебные — в отчеты\служебные, отчёты .bat — в отчеты\операции). Ничего не удалено." -f @($layoutInit.Migration.Moves).Count)
}
# старые значения из ui_settings.txt (пути, Google, режим вставки, «Выборочно») — в проект, к которому они относятся; один раз, ничего не теряя
$migratedProject = Invoke-StartupSettingsMigration $saved
if ($migratedProject -and -not $saved.LastProject) { $saved.LastProject = $migratedProject }
$ui.CollapseDay.Checked = ($saved.CollapseDay -eq "1")
Set-AsgScopeEnabled
if ($saved.DynDays -match '^(\d+)') { $ui.DynDays.SelectedIndex = 0 }
if ($saved.DynDays -eq "14") { $ui.DynDays.SelectedIndex = 1 }
if ($saved.DynDays -eq "30") { $ui.DynDays.SelectedIndex = 2 }
if ($ui.DynDays.SelectedIndex -lt 0) { $ui.DynDays.SelectedIndex = 0 }
if ($saved.DynCustom -eq "1") { $ui.DynCustom.Checked = $true }
if (-not [string]::IsNullOrWhiteSpace($saved.DynFrom)) {
    try { $ui.DynFrom.Value = [datetime]::Parse($saved.DynFrom).Date } catch { }
}
if (-not [string]::IsNullOrWhiteSpace($saved.DynTo)) {
    try { $ui.DynTo.Value = [datetime]::Parse($saved.DynTo).Date } catch { }
}
$ui.DynTo.Value = (Get-Date).Date
$ui.DynFrom.Value = (Get-Date).Date.AddDays(-6)
[void](Update-SectionCatalog (Get-SectionCatalogPath $ReportDir))      # справочник разделов: создаётся при первом запуске, подсветка в таблице
# проект: последний открытый (lastproject=), иначе первый в списке; проектов нет — пустое окно и «Новый проект»
$projectNames = @(Get-ProjectNames $ReportDir)
$open = ""
if ($saved.LastProject -and $projectNames -contains $saved.LastProject) { $open = $saved.LastProject }
elseif ($projectNames.Count -gt 0) { $open = $projectNames[0] }
# при открытии проекта прежний выбор_<проект>.txt переносится в [inproject] (один раз, с резервной копией отчеты\проект\_backup_<время>)
Load-DistProject $(if ($open) { Get-ProjectConfigPath $ReportDir $open } else { "" })
Update-ProjectList
if (-not $open) { $ui.Status.Text = "Нет проекта: нажмите «Новый проект»." }
Add-DistStateMigrationNote
Update-ActionState

[void]$form.ShowDialog()
