# UTF-8 WITH BOM required (PowerShell 5.1).
# Настройка проекта — ЕДИНСТВЕННЫЙ источник «какие спецификации в какие модели/разделы» (вкладка «Распределение»).
# Файл: отчеты\проект\распределение_<префикс моделей>.txt (префикс — часть имени моделей до первого «_», например PROJ), рядом с отчётами проекта,
# не в папке с кодом. Читают: ui_schedules.ps1, assignment_partial_ui.ps1, insert_schedules.py.
#
# Формат (UTF-8 с BOM, строки «#» — комментарии; данные — через TAB):
#   [project]   stage=<П|РД|…>, template=<общий файл спецификаций>, models_dir=<папка с моделями>, summary_model=<сводная модель>, google_link=<ссылка на таблицу для чата>,
#               summary_done=<0|1> («В сводную» выполнена), insert_mode=<add_diff|replace>, asg_scope=<all|partial>  (всё, что относится к проекту, хранится здесь, а не в ui_settings.txt)
#               classifier=<UNC-путь файла классификатора для кнопки «Классификатор»> (пусто — для PROJ официальный путь, см. Get-ProjectClassifierPath)
#   [asg_models] имена моделей, отмеченных в «Выборочно»;  [asg_pairs] пары «модель|спека» (раздел есть — выбор делали)
#   [sections]  <код раздела>TAB<название>            (справочно; раздел модели всегда определяется по имени файла, см. Get-ModelSectionFromName)
#   [models]    <имя модели без .rvt>TAB<код раздела>  (только модели целевых разделов из $script:PcSections; после «Взять выбор из сводной» — только такие связи сводной.
#               Строка с чужим кодом, в том числе старая UNKNOWN, при чтении не показывается и не роняет разбор; в файле остаётся до явного «Сохранить»)
#   [specs]     <имя спецификации в шаблоне> (порядок = порядок в шаблоне)
#   [groups]    <имя спецификации>TAB<группа>         (только ручные переназначения группы)
#   [inproject] <имя спецификации>                     (спецификации, отмеченные «В проекте»; нет раздела — в проекте все). ЕДИНСТВЕННОЕ место галочек «В проекте»:
#               прежний отдельный выбор_<проект>.txt переносится сюда один раз (Invoke-ProjectStateMigration) и переименовывается в выбор_<проект>.migrated_<дата>.txt.
#               Имена, которых нет в [specs] (нет в общем файле), тоже хранятся здесь и не теряются при сохранении.
#   [summary]   <имя спецификации>                     (какие из них идут в сводную; нет раздела — все спецификации проекта)
#   [categories] <имя спецификации>TAB<категория Revit>  (кэш: читается из общего файла при «Загрузить спецификации и модели»)
#   [cells]     <модель>TAB<спека>TAB<1|0>TAB<a|m|x|c>TAB<число элементов или пусто>
#               a — авто (по справочнику «раздел → категории»), m — ручная правка, c — только число элементов («Посчитать элементы»), x — устарело (не удалось определить)
# Справочник «раздел → номера/категории» — отдельный файл отчеты\проект\справочник_разделов.txt, один на все проекты (см. Read-SectionCatalog).
#
# Профиль при вставке: в значениях фильтров шаблона {СТАДИЯ}/{STAGE} заменяются стадией проекта; имя вида в Revit = стадия + 8 пробелов + название.

$script:PcStagePad = 8      # имя вида в Revit: <стадия> + 8 пробелов + название (папка «П» в диспетчере проекта задаётся пробелами)

# ЕДИНСТВЕННОЕ правило «раздел по имени файла»: код ищется только в виде «_КОД_» (регистр не важен), кириллица и латиница.
# Список — разделы.txt (bvs_sections.ps1). Порядок $script:PcSections = порядок столбцов окна.
# Нет файла — тот же встроенный список и предупреждение, запуск не падает.
try {
    if (-not (Get-Command Get-BvsPcSections -ErrorAction SilentlyContinue)) {
        $__bvsSections = Join-Path $PSScriptRoot "bvs_sections.ps1"
        if (-not $PSScriptRoot) { $__bvsSections = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "bvs_sections.ps1" }
        if (Test-Path -LiteralPath $__bvsSections) { . $__bvsSections } else { Write-Warning "нет bvs_sections.ps1, встроенный список разделов" }
    }
}
catch { Write-Warning ("разделы: " + $_.Exception.Message) }
if (Get-Command Get-BvsPcSections -ErrorAction SilentlyContinue) {
    $script:PcSections = @(Get-BvsPcSections)
}
else {
    Write-Warning "встроенный список разделов (bvs_sections.ps1 не подключён)"
    $script:PcSections = @(
        @{ Code = "OV"; Title = "ОВ"; Names = @("ОВ", "OV") }, @{ Code = "PD"; Title = "ПД"; Names = @("ПД", "PD") },
        @{ Code = "VK"; Title = "ВК"; Names = @("ВК", "VK") }, @{ Code = "PT"; Title = "ПТ"; Names = @("ПТ", "PT") },
        @{ Code = "OT"; Title = "ОТ"; Names = @("ОТ", "OT") }, @{ Code = "HS"; Title = "ХС"; Names = @("ХС", "HS") },
        @{ Code = "ITP"; Title = "ИТП"; Names = @("ИТП", "ITP") }, @{ Code = "EOM"; Title = "ЭОМ"; Names = @("ЭОМ", "EOM") },
        @{ Code = "SS"; Title = "СС"; Names = @("СС", "SS") },
        @{ Code = "AVT"; Title = "АВТ"; Names = @("АВТ", "AVT") }, @{ Code = "HC"; Title = "ХЦ"; Names = @("ХЦ", "HC") }
    )
}
# В таблицу попадают только коды $script:PcSections. Похожий, но чужой код (…_ORG_BF_…, …_IOS_GU_…) — не колонка и не модель.
$script:PcNonSectionTokens = @("TASK", "AR", "KR", "TH", "BASE", "SVOD", "IOS", "ORG", "СВОД", "BC")
$script:PcDefaultTitles = @{}
foreach ($__s in $script:PcSections) { $script:PcDefaultTitles[$__s.Code] = $__s.Title }

function New-ProjectConfig([string]$Stage = "П") {
    return [pscustomobject]@{
        Stage = $Stage; Template = ""; ModelsDir = ""
        ModelsRsn = New-Object System.Collections.Generic.List[string]     # модели, выбранные с Revit Server (RSN://); при ModelsDir = RSN://…
        GroupsCache = $null                                                 # результат Get-ProjectGroups; сбрасывается при смене спецификаций и групп
        InProject = @{}                                                     # ключ спеки -> $true; $null (файл без [inproject]) — в проекте все
        Summary = @{}                                                       # ключ спеки -> $true: идёт в сводную (подмножество «В проекте»); $null (файл без [summary]) — все спецификации проекта
        Sections = New-Object System.Collections.Generic.List[object]     # @{ Code; Title }
        Models = New-Object System.Collections.Generic.List[object]       # @{ Name; Section }
        Specs = New-Object System.Collections.Generic.List[string]
        GroupOverride = @{}                                                 # спека -> группа
        Categories = @{}                                                    # ключ спеки -> категория Revit (читается из общего файла при «Загрузить…», кэш)
        SummaryModel = ""; GoogleLink = ""; SummaryDone = $false; InsertMode = ""     # путь сводной модели, ссылка на Google-таблицу для чата, «В сводную» выполнена, режим вставки (add_diff / replace)
        Classifier = ""                                                     # путь классификатора проекта (classifier=); пусто — значение по умолчанию (Get-ProjectClassifierPath)
        AsgScope = ""; AsgModels = $null; AsgPairs = $null                          # «Выборочно»: all|partial, отмеченные модели (имя -> $true), пары «модель|спека» ($null — выбор не делали)
        SelectionUnknown = New-Object System.Collections.Generic.List[string]   # имена «В проекте», которых нет в общем файле: пишутся в [inproject] как есть
        Cells = @{}                                                         # "модель`tспека" -> @{ State; Origin; Count }
        HeldModels = (New-Object System.Collections.Generic.List[object])   # строки [models] без целевого раздела: не в таблице, пишутся обратно до явного «Сохранить»
        HeldCells = (New-Object System.Collections.Generic.List[string])    # их строки [cells], как в файле
    }
}

# Копия карты ($null остаётся $null — это другое состояние, чем пустая карта)
function Copy-ProjectMap($Map) {
    if ($null -eq $Map) { return $null }
    $c = @{}
    foreach ($k in @($Map.Keys)) { $c[$k] = $Map[$k] }
    return $c
}

# Независимая копия настройки (для отмены массового действия). Кэш групп не копируется.
function Copy-ProjectConfig($Cfg) {
    if ($null -eq $Cfg) { return $null }
    $n = New-ProjectConfig $(if ($Cfg.Stage) { [string]$Cfg.Stage } else { "П" })
    $n.Template = [string]$Cfg.Template
    $n.ModelsDir = [string]$Cfg.ModelsDir
    $n.SummaryModel = [string]$Cfg.SummaryModel
    $n.GoogleLink = [string]$Cfg.GoogleLink
    $n.SummaryDone = [bool]$Cfg.SummaryDone
    $n.InsertMode = [string]$Cfg.InsertMode
    $n.Classifier = [string]$Cfg.Classifier
    $n.AsgScope = [string]$Cfg.AsgScope
    $n.InProject = Copy-ProjectMap $Cfg.InProject
    $n.Summary = Copy-ProjectMap $Cfg.Summary
    $n.GroupOverride = Copy-ProjectMap $Cfg.GroupOverride
    $n.Categories = Copy-ProjectMap $Cfg.Categories
    $n.AsgModels = Copy-ProjectMap $Cfg.AsgModels
    $n.AsgPairs = Copy-ProjectMap $Cfg.AsgPairs
    $n.ModelsRsn = New-Object System.Collections.Generic.List[string]
    if ($Cfg.ModelsRsn) { foreach ($m in $Cfg.ModelsRsn) { if ($m) { [void]$n.ModelsRsn.Add([string]$m) } } }
    $n.Sections = New-Object System.Collections.Generic.List[object]
    if ($Cfg.Sections) { foreach ($s in $Cfg.Sections) { [void]$n.Sections.Add([pscustomobject]@{ Code = [string]$s.Code; Title = [string]$s.Title }) } }
    $n.Models = New-Object System.Collections.Generic.List[object]
    if ($Cfg.Models) { foreach ($m in $Cfg.Models) { [void]$n.Models.Add([pscustomobject]@{ Name = [string]$m.Name; Section = [string]$m.Section }) } }
    $n.Specs = New-Object System.Collections.Generic.List[string]
    if ($Cfg.Specs) { foreach ($s in $Cfg.Specs) { [void]$n.Specs.Add([string]$s) } }
    $n.SelectionUnknown = New-Object System.Collections.Generic.List[string]
    if ($Cfg.SelectionUnknown) { foreach ($u in $Cfg.SelectionUnknown) { if ($u) { [void]$n.SelectionUnknown.Add([string]$u) } } }
    $n.Cells = @{}
    if ($Cfg.Cells) {
        foreach ($k in @($Cfg.Cells.Keys)) {
            $c = $Cfg.Cells[$k]
            if ($null -eq $c) { continue }
            $n.Cells[[string]$k] = [pscustomobject]@{ State = [int]$c.State; Origin = [string]$c.Origin; Count = [int]$c.Count }
        }
    }
    $n.HeldModels = New-Object System.Collections.Generic.List[object]
    if ($Cfg.HeldModels) { foreach ($m in $Cfg.HeldModels) { [void]$n.HeldModels.Add([pscustomobject]@{ Name = [string]$m.Name; Section = [string]$m.Section }) } }
    $n.HeldCells = New-Object System.Collections.Generic.List[string]
    if ($Cfg.HeldCells) { foreach ($hc in $Cfg.HeldCells) { if ($null -ne $hc) { [void]$n.HeldCells.Add([string]$hc) } } }
    $n.GroupsCache = $null
    return $n
}

# ---------- пути ----------
function Get-ProjectConfigDir([string]$ReportDir) { return (Join-Path $ReportDir "проект") }
function Get-ProjectConfigPath([string]$ReportDir, [string]$Prefix) {
    $c = ([string]$Prefix).Trim()
    if (-not $c) { $c = "PROJ" }
    $c = ($c -replace '[\\/:*?"<>|]', '_')
    return (Join-Path (Get-ProjectConfigDir $ReportDir) ("распределение_" + $c + ".txt"))
}

# Префикс проекта = часть имени моделей до первого «_» (самый частый среди имён); нет моделей — пусто
# Кэш спецификаций проекта для показа таблицы без Revit при открытии окна: сохранённый путь настройки, если в нём есть спецификации; иначе
# самый свежий файл распределение_*.txt, у которого общий файл (template=) совпадает с указанным в настройках окна. Нет кэша — пустая строка.
function Find-ProjectCachePath([string]$ReportDir, [string]$SavedPath, [string]$Template) {
    if ($SavedPath -and (Test-Path -LiteralPath $SavedPath)) {
        $c = Read-ProjectConfig $SavedPath
        if ($null -ne $c -and $c.Specs.Count -gt 0) { return $SavedPath }
    }
    $want = ([string]$Template).Trim().ToLowerInvariant()
    if (-not $want) { return "" }
    $dir = Get-ProjectConfigDir $ReportDir
    if (-not (Test-Path -LiteralPath $dir)) { return "" }
    foreach ($f in @(Get-ChildItem -LiteralPath $dir -Filter "распределение_*.txt" -File -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending)) {
        $c = Read-ProjectConfig $f.FullName
        if ($null -ne $c -and $c.Specs.Count -gt 0 -and ([string]$c.Template).Trim().ToLowerInvariant() -eq $want) { return $f.FullName }
    }
    return ""
}

# ---------- список проектов и перенос старых настроек окна ----------
# Проект = файл отчеты\проект\распределение_<имя>.txt. Возвращает имена (без «распределение_» и «.txt») по алфавиту
function Get-ProjectNames([string]$ReportDir) {
    $dir = Get-ProjectConfigDir $ReportDir
    if (-not (Test-Path -LiteralPath $dir)) { return @() }
    $names = New-Object System.Collections.Generic.List[string]
    foreach ($f in @(Get-ChildItem -LiteralPath $dir -Filter "распределение_*.txt" -File -ErrorAction SilentlyContinue)) {
        if ($f.Name -match '^распределение_(.+)\.txt$') { $names.Add($Matches[1]) }
    }
    return @($names | Sort-Object)
}

# Имя проекта для пути «распределение_<имя>.txt»: без символов, недопустимых в имени файла; пустое / занятое — $null
function Test-NewProjectName([string]$ReportDir, [string]$Name) {
    $n = ([string]$Name).Trim()
    if (-not $n) { return "Укажите имя проекта." }
    if ($n -match '[\\/:*?"<>|]') { return "В имени проекта нельзя использовать символы \ / : * ? `" < > |" }
    if (Test-Path -LiteralPath (Get-ProjectConfigPath $ReportDir $n)) { return ("Проект «{0}» уже есть." -f $n) }
    return $null
}

# Старые настройки окна (ui_settings.txt: пути, Google, режим вставки, «Выборочно» …) -> в тот проект, к которому они относятся.
# Legacy — хэш Read-Settings (Source, ProjStage, ProjFile, ProjModelsDir, ProjRsn, Summary, Google, SummaryDone, InsertMode, AsgScope, AsgModels, AsgPairs).
# Проект: имя из прежнего projfile=, иначе префикс моделей по сводной (PROJ_IOS_SVOD_… -> PROJ), иначе по папке моделей, иначе PROJ.
# В проекте заполняются только пустые поля (ничего не затирается); повторный вызов ничего не меняет. Возвращает @{ Name; Path; Changed }
function Get-LegacyProjectName($Legacy, [string]$ReportDir) {
    $leaf = ""
    if ($Legacy.ProjFile) { $leaf = Split-Path -Leaf ([string]$Legacy.ProjFile).Replace("\", "/") }
    if ($leaf -match '^распределение_(.+)\.txt$') { return $Matches[1] }
    foreach ($cand in @($Legacy.Summary, $Legacy.ProjModelsDir, $Legacy.Source)) {
        if ($cand) {
            $name = [System.IO.Path]::GetFileNameWithoutExtension((([string]$cand).TrimEnd("/", "\") -replace '^.*[\\/]', ''))
            if ($name -match '^([A-Za-z0-9]+)_IOS_' -or $name -match '^([A-Za-z0-9]+)_') { return $Matches[1] }
        }
    }
    return "PROJ"
}

function Invoke-LegacySettingsMigration($Legacy, [string]$ReportDir) {
    $any = $false
    foreach ($k in @("Source", "Summary", "ProjModelsDir", "Google", "ProjFile")) { if ($Legacy[$k]) { $any = $true } }
    if (@($Legacy.ProjRsn).Count -gt 0 -or $null -ne $Legacy.AsgModels -or $null -ne $Legacy.AsgPairs) { $any = $true }
    if (-not $any) { return @{ Name = ""; Path = ""; Changed = $false } }
    $name = Get-LegacyProjectName $Legacy $ReportDir
    $path = Get-ProjectConfigPath $ReportDir $name
    $cfg = Read-ProjectConfig $path
    $existed = ($null -ne $cfg)
    if (-not $existed) { $cfg = New-ProjectConfig $(if ($Legacy.ProjStage) { [string]$Legacy.ProjStage } else { "П" }); $cfg.InProject = $null; $cfg.Summary = $null }
    $changed = -not $existed
    $fill = {
        param([string]$Prop, $Value)
        if ($Value -and -not [string]$cfg.$Prop) { $cfg.$Prop = [string]$Value; $script:__mig = $true }
    }
    $script:__mig = $false
    & $fill "Template" $Legacy.Source
    & $fill "ModelsDir" $Legacy.ProjModelsDir
    & $fill "SummaryModel" $Legacy.Summary
    & $fill "GoogleLink" $Legacy.Google
    & $fill "InsertMode" $Legacy.InsertMode
    & $fill "AsgScope" $Legacy.AsgScope
    if ($Legacy.ProjStage -and -not $cfg.Stage) { $cfg.Stage = [string]$Legacy.ProjStage; $script:__mig = $true }
    if ($Legacy.SummaryDone -and -not $cfg.SummaryDone) { $cfg.SummaryDone = $true; $script:__mig = $true }
    if (@($Legacy.ProjRsn).Count -gt 0 -and $cfg.ModelsRsn.Count -eq 0) { foreach ($m in @($Legacy.ProjRsn)) { $cfg.ModelsRsn.Add([string]$m) }; $script:__mig = $true }
    if ($null -ne $Legacy.AsgModels -and $null -eq $cfg.AsgModels) { $cfg.AsgModels = $Legacy.AsgModels; $script:__mig = $true }
    if ($null -ne $Legacy.AsgPairs -and $null -eq $cfg.AsgPairs) { $cfg.AsgPairs = $Legacy.AsgPairs; $script:__mig = $true }
    if ($script:__mig) { $changed = $true }
    if ($changed) { Write-ProjectConfig $cfg $path }
    return @{ Name = $name; Path = $path; Changed = $changed }
}

function Get-ProjectPrefix($ModelNames) {
    $count = @{}
    foreach ($m in @($ModelNames)) {
        $n = Get-ProjectModelName ([string]$m)
        if (-not $n) { continue }
        $p = $n.Split("_")[0]
        if ($p) { if ($count.ContainsKey($p)) { $count[$p]++ } else { $count[$p] = 1 } }
    }
    $best = ""; $bn = 0
    foreach ($k in @($count.Keys | Sort-Object)) { if ($count[$k] -gt $bn) { $best = $k; $bn = $count[$k] } }
    return $best
}

# Папка моделей: принимается только папка. Возвращает причину отказа или пустую строку.
function Get-ModelsDirProblem([string]$Path) {
    $t = ([string]$Path).Trim()
    if (-not $t) { return "Укажите папку с моделями." }
    if (Test-RsnModelsSource $t) { return "" }
    if ($t.TrimEnd([char[]]@([char]92, [char]47)).ToLowerInvariant().EndsWith(".rvt")) { return "Нужна папка с моделями, а указан файл модели (.rvt). Укажите папку, в которой лежат модели." }
    if ((Test-Path -LiteralPath $t -PathType Leaf)) { return "Нужна папка с моделями, а указан файл. Укажите папку, в которой лежат модели." }
    return ""
}

# Пустой путь настройки — запускать скрипты нельзя. Возвращает текст причины или пустую строку.
# Источник моделей — Revit Server (поле «Папка с моделями» = RSN://сервер/папка)
function Test-RsnModelsSource([string]$Text) { return (([string]$Text).Trim() -match '^(?i)RSN://') }

# Подпись источника по выбранным моделям: общая папка RSN://сервер/папка
function Get-RsnSourceLabel([string[]]$Paths) {
    $dirs = @($Paths | Where-Object { $_ -match '^(?i)RSN://' } | ForEach-Object { $t = $_.Trim().Replace("\", "/"); $t.Substring(0, $t.LastIndexOf("/")) })
    if ($dirs.Count -eq 0) { return "" }
    $common = $dirs[0]
    foreach ($d in $dirs) {
        while ($common.Length -gt 6 -and -not (($d + "/").ToLowerInvariant().StartsWith(($common + "/").ToLowerInvariant()))) { $common = $common.Substring(0, $common.LastIndexOf("/")) }
    }
    return $common
}

# Отбор инженерных моделей: по коду раздела, без общего файла и сводной
# Источник моделей по прежнему списку шага «Разделы» (rvt_list.txt): RSN:// — общая папка сервера и сами пути; локальные — общая родительская папка
# Возвращает @{ Source = текст поля «Модели»; Rsn = пути RSN:// }
function Get-LegacyModelsSource([string[]]$Paths) {
    $list = @($Paths | Where-Object { $_ -and $_.Trim() } | ForEach-Object { $_.Trim() })
    if ($list.Count -eq 0) { return @{ Source = ""; Rsn = @() } }
    $rsn = @($list | Where-Object { $_ -match '^(?i)RSN://' })
    if ($rsn.Count -gt 0) { return @{ Source = (Get-RsnSourceLabel $rsn); Rsn = $rsn } }
    $common = $null
    foreach ($p in $list) {
        $parts = @($p.Replace("/", "\").Split([char]'\'))
        $dir = @($parts[0..([Math]::Max(0, $parts.Count - 2))])
        if ($parts.Count -le 1) { $dir = @() }
        if ($null -eq $common) { $common = $dir; continue }
        $n = [Math]::Min($common.Count, $dir.Count); $k = 0
        while ($k -lt $n -and [string]::Equals($common[$k], $dir[$k], [System.StringComparison]::OrdinalIgnoreCase)) { $k++ }
        $common = if ($k -gt 0) { @($common[0..($k - 1)]) } else { @() }
    }
    $text = ($common -join "\")
    return @{ Source = $text; Rsn = @() }
}

function Select-DistModelPaths([string[]]$Paths, [string]$Template, [string]$Summary) {
    $tpl = (Get-ProjectModelName $Template).ToLowerInvariant()
    $sum = (Get-ProjectModelName $Summary).ToLowerInvariant()
    return @($Paths | Where-Object {
            $n = (Get-ProjectModelName $_).ToLowerInvariant()
            $n -and ($n -ne $tpl) -and ($n -ne $sum) -and (Test-ProjectEngineeringModel $n)
        })
}

function Get-ProjectRunBlocker([string]$ProjectPath) {
    if (-not ([string]$ProjectPath).Trim()) { return "Не задан путь к файлу настройки проекта. Откройте вкладку «Распределение» и нажмите «Загрузить спецификации и модели»." }
    if (-not (Test-Path -LiteralPath $ProjectPath)) { return ("Файл настройки проекта не найден: {0}. Нажмите «Загрузить спецификации и модели» на вкладке «Распределение»." -f $ProjectPath) }
    return ""
}

# ---------- имена ----------
function Get-ProjectModelName([string]$PathOrName) {
    $t = ([string]$PathOrName).Trim().Replace("\", "/")
    $leaf = $t.Split("/")[-1]
    if ($leaf.ToLowerInvariant().EndsWith(".rvt")) { $leaf = $leaf.Substring(0, $leaf.Length - 4) }
    return $leaf
}

# Ключи нечувствительны к регистру, пробелам, префиксу стадии («П        5.1_…» = «5.1_…») и опечатке «значния»
# Ключи вычисляются на каждую ячейку (модели × спецификации — тысячи обращений на клик), поэтому результат запоминается: ключ зависит только от строки
$script:PcSpecKeyCache = @{}
$script:PcModelKeyCache = @{}
function Get-ProjectSpecKey([string]$Spec) {
    $hit = $script:PcSpecKeyCache[$Spec]
    if ($null -ne $hit) { return $hit }
    $t = ConvertTo-CanonicalSpecName-Pc (Remove-SpecStagePrefix $Spec)
    $key = ((@($t.Trim() -split '\s+') -join " ").ToLowerInvariant())
    $script:PcSpecKeyCache[$Spec] = $key
    return $key
}
function Get-ProjectModelKey([string]$Model) {
    $hit = $script:PcModelKeyCache[$Model]
    if ($null -ne $hit) { return $hit }
    $key = (Get-ProjectModelName $Model).ToLowerInvariant()
    $script:PcModelKeyCache[$Model] = $key
    return $key
}
function Get-ProjectCellKey([string]$Model, [string]$Spec) { return ((Get-ProjectModelKey $Model) + "`t" + (Get-ProjectSpecKey $Spec)) }

# Префикс стадии («П», «РД», «ПП») перед пробелом или цифрой; «Пров …» префиксом не является
function Remove-SpecStagePrefix([string]$Name) {
    $t = ([string]$Name).Trim()
    if ($t -match '^(РД|ПП|П)(?=\s|\d)\s*(.*)$') { return $Matches[2] }
    return $t
}

# Имя спецификации = группа + хвост. Группа — имя без хвоста и без префикса стадии: «7.1_Диаметр_Воздуховоды» и «7.1_Диаметр_Детали возд-в» — разные группы,
# «5.3_Ширина_Отопит приборы» — отдельная; хвост — «Все», «Без значения» (в т.ч. «Без значния»), «Пустые», «Уникальные_…», «999»:
# 1_Классификатор_Все, 1_Классификатор_999, 1_Классификатор_Уникальные_* -> группа «1_Классификатор»; 99_Толщина стенки труб и 99_Толщина стенки труб_Без значения ->
# группа «99_Толщина стенки труб». Хвост ищется после названия и объекта (с третьего звена имени); нет хвоста — группа = всё имя.
$script:PcTailPattern = '^(?i)(Все|Пустые|Без\s+знач\w*|Уникальные|999)$'
function Split-SpecName([string]$Name) {
    $t = ((ConvertTo-CanonicalSpecName-Pc (Remove-SpecStagePrefix $Name)).Trim()) -replace '\s+', ' '
    if ($t -match '^\d+(?:\.\d+)*_') {
        $seg = @($t -split '_')
        for ($i = 2; $i -lt $seg.Count; $i++) {
            if ($seg[$i].Trim() -match $script:PcTailPattern) {
                return [pscustomobject]@{ Group = (@($seg[0..($i - 1)]) -join "_").Trim(); Tail = (@($seg[$i..($seg.Count - 1)]) -join "_").Trim() }
            }
        }
        return [pscustomobject]@{ Group = $t; Tail = "" }
    }
    $g = ($t -replace '(?i)[_\s]+(Пустые|Без\s+знач\w*|Все)\s*$', '').Trim()
    if (-not $g) { return [pscustomobject]@{ Group = $t; Tail = "" } }
    return [pscustomobject]@{ Group = $g; Tail = $t.Substring($g.Length).Trim([char]'_', [char]' ') }
}

function Get-SpecGroupName([string]$Name) { return (Split-SpecName $Name).Group }

# Название без хвоста — для порядка (оставлено для совместимости; группа и есть имя без хвоста)
function Get-SpecStrippedName([string]$Name) { return (Get-SpecGroupName $Name).ToLowerInvariant() }

# Подпись строки спецификации внутри раскрытой группы — только хвост («Все», «Без значения», «Пустые» …); спецификация без хвоста — «(основная)».
# Спецификация, вручную перенесённая в другую группу, подписана полным именем
function Get-SpecRowLabel([string]$Spec, [string]$GroupName) {
    $p = Split-SpecName $Spec
    if ($p.Group -ieq ([string]$GroupName).Trim()) {
        if ($p.Tail) { return $p.Tail }
        return "(основная)"
    }
    return (($p.Group + $(if ($p.Tail) { "_" + $p.Tail } else { "" })).Trim())
}

function ConvertTo-CanonicalSpecName-Pc([string]$Name) { return ([string]$Name).Replace("значния", "значения").Replace(([string][char]0x43 + "антехприборы"), ([string][char]0x421 + "антехприборы")).Replace(([string][char]0x63 + "антехприборы"), ([string][char]0x441 + "антехприборы")) }

function Get-SpecTailRank([string]$Name) {
    $tail = (Split-SpecName $Name).Tail
    if (-not $tail) { return -1 }                 # без хвоста (99_Толщина стенки труб) — первой в группе
    if ($tail -match '(?i)^Все$') { return 0 }
    if ($tail -match '(?i)^Пустые$') { return 1 }
    if ($tail -match '(?i)^Без\s+знач') { return 2 }
    if ($tail -match '^999') { return 3 }
    if ($tail -match '(?i)^Уникальные') { return 4 }
    return 5
}

# Ключ сортировки: ведущий номер как версия (0 < 1 < 2.1 < 5.7 < 13 < 14.7 < 99), дальше имя без регистра; без номера — в конец
function Get-ProjectSortKey([string]$Name) {
    $t = Remove-SpecStagePrefix $Name
    if ($t -match '^(\d+(?:\.\d+)*)(.*)$') {
        $parts = @($Matches[1].Split(".") | ForEach-Object { $_.PadLeft(6, "0") })
        return (($parts -join ".") + "|" + $Matches[2].ToLowerInvariant())
    }
    return ("999999|" + $t.ToLowerInvariant())
}

function Get-SpecGroup($Cfg, [string]$Spec) {
    $k = Get-ProjectSpecKey $Spec
    if ($Cfg.GroupOverride.ContainsKey($k)) { return [string]$Cfg.GroupOverride[$k] }
    return (Get-SpecGroupName $Spec)
}

# Группы по порядку номеров; в группе только хвосты: «Все», «Пустые», «Без значения» …
function Get-ProjectGroups($Cfg) {
    if ($null -ne $Cfg.GroupsCache) { return $Cfg.GroupsCache }
    $by = @{}
    $order = New-Object System.Collections.Generic.List[string]
    foreach ($s in $Cfg.Specs) {
        $g = Get-SpecGroup $Cfg $s
        if (-not $by.ContainsKey($g)) { $by[$g] = New-Object System.Collections.Generic.List[string]; $order.Add($g) }
        $by[$g].Add($s)
    }
    $out = New-Object System.Collections.Generic.List[object]
    foreach ($g in @($order | Sort-Object { Get-ProjectSortKey $_ })) {
        $specs = @($by[$g] | Sort-Object @{ Expression = { Get-SpecTailRank $_ } }, @{ Expression = { $_.ToLowerInvariant() } })
        $out.Add([pscustomobject]@{ Name = $g; Specs = $specs })
    }
    $Cfg.GroupsCache = $out.ToArray()
    return $Cfg.GroupsCache
}

# ---------- разделы и модели ----------
# Код раздела по имени файла — только «_КОД_» (без учёта регистра; кириллица и латиница). Признаки IOS/ИОС/ВИС не используются:
# они есть и у сводной, и у TASK, и у граничных условий. Сводная (SVOD/СВОД) — не раздел. Нет кода — пустая строка.
function Get-ModelSectionFromName([string]$Name) {
    $n = (Get-ProjectModelName $Name).ToUpperInvariant()
    if ($n -match '_SVOD_|_СВОД_|СВОДН') { return "" }
    $best = ""; $bestPos = [int]::MaxValue
    foreach ($s in $script:PcSections) {
        foreach ($v in $s.Names) {
            $pos = $n.IndexOf("_" + $v + "_", [System.StringComparison]::Ordinal)
            if ($pos -ge 0 -and $pos -lt $bestPos) { $best = [string]$s.Code; $bestPos = $pos }
        }
    }
    # запасной вариант — по имени папки, где лежит модель (если в имени файла кода нет и это не модель задания/базовая)
    if (-not $best -and $n -notmatch '_(TASK|AR|KR|TH|BASE)_') { return (Get-ModelSectionFromFolder $Name) }
    return $best
}

# Раздел по папке модели: ближайшая папка (папка RVT пропускается), её имя целиком или одно из слов имени равно коду раздела («АВТ», «HC» …). Нет пути/кода — ""
function Get-ModelSectionFromFolder([string]$PathOrName) {
    $t = ([string]$PathOrName).Trim().Replace("/", "\")
    $parts = @($t.Split([char]'\') | Where-Object { $_ })
    if ($parts.Count -lt 2) { return "" }
    $i = $parts.Count - 2
    while ($i -ge 0 -and $parts[$i].ToUpperInvariant() -eq "RVT") { $i-- }
    if ($i -lt 0) { return "" }
    $tokens = @(($parts[$i].ToUpperInvariant() -replace '[_\-\.\(\)]', ' ') -split '\s+' | Where-Object { $_ })
    foreach ($s in $script:PcSections) { foreach ($v in $s.Names) { if ($tokens -contains $v) { return [string]$s.Code } } }
    return ""
}

function Test-TargetSectionCode([string]$Code) {
    if (-not $Code) { return $false }
    foreach ($s in $script:PcSections) { if ([string]$s.Code -eq $Code) { return $true } }
    return $false
}

# Код в имени вида «_IOS_BF_» / «_ORG_GU_», которого нет среди целевых разделов. В таблицу такая модель не входит.
function Get-NonTargetSectionToken([string]$PathOrName) {
    if (Get-ModelSectionFromName $PathOrName) { return "" }
    $n = (Get-ProjectModelName $PathOrName).ToUpperInvariant()
    if ($n -match '_SVOD_|_СВОД_|СВОДН') { return "" }
    if ($n -match '_(?:IOS|ORG)_([A-ZА-Я]{2,5})_') {
        $tok = $Matches[1]
        if ($script:PcNonSectionTokens -notcontains $tok) { return $tok }
    }
    return ""
}

function Measure-NonTargetModels($Names) {
    $n = 0
    $seen = @{}
    foreach ($raw in @($Names)) {
        $name = Get-ProjectModelName $raw
        if (-not $name) { continue }
        $key = Get-ProjectModelKey $name
        if (-not $key -or $seen.ContainsKey($key)) { continue }
        if (Get-NonTargetSectionToken $name) { $seen[$key] = $true; $n++ }
    }
    return $n
}

function Format-SkippedModelsStatus([int]$Count) {
    if ($Count -le 0) { return "" }
    return ("Пропущено {0} моделей без целевого раздела" -f $Count)
}

function Clear-ProjectHeld($Cfg) {
    if ($null -eq $Cfg) { return }
    if ($Cfg.HeldModels) { $Cfg.HeldModels.Clear() } else { $Cfg.HeldModels = New-Object System.Collections.Generic.List[object] }
    if ($Cfg.HeldCells) { $Cfg.HeldCells.Clear() } else { $Cfg.HeldCells = New-Object System.Collections.Generic.List[string] }
}

function Add-ProjectHeldModel($Cfg, [string]$Name, [string]$Section) {
    if (-not $Name) { return }
    if (-not $Cfg.HeldModels) { $Cfg.HeldModels = New-Object System.Collections.Generic.List[object] }
    $key = Get-ProjectModelKey $Name
    foreach ($m in $Cfg.HeldModels) { if ((Get-ProjectModelKey $m.Name) -eq $key) { return } }
    $Cfg.HeldModels.Add([pscustomobject]@{ Name = $Name; Section = $Section })
}

function Test-ProjectHeldModel($Cfg, [string]$Name) {
    if (-not $Cfg -or -not $Cfg.HeldModels -or -not $Name) { return $false }
    $key = Get-ProjectModelKey $Name
    foreach ($m in $Cfg.HeldModels) { if ((Get-ProjectModelKey $m.Name) -eq $key) { return $true } }
    return $false
}

# Инженерная модель — только целевой раздел из $script:PcSections (имя или папка). Сводная, шаблон, TASK и чужой код (BF, GU…) в распределении не участвуют.
function Test-ProjectEngineeringModel([string]$PathOrName) { return ((Get-ModelSectionFromName $PathOrName) -ne "") }

function Find-ProjectModel($Cfg, [string]$PathOrName) {
    $n = (Get-ProjectModelName $PathOrName).ToLowerInvariant()
    foreach ($m in $Cfg.Models) { if ($m.Name.ToLowerInvariant() -eq $n) { return $m } }
    return $null
}

# Раздел модели — всегда по имени файла
function Get-ProjectModelSection($Cfg, [string]$PathOrName) { return (Get-ModelSectionFromName $PathOrName) }

# Прочитанные из папки инженерные модели: новые добавляются. Нет целевого раздела — модель не принимается.
function Update-ProjectModels($Cfg, $ModelNames) {
    foreach ($raw in @($ModelNames)) {
        $name = Get-ProjectModelName $raw
        if (-not $name) { continue }
        $code = Get-ModelSectionFromName $raw                     # полный путь: раздел может определиться по папке
        if (-not (Test-TargetSectionCode $code)) { continue }
        if (Find-ProjectModel $Cfg $name) { continue }
        $Cfg.Models.Add([pscustomobject]@{ Name = $name; Section = $code })
    }
    Update-ProjectSections $Cfg
}

# Разделы настройки — коды моделей в порядке таблицы $script:PcSections. Чужой код колонки не получает.
function Update-ProjectSections($Cfg) {
    $Cfg.Sections.Clear()
    foreach ($s in $script:PcSections) {
        if (@($Cfg.Models | Where-Object { $_.Section -eq $s.Code }).Count -gt 0) { $Cfg.Sections.Add([pscustomobject]@{ Code = [string]$s.Code; Title = [string]$s.Title }) }
    }
}

# Имена связей сводной, из которых строятся колонки: без сводной и служебных (TASK/AR/…/SVOD), только целевой раздел.
# Чужой код не входит. Пустой результат — модели настройки не трогать.
function Select-SummaryLinkNames($Names) {
    $out = New-Object System.Collections.Generic.List[string]
    $seen = @{}
    foreach ($raw in @($Names)) {
        $name = Get-ProjectModelName $raw
        if (-not $name) { continue }
        $key = Get-ProjectModelKey $name
        if (-not $key -or $seen.ContainsKey($key)) { continue }
        $upper = $name.ToUpperInvariant()
        if ($upper -match '_SVOD_|_СВОД_|СВОДН') { continue }
        if ($script:PcNonSectionTokens -contains $upper) { continue }
        if (-not (Test-ProjectEngineeringModel $name)) { continue }
        $seen[$key] = $true
        $out.Add($name)
    }
    return $out.ToArray()
}

# Имена локальных копий Revit: текущий пользователь Windows (USERNAME / USER).
function Get-RevitLocalUserNames {
    $out = New-Object System.Collections.Generic.List[string]
    $seen = @{}
    foreach ($envName in @("USERNAME", "USER")) {
        $v = ""
        try { $v = [Environment]::GetEnvironmentVariable($envName) } catch { $v = "" }
        if (-not $v) { continue }
        $v = $v.Trim()
        if (-not $v) { continue }
        $k = $v.ToLowerInvariant()
        if ($seen.ContainsKey($k)) { continue }
        $seen[$k] = $true
        $out.Add($v)
    }
    return $out.ToArray()
}

# Имя связи или файла без пути, «.rvt», «.rvt : N», суффикса «_отсоединено» и «_<пользователь>».
function Get-LinkCanonicalName([string]$Name) {
    $t = Get-ProjectModelName $Name
    if (-not $t) { return "" }
    $cut = $t.IndexOf(".rvt", [System.StringComparison]::OrdinalIgnoreCase)
    if ($cut -ge 0) { $t = $t.Substring(0, $cut).Trim() }
    $t = [regex]::Replace($t, '\s+:\s*\d+\s*$', '')
    $guard = 0
    do {
        $prev = $t
        $suf = "_отсоединено"
        if ($t.Length -gt $suf.Length -and $t.EndsWith($suf, [System.StringComparison]::OrdinalIgnoreCase)) { $t = $t.Substring(0, $t.Length - $suf.Length).Trim() }
        foreach ($user in @(Get-RevitLocalUserNames)) {
            if (-not $user) { continue }
            $us = "_" + $user
            if ($t.Length -gt $us.Length -and $t.EndsWith($us, [System.StringComparison]::OrdinalIgnoreCase)) { $t = $t.Substring(0, $t.Length - $us.Length).Trim() }
        }
        $guard++
    } while ($t -ne $prev -and $guard -lt 4)
    return $t
}

# Ключ сопоставления связь ↔ файл ↔ модель таблицы: без учёта регистра.
function Get-LinkMatchKey([string]$Name) {
    $t = Get-LinkCanonicalName $Name
    if (-not $t) { return "" }
    return $t.ToLowerInvariant()
}

# Связей меньше половины моделей таблицы — колонки сами не пересобирать.
function Test-SummaryLinksThin([int]$Models, [int]$Links) {
    if ($Models -le 0) { return $false }
    return ($Links -lt ($Models / 2.0))
}

# План замены моделей: имя из таблицы сохраняется, если ключ совпал со связью; нет файла — модель всё равно остаётся.
# FolderPaths = $null — файлы не проверялись (Missing пустой).
function Get-SummaryLinkPlan($Cfg, $Names, $FolderPaths) {
    $links = @(Select-SummaryLinkNames $Names)
    $existing = @{}
    foreach ($m in $Cfg.Models) {
        $k = Get-LinkMatchKey $m.Name
        if ($k -and -not $existing.ContainsKey($k)) { $existing[$k] = [string]$m.Name }
    }
    $files = @{}
    $knowFiles = ($null -ne $FolderPaths)
    if ($knowFiles) {
        foreach ($p in @($FolderPaths)) {
            if (-not $p) { continue }
            $k = Get-LinkMatchKey $p
            if ($k -and -not $files.ContainsKey($k)) { $files[$k] = (Get-ProjectModelName $p) }
        }
    }
    $chosen = @{}
    $order = New-Object System.Collections.Generic.List[string]
    $missing = New-Object System.Collections.Generic.List[string]
    foreach ($link in $links) {
        $k = Get-LinkMatchKey $link
        if (-not $k -or $chosen.ContainsKey($k)) { continue }
        $name = Get-LinkCanonicalName $link
        if ($existing.ContainsKey($k)) { $name = $existing[$k] }
        elseif ($files.ContainsKey($k)) { $name = $files[$k] }
        if (-not $name) { continue }
        $chosen[$k] = $name
        $order.Add($k)
        if ($knowFiles -and -not $files.ContainsKey($k)) { $missing.Add($name) }
    }
    $names = New-Object System.Collections.Generic.List[string]
    foreach ($k in $order) { $names.Add($chosen[$k]) }
    $leaving = New-Object System.Collections.Generic.List[string]
    foreach ($m in $Cfg.Models) {
        $k = Get-LinkMatchKey $m.Name
        if (-not $chosen.ContainsKey($k)) { $leaving.Add([string]$m.Name) }
    }
    return [pscustomobject]@{ Names = $names.ToArray(); Missing = $missing.ToArray(); Leaving = $leaving.ToArray() }
}

# Убирает из карты «Выборочно» модели, которых больше нет. Ключ пары — «модель|спека» или «модель|*». $null не трогаем.
function Remove-AsgMapModels($Map, $Keep) {
    if ($null -eq $Map) { return }
    $keys = New-Object System.Collections.Generic.List[string]
    foreach ($k in $Map.Keys) { $keys.Add([string]$k) }
    foreach ($k in $keys) {
        $model = $k.Split("|")[0]
        if (-not $Keep.ContainsKey((Get-ProjectModelKey $model))) { [void]$Map.Remove($k) }
    }
}

# Заменяет модели настройки списком связей. Связей нет — настройка не меняется (Applied = $false).
# Модель без файла в папке не выбрасывается. Совпавшая по ключу модель сохраняет прежнее имя, галочки и пары «Выборочно».
function Set-ProjectModelsFromLinkNames($Cfg, $Names, $FolderPaths) {
    $plan = Get-SummaryLinkPlan $Cfg $Names $FolderPaths
    $skipped = Measure-NonTargetModels $Names
    if (@($plan.Names).Count -eq 0) { return [pscustomobject]@{ Applied = $false; Names = @(); Skipped = $skipped; Unknown = ""; Missing = @(); Leaving = @($plan.Leaving) } }
    $snapshot = New-Object System.Collections.Generic.List[object]
    foreach ($m in $Cfg.Models) { $snapshot.Add($m) }
    $Cfg.Models.Clear()
    Update-ProjectModels $Cfg @($plan.Names)
    if ($Cfg.Models.Count -eq 0) {
        foreach ($m in $snapshot) { $Cfg.Models.Add($m) }
        Update-ProjectSections $Cfg
        return [pscustomobject]@{ Applied = $false; Names = @(); Skipped = $skipped; Unknown = ""; Missing = @(); Leaving = @() }
    }
    $keep = @{}
    foreach ($m in $Cfg.Models) { $keep[(Get-ProjectModelKey $m.Name)] = $true }
    $cellKeys = New-Object System.Collections.Generic.List[string]
    foreach ($k in $Cfg.Cells.Keys) { $cellKeys.Add([string]$k) }
    foreach ($k in $cellKeys) {
        $mk = $k.Split("`t")[0]
        if (-not $keep.ContainsKey($mk)) { [void]$Cfg.Cells.Remove($k) }
    }
    Remove-AsgMapModels $Cfg.AsgModels $keep
    Remove-AsgMapModels $Cfg.AsgPairs $keep
    $names = New-Object System.Collections.Generic.List[string]
    foreach ($m in $Cfg.Models) { $names.Add([string]$m.Name) }
    return [pscustomobject]@{ Applied = $true; Names = $names.ToArray(); Skipped = $skipped; Unknown = ""; Missing = @($plan.Missing); Leaving = @($plan.Leaving) }
}

# Имена из списка, которых нет среди путей: тот же ключ, что у связи (без регистра, без .rvt и локальных суффиксов).
function Get-NamesMissingInPaths($Names, $FolderPaths) {
    $have = @{}
    foreach ($p in @($FolderPaths)) {
        $k = Get-LinkMatchKey $p
        if ($k) { $have[$k] = $true }
    }
    $miss = New-Object System.Collections.Generic.List[string]
    foreach ($raw in @($Names)) {
        $name = Get-LinkCanonicalName $raw
        $k = Get-LinkMatchKey $name
        if ($k -and -not $have.ContainsKey($k)) { $miss.Add($name) }
    }
    return $miss.ToArray()
}

function Get-ProjectSectionTitle($Cfg, [string]$Code) {
    if ($script:PcDefaultTitles.ContainsKey($Code)) { return $script:PcDefaultTitles[$Code] }
    return $Code
}

function Get-ProjectSectionModels($Cfg, [string]$Code) {
    $out = New-Object System.Collections.Generic.List[string]
    foreach ($m in $Cfg.Models) { if ([string]$m.Section -eq $Code) { $out.Add([string]$m.Name) } }
    return $out.ToArray()
}

function Get-ProjectViewSections($Cfg) {
    $out = New-Object System.Collections.Generic.List[object]
    foreach ($s in $Cfg.Sections) { $out.Add([pscustomobject]@{ Code = [string]$s.Code; Title = [string]$s.Title }) }
    return $out.ToArray()
}

# ---------- «В проекте» (какие спецификации вообще участвуют в проекте, отмечает BIM-менеджер по BEP; включая сводную) ----------
function Test-ProjectSpecIn($Cfg, [string]$Spec) {
    if ($null -eq $Cfg.InProject) { return $true }
    return $Cfg.InProject.ContainsKey((Get-ProjectSpecKey $Spec))
}

# Состояние группы: Checked — все её спеки в проекте, Indeterminate — часть, Unchecked — ни одной
function Get-ProjectGroupInState($Cfg, $Specs) {
    $n = 0; $on = 0
    foreach ($s in @($Specs)) { $n++; if (Test-ProjectSpecIn $Cfg ([string]$s)) { $on++ } }
    if ($n -gt 0 -and $on -eq $n) { return "Checked" }
    if ($on -gt 0) { return "Indeterminate" }
    return "Unchecked"
}

# Сводная: какие спецификации проекта идут в сводную модель (подмножество «В проекте»). Нет раздела [summary] — идут все спецификации проекта
function Test-ProjectSpecInSummary($Cfg, [string]$Spec) {
    if ($null -eq $Cfg.Summary) { return $true }
    return $Cfg.Summary.ContainsKey((Get-ProjectSpecKey $Spec))
}

function Initialize-ProjectSummary($Cfg) {
    if ($null -eq $Cfg.Summary) {
        $Cfg.Summary = @{}
        foreach ($s in $Cfg.Specs) { if (Test-ProjectSpecIn $Cfg $s) { $Cfg.Summary[(Get-ProjectSpecKey $s)] = $true } }
    }
}

function Initialize-ProjectInProject($Cfg) {
    if ($null -eq $Cfg.InProject) { $Cfg.InProject = @{}; foreach ($s in $Cfg.Specs) { $Cfg.InProject[(Get-ProjectSpecKey $s)] = $true } }
}

# Отметка «В проекте» у спецификаций: попавшая в проект спецификация сразу идёт и в сводную (снять можно в столбце «Сводная»); снятая — уходит и оттуда
function Set-ProjectInProjectBulk($Cfg, $Specs, [bool]$On) {
    Initialize-ProjectInProject $Cfg
    foreach ($s in @($Specs)) {
        $k = Get-ProjectSpecKey ([string]$s)
        if ($On) { $Cfg.InProject[$k] = $true; if ($null -ne $Cfg.Summary) { $Cfg.Summary[$k] = $true } }
        else { $Cfg.InProject.Remove($k); if ($null -ne $Cfg.Summary) { $Cfg.Summary.Remove($k) } }
    }
}

# Группа: из «снято/частично» -> отметить все её спецификации, из «отмечено» -> снять все
function Invoke-ProjectGroupInClick($Cfg, $Specs) {
    Initialize-ProjectInProject $Cfg
    $state = Get-ProjectGroupInState $Cfg $Specs
    Set-ProjectInProjectBulk $Cfg $Specs ($state -ne "Checked")
}

# Одна спецификация внутри раскрытой группы
function Invoke-ProjectSpecInClick($Cfg, [string]$Spec) {
    Initialize-ProjectInProject $Cfg
    Set-ProjectInProjectBulk $Cfg @($Spec) (-not (Test-ProjectSpecIn $Cfg $Spec))
}

# Спецификации из Specs, участвующие в проекте
function Get-ProjectActiveSpecs($Cfg, $Specs) {
    return @(@($Specs) | Where-Object { Test-ProjectSpecIn $Cfg ([string]$_) })
}

# Состояние «Сводной» у набора спецификаций (учитываются только спецификации проекта): Checked / Indeterminate / Unchecked
function Get-ProjectSummaryState($Cfg, $Specs) {
    $act = @(Get-ProjectActiveSpecs $Cfg $Specs)
    $on = @($act | Where-Object { Test-ProjectSpecInSummary $Cfg ([string]$_) }).Count
    if ($act.Count -gt 0 -and $on -eq $act.Count) { return "Checked" }
    if ($on -gt 0) { return "Indeterminate" }
    return "Unchecked"
}

# Щелчок по «Сводной»: из «снято/частично» -> все спецификации проекта набора идут в сводную, из «отмечено» -> ни одна
function Invoke-ProjectSummaryClick($Cfg, $Specs) {
    $act = @(Get-ProjectActiveSpecs $Cfg $Specs)
    if ($act.Count -eq 0) { return }
    $state = Get-ProjectSummaryState $Cfg $Specs
    Initialize-ProjectSummary $Cfg
    foreach ($s in $act) {
        $k = Get-ProjectSpecKey ([string]$s)
        if ($state -eq "Checked") { $Cfg.Summary.Remove($k) } else { $Cfg.Summary[$k] = $true }
    }
}

# ---------- спецификации и ячейки ----------
# Имена спецификаций в настройке — без префикса стадии (стадия — свойство проекта); старые имена с префиксом приводятся
function Update-ProjectSpecs($Cfg, $Names) {
    $seen = @{}
    $list = New-Object System.Collections.Generic.List[string]
    foreach ($n in @($Names)) {
        $t = (Remove-SpecStagePrefix ([string]$n)).Trim()
        $k = Get-ProjectSpecKey $t
        if ($t -and -not $seen.ContainsKey($k)) { $seen[$k] = $true; $list.Add($t) }
    }
    $oldKeys = @{}
    foreach ($o in $Cfg.Specs) { $oldKeys[(Get-ProjectSpecKey $o)] = $true }
    $Cfg.Specs = $list
    $Cfg.GroupsCache = $null
    foreach ($k in @($Cfg.Cells.Keys)) { $sp = $k.Split("`t", 2)[1]; if (-not $seen.ContainsKey($sp)) { $Cfg.Cells.Remove($k) } }
    foreach ($k in @($Cfg.GroupOverride.Keys)) { if (-not $seen.ContainsKey($k)) { $Cfg.GroupOverride.Remove($k) } }
    foreach ($k in @($Cfg.Categories.Keys)) { if (-not $seen.ContainsKey($k)) { $Cfg.Categories.Remove($k) } }
    foreach ($map in @($Cfg.InProject, $Cfg.Summary)) {
        if ($null -ne $map) { foreach ($k in @($map.Keys)) { if (-not $seen.ContainsKey($k)) { $map.Remove($k) } } }
    }
    # новая спецификация попадает в проект, если в её группе все прежние спецификации уже в проекте (отметки по отдельным спецификациям не сбрасываются)
    if ($null -ne $Cfg.InProject) {
        foreach ($g in @(Get-ProjectGroups $Cfg)) {
            $old = @($g.Specs | Where-Object { $oldKeys.ContainsKey((Get-ProjectSpecKey $_)) })
            $new = @($g.Specs | Where-Object { -not $oldKeys.ContainsKey((Get-ProjectSpecKey $_)) })
            if ($old.Count -gt 0 -and $new.Count -gt 0 -and @($old | Where-Object { -not $Cfg.InProject.ContainsKey((Get-ProjectSpecKey $_)) }).Count -eq 0) {
                foreach ($sp in $new) { $k = Get-ProjectSpecKey $sp; $Cfg.InProject[$k] = $true; if ($null -ne $Cfg.Summary) { $Cfg.Summary[$k] = $true } }
            }
        }
    }
}

function Set-ProjectSpecGroup($Cfg, [string]$Spec, [string]$Group) {
    $g = ([string]$Group).Trim()
    $k = Get-ProjectSpecKey $Spec
    $Cfg.GroupsCache = $null
    if (-not $g -or $g -eq (Get-SpecGroupName $Spec)) { if ($Cfg.GroupOverride.ContainsKey($k)) { $Cfg.GroupOverride.Remove($k) }; return }
    $Cfg.GroupOverride[$k] = $g
}

function Get-ProjectCell($Cfg, [string]$Model, [string]$Spec) {
    $k = Get-ProjectCellKey $Model $Spec
    if ($Cfg.Cells.ContainsKey($k)) { return $Cfg.Cells[$k] }
    return $null
}

function Set-ProjectCell($Cfg, [string]$Model, [string]$Spec, [int]$State, [string]$Origin, [int]$Count = -1) {
    $Cfg.Cells[(Get-ProjectCellKey $Model $Spec)] = [pscustomobject]@{ State = $State; Origin = $Origin; Count = $Count }
}

# Агрегат по набору спецификаций × набору моделей: сколько отмечено, есть ли ручные/неопределённые, сумма элементов
function Get-ProjectAggregate($Cfg, $Models, $Specs) {
    $total = 0; $on = 0; $manual = 0; $undef = 0; $sum = 0; $known = 0
    foreach ($m in @($Models)) {
        foreach ($s in @($Specs)) {
            $total++
            $c = Get-ProjectCell $Cfg ([string]$m) ([string]$s)
            if ($null -eq $c) { continue }
            if ($c.State -eq 1) { $on++ }
            if ($c.Origin -eq "m") { $manual++ }
            if ($c.Origin -eq "x") { $undef++ }
            if ($c.Count -ge 0) { $sum += $c.Count; $known++ }
        }
    }
    $state = "Unchecked"
    if ($total -gt 0 -and $on -eq $total) { $state = "Checked" } elseif ($on -gt 0) { $state = "Indeterminate" }
    return [pscustomobject]@{ State = $state; Total = $total; On = $on; Manual = $manual; Undefined = $undef; Count = $sum; HasCount = ($known -gt 0) }
}

# Цвет ячейки/узла: красный — не определено, серый — частично, жёлтый — есть ручные правки, зелёный — авто и отмечено, иначе без цвета
function Get-ProjectAggregateColor($Agg) {
    if ($Agg.Undefined -gt 0) { return "red" }
    if ($Agg.State -eq "Indeterminate") { return "gray" }
    if ($Agg.Manual -gt 0) { return "yellow" }
    if ($Agg.State -eq "Checked") { return "green" }
    return "none"
}

function Format-ProjectAggregate($Agg) {
    $glyph = switch ($Agg.State) { "Checked" { [string][char]0x2611 } "Indeterminate" { [string][char]0x25EA } default { [string][char]0x2610 } }
    if ($Agg.Undefined -gt 0 -and $Agg.Total -eq $Agg.Undefined) { return "? " + [char]0x2014 }
    if ($Agg.HasCount) { return ("{0} {1}" -f $glyph, $Agg.Count) }
    return $glyph
}

# Число элементов (результат пробной вставки, «Посчитать элементы») — только для информации: галочки и признак ячейки не меняются.
# Ячейки, которых ещё нет, создаются неотмеченными с признаком «c» (только число). Counts: объекты @{ Model; Spec; Count } (Count < 0 — определить не удалось)
function Set-ProjectCellsFromCounts($Cfg, $Counts) {
    foreach ($r in @($Counts)) {
        $model = [string]$r.Model; $spec = [string]$r.Spec; $n = [int]$r.Count
        $old = Get-ProjectCell $Cfg $model $spec
        if ($old) { Set-ProjectCell $Cfg $model $spec $old.State $old.Origin $n }
        else { Set-ProjectCell $Cfg $model $spec 0 "c" $n }
    }
}

# Распределение уже выполнено: есть ячейки, размеченные справочником («a») или вручную («m»); ячейки «только число» («c») не считаются
function Test-ProjectDistributed($Cfg) {
    foreach ($c in $Cfg.Cells.Values) { if ($c.Origin -eq "a" -or $c.Origin -eq "m") { return $true } }
    return $false
}

# ---------- «В проекте»: только [inproject] настройки проекта; перенос прежнего файла выбора ----------
# Раньше галочки «В проекте» хранились ещё и в отдельном файле отчеты\проект\выбор_<проект>.txt (по строке на спецификацию).
# Скрипты Revit всегда читали только [inproject], поэтому теперь источник один — [inproject]. Старый файл выбора при первом открытии окна
# переносится (Invoke-ProjectStateMigration) и переименовывается в выбор_<проект>.migrated_<yyyy-MM-dd>.txt. Он не удаляется.
function Get-ProjectSelectionPath([string]$ProjectConfigPath) {
    if (-not $ProjectConfigPath) { return "" }
    $dir = Split-Path -Parent $ProjectConfigPath
    $leaf = Split-Path -Leaf $ProjectConfigPath
    if ($leaf -match '^распределение_(.+)$') { $leaf = "выбор_" + $Matches[1] } else { $leaf = "выбор_" + $leaf }
    return (Join-Path $dir $leaf)
}

# Имена из прежнего файла выбора; файла нет — $null (пустой файл — пустой массив). Только чтение (перенос).
function Read-ProjectSelection([string]$Path) {
    if (-not $Path -or -not (Test-Path -LiteralPath $Path)) { return $null }
    $text = [System.IO.File]::ReadAllText($Path, (New-Object System.Text.UTF8Encoding $true))
    $out = New-Object System.Collections.Generic.List[string]
    foreach ($line in @($text -split "`r?`n")) {
        $t = $line.Trim()
        if ($t -and -not $t.StartsWith("#")) { $out.Add((Remove-SpecStagePrefix $t).Trim()) }
    }
    return $out.ToArray()
}

# Имена «В проекте» так, как они лягут в [inproject]: отмеченные спецификации в порядке общего файла + имена не из общего файла.
# Нет раздела [inproject] (в проекте все) — $null.
function Get-ProjectInProjectNames($Cfg) {
    if ($null -eq $Cfg -or $null -eq $Cfg.InProject) { return $null }
    $L = New-Object System.Collections.Generic.List[string]
    $have = @{}
    foreach ($sp in $Cfg.Specs) { if ($Cfg.InProject.ContainsKey((Get-ProjectSpecKey ([string]$sp)))) { $L.Add([string]$sp); $have[(Get-ProjectSpecKey ([string]$sp))] = $true } }
    if ($Cfg.SelectionUnknown) { foreach ($u in $Cfg.SelectionUnknown) { $k = Get-ProjectSpecKey ([string]$u); if ($u -and -not $have.ContainsKey($k)) { $L.Add([string]$u); $have[$k] = $true } } }
    return , $L.ToArray()
}

# Есть ли уже галочки «В проекте» (нет раздела [inproject] — в проекте все, это тоже выбор)
function Test-ProjectHasSelection($Cfg) {
    if ($null -eq $Cfg.InProject) { return ($Cfg.Specs.Count -gt 0) }
    return (($Cfg.InProject.Count + @($Cfg.SelectionUnknown).Count) -gt 0)
}

# Резервная копия перед любой записью переноса: всё из отчеты\проект\ (кроме прежних _backup_*) и история/журналы из отчеты\история\
# -> отчеты\проект\_backup_<yyyy-MM-dd_HH-mm>\ (история — в подпапку «история»). Имя занято — _2, _3 … Не удалось скопировать — исключение,
# перенос тогда ничего не пишет. Автоочистка (report_layout.ps1) папку проект\ и _backup_* не трогает.
function New-ProjectStateBackup([string]$ReportDir, $Now = $null) {
    if ($null -eq $Now) { $Now = Get-Date }
    $projDir = Get-ProjectConfigDir $ReportDir
    if (-not (Test-Path -LiteralPath $projDir)) { New-Item -ItemType Directory -Path $projDir -Force | Out-Null }
    $base = Join-Path $projDir ("_backup_" + ([datetime]$Now).ToString("yyyy-MM-dd_HH-mm", [System.Globalization.CultureInfo]::InvariantCulture))
    $dir = $base; $n = 2
    while (Test-Path -LiteralPath $dir) { $dir = $base + "_" + $n; $n++ }
    New-Item -ItemType Directory -Path $dir -Force | Out-Null
    foreach ($item in @(Get-ChildItem -LiteralPath $projDir -Force -ErrorAction Stop)) {
        if ($item.Name -like "_backup_*") { continue }
        Copy-Item -LiteralPath $item.FullName -Destination (Join-Path $dir $item.Name) -Recurse -Force -ErrorAction Stop
    }
    $hist = Join-Path $ReportDir "история"
    if (Test-Path -LiteralPath $hist) {
        $hdst = Join-Path $dir "история"
        foreach ($f in @(Get-ChildItem -LiteralPath $hist -File -Force -ErrorAction Stop)) {
            if ($f.Name -notmatch '^(история_.+|задание_история(_.+)?|задание_журнал(_.+)?)\.txt$') { continue }
            if (-not (Test-Path -LiteralPath $hdst)) { New-Item -ItemType Directory -Path $hdst -Force | Out-Null }
            Copy-Item -LiteralPath $f.FullName -Destination (Join-Path $hdst $f.Name) -Force -ErrorAction Stop
        }
    }
    return $dir
}

# Новое имя прежнего файла выбора: выбор_<проект>.migrated_<yyyy-MM-dd>.txt (занято — _2, _3 …). Ничего не затирается.
function Get-ProjectSelectionMigratedPath([string]$SelectionPath, $Now = $null) {
    if ($null -eq $Now) { $Now = Get-Date }
    $dir = Split-Path -Parent $SelectionPath
    $stem = [System.IO.Path]::GetFileNameWithoutExtension($SelectionPath)
    $day = ([datetime]$Now).ToString("yyyy-MM-dd", [System.Globalization.CultureInfo]::InvariantCulture)
    $path = Join-Path $dir ("{0}.migrated_{1}.txt" -f $stem, $day)
    $n = 2
    while (Test-Path -LiteralPath $path) { $path = Join-Path $dir ("{0}.migrated_{1}_{2}.txt" -f $stem, $day, $n); $n++ }
    return $path
}

function Get-ProjectNameKeySet($Names) {
    $set = @{}
    foreach ($n in @($Names)) { $t = ([string]$n).Trim(); if ($t) { $set[(Get-ProjectSpecKey $t)] = $true } }
    return $set
}

# Перенос одного проекта: выбор_<проект>.txt -> [inproject] распределение_<проект>.txt.
#   [inproject] нет или пуст, в файле выбора есть имена -> имена переносятся в [inproject] (Action = imported);
#   оба непусты и различаются -> остаётся [inproject] (его и читали скрипты Revit), разница — в журнал (Action = kept);
#   совпадают или файл выбора пуст -> только переименование (Action = renamed).
# Файл выбора всегда переименовывается (не удаляется). Нет файла выбора или нет настройки проекта -> ничего (Action = none).
# EnsureBackup — scriptblock, возвращает папку резервной копии; вызывается ДО первой записи. Возвращает @{ Action; Project; Renamed; Log; Note }
function Invoke-ProjectSelectionMigration([string]$ConfigPath, [scriptblock]$EnsureBackup, $Now = $null) {
    $res = @{ Action = "none"; Project = ""; Renamed = ""; Log = (New-Object System.Collections.Generic.List[string]); Note = "" }
    if (-not $ConfigPath) { return $res }
    $res.Project = ([System.IO.Path]::GetFileNameWithoutExtension($ConfigPath) -replace '^распределение_', '')
    $sel = Get-ProjectSelectionPath $ConfigPath
    if (-not $sel -or -not (Test-Path -LiteralPath $sel -PathType Leaf)) { return $res }
    if (-not (Test-Path -LiteralPath $ConfigPath -PathType Leaf)) {
        $res.Log.Add(("{0}: есть {1}, но нет {2} — файл выбора оставлен как есть" -f $res.Project, (Split-Path -Leaf $sel), (Split-Path -Leaf $ConfigPath)))
        return $res
    }
    $cfg = Read-ProjectConfig $ConfigPath
    if ($null -eq $cfg) { return $res }
    $names = @(Read-ProjectSelection $sel)
    $cur = Get-ProjectInProjectNames $cfg
    $curCount = 0
    if ($null -ne $cur) { $curCount = @($cur).Count }
    $selSet = Get-ProjectNameKeySet $names
    $curSet = Get-ProjectNameKeySet $cur
    $action = "renamed"
    if ($selSet.Count -gt 0 -and $curCount -eq 0) { $action = "imported" }
    elseif ($selSet.Count -gt 0) {
        $onlySel = @($names | Where-Object { $_ -and -not $curSet.ContainsKey((Get-ProjectSpecKey $_)) })
        $onlyCur = @($cur | Where-Object { $_ -and -not $selSet.ContainsKey((Get-ProjectSpecKey $_)) })
        if ($onlySel.Count -gt 0 -or $onlyCur.Count -gt 0) {
            $action = "kept"
            $res.Log.Add(("{0}: [inproject] и {1} различаются, оставлен [inproject] ({2} спек). Только в файле выбора ({3}): {4}. Только в [inproject] ({5}): {6}." -f $res.Project, (Split-Path -Leaf $sel), $curCount, $onlySel.Count, ($onlySel -join "; "), $onlyCur.Count, ($onlyCur -join "; ")))
        }
    }
    $backup = & $EnsureBackup
    if (-not $backup) { throw "нет резервной копии — перенос не выполнен" }
    if ($action -eq "imported") {
        $r = Set-ProjectSelectionFromNames $cfg $names
        Write-ProjectConfig $cfg $ConfigPath
        $res.Log.Add(("{0}: [inproject] был пуст — перенесено из {1}: {2} спек (из них нет в общем файле: {3})." -f $res.Project, (Split-Path -Leaf $sel), @($names).Count, @($r.Unknown).Count))
        $res.Note = ("Выбор «В проекте» проекта {0} перенесён в настройку проекта (технический шаг, один раз)." -f $res.Project)
    }
    elseif ($action -eq "renamed") {
        $res.Log.Add(("{0}: {1} совпадает с [inproject] (или пуст) — только переименован." -f $res.Project, (Split-Path -Leaf $sel)))
    }
    $to = Get-ProjectSelectionMigratedPath $sel $Now
    [System.IO.File]::Move($sel, $to)
    $res.Renamed = $to
    $res.Log.Add(("{0}: {1} -> {2}" -f $res.Project, (Split-Path -Leaf $sel), (Split-Path -Leaf $to)))
    $res.Action = $action
    return $res
}

# Перенос всех проектов при открытии окна (повторный запуск ничего не делает: файлов выбор_<проект>.txt уже нет).
# Перед первой записью — одна резервная копия (New-ProjectStateBackup), в неё же журнал переноса «перенос_выбора.txt».
# Возвращает @{ Backup; Results; Notes; Errors }
function Invoke-ProjectStateMigration([string]$ReportDir, $Now = $null) {
    $out = @{ Backup = ""; Results = (New-Object System.Collections.Generic.List[object]); Notes = (New-Object System.Collections.Generic.List[string]); Errors = (New-Object System.Collections.Generic.List[string]) }
    $projDir = Get-ProjectConfigDir $ReportDir
    $sels = @()
    if (Test-Path -LiteralPath $projDir) {
        $sels = @(Get-ChildItem -LiteralPath $projDir -Filter "выбор_*.txt" -File -ErrorAction SilentlyContinue | Where-Object { $_.Name -notmatch '\.migrated_' -and $_.Name -match '^выбор_(.+)\.txt$' } | Sort-Object Name)
    }
    # фаза 7: прежние файлы состояния «Выборочно» (<история>_буфер/_проверки/_каталог.txt) -> <история>_состояние.txt
    $histDir = Join-Path $ReportDir "история"
    $hasOldState = $false
    if ((Get-Command Invoke-PartialStateMigrationAll -ErrorAction SilentlyContinue) -and (Test-Path -LiteralPath $histDir)) {
        $hasOldState = @(Get-ChildItem -LiteralPath $histDir -File -ErrorAction SilentlyContinue | Where-Object { $_.Name -notmatch '\.migrated_' -and $_.Name -match '^.+_(буфер|проверки|каталог)\.txt$' }).Count -gt 0
    }
    if ($sels.Count -eq 0 -and -not $hasOldState) { return $out }
    $state = @{ Dir = "" }
    $backupRoot = $ReportDir; $backupNow = $Now
    $ensure = { if (-not $state.Dir) { $state.Dir = New-ProjectStateBackup $backupRoot $backupNow }; return $state.Dir }      # одна копия на все проекты, до первой записи
    $log = New-Object System.Collections.Generic.List[string]
    foreach ($f in $sels) {
        $name = ($f.Name -replace '^выбор_', '') -replace '\.txt$', ''
        $cfgPath = Join-Path $projDir ("распределение_" + $name + ".txt")
        try {
            $r = Invoke-ProjectSelectionMigration $cfgPath $ensure $Now
            $out.Results.Add($r)
            foreach ($l in $r.Log) { $log.Add($l) }
            if ($r.Note) { $out.Notes.Add($r.Note) }
        }
        catch { $out.Errors.Add(("{0}: {1}" -f $name, $_.Exception.Message)); $log.Add(("{0}: ОШИБКА {1} — файл выбора оставлен как есть" -f $name, $_.Exception.Message)) }
    }
    $stateLog = New-Object System.Collections.Generic.List[string]
    if ($hasOldState) {
        $ps = Invoke-PartialStateMigrationAll $histDir $ensure $Now
        foreach ($l in $ps.Log) { $stateLog.Add($l) }
        foreach ($e in $ps.Errors) { $out.Errors.Add($e) }
        $merged = @($ps.Results | Where-Object { $_.Action -eq "merged" }).Count
        if ($merged -gt 0) { $out.Notes.Add(("Файлы «Выборочно» объединены: буфер, проверки и каталог теперь в одном файле <история>_состояние.txt ({0} шт.; технический шаг, один раз)." -f $merged)) }
    }
    $out.Backup = $state.Dir
    $stamp = $(if ($null -eq $Now) { Get-Date } else { [datetime]$Now }).ToString("yyyy-MM-dd HH:mm:ss")
    if ($log.Count -gt 0 -and $state.Dir) {                 # журнал — только в папку резервной копии (без переноса ничего не пишется)
        $lines = @(("# Перенос «В проекте» из выбор_<проект>.txt в [inproject], {0}. Резервная копия: {1}" -f $stamp, $state.Dir)) + @($log)
        try { [System.IO.File]::AppendAllText((Join-Path $state.Dir "перенос_выбора.txt"), (($lines -join "`r`n") + "`r`n"), (New-Object System.Text.UTF8Encoding $true)) } catch { }
    }
    if ($stateLog.Count -gt 0 -and $state.Dir) {
        $lines = @(("# Объединение файлов состояния «Выборочно» в <история>_состояние.txt, {0}. Резервная копия: {1}" -f $stamp, $state.Dir)) + @($stateLog)
        try { [System.IO.File]::AppendAllText((Join-Path $state.Dir "перенос_состояния.txt"), (($lines -join "`r`n") + "`r`n"), (New-Object System.Text.UTF8Encoding $true)) } catch { }
    }
    return $out
}

# ---------- путь классификатора проекта (кнопка «Классификатор») ----------
# Хранится в настройке проекта: строка classifier= в [project]. Пусто — значение по умолчанию: для PROJ — путь-заглушка
# (тот же, что CLASSIFIER_PATH в set_assembly_code.py). Для других проектов по умолчанию пусто.
$script:PcDefaultClassifiers = @{
    "PROJ" = "\\server\share\BIM\classifier.txt"
}
function Get-ProjectClassifierPath($Cfg, [string]$ProjectName) {
    if ($null -ne $Cfg -and ([string]$Cfg.Classifier).Trim()) { return ([string]$Cfg.Classifier).Trim() }
    $k = ([string]$ProjectName).Trim().ToUpperInvariant()
    if ($k -and $script:PcDefaultClassifiers.ContainsKey($k)) { return [string]$script:PcDefaultClassifiers[$k] }
    return ""
}

# «В проекте» ровно по списку имён (файл выбора или содержимое сводной): совпавшие отмечаются, остальные снимаются.
# Возвращает @{ Marked = число совпавших; Unknown = имена, которых нет в общем файле }. Таблицу и список спецификаций не меняет.
function Set-ProjectSelectionFromNames($Cfg, $Names) {
    $byKey = @{}
    foreach ($sp in $Cfg.Specs) { $byKey[(Get-ProjectSpecKey ([string]$sp))] = [string]$sp }
    $want = New-Object System.Collections.Generic.List[string]
    $unknown = New-Object System.Collections.Generic.List[string]
    $seen = @{}
    foreach ($n in @($Names)) {
        $t = (Remove-SpecStagePrefix ([string]$n)).Trim()
        if (-not $t) { continue }
        $k = Get-ProjectSpecKey $t
        if ($seen.ContainsKey($k)) { continue }
        $seen[$k] = $true
        if ($byKey.ContainsKey($k)) { $want.Add($byKey[$k]) } else { $unknown.Add($t) }
    }
    Initialize-ProjectInProject $Cfg
    $off = @($Cfg.Specs | Where-Object { -not $seen.ContainsKey((Get-ProjectSpecKey ([string]$_))) })
    Set-ProjectInProjectBulk $Cfg $off $false
    $new = @($want | Where-Object { -not $Cfg.InProject.ContainsKey((Get-ProjectSpecKey $_)) })
    Set-ProjectInProjectBulk $Cfg $new $true
    $Cfg.SelectionUnknown = New-Object System.Collections.Generic.List[string]
    foreach ($u in $unknown) { $Cfg.SelectionUnknown.Add($u) }
    return @{ Marked = $want.Count; Unknown = $unknown.ToArray() }
}

# ---------- справочник «раздел -> номера / категории» и распределение по категориям ----------
# Файл отчеты\проект\справочник_разделов.txt — один на все проекты, рядом с настройками проектов; в окне его нет, правится текстом.
# Формат (UTF-8, «#» — комментарий): блок [РАЗДЕЛ] (ОВ, ПД, ВК, ПТ, ОТ, ХС, ИТП, ЭОМ, СС; можно латиницей) и строки «ключ = значения»:
#   номера     = 0, 1, 16.0, 16.1-18.1     номера спецификаций (ведущий номер имени), через «,» или «;», допускаются диапазоны «16.1-18.1»
#   категории  = Воздуховоды; Оборудование  категории Revit из общего файла, через «;»
#   кроме      = 15.1                       номера, исключаемые из раздела (сильнее двух предыдущих строк)
# Спецификация входит в раздел, если её номер указан в «номера» ИЛИ её категория — в «категории», И её номер не в «кроме».
function Get-SectionCatalogPath([string]$ReportDir) { return (Join-Path (Get-ProjectConfigDir $ReportDir) "справочник_разделов.txt") }

function Get-DefaultSectionCatalogText {
    $common = "0, 1, 16.0, 17.0, 16.1, 17.1, 18.1"
    $duct = "Воздуховоды; Соединительные детали воздуховодов"
    $pipe = "Трубы; Соединительные детали трубопроводов; Оборудование"
    $tray = "Кабельные лотки; Соединительные детали кабельных лотков; Короба"
    $L = @(
        "# Справочник «раздел -> что входит в раздел». Один на все проекты. Правится текстом, в окне его нет.",
        "# номера    — номера спецификаций (ведущий номер имени): 0, 1, 16.0 или диапазон 16.1-18.1",
        "# категории — категории Revit спецификаций (читаются из общего файла), через «;»",
        "# кроме     — номера, которые в раздел не входят",
        "# Спецификация входит в раздел: (её номер в «номера» ИЛИ её категория в «категории») И её номер не в «кроме».",
        "# Общие: 0, 1, 16.0, 17.0. 16.1, 17.1, 18.1 — Корпус/Секция/Этаж для оборудования. 13 и 99 входят через категории (трубы / воздуховоды).",
        "# Спецификации без категории, которых нет в «номера», и категории, которых нет ни в одном разделе, в окне подсвечиваются и не отмечаются.",
        "",
        "[ОВ]", "номера = $common", "категории = $duct; Оборудование", "кроме = 15.1", "",
        "[ПД]", "номера = 0, 1, 16.0, 17.0", "категории = $duct", "",
        "[ВК]", "номера = $common", "категории = $pipe", "кроме = 15.1", "",
        "[ПТ]", "номера = $common", "категории = $pipe", "кроме = 15.1", "",
        "[ОТ]", "номера = $common", "категории = $pipe", "",
        "[ХС]", "номера = $common", "категории = $pipe", "",
        "[ИТП]", "номера = $common", "категории = $pipe", "",
        "[ЭОМ]", "номера = $common", "категории = $tray", "",
        "[СС]", "номера = $common", "категории = $tray", "",
        "[АВТ]", "номера = $common", "категории = Кабельные лотки; Соединительные детали кабельных лотков", "",
        "[ХЦ]", "номера = $common", "категории = Трубы; Соединительные детали трубопроводов; Арматура трубопроводов; Оборудование", ""
    )
    return ($L -join "`r`n")
}

# Нет файла — создаётся со значениями по умолчанию (существующий файл не перезаписывается)
# Блоки разделов, добавленных позже (АВТ, ХЦ): в существующий файл дописываются один раз (файл не перезаписывается, правки пользователя сохраняются).
# Нет блока [АВТ]/[AVT] или [ХЦ]/[HC] и нет отметки «добавлены разделы» — дописывается недостающий; отметка не даёт вернуть блок, который пользователь удалил
$script:SectionCatalogAddedMarker = "# добавлены разделы АВТ и ХЦ (дописано один раз)"
function Add-MissingSectionBlocks([string]$Path) {
    $text = [System.IO.File]::ReadAllText($Path, (New-Object System.Text.UTF8Encoding $true))
    if ($text.Contains($script:SectionCatalogAddedMarker)) { return $false }
    $common = "0, 1, 16.0, 17.0, 16.1, 17.1, 18.1"
    $add = New-Object System.Collections.Generic.List[string]
    if ($text -notmatch '(?mi)^\s*\[(АВТ|AVT)\]') { $add += @("[АВТ]", "номера = $common", "категории = Кабельные лотки; Соединительные детали кабельных лотков", "") }
    if ($text -notmatch '(?mi)^\s*\[(ХЦ|HC)\]') { $add += @("[ХЦ]", "номера = $common", "категории = Трубы; Соединительные детали трубопроводов; Арматура трубопроводов; Оборудование", "") }
    if ($add.Count -eq 0) { return $false }
    $tail = $(if ($text.EndsWith("`n")) { "" } else { "`r`n" })
    [System.IO.File]::AppendAllText($Path, ($tail + "`r`n" + $script:SectionCatalogAddedMarker + "`r`n" + ($add -join "`r`n") + "`r`n"), (New-Object System.Text.UTF8Encoding $true))
    return $true
}

function Initialize-SectionCatalogFile([string]$Path) {
    if (Test-Path -LiteralPath $Path) { [void](Add-MissingSectionBlocks $Path); return }
    $dir = Split-Path -Parent $Path
    if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    [System.IO.File]::WriteAllText($Path, (Get-DefaultSectionCatalogText) + "`r`n", (New-Object System.Text.UTF8Encoding $true))
}

# Номер как массив чисел для сравнения (16.1 < 17.0 < 18.1); не номер — $null
function ConvertTo-SpecNumber([string]$Text) {
    $t = ([string]$Text).Trim()
    if ($t -notmatch '^\d+(?:\.\d+)*$') { return $null }
    return @($t.Split(".") | ForEach-Object { [int]$_ })
}
function Compare-SpecNumber($A, $B) {
    $n = [Math]::Max(@($A).Count, @($B).Count)
    for ($i = 0; $i -lt $n; $i++) {
        $x = 0; $y = 0
        if ($i -lt @($A).Count) { $x = [int]@($A)[$i] }
        if ($i -lt @($B).Count) { $y = [int]@($B)[$i] }
        if ($x -lt $y) { return -1 }
        if ($x -gt $y) { return 1 }
    }
    return 0
}
function Get-SpecNumberText([string]$Spec) {
    $t = Remove-SpecStagePrefix $Spec
    if ($t -match '^(\d+(?:\.\d+)*)') { return $Matches[1] }
    return ""
}

# Разбор списка номеров «0, 1, 16.1-18.1» -> @(@{ From; To }); ошибка -> исключение с текстом
function ConvertTo-NumberRanges([string]$Text) {
    $out = New-Object System.Collections.Generic.List[object]
    foreach ($tok in @(([string]$Text) -split '[,;]')) {
        $t = $tok.Trim()
        if (-not $t) { continue }
        $parts = @($t -split '\s*[-–—]\s*')
        if ($parts.Count -eq 1) { $a = ConvertTo-SpecNumber $parts[0]; $b = $a }
        elseif ($parts.Count -eq 2) { $a = ConvertTo-SpecNumber $parts[0]; $b = ConvertTo-SpecNumber $parts[1] }
        else { $a = $null; $b = $null }
        if ($null -eq $a -or $null -eq $b) { throw ("«{0}» — не номер спецификации (ожидается 16.1 или диапазон 16.1-18.1)" -f $t) }
        $out.Add([pscustomobject]@{ From = $a; To = $b })
    }
    return $out.ToArray()
}

function Get-SectionCodeByTitle([string]$Text) {
    $u = ([string]$Text).Trim().ToUpperInvariant()
    foreach ($s in $script:PcSections) { foreach ($n in $s.Names) { if ($n.ToUpperInvariant() -eq $u) { return [string]$s.Code } } }
    return ""
}

# Чтение справочника. Результат: @{ Ok; Error; Path; Sections = @{ код = @{ Numbers; Categories; Except } } }. Ошибки файла не бросают исключений — Ok = $false и понятный текст
function Read-SectionCatalog([string]$Path) {
    $res = @{ Ok = $false; Error = ""; Path = $Path; Sections = @{} }
    try {
        if (-not (Test-Path -LiteralPath $Path)) { $res.Error = ("Нет файла справочника разделов: {0}" -f $Path); return $res }
        $text = [System.IO.File]::ReadAllText($Path, (New-Object System.Text.UTF8Encoding $true))
    }
    catch { $res.Error = ("Справочник разделов не прочитан ({0}): {1}" -f $Path, $_.Exception.Message); return $res }
    $code = ""; $n = 0
    foreach ($raw in @($text -split "`r?`n")) {
        $n++
        $line = $raw.Trim()
        if (-not $line -or $line.StartsWith("#")) { continue }
        try {
            if ($line -match '^\[(.*)\]$') {
                $code = Get-SectionCodeByTitle $Matches[1]
                if (-not $code) { throw ("раздел «{0}» не распознан (допустимо: {1})" -f $Matches[1].Trim(), (($script:PcSections | ForEach-Object { $_.Title }) -join ", ")) }
                if ($res.Sections.ContainsKey($code)) { throw ("раздел «{0}» указан дважды" -f $Matches[1].Trim()) }
                $res.Sections[$code] = @{ Numbers = @(); Categories = @(); Except = @() }
                continue
            }
            if (-not $code) { throw "строка вне блока [РАЗДЕЛ]" }
            if ($line -notmatch '^([^=]+?)\s*=\s*(.*)$') { throw ("нет знака «=» («{0}»)" -f $line) }
            $key = $Matches[1].Trim().ToLowerInvariant(); $val = $Matches[2]
            switch ($key) {
                "номера" { $res.Sections[$code].Numbers = @(ConvertTo-NumberRanges $val) }
                "кроме" { $res.Sections[$code].Except = @(ConvertTo-NumberRanges $val) }
                "категории" { $res.Sections[$code].Categories = @($val -split ';' | ForEach-Object { $_.Trim() } | Where-Object { $_ }) }
                default { throw ("неизвестный ключ «{0}» (допустимо: номера, категории, кроме)" -f $Matches[1].Trim()) }
            }
        }
        catch { $res.Error = ("Справочник разделов, строка {0}: {1}. Файл: {2}" -f $n, $_.Exception.Message, $Path); $res.Sections = @{}; return $res }
    }
    if ($res.Sections.Count -eq 0) { $res.Error = ("Справочник разделов пуст: {0}" -f $Path); return $res }
    $res.Ok = $true
    return $res
}

function Test-NumberInRanges([string]$NumberText, $Ranges) {
    $v = ConvertTo-SpecNumber $NumberText
    if ($null -eq $v) { return $false }
    foreach ($r in @($Ranges)) { if ((Compare-SpecNumber $v $r.From) -ge 0 -and (Compare-SpecNumber $v $r.To) -le 0) { return $true } }
    return $false
}

# Категория спецификации из кэша настройки (пусто — не прочитана / у спецификации нет категории)
function Get-ProjectSpecCategory($Cfg, [string]$Spec) {
    $k = Get-ProjectSpecKey $Spec
    if ($Cfg.Categories.ContainsKey($k)) { return [string]$Cfg.Categories[$k] }
    return ""
}

# Входит ли спецификация в раздел по справочнику
function Test-SectionTakesSpec($Catalog, [string]$Code, [string]$Number, [string]$Category) {
    if (-not $Catalog -or -not $Catalog.Ok -or -not $Catalog.Sections.ContainsKey($Code)) { return $false }
    $sec = $Catalog.Sections[$Code]
    if ($Number -and (Test-NumberInRanges $Number $sec.Except)) { return $false }
    if ($Number -and (Test-NumberInRanges $Number $sec.Numbers)) { return $true }
    if ($Category) { foreach ($c in $sec.Categories) { if ([string]::Equals($c, $Category, [System.StringComparison]::OrdinalIgnoreCase)) { return $true } } }
    return $false
}

# Проблема спецификации для подсветки: "" — нормально; "nocat" — нет категории и номера нет в справочнике; "nosection" — категория не входит ни в один раздел справочника
function Get-ProjectSpecIssue($Cfg, $Catalog, [string]$Spec) {
    if (-not $Catalog -or -not $Catalog.Ok) { return "" }
    $num = Get-SpecNumberText $Spec
    $cat = Get-ProjectSpecCategory $Cfg $Spec
    foreach ($code in @($Catalog.Sections.Keys)) {
        if (Test-SectionTakesSpec $Catalog $code $num $cat) { return "" }
    }
    if (-not $cat) { return "nocat" }
    return "nosection"
}

$script:PcCatalog = $null      # последний прочитанный справочник (для подсветки таблицы)
function Update-SectionCatalog([string]$Path) {
    Initialize-SectionCatalogFile $Path
    $script:PcCatalog = Read-SectionCatalog $Path
    return $script:PcCatalog
}

# «Распределить»: модели не открываются. Галочка (модель × спека) = спека «В проекте» И (номер или категория) входит в раздел модели по справочнику.
# Ручные правки («m») не меняются; число элементов сохраняется. Модель с разделом, которого нет в справочнике, и спеки без категории не отмечаются — их список в результате.
# Возвращает @{ Ticked; Models; UnknownSections = коды; NoCategory = спеки; NoSection = спеки }
function Set-ProjectDistribution($Cfg, $Catalog) {
    $res = @{ Ticked = 0; Models = 0; UnknownSections = @(); NoCategory = @(); NoSection = @() }
    if (-not $Catalog -or -not $Catalog.Ok) { throw "Справочник разделов недоступен" }
    $unknown = New-Object System.Collections.Generic.List[string]
    foreach ($s in $Cfg.Specs) {
        if (-not (Test-ProjectSpecIn $Cfg $s)) { continue }
        $issue = Get-ProjectSpecIssue $Cfg $Catalog $s
        if ($issue -eq "nocat") { $res.NoCategory += [string]$s } elseif ($issue -eq "nosection") { $res.NoSection += [string]$s }
    }
    foreach ($m in $Cfg.Models) {
        $code = [string]$m.Section
        if (-not $Catalog.Sections.ContainsKey($code)) { if (-not $unknown.Contains($code)) { $unknown.Add($code) }; continue }
        $res.Models++
        foreach ($s in $Cfg.Specs) {
            $old = Get-ProjectCell $Cfg $m.Name $s
            if ($old -and $old.Origin -eq "m") { if ($old.State -eq 1) { $res.Ticked++ }; continue }
            $on = 0
            if ((Test-ProjectSpecIn $Cfg $s) -and (Test-SectionTakesSpec $Catalog $code (Get-SpecNumberText $s) (Get-ProjectSpecCategory $Cfg $s))) { $on = 1 }
            $count = -1
            if ($old) { $count = $old.Count }
            Set-ProjectCell $Cfg $m.Name $s $on "a" $count
            $res.Ticked += $on
        }
    }
    $res.UnknownSections = $unknown.ToArray()
    return $res
}

# ---------- запросы потребителей ----------
# Включённые (отмеченные) спецификации модели, в порядке шаблона; только спецификации групп «В проекте»
function Get-ProjectEnabledSpecs($Cfg, [string]$Model) {
    $out = New-Object System.Collections.Generic.List[string]
    foreach ($s in $Cfg.Specs) {
        if (-not (Test-ProjectSpecIn $Cfg $s)) { continue }
        $c = Get-ProjectCell $Cfg $Model $s
        if ($c -and $c.State -eq 1) { $out.Add([string]$s) }
    }
    return $out.ToArray()
}

# Спецификации раздела = объединение по его моделям
function Get-ProjectSectionSpecs($Cfg, [string]$Code) {
    $models = @(Get-ProjectSectionModels $Cfg $Code)
    $out = New-Object System.Collections.Generic.List[string]
    foreach ($s in $Cfg.Specs) {
        if (-not (Test-ProjectSpecIn $Cfg $s)) { continue }
        foreach ($m in $models) { $c = Get-ProjectCell $Cfg $m $s; if ($c -and $c.State -eq 1) { $out.Add([string]$s); break } }
    }
    return $out.ToArray()
}

# Спецификации сводной: спецификации проекта, отмеченные в столбце «Сводная» (нет раздела [summary] — все спецификации проекта;
# нет и отметок «В проекте» — объединение отмеченных ячеек, как в файлах прежних версий)
function Get-ProjectSummarySpecs($Cfg) {
    $out = New-Object System.Collections.Generic.List[string]
    foreach ($s in $Cfg.Specs) {
        if ($null -ne $Cfg.InProject) {
            if ((Test-ProjectSpecIn $Cfg $s) -and (Test-ProjectSpecInSummary $Cfg $s)) { $out.Add([string]$s) }
            continue
        }
        foreach ($m in $Cfg.Models) { $c = Get-ProjectCell $Cfg $m.Name $s; if ($c -and $c.State -eq 1) { $out.Add([string]$s); break } }
    }
    return $out.ToArray()
}

# Набор спецификаций модели для вставки/проверки/«Выборочно»: @{ Specs = string[] | $null; Why = текст }
#  1) модель есть в настройке — её отметки; 2) иначе раздел по имени файла и у раздела есть модели — объединение раздела; иначе $null + причина
function Get-ProjectModelSpecs($Cfg, [string]$PathOrName) {
    $m = Find-ProjectModel $Cfg $PathOrName
    if ($m) {
        $any = $false
        foreach ($s in $Cfg.Specs) { if (Get-ProjectCell $Cfg $m.Name $s) { $any = $true; break } }
        if ($any) { return @{ Specs = @(Get-ProjectEnabledSpecs $Cfg $m.Name); Why = "" } }
        return @{ Specs = $null; Why = ("модель «{0}» есть в распределении, но спецификации для неё не размечены (вкладка «Спецификации», кнопка «Распределить»)" -f $m.Name) }
    }
    $code = Get-ModelSectionFromName $PathOrName
    if (-not $code) { return @{ Specs = $null; Why = "модели нет в распределении, раздел не определён по имени (ожидается «_ОВ_», «_VK_» и т. п.: ОВ/OV, ПД/PD, ВК/VK, ПТ/PT, ОТ/OT, ХС/HS, ИТП/ITP, ЭОМ/EOM, СС/SS)" } }
    $secSpecs = @(Get-ProjectSectionSpecs $Cfg $code)
    if ($secSpecs.Count -eq 0) { return @{ Specs = $null; Why = ("модели нет в распределении, а у раздела {0} нет размеченных моделей" -f $code) } }
    return @{ Specs = $secSpecs; Why = "" }
}

# ---------- таблица (вид «Распределение») ----------
function Get-ProjectModelShort($Cfg, [string]$Model, [string]$Code) {
    $n = [string]$Model
    if ($Code -and $n -match ('(?i)_' + [regex]::Escape($Code) + '_(?:ORG_)?(.+)$')) { return $Matches[1] }
    return $n
}

# Expand: @{ Sections = @{ код = $true }; Groups = @{ группа = $true }; Hide = $true|$false }.
# Столбцы: «Группа / спецификация», «В проекте», «Сводная»; дальше у каждого раздела — столбец раздела (агрегат), у раскрытого столбцы его моделей;
# между разделами — разделитель (окно его не рисует: границы — тонкая сетка). Строки: группа (у группы из нескольких спеков Expandable — значок раскрытия рисует окно,
# в тексте его нет; Open — раскрыта), у раскрытой — её спецификации (подпись — хвост), у каждой своя галочка «В проекте» и «Сводная».
# Issue у строки/столбца — подсветка: nocat / nosection (спека без категории / категория не входит ни в один раздел справочника), unknown (раздел модели не в справочнике).
# «Скрыть не участвующие» (Hide) убирает группы и спецификации, не отмеченные «В проекте».

# Состояние строки (группа или спецификация) по настройке: In — «В проекте», Active — спецификации проекта в строке, Sum — «Сводная»
function Update-ProjectGridRowState($Cfg, $Row) {
    $Row.In = Get-ProjectGroupInState $Cfg $Row.Specs
    $Row.Active = @(Get-ProjectActiveSpecs $Cfg $Row.Specs)
    $Row.Sum = Get-ProjectSummaryState $Cfg $Row.Specs
    # подсветка по справочнику разделов: только спецификации проекта; у группы — если проблема есть у любой из них
    $Row.Issue = ""; $Row.IssueText = ""
    if ($script:PcCatalog -and $script:PcCatalog.Ok) {
        foreach ($sp in @($Row.Active)) {
            $i = Get-ProjectSpecIssue $Cfg $script:PcCatalog ([string]$sp)
            if ($i) {
                if ($Row.Issue -ne "nocat") { $Row.Issue = $i }
                $Row.IssueText = $(if ($Row.Issue -eq "nocat") { "Нет категории в общем файле, номера нет в справочнике разделов — не отмечается" } else { "Категория не входит ни в один раздел справочника — не отмечается" })
            }
        }
    }
}

function Get-ProjectGridModel($Cfg, $Expand) {
    $cols = New-Object System.Collections.Generic.List[object]
    $cols.Add([pscustomobject]@{ Kind = "name"; Section = ""; Model = ""; Models = @(); Title = "Группа / спецификация"; Tint = -1; Open = $false; Issue = "" })
    $cols.Add([pscustomobject]@{ Kind = "inproject"; Section = ""; Model = ""; Models = @(); Title = "В проекте"; Tint = -1; Open = $false; Issue = "" })
    $cols.Add([pscustomobject]@{ Kind = "summary"; Section = ""; Model = ""; Models = @(); Title = "Сводная"; Tint = -1; Open = $false; Issue = "" })
    $tint = 0
    foreach ($sec in @(Get-ProjectViewSections $Cfg)) {
        $models = @(Get-ProjectSectionModels $Cfg $sec.Code)
        $open = $Expand.Sections.ContainsKey($sec.Code) -and $Expand.Sections[$sec.Code]
        $secIssue = ""
        if ($script:PcCatalog -and $script:PcCatalog.Ok -and -not $script:PcCatalog.Sections.ContainsKey([string]$sec.Code)) { $secIssue = "unknown" }
        $cols.Add([pscustomobject]@{ Kind = "sep"; Section = $sec.Code; Model = ""; Models = @(); Title = ""; Tint = -1; Open = $false; Issue = "" })
        $cols.Add([pscustomobject]@{ Kind = "section"; Section = $sec.Code; Model = ""; Models = $models; Title = [string]$sec.Title; Tint = $tint; Open = [bool]$open; Issue = $secIssue })
        if ($open) {
            foreach ($m in $models) { $cols.Add([pscustomobject]@{ Kind = "model"; Section = $sec.Code; Model = $m; Models = @($m); Title = (Get-ProjectModelShort $Cfg $m $sec.Code); Tint = $tint; Open = $false; Issue = $secIssue }) }
        }
        $tint++
    }
    $hide = $Expand.ContainsKey("Hide") -and $Expand.Hide
    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($g in @(Get-ProjectGroups $Cfg)) {
        $grow = [pscustomobject]@{ Kind = "group"; Group = $g.Name; Spec = ""; Specs = @($g.Specs); Title = [string]$g.Name; Expandable = (@($g.Specs).Count -gt 1); Open = $false; In = ""; Active = @(); Sum = ""; Issue = ""; IssueText = "" }
        Update-ProjectGridRowState $Cfg $grow
        if ($hide -and $grow.In -eq "Unchecked") { continue }
        $open = $grow.Expandable -and $Expand.Groups.ContainsKey($g.Name) -and $Expand.Groups[$g.Name]
        $grow.Open = [bool]$open
        $rows.Add($grow)
        if ($open) {
            foreach ($s in $g.Specs) {
                $srow = [pscustomobject]@{ Kind = "spec"; Group = $g.Name; Spec = [string]$s; Specs = @([string]$s); Title = (Get-SpecRowLabel $s $g.Name); Expandable = $false; Open = $false; In = ""; Active = @(); Sum = ""; Issue = ""; IssueText = "" }
                Update-ProjectGridRowState $Cfg $srow
                if ($hide -and $srow.In -eq "Unchecked") { continue }
                $rows.Add($srow)
            }
        }
    }
    return [pscustomobject]@{ Columns = $cols.ToArray(); Rows = $rows.ToArray() }
}

function Get-ProjectGlyph([string]$State) {
    switch ($State) { "Checked" { return [string][char]0x2611 } "Indeterminate" { return [string][char]0x25EA } default { return [string][char]0x2610 } }
}

function Get-ProjectGridCell($Cfg, $Row, $Col) {
    if ($Col.Kind -eq "inproject") { return [pscustomobject]@{ Text = (Get-ProjectGlyph $Row.In); Color = "none"; State = $Row.In; Aggregate = $null } }
    $none = [pscustomobject]@{ Text = ""; Color = "none"; State = "Unchecked"; Aggregate = $null }
    if (@($Row.Active).Count -eq 0) { return $none }                 # вне проекта — не идёт ни в сводную, ни в модели
    if ($Col.Kind -eq "summary") { return [pscustomobject]@{ Text = (Get-ProjectGlyph $Row.Sum); Color = "none"; State = $Row.Sum; Aggregate = $null } }
    $agg = Get-ProjectAggregate $Cfg $Col.Models $Row.Active
    return [pscustomobject]@{ Text = (Format-ProjectAggregate $agg); Color = (Get-ProjectAggregateColor $agg); State = $agg.State; Aggregate = $agg }
}

# Щелчок по ячейке. «В проекте»: группа — все спецификации, спецификация — только она. «Сводная»: спецификации проекта строки.
# Раздел/модель: «отмечено везде» -> снять у всех, иначе -> отметить у всех (только спецификации проекта); затронутые ячейки получают признак «ручная правка»
function Invoke-ProjectGridClick($Cfg, $Row, $Col) {
    if ($Col.Kind -eq "inproject") {
        if ($Row.Kind -eq "group") { Invoke-ProjectGroupInClick $Cfg $Row.Specs } else { Invoke-ProjectSpecInClick $Cfg ([string]$Row.Spec) }
        return
    }
    if ($Col.Kind -eq "summary") { Invoke-ProjectSummaryClick $Cfg $Row.Specs; return }
    if ($Col.Kind -ne "section" -and $Col.Kind -ne "model") { return }
    $specs = @(Get-ProjectActiveSpecs $Cfg $Row.Specs)
    if ($specs.Count -eq 0) { return }                       # распределяются только спецификации проекта
    $agg = Get-ProjectAggregate $Cfg $Col.Models $specs
    $new = 1
    if ($agg.State -eq "Checked") { $new = 0 }
    foreach ($m in @($Col.Models)) {
        foreach ($s in $specs) {
            $c = Get-ProjectCell $Cfg ([string]$m) ([string]$s)
            $count = -1
            if ($c) { $count = $c.Count }
            if ($c -and $c.State -eq $new -and $c.Origin -ne "x") { continue }      # уже так — признак не трогаем
            Set-ProjectCell $Cfg ([string]$m) ([string]$s) $new "m" $count
        }
    }
}

# ---------- чтение / запись файла ----------
function Read-ProjectConfig([string]$Path) {
    $cfg = New-ProjectConfig
    $cfg.InProject = $null              # нет раздела [inproject] — в проекте все спецификации (файлы прежних версий)
    $cfg.Summary = $null                # нет раздела [summary] — в сводную идут все спецификации проекта
    if (-not (Test-Path -LiteralPath $Path)) { return $null }
    $text = [System.IO.File]::ReadAllText($Path, (New-Object System.Text.UTF8Encoding $true))
    $section = ""
    $inRaw = New-Object System.Collections.Generic.List[string]      # имена [inproject] как в файле: те, которых нет в [specs], — в SelectionUnknown
    foreach ($raw in @($text -split "`r?`n")) {
        $line = $raw.TrimEnd()
        if (-not $line.Trim() -or $line.TrimStart().StartsWith("#")) { continue }
        if ($line -match '^\[(.+)\]$') {
            $section = $Matches[1]
            if ($section -eq "asg_models" -and $null -eq $cfg.AsgModels) { $cfg.AsgModels = @{} }
            if ($section -eq "asg_pairs" -and $null -eq $cfg.AsgPairs) { $cfg.AsgPairs = @{} }
            # раздел есть, но пуст — «ничего не отмечено» (так его читает и insert_schedules.py); раздела нет — прежнее «все» ($null)
            if ($section -eq "inproject" -and $null -eq $cfg.InProject) { $cfg.InProject = @{} }
            if ($section -eq "summary" -and $null -eq $cfg.Summary) { $cfg.Summary = @{} }
            continue
        }
        switch ($section) {
            "project" {
                if ($line -match '^([A-Za-z_]+)=(.*)$') {
                    $v = $Matches[2].Trim()
                    switch ($Matches[1].ToLowerInvariant()) {
                        "stage" { $cfg.Stage = $v } "template" { $cfg.Template = $v } "models_dir" { $cfg.ModelsDir = $v }
                        "summary_model" { $cfg.SummaryModel = $v } "google_link" { $cfg.GoogleLink = $v } "summary_done" { $cfg.SummaryDone = ($v -eq "1") }
                        "insert_mode" { $cfg.InsertMode = $v } "asg_scope" { $cfg.AsgScope = $v } "classifier" { $cfg.Classifier = $v }
                    }
                }
            }
            "asg_models" { if ($null -eq $cfg.AsgModels) { $cfg.AsgModels = @{} }; $cfg.AsgModels[$line.Trim()] = $true }
            "asg_pairs" { if ($null -eq $cfg.AsgPairs) { $cfg.AsgPairs = @{} }; $cfg.AsgPairs[$line.Trim()] = $true }
            "models_rsn" { if ($line.Trim() -match '^(?i)RSN://') { $cfg.ModelsRsn.Add($line.Trim()) } }
            "sections" { }          # справочно: разделы берутся из имён моделей
            "models" {
                $pm = $line.Split("`t"); $nm = $pm[0].Trim()
                if (-not $nm) { continue }
                $stored = ""; if ($pm.Count -ge 2) { $stored = $pm[1].Trim() }
                $sec = Get-ModelSectionFromName $nm
                if (-not $sec -and (Test-TargetSectionCode $stored)) { $sec = $stored }      # раздел по папке — как записан, если это целевой код
                if ($sec -and (Test-TargetSectionCode $sec)) { $cfg.Models.Add([pscustomobject]@{ Name = $nm; Section = $sec }) }
                elseif ($stored) { Add-ProjectHeldModel $cfg $nm $stored }                   # UNKNOWN и чужой код: не показывать, не падать, вернуть в файл до «Сохранить»
            }
            "inproject" { if ($null -eq $cfg.InProject) { $cfg.InProject = @{} }; $cfg.InProject[(Get-ProjectSpecKey $line.Trim())] = $true; $inRaw.Add((Remove-SpecStagePrefix $line.Trim()).Trim()) }
            "summary" { if ($null -eq $cfg.Summary) { $cfg.Summary = @{} }; $cfg.Summary[(Get-ProjectSpecKey $line.Trim())] = $true }
            "specs" { $cfg.Specs.Add($line.Trim()) }
            "groups" { $p = $line.Split("`t"); if ($p.Count -ge 2) { $cfg.GroupOverride[(Get-ProjectSpecKey $p[0])] = $p[1].Trim() } }
            "categories" { $p = $line.Split("`t"); if ($p.Count -ge 2 -and $p[1].Trim()) { $cfg.Categories[(Get-ProjectSpecKey $p[0])] = $p[1].Trim() } }
            "cells" {
                $p = $line.Split("`t")
                if ($p.Count -ge 4) {
                    if (Test-ProjectHeldModel $cfg $p[0].Trim()) {
                        if (-not $cfg.HeldCells) { $cfg.HeldCells = New-Object System.Collections.Generic.List[string] }
                        $cfg.HeldCells.Add($line)
                        continue
                    }
                    $n = -1
                    if ($p.Count -ge 5 -and $p[4].Trim()) { [void][int]::TryParse($p[4].Trim(), [ref]$n) }
                    $st = 0; [void][int]::TryParse($p[2].Trim(), [ref]$st)
                    Set-ProjectCell $cfg $p[0].Trim() $p[1].Trim() $st $p[3].Trim() $n
                }
            }
        }
    }
    if ($null -ne $cfg.InProject -and $inRaw.Count -gt 0) {
        $known = @{}
        foreach ($sp in $cfg.Specs) { $known[(Get-ProjectSpecKey ([string]$sp))] = $true }
        $seenU = @{}
        foreach ($nm in $inRaw) {
            $k = Get-ProjectSpecKey $nm
            if (-not $nm -or $known.ContainsKey($k) -or $seenU.ContainsKey($k)) { continue }
            $seenU[$k] = $true
            $cfg.InProject.Remove($k)
            $cfg.SelectionUnknown.Add($nm)
        }
    }
    Update-ProjectSections $cfg
    $cfg.GroupsCache = $null
    return $cfg
}

# Пустой раздел [inproject] / [summary] = «ни одна не отмечена»; раздела нет = «все» (файлы прежних версий). Строка-пометка «# пусто…» —
# комментарий: её пропускают и окно, и insert_schedules.py, и прежние версии.
function Write-ProjectConfig($Cfg, [string]$Path) {
    $dir = Split-Path -Parent $Path
    if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $L = New-Object System.Collections.Generic.List[string]
    $L.Add("# Настройка проекта: распределение спецификаций по моделям и разделам (вкладка «Распределение»). Править можно в окне.")
    $L.Add("[project]")
    $L.Add("stage=" + $Cfg.Stage); $L.Add("template=" + $Cfg.Template); $L.Add("models_dir=" + $Cfg.ModelsDir)
    $L.Add("summary_model=" + $Cfg.SummaryModel); $L.Add("google_link=" + $Cfg.GoogleLink)
    $L.Add("summary_done=" + $(if ($Cfg.SummaryDone) { "1" } else { "0" })); $L.Add("insert_mode=" + $Cfg.InsertMode); $L.Add("asg_scope=" + $Cfg.AsgScope)
    if ($Cfg.Classifier) { $L.Add("classifier=" + $Cfg.Classifier) }
    if ($null -ne $Cfg.AsgModels) { $L.Add("[asg_models]"); foreach ($k in @($Cfg.AsgModels.Keys | Sort-Object)) { $L.Add([string]$k) } }
    if ($null -ne $Cfg.AsgPairs) { $L.Add("[asg_pairs]"); foreach ($k in @($Cfg.AsgPairs.Keys | Sort-Object)) { $L.Add([string]$k) } }
    if ($Cfg.ModelsRsn -and $Cfg.ModelsRsn.Count -gt 0) {
        $L.Add("[models_rsn]")
        foreach ($m in $Cfg.ModelsRsn) { $L.Add([string]$m) }
    }
    $L.Add("[sections]")
    foreach ($s in $Cfg.Sections) { $L.Add(("{0}`t{1}" -f $s.Code, $s.Title)) }
    $L.Add("[models]")
    $seenModels = @{}
    foreach ($m in $Cfg.Models) { $seenModels[(Get-ProjectModelKey $m.Name)] = $true; $L.Add(("{0}`t{1}" -f $m.Name, $m.Section)) }
    if ($Cfg.HeldModels) {
        foreach ($m in $Cfg.HeldModels) {
            $hk = Get-ProjectModelKey $m.Name
            if (-not $hk -or $seenModels.ContainsKey($hk)) { continue }
            $seenModels[$hk] = $true
            $L.Add(("{0}`t{1}" -f $m.Name, $m.Section))
        }
    }
    $L.Add("[specs]")
    foreach ($s in $Cfg.Specs) { $L.Add([string]$s) }
    if ($null -ne $Cfg.InProject) {
        $L.Add("[inproject]")
        $inSeen = @{}
        foreach ($s in $Cfg.Specs) { $k = Get-ProjectSpecKey $s; if ($Cfg.InProject.ContainsKey($k)) { $L.Add([string]$s); $inSeen[$k] = $true } }
        # имена не из общего файла (например, из прежнего файла выбора) — как есть, чтобы не терялись при следующем сохранении
        if ($Cfg.SelectionUnknown) { foreach ($u in $Cfg.SelectionUnknown) { $k = Get-ProjectSpecKey ([string]$u); if ($u -and -not $inSeen.ContainsKey($k)) { $L.Add([string]$u); $inSeen[$k] = $true } } }
        if ($inSeen.Count -eq 0) { $L.Add("# пусто: ни одна спецификация не отмечена «В проекте»") }          # пометка для человека; признак «пусто» — сам раздел без строк
    }
    if ($null -ne $Cfg.Summary) {
        $L.Add("[summary]")
        $sumN = 0
        foreach ($s in $Cfg.Specs) { if ($Cfg.Summary.ContainsKey((Get-ProjectSpecKey $s))) { $L.Add([string]$s); $sumN++ } }
        if ($sumN -eq 0) { $L.Add("# пусто: в сводную не отмечена ни одна спецификация") }
    }
    $L.Add("[groups]")
    foreach ($s in $Cfg.Specs) {
        $k = Get-ProjectSpecKey $s
        if ($Cfg.GroupOverride.ContainsKey($k)) { $L.Add(("{0}`t{1}" -f $s, $Cfg.GroupOverride[$k])) }
    }
    if ($Cfg.Categories.Count -gt 0) {
        $L.Add("[categories]")
        foreach ($s in $Cfg.Specs) {
            $k = Get-ProjectSpecKey $s
            if ($Cfg.Categories.ContainsKey($k)) { $L.Add(("{0}`t{1}" -f $s, $Cfg.Categories[$k])) }
        }
    }
    $L.Add("[cells]")
    $specKeys = @{}
    foreach ($s in $Cfg.Specs) { $specKeys[[string]$s] = Get-ProjectSpecKey $s }
    foreach ($m in $Cfg.Models) {
        $mk = Get-ProjectModelKey $m.Name
        foreach ($s in $Cfg.Specs) {
            $c = $Cfg.Cells[$mk + "`t" + $specKeys[[string]$s]]
            if (-not $c) { continue }
            $cnt = ""
            if ($c.Count -ge 0) { $cnt = [string]$c.Count }
            $L.Add(("{0}`t{1}`t{2}`t{3}`t{4}" -f $m.Name, $s, $c.State, $c.Origin, $cnt))
        }
    }
    if ($Cfg.HeldCells) { foreach ($hc in $Cfg.HeldCells) { if ($hc) { $L.Add([string]$hc) } } }
    $tmp = $Path + ".tmp"
    [System.IO.File]::WriteAllLines($tmp, [string[]]$L.ToArray(), (New-Object System.Text.UTF8Encoding $true))
    if (Test-Path -LiteralPath $Path) { [System.IO.File]::Delete($Path) }
    [System.IO.File]::Move($tmp, $Path)
}

# Список спецификаций для вкладок «Сводная» и «Разделы» (замена прежнего список_спецификаций.txt): имена с префиксом стадии проекта
function Get-ProjectSpecDisplayName($Cfg, [string]$Spec) {
    return ($Cfg.Stage + (" " * $script:PcStagePad) + (Remove-SpecStagePrefix $Spec))
}

# Профиль: подстановка {СТАДИЯ} в значение фильтра (то же делает insert_schedules.py)
function Expand-ProjectProfile($Cfg, [string]$Value) {
    $v = [string]$Value
    $v = [regex]::Replace($v, '\{(?:СТАДИЯ|STAGE)\}', [string]$Cfg.Stage, "IgnoreCase")
    return $v
}

# Префикс проекта для «Задания» (имена связей <префикс>_IOS_<РАЗДЕЛ>_…): из имени файла настройки распределение_<префикс>.txt,
# иначе из имени сводной модели (часть до первого «_»); нет — пусто (подходит любой префикс)
function Get-AssignProjectPrefix([string]$ConfigPath, [string]$SummaryPath) {
    $leaf = ([string]$ConfigPath).Replace("\", "/").Split("/")[-1]
    if ($leaf -match '^распределение_(.+)\.txt$') {
        if ($Matches[1] -ne "PROJ") { return $Matches[1] }
    }
    $n = Get-ProjectModelName $SummaryPath
    if ($n -match '^([A-Za-z0-9]+)_') { return $Matches[1] }
    return ""
}

# ---------- план вставки (plan.txt прогона «В сводную» / «В разделы») ----------
# Окно заранее решает, какие спецификации вставлять, и пишет готовый план в служебную папку прогона; insert_schedules.py берёт его
# и файл настройки проекта сам не разбирает (прежний разбор там остался запасным путём на случай, если плана нет или он повреждён).
# Правила — ровно те, по которым insert_schedules.py решал раньше (load_project / project_model_specs / project_summary_specs):
#  * модели — только с кодом раздела в имени файла (модели, у которых раздел известен лишь по папке, ищутся как «нет в распределении»);
#  * модель с ячейками — её отмеченные спецификации «В проекте»; модель без ячеек — «не размечена» (вставки нет);
#  * модели нет — объединение отметок её раздела (sspec); пустое объединение — вставки нет;
#  * сводная: при [inproject] — «В проекте» ∩ «Сводная»; без [inproject] — объединение отмеченных ячеек всех моделей.
# Формат (UTF-8 с BOM, строка = поля через TAB): version 1 | stage | summary <спец> | model <имя> marked|unmarked | mspec <имя> <спец> |
# sspec <код> <спец> | end <число строк данных>. Без строки end с верным числом план не используется.
$script:InsertPlanVersion = "1"
function Get-InsertPlanLines([string]$ProjectPath) {
    $cfg = Read-ProjectConfig $ProjectPath
    if ($null -eq $cfg) { throw ("Нет файла настройки проекта: " + $ProjectPath) }
    $text = [System.IO.File]::ReadAllText($ProjectPath, (New-Object System.Text.UTF8Encoding $true))
    # пустой раздел [inproject]/[summary] — «ничего не отмечено», раздела нет — «все»: так читают и Read-ProjectConfig, и insert_schedules.py.
    # Нет строки stage= — стадии нет (у Read-ProjectConfig по умолчанию «П»)
    $inProject = $cfg.InProject
    $summaryMap = $cfg.Summary
    $stage = [string]$cfg.Stage
    if ($text -notmatch '(?mi)^[ \t]*stage[ \t]*=') { $stage = "" }
    $specs = @($cfg.Specs | ForEach-Object { [string]$_ })
    $specKeys = @{}
    foreach ($s in $specs) { $specKeys[$s] = Get-ProjectSpecKey $s }
    $isIn = { param($s) ($null -eq $inProject) -or $inProject.ContainsKey($specKeys[$s]) }
    $named = New-Object System.Collections.Generic.List[object]          # модели с кодом раздела в имени, в порядке файла
    foreach ($m in $cfg.Models) {
        $code = Get-ModelSectionFromName ([string]$m.Name)          # у имени без папки запасного «по папке» нет: только код в имени
        if ($code) { $named.Add([pscustomobject]@{ Name = [string]$m.Name; Key = Get-ProjectModelKey ([string]$m.Name); Section = $code }) }
    }
    $L = New-Object System.Collections.Generic.List[string]
    $L.Add("version`t" + $script:InsertPlanVersion)
    $L.Add("stage`t" + $stage)
    $enabled = @{}          # ключ модели -> отмеченные спецификации «В проекте» (в порядке [specs])
    foreach ($m in $named) {
        if ($enabled.ContainsKey($m.Key)) { continue }
        $on = New-Object System.Collections.Generic.List[string]
        foreach ($s in $specs) {
            $c = $cfg.Cells[($m.Key + "`t" + $specKeys[$s])]
            if ($null -ne $c -and [int]$c.State -eq 1 -and (& $isIn $s)) { $on.Add($s) }
        }
        $enabled[$m.Key] = $on
    }
    $sumSpecs = New-Object System.Collections.Generic.List[string]
    foreach ($s in $specs) {
        if ($null -ne $inProject) {
            if ((& $isIn $s) -and ($null -eq $summaryMap -or $summaryMap.ContainsKey($specKeys[$s]))) { $sumSpecs.Add($s) }
        }
        else {
            foreach ($m in $named) {
                $c = $cfg.Cells[($m.Key + "`t" + $specKeys[$s])]
                if ($null -ne $c -and [int]$c.State -eq 1) { $sumSpecs.Add($s); break }
            }
        }
    }
    foreach ($s in $sumSpecs) { $L.Add("summary`t" + $s) }
    $done = @{}
    foreach ($m in $named) {
        if ($done.ContainsKey($m.Key)) { continue }
        $done[$m.Key] = $true
        $hasCells = $false
        foreach ($s in $specs) { if ($cfg.Cells.ContainsKey($m.Key + "`t" + $specKeys[$s])) { $hasCells = $true; break } }
        if (-not $hasCells) { $L.Add("model`t" + $m.Name + "`tunmarked"); continue }
        $L.Add("model`t" + $m.Name + "`tmarked")
        foreach ($s in $enabled[$m.Key]) { $L.Add("mspec`t" + $m.Name + "`t" + $s) }
    }
    $codes = New-Object System.Collections.Generic.List[string]
    foreach ($m in $named) { if (-not $codes.Contains($m.Section)) { $codes.Add($m.Section) } }
    foreach ($code in $codes) {
        $keys = @($named | Where-Object { $_.Section -eq $code } | ForEach-Object { $_.Key })
        foreach ($s in $specs) {
            if (-not (& $isIn $s)) { continue }
            foreach ($k in $keys) {
                $c = $cfg.Cells[($k + "`t" + $specKeys[$s])]
                if ($null -ne $c -and [int]$c.State -eq 1) { $L.Add("sspec`t" + $code + "`t" + $s); break }
            }
        }
    }
    $L.Add("end`t" + $L.Count)
    return , $L.ToArray()
}

function Write-InsertPlan([string]$ProjectPath, [string]$OutPath) {
    $lines = Get-InsertPlanLines $ProjectPath
    $dir = Split-Path -Parent $OutPath
    if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $head = "# План вставки для insert_schedules.py (служебный файл прогона; пишет окно из " + (Split-Path -Leaf $ProjectPath) + ")"
    [System.IO.File]::WriteAllText($OutPath, (@($head) + @($lines) -join "`r`n") + "`r`n", (New-Object System.Text.UTF8Encoding $true))
    return $OutPath
}
