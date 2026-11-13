# UTF-8 WITH BOM required (PowerShell 5.1).
# Окно: списки моделей и спецификаций для режима «Выборочно» (без Revit — из файлов состояния прошлых сборок)
# и строки run.cfg для выгрузки. Требует assignment_partial.ps1 и assignment_export_stream.ps1 (Sort-CheckNames).

try {
    if (-not (Get-Command Get-BvsReportCodes -ErrorAction SilentlyContinue)) {
        $__bvsSections = if ($PSScriptRoot) { Join-Path $PSScriptRoot "bvs_sections.ps1" } else { Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "bvs_sections.ps1" }
        if (Test-Path -LiteralPath $__bvsSections) { . $__bvsSections }
    }
}
catch { Write-Warning ("разделы: " + $_.Exception.Message) }
if (Get-Command Get-BvsReportCodes -ErrorAction SilentlyContinue) {
    $script:PartialSectionOrder = @(Get-BvsReportCodes)
    $script:PartialSectionTitle = Get-BvsTitleMap
}
else {
    Write-Warning "встроенный порядок «Выборочно» (bvs_sections.ps1 не подключён)"
    $script:PartialSectionOrder = @("EOM", "SS", "OV", "PD", "PT", "VK", "OT", "HS", "ITP", "AVT", "HC")
    $script:PartialSectionTitle = @{ EOM = "ЭОМ"; SS = "СС"; OV = "ОВ"; PD = "ПД"; PT = "ПТ"; VK = "ВК"; OT = "ОТ"; HS = "ХС"; ITP = "ИТП"; AVT = "АВТ"; HC = "ХЦ" }
}

# Раздел модели — единственное правило из project_config.ps1 (только «_КОД_» в имени файла)
function Get-PartialModelSection([string]$Name) { return (Get-ModelSectionFromName $Name) }

# Имена из старых файлов: журнал (Model — 1-я колонка, Check — 3-я), история (Model/Check — 3-я/4-я колонки, у самых старых 2-я/3-я)
function Read-NamesFromLedgerAndHistory([string]$LedgerPath, [string]$HistoryPath) {
    $res = @{ Models = @{}; Specs = @{} }
    if ($LedgerPath -and (Test-Path -LiteralPath $LedgerPath)) {
        foreach ($line in @(Get-Content -LiteralPath $LedgerPath -Encoding UTF8)) {
            $p = @($line.Split([char]9))
            if ($p.Count -ge 11) { $res.Models[$p[0]] = $true; $res.Specs[(ConvertTo-CanonicalSpecName $p[2])] = $true }
        }
    }
    if ($HistoryPath -and (Test-Path -LiteralPath $HistoryPath)) {
        foreach ($line in @(Get-Content -LiteralPath $HistoryPath -Encoding UTF8)) {
            $p = @($line.Split([char]9))
            if ($p.Count -ge 6) { $res.Models[$p[2].Trim()] = $true; $res.Specs[(ConvertTo-CanonicalSpecName $p[3].Trim())] = $true }
            elseif ($p.Count -ge 4) { $res.Models[$p[1].Trim()] = $true; $res.Specs[(ConvertTo-CanonicalSpecName $p[2].Trim())] = $true }
        }
    }
    return $res
}

# Что можно выбрать. Источники (по убыванию надёжности): каталог последней выгрузки (все спецификации и связи сводной,
# со статусом загрузки), буфер и даты проверок, журнал замечаний и история. Revit для этого не нужен.
function Get-PartialChoices([string]$HistoryPath, [string]$LedgerPath) {
    $paths = Get-PartialStatePaths $HistoryPath
    $models = @{}      # norm -> @{ Name; Loaded }
    $specs = @{}
    $addModel = {
        param($name, $loaded)
        $n = Get-ModelNorm $name
        if (-not $n) { return }
        # имя как в отчёте (Get-ModelFileName): без пути и без .rvt — оно же уходит в пары проверки
        $shown = ([string]$name).Trim()
        if ($shown -match '[^\\/]+$') { $shown = $Matches[0] }
        $shown = $shown -replace '(?i)\.rvt$', ''
        if (-not $models.ContainsKey($n)) { $models[$n] = @{ Name = $shown; Loaded = $loaded } }
    }
    $state = Read-PartialState $paths
    $cat = $state.Catalog
    foreach ($k in @($cat.Specs.Keys)) { $specs[$k] = $true }
    foreach ($n in @($cat.Models.Keys)) { & $addModel $cat.ModelNames[$n] ($cat.Models[$n] -ne "unloaded") }
    $buf = $state.Buffer
    foreach ($r in @($buf)) { & $addModel $r.Model $true; $specs[$r.Check] = $true }
    foreach ($k in @($state.Checks.Keys)) {
        $pair = @($k.Split($script:PartialSep))
        if ($pair[0] -ne "*") { & $addModel $pair[0] $true }
        $specs[$pair[1]] = $true
    }
    $old = Read-NamesFromLedgerAndHistory $LedgerPath $HistoryPath
    foreach ($m in @($old.Models.Keys)) { & $addModel $m $true }
    foreach ($s in @($old.Specs.Keys)) { $specs[$s] = $true }
    # спецификации, которых нет в каталоге (если он есть), уже удалены из сводной — не предлагаем
    $specList = @($specs.Keys | Where-Object { -not $cat.Has -or $cat.Specs.Count -lt 1 -or $cat.Specs.ContainsKey($_) })
    $modelList = @($models.Values | Where-Object { -not $cat.Has -or $cat.Models.Count -lt 1 -or $cat.Models.ContainsKey((Get-ModelNorm $_.Name)) })
    $rank = @{}
    for ($i = 0; $i -lt $script:PartialSectionOrder.Count; $i++) { $rank[$script:PartialSectionOrder[$i]] = $i }
    $out = New-Object System.Collections.Generic.List[object]
    foreach ($m in @($modelList | Where-Object { Get-PartialModelSection $_.Name } | Sort-Object { $s = Get-PartialModelSection $_.Name; if ($rank.ContainsKey($s)) { $rank[$s] } else { 99 } }, { $_.Name })) {
        $sec = Get-PartialModelSection $m.Name
        $title = "Прочее"
        if ($script:PartialSectionTitle.ContainsKey($sec)) { $title = $script:PartialSectionTitle[$sec] }
        elseif ($sec) { $title = $sec }
        $out.Add([pscustomobject]@{ Name = $m.Name; Section = $sec; SectionTitle = $title; Loaded = [bool]$m.Loaded })
    }
    return [pscustomobject]@{ Models = $out.ToArray(); Specs = @(Sort-CheckNames $specList) }
}

# Выборочно можно работать, только если есть с чем сравнивать: буфер или (журнал/история) прошлых сборок
function Test-PartialBaseReady([string]$HistoryPath, [string]$LedgerPath) {
    $paths = Get-PartialStatePaths $HistoryPath
    if ($null -ne (Read-PartialStateLines $paths).buffer) { return $true }          # буфер есть (в т. ч. пустой), как прежде «есть файл _буфер.txt»
    if ($HistoryPath -and (Test-Path -LiteralPath $HistoryPath)) { return $true }
    if ($LedgerPath -and (Test-Path -LiteralPath $LedgerPath)) { return $true }
    return $false
}

# Строки run.cfg для export_schedules.py; имена Revit и файлов не содержат «|»
function ConvertTo-PartialCfgLines([string[]]$Specs, [string[]]$Models, [string]$PairsFile = "") {
    $lines = @(
        "assign_scope=partial",
        ("assign_specs=" + (@($Specs) -join "|")),
        ("assign_models=" + (@($Models) -join "|"))
    )
    # точные пары «модель × спецификация» — в файле (export_schedules.py читает его через .NET); без него — прежнее пересечение списков
    if ($PairsFile) { $lines += ("assign_pairs_file=" + $PairsFile) }
    return $lines
}

# Сохранённый выбор (ui_settings.txt): имена через «|»
function ConvertTo-PartialSelectionText([string[]]$Names) { return (@($Names) -join "|") }
function ConvertFrom-PartialSelectionText([string]$Text) {
    $set = @{}
    foreach ($x in @(([string]$Text).Split([char]'|'))) { if ($x.Trim()) { $set[$x.Trim()] = $true } }
    return $set
}

# ---------- пары «модель × спецификация» (режим «Выборочно») ----------
# Набор спецификаций модели — из настройки проекта (вкладка «Распределение», project_config.ps1): отметки этой модели; модели нет в настройке —
# набор её раздела (раздел из имени …_IOS_<КОД>_…). Набор определить нельзя — у модели «все спецификации», причина фиксируется (Why)
# и показывается в окне. Зашитых таблиц «раздел -> пакеты» здесь больше нет.

# Ключ спецификации для сравнения: опечатка «значния», пробелы любого вида, регистр и префикс стадии не важны (имена из истории — «П        5.1_…»)
function Get-PartialSpecKey([string]$Name) { return (Get-ProjectSpecKey $Name) }

# Для каждой модели: набор (ключи спецификаций) или $null + причина, почему набор не определён.
# Project — настройка проекта (Read-ProjectConfig) или $null (файла нет).
# Возвращает @{ Sets = @{модель = @{ключ=$true} | $null}; Why = @{модель = текст причины (только для $null)} }
function Get-PartialModelSpecInfo($ModelNames, $Project) {
    $sets = @{}
    $why = @{}
    foreach ($m in @($ModelNames)) {
        $model = [string]$m
        $set = $null
        $reason = ""
        if ($null -eq $Project) {
            $reason = "нет настройки проекта: откройте вкладку «Распределение» и сохраните настройку"
        }
        else {
            $r = Get-ProjectModelSpecs $Project $model
            if ($null -eq $r.Specs) { $reason = [string]$r.Why }
            else {
                $set = @{}
                foreach ($n in @($r.Specs)) { $set[(Get-PartialSpecKey $n)] = $true }
            }
        }
        $sets[$model] = $set
        if ($null -eq $set) { $why[$model] = $reason }
    }
    return @{ Sets = $sets; Why = $why }
}

function Test-PartialModelHasSpec($ModelSets, [string]$Model, [string]$Spec) {
    if (-not $ModelSets.ContainsKey($Model)) { return $true }
    $set = $ModelSets[$Model]
    if ($null -eq $set) { return $true }
    return $set.ContainsKey((Get-PartialSpecKey $Spec))
}

# Спецификации модели (в порядке общего списка); набор не определён — все
function Get-PartialModelSpecList([string]$Model, $AllSpecs, $ModelSets) {
    $out = New-Object System.Collections.Generic.List[string]
    foreach ($sp in @($AllSpecs)) { if (Test-PartialModelHasSpec $ModelSets $Model ([string]$sp)) { $out.Add([string]$sp) } }
    return $out.ToArray()
}

# Спецификации узла = объединение спецификаций его моделей (одна модель или все модели раздела)
function Get-PartialNodeSpecs($NodeModels, $AllSpecs, $ModelSets) {
    $out = New-Object System.Collections.Generic.List[string]
    foreach ($sp in @($AllSpecs)) {
        foreach ($m in @($NodeModels)) { if (Test-PartialModelHasSpec $ModelSets ([string]$m) ([string]$sp)) { $out.Add([string]$sp); break } }
    }
    return $out.ToArray()
}

# Отметки пар: @{ модель = @{ спецификация = $true/$false } } только по спецификациям модели. Saved: $null — всё отмечено;
# @{ ключ "модель|спецификация" = $true } — сохранённые пары (формат asgpair=; «модель|*» — модель была в выборе, иначе она новая и отмечена вся); SavedSpecs — прежний формат asgspecs= (одни отметки на все модели)
function New-PartialPicks($Models, $AllSpecs, $ModelSets, $SavedPairs = $null, $SavedSpecs = $null) {
    $picks = @{}
    foreach ($m in @($Models)) {
        $model = [string]$m
        $row = @{}
        foreach ($sp in @(Get-PartialModelSpecList $model $AllSpecs $ModelSets)) {
            if ($null -ne $SavedPairs -and $SavedPairs.ContainsKey($model + "|*")) { $row[$sp] = $SavedPairs.ContainsKey($model + "|" + $sp) }   # «модель|*» — модель уже была в сохранённом выборе
            elseif ($null -ne $SavedSpecs) { $row[$sp] = $SavedSpecs.ContainsKey($sp) }
            else { $row[$sp] = $true }
        }
        $picks[$model] = $row
    }
    return $picks
}

# Состояние спецификации у узла: Checked (у всех моделей узла, где она есть), Unchecked (ни у одной), Indeterminate (частично)
function Get-PartialSpecState($Picks, $NodeModels, $ModelSets, [string]$Spec) {
    $on = 0; $all = 0
    foreach ($m in @($NodeModels)) {
        $model = [string]$m
        if (-not (Test-PartialModelHasSpec $ModelSets $model $Spec)) { continue }
        $all++
        if ($Picks.ContainsKey($model) -and $Picks[$model].ContainsKey($Spec) -and $Picks[$model][$Spec]) { $on++ }
    }
    if ($all -eq 0 -or $on -eq 0) { return "Unchecked" }
    if ($on -eq $all) { return "Checked" }
    return "Indeterminate"
}

# Отметить/снять спецификацию у всех моделей узла, где она есть
function Set-PartialSpecPick($Picks, $NodeModels, $ModelSets, [string]$Spec, [bool]$On) {
    foreach ($m in @($NodeModels)) {
        $model = [string]$m
        if (-not (Test-PartialModelHasSpec $ModelSets $model $Spec)) { continue }
        if (-not $Picks.ContainsKey($model)) { $Picks[$model] = @{} }
        $Picks[$model][$Spec] = $On
    }
}

# «все»/«ничего» — только для спецификаций узла
function Set-PartialNodeAll($Picks, $NodeModels, $AllSpecs, $ModelSets, [bool]$On) {
    foreach ($sp in @(Get-PartialNodeSpecs $NodeModels $AllSpecs $ModelSets)) { Set-PartialSpecPick $Picks $NodeModels $ModelSets $sp $On }
}

# Модель отмечена слева («выгружать эту модель»): её спецификации по умолчанию отмечаются все. Снятие галочки отметки спек не трогает.
function Update-PartialPicksOnModelCheck($Picks, $AllSpecs, $ModelSets, $PrevModels, $NewModels) {
    $prev = @{}; foreach ($m in @($PrevModels)) { $prev[[string]$m] = $true }
    foreach ($m in @($NewModels)) {
        $model = [string]$m
        if ($prev.ContainsKey($model)) { continue }
        $Picks[$model] = @{}
        foreach ($sp in @(Get-PartialModelSpecList $model $AllSpecs $ModelSets)) { $Picks[$model][$sp] = $true }
    }
}

# Выгружаемые пары: отмеченные модели × отмеченные у них спецификации. Порядок: модели как в $CheckedModels, спецификации как в общем списке
function Get-PartialPairs($Picks, $CheckedModels, $AllSpecs, $ModelSets) {
    $out = New-Object System.Collections.Generic.List[object]
    foreach ($m in @($CheckedModels)) {
        $model = [string]$m
        if (-not $Picks.ContainsKey($model)) { continue }
        foreach ($sp in @(Get-PartialModelSpecList $model $AllSpecs $ModelSets)) {
            if ($Picks[$model].ContainsKey($sp) -and $Picks[$model][$sp]) { $out.Add([pscustomobject]@{ Model = $model; Spec = $sp }) }
        }
    }
    return $out.ToArray()
}

# «<модель или раздел>: выбрано N из M»; N — спецификации, отмеченные у всех моделей узла; у раздела «частично» — отмечены не везде
function Format-PartialNodeCaption([string]$Title, $Picks, $NodeModels, $AllSpecs, $ModelSets) {
    $specs = @(Get-PartialNodeSpecs $NodeModels $AllSpecs $ModelSets)
    $n = 0; $part = 0
    foreach ($sp in $specs) {
        $st = Get-PartialSpecState $Picks $NodeModels $ModelSets $sp
        if ($st -eq "Checked") { $n++ } elseif ($st -eq "Indeterminate") { $part++ }
    }
    $text = ("{0}: выбрано {1} из {2}" -f $Title, $n, $specs.Count)
    if ($part -gt 0) { $text += (" (частично: {0})" -f $part) }
    return $text
}

# «Всего к выгрузке: P пар, K моделей» (K — отмеченные модели, у которых есть хотя бы одна пара)
function Format-PartialTotalCaption($Pairs) {
    $models = @{}
    foreach ($p in @($Pairs)) { $models[[string]$p.Model] = $true }
    return ("Всего к выгрузке: {0} пар, {1} моделей" -f @($Pairs).Count, $models.Count)
}

# Сохранение: строки asgpair=<модель>|<спецификация> для всех отмеченных пар (у любых моделей — выбор переживает снятие галочки модели)
function ConvertTo-PartialPairLines($Picks) {
    $out = New-Object System.Collections.Generic.List[string]
    foreach ($model in @($Picks.Keys | Sort-Object)) {
        foreach ($sp in @($Picks[$model].Keys | Sort-Object)) { if ($Picks[$model][$sp]) { $out.Add(("asgpair={0}|{1}" -f $model, $sp)) } }
    }
    return $out.ToArray()
}

# Файл пар для export_schedules.py: «модель<TAB>спецификация», UTF-8 с BOM (через .NET: пути и имена с кириллицей)
function Write-PartialPairsFile([string]$Path, $Pairs) {
    $lines = @(@($Pairs) | ForEach-Object { "{0}`t{1}" -f $_.Model, $_.Spec })
    [System.IO.File]::WriteAllLines($Path, [string[]]$lines, (New-Object System.Text.UTF8Encoding $true))
}
