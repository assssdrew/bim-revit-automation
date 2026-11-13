# UTF-8 WITH BOM required (PowerShell 5.1).
# Вкладка «Спецификации»: пути (общий файл, модели, сводная), стадия, таблица распределения и кнопки: «Загрузить спецификации и модели», «Распределить», «В сводную» и «В разделы» (независимы, обе после распределения); отдельно «Посчитать элементы».
# Данные и логика — project_config.ps1; здесь чтение результатов RBP-скриптов (общий файл, распределение по моделям) и окно. Подключается из ui_schedules.ps1.
# Любая правка распределения держится в памяти ($script:DistDirty), пока не нажаты «Сохранить», «Сохранить как новый проект» или «Отменить».
# Автосохранения правок пользователя нет. Служебная запись (summary_done, новый проект, числа после «Посчитать») не подменяет несохранённые правки.

$script:DistExpand = @{ Sections = @{}; Groups = @{}; Hide = $false }     # что раскрыто в таблице (по умолчанию всё свёрнуто)
$script:DistModelPaths = @{}
$script:DistRsnModels = @()                                                # модели, выбранные «С сервера…» (RSN://)                                              # имя модели -> полный путь (из папки)
$script:DistBusy = $false
$script:DistLoading = $false
$script:DistSaveTimer = $null
$script:StateMigrationNote = ""   # одна строка о переносе прежнего выбор_<проект>.txt в [inproject] (Invoke-DistStateMigration)
$script:DistDirty = $false          # после любой правки пользователя данные только в памяти, в файлы не пишутся
$script:DistDirtySnapshot = $null   # копия настройки до первой несохранённой правки
$script:DistRestoring = $false      # «Отменить» подставляет снимок и не считает это новой правкой
$script:DistDirtyPaths = $null      # копия путей моделей (имя -> файл) на тот же момент
$script:DistDirtyBypass = $false    # явное «Сохранить» пишет файлы, несмотря на DistDirty
$script:DistDirtyChoice = ""        # save | discard | cancel; пусто — спросить в окне (проверка без диалога)
$script:DistColumnChoice = ""       # yes | no; пусто — спросить, пересобирать ли колонки по связям
$script:DistThinChoice = ""         # yes | no; пусто — спросить, если связей меньше половины моделей

# ---------- чтение результатов RBP (без окна, тестируется) ----------
# template_specs.tsv (read_template_specs.py): имя TAB категория -> имена спецификаций общего файла
function Read-TemplateSpecsFile([string]$Path) {
    $names = New-Object System.Collections.Generic.List[string]
    if (-not (Test-Path -LiteralPath $Path)) { return $names.ToArray() }
    $text = [System.IO.File]::ReadAllText($Path, (New-Object System.Text.UTF8Encoding $true))
    foreach ($line in @($text -split "`r?`n")) {
        if (-not $line.Trim()) { continue }
        $n = $line.Split("`t")[0].Trim()
        if ($n) { $names.Add($n) }
    }
    return $names.ToArray()
}

# summary_links.txt (read_template_specs.py, links_path): имя связи на строку, без .rvt. Нет файла — пусто, без ошибки.
function Read-SummaryLinksFile([string]$Path) {
    $names = New-Object System.Collections.Generic.List[string]
    if (-not $Path -or -not (Test-Path -LiteralPath $Path)) { return $names.ToArray() }
    $text = [System.IO.File]::ReadAllText($Path, (New-Object System.Text.UTF8Encoding $true))
    foreach ($line in @($text -split "`r?`n")) {
        $t = $line.Trim()
        if (-not $t -or $t.StartsWith("#")) { continue }
        $i = $t.IndexOf(".rvt", [System.StringComparison]::OrdinalIgnoreCase)
        if ($i -ge 0) { $t = $t.Substring(0, $i).Trim() }
        $name = Get-ProjectModelName $t
        if ($name) { $names.Add($name) }
    }
    return $names.ToArray()
}

function Format-DistMissingFiles($Names) {
    $list = New-Object System.Collections.Generic.List[string]
    foreach ($x in @($Names)) { if ($x) { $list.Add([string]$x) } }
    if ($list.Count -eq 0) { return "" }
    return ("нет файла в папке: {0}." -f ($list -join ", "))
}

# «Уйдут модели (N): …» — до 10 имён, дальше «и ещё N».
function Format-LeavingModels($Names) {
    $list = New-Object System.Collections.Generic.List[string]
    foreach ($x in @($Names)) { if ($x) { $list.Add([string]$x) } }
    if ($list.Count -eq 0) { return "Уйдут модели: нет." }
    $shown = New-Object System.Collections.Generic.List[string]
    $i = 0
    foreach ($n in $list) {
        if ($i -ge 10) { break }
        $shown.Add($n)
        $i++
    }
    $text = ("Уйдут модели ({0}): {1}" -f $list.Count, ($shown -join ", "))
    if ($list.Count -gt 10) { $text += (" и ещё {0}" -f ($list.Count - 10)) }
    return $text + "."
}

function Get-DistColumnRebuildPrompt($LeavingNames) {
    return ("Колонки моделей пересоберутся по связям сводной. Настройки (галочки и числа) моделей, которых нет среди связей, будут удалены при «Сохранить». «Отменить» вернёт всё.`r`n`r`n" + (Format-LeavingModels $LeavingNames))
}

function Get-DistThinLinksPrompt([int]$Links, [int]$Models, $LeavingNames) {
    $head = ("Связей сводной меньше половины моделей таблицы ({0} из {1}). Пересобрать колонки по связям? Настройки (галочки и числа) моделей, которых нет среди связей, будут удалены при «Сохранить». «Отменить» вернёт всё." -f $Links, $Models)
    return ($head + "`r`n`r`n" + (Format-LeavingModels $LeavingNames))
}

# yes/no — ответ без окна (проверка). Пусто — вопрос, по умолчанию «Нет».
function Confirm-DistYes([string]$Preset, [string]$Text, [string]$Title) {
    $c = ([string]$Preset).Trim().ToLowerInvariant()
    if ($c -eq "yes" -or $c -eq "да") { return $true }
    if ($c -eq "no" -or $c -eq "нет") { return $false }
    $ask = [System.Windows.Forms.MessageBox]::Show($form, $Text, $Title, [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question, [System.Windows.Forms.MessageBoxDefaultButton]::Button2)
    return ($ask -eq [System.Windows.Forms.DialogResult]::Yes)
}

function Copy-DistModelPaths($Map) {
    $n = @{}
    if ($null -eq $Map) { return $n }
    foreach ($k in $Map.Keys) { $n[[string]$k] = [string]$Map[$k] }
    return $n
}

# «Выборочно»: убрать модели, которых больше нет в таблице (те же карты, что у проекта, либо отдельные).
function Update-AsgSelForProjectModels($Cfg) {
    if ($null -eq $script:AsgSel -or $null -eq $Cfg) { return }
    $keep = @{}
    foreach ($m in $Cfg.Models) { $keep[(Get-ProjectModelKey $m.Name)] = $true }
    Remove-AsgMapModels $script:AsgSel.Models $keep
    Remove-AsgMapModels $script:AsgSel.Pairs $keep
}

# Кэш путей только для переданных моделей: файл ищется в папке моделей. Нет файла — имя в результате.
function Update-DistModelPathsForNames($Names) {
    $files = @()
    try { $files = @(Get-DistModelFiles) } catch { $files = @() }
    $by = @{}
    foreach ($f in $files) {
        $k = Get-LinkMatchKey $f
        if ($k -and -not $by.ContainsKey($k)) { $by[$k] = [string]$f }
    }
    $script:DistModelPaths = @{}
    $missing = New-Object System.Collections.Generic.List[string]
    foreach ($raw in @($Names)) {
        $name = Get-ProjectModelName $raw
        $k = Get-LinkMatchKey $name
        if (-not $k) { continue }
        if ($by.ContainsKey($k)) { $script:DistModelPaths[$name] = $by[$k] } else { $missing.Add($name) }
    }
    return $missing.ToArray()
}

# template_specs.tsv: категория спецификации (вторая колонка) -> @{ ключ спеки = категория }; пустая категория не записывается
function Read-TemplateSpecCategories([string]$Path) {
    $map = @{}
    if (-not (Test-Path -LiteralPath $Path)) { return $map }
    $text = [System.IO.File]::ReadAllText($Path, (New-Object System.Text.UTF8Encoding $true))
    foreach ($line in @($text -split "`r?`n")) {
        if (-not $line.Trim()) { continue }
        $p = $line.Split("`t")
        if ($p.Count -ge 2 -and $p[0].Trim() -and $p[1].Trim()) { $map[(Get-ProjectSpecKey $p[0].Trim())] = $p[1].Trim() }
    }
    return $map
}

# prefill.tsv (insert_schedules.py, mode=prefill): модель TAB спека TAB число TAB секунды TAB способ/ошибка; итоги по моделям — строки «#МОДЕЛЬ»:
# модель, секунды, спеки, ошибки, способы подсчёта, пачка|по одной, секунды: открытие источника, вставка, профиль, подсчёт, откат
# Возвращает @{ Counts = @(@{Model;Spec;Count;Seconds;Note}); Models = @(@{Model;Seconds;Specs;Errors;Methods;Mode;Open;Paste;Profile;Count;Rollback}); Errors = @(строки «#ОШИБКА») }
function Read-PrefillResults([string]$Path) {
    $counts = New-Object System.Collections.Generic.List[object]
    $models = New-Object System.Collections.Generic.List[object]
    $errors = New-Object System.Collections.Generic.List[string]
    if (Test-Path -LiteralPath $Path) {
        $text = [System.IO.File]::ReadAllText($Path, (New-Object System.Text.UTF8Encoding $true))
        $inv = [System.Globalization.CultureInfo]::InvariantCulture
        $num = { param($p, $i) $v = 0.0; if ($p.Count -gt $i) { [void][double]::TryParse($p[$i], [System.Globalization.NumberStyles]::Float, $inv, [ref]$v) }; return $v }
        foreach ($line in @($text -split "`r?`n")) {
            if (-not $line.Trim()) { continue }
            $p = $line.Split("`t")
            if ($p[0] -eq "#МОДЕЛЬ") {
                $sp = 0; $er = 0
                if ($p.Count -gt 3) { [void][int]::TryParse($p[3], [ref]$sp) }
                if ($p.Count -gt 4) { [void][int]::TryParse($p[4], [ref]$er) }
                $m = ""; if ($p.Count -gt 5) { $m = $p[5] }
                $mode = ""; if ($p.Count -gt 6) { $mode = $p[6] }
                $models.Add([pscustomobject]@{ Model = $p[1]; Seconds = (& $num $p 2); Specs = $sp; Errors = $er; Methods = $m; Mode = $mode
                        Open = (& $num $p 7); Paste = (& $num $p 8); Profile = (& $num $p 9); Count = (& $num $p 10); Rollback = (& $num $p 11) })
            }
            elseif ($p[0] -eq "#ОШИБКА") { $errors.Add(("{0}: {1}" -f $p[1], $(if ($p.Count -gt 2) { $p[2] } else { "" }))) }
            elseif ($p.Count -ge 3) {
                $n = -1; [void][int]::TryParse($p[2], [ref]$n)
                $note = ""; if ($p.Count -gt 4) { $note = $p[4] }
                $counts.Add([pscustomobject]@{ Model = $p[0]; Spec = $p[1]; Count = $n; Seconds = (& $num $p 3); Note = $note })
            }
        }
    }
    return @{ Counts = $counts.ToArray(); Models = $models.ToArray(); Errors = $errors.ToArray() }
}

function Format-DistClock([double]$Seconds) {
    $t = [TimeSpan]::FromSeconds($Seconds)
    return ("{0:00}:{1:00}:{2:00}" -f [int][math]::Floor($t.TotalHours), $t.Minutes, $t.Seconds)
}

# Применяет результат к настройке (только числа элементов — галочки не меняются); возвращает текст итога со временем прогона (в окно)
function Merge-PrefillIntoProject($Cfg, $Result, [double]$WallSeconds) {
    Set-ProjectCellsFromCounts $Cfg $Result.Counts
    $ok = @($Result.Counts | Where-Object { $_.Count -ge 0 }).Count
    $bad = @($Result.Counts | Where-Object { $_.Count -lt 0 }).Count
    $nModels = @($Result.Models).Count
    $per = @($Result.Models | ForEach-Object { "{0}: {1:N1} с" -f $_.Model, $_.Seconds }) -join "; "
    $text = ("Посчитано: пар {0}, не определено {1}, моделей {2}. Время прогона {3}" -f $ok, $bad, $nModels, (Format-DistClock $WallSeconds))
    if ($per) { $text += ". По моделям — " + $per }
    return $text
}

# Журнал прогона: по каждой модели — способ (пачка / по одной), время открытия источника, вставки, профиля, подсчёта, отката; в конце общее время
function Get-DistLogLines($Result, [double]$WallSeconds) {
    $L = New-Object System.Collections.Generic.List[string]
    $sum = 0.0
    foreach ($m in @($Result.Models)) {
        $sum += $m.Seconds
        $L.Add(("{0}: {1}; спецификаций {2}, ошибок {3}; открытие источника {4:N1} с, вставка {5:N1} с, профиль {6:N1} с, подсчёт {7:N1} с, откат {8:N1} с; всего {9:N1} с" -f
                $m.Model, $(if ($m.Mode) { $m.Mode } else { "-" }), $m.Specs, $m.Errors, $m.Open, $m.Paste, $m.Profile, $m.Count, $m.Rollback, $m.Seconds))
        if ($m.Mode -eq "по одной") { $L.Add(("{0}: пачка не вставилась, спецификации вставлены по одной (причина — в журнале RBP: WARN prefill)" -f $m.Model)) }
    }
    foreach ($e in @($Result.Errors)) { $L.Add("ОШИБКА " + $e) }
    $L.Add(("Общее время прогона: {0} ({1:N1} с); сумма по моделям в скриптах: {2:N1} с" -f (Format-DistClock $WallSeconds), $WallSeconds, $sum))
    return $L.ToArray()
}

function Get-DistSummaryText($Cfg) {
    $groups = @(Get-ProjectGroups $Cfg)
    $inProj = @($groups | Where-Object { (Get-ProjectGroupInState $Cfg $_.Specs) -ne "Unchecked" }).Count
    $undef = 0; $on = 0
    foreach ($c in $Cfg.Cells.Values) { if ($c.Origin -eq "x") { $undef++ }; if ($c.State -eq 1) { $on++ } }
    $nSpecs = [int]$Cfg.Specs.Count; $nModels = [int]$Cfg.Models.Count; $nSections = [int]$Cfg.Sections.Count
    $nSum = [int]@(Get-ProjectSummarySpecs $Cfg).Count
    return ("Спецификаций {0} · групп {1} · в проекте {2} · в сводную {3} · моделей {4} · пар {5} · не определено {6}" -f $nSpecs, $groups.Count, $inProj, $nSum, $nModels, $on, $undef)
}

# ---------- окно ----------
function Get-DistColor([string]$Name) {
    switch ($Name) {
        "green" { return [System.Drawing.Color]::FromArgb(198, 239, 206) }
        "yellow" { return [System.Drawing.Color]::FromArgb(255, 235, 156) }
        "red" { return [System.Drawing.Color]::FromArgb(255, 199, 206) }
        "gray" { return [System.Drawing.Color]::FromArgb(217, 217, 217) }
        default { return [System.Drawing.Color]::White }
    }
}

# Цвета разделов: полоса заголовка раздела — насыщеннее, столбцы его моделей — светлее; Level 0 — раздел, 1 — модель, 2 — фон пустых ячеек
$script:DistTints = @(@(91, 155, 213), @(112, 173, 71), @(237, 125, 49), @(165, 105, 189), @(0, 150, 136), @(191, 144, 0), @(200, 90, 120), @(110, 130, 150), @(150, 100, 60))
function Get-DistTint([int]$Index, [int]$Level) {
    $b = $script:DistTints[$Index % $script:DistTints.Count]
    $mix = @(0.45, 0.72, 0.9)[$Level]
    return [System.Drawing.Color]::FromArgb([int]($b[0] + (255 - $b[0]) * $mix), [int]($b[1] + (255 - $b[1]) * $mix), [int]($b[2] + (255 - $b[2]) * $mix))
}

function Get-ProjectFilePath {
    if ($script:ProjectPath) { return $script:ProjectPath }
    return ""
}

# Столбцы таблицы в окне: разделитель между разделами не рисуется (границы — тонкая сетка одного цвета; цветом отделены заголовки разделов)
function Get-DistViewColumns($View) { return @($View.Columns | Where-Object { $_.Kind -ne "sep" }) }

# Подсветка спецификаций и моделей, которые справочник разделов не распознал: светло-оранжевая заливка названия / заголовка
function Get-DistIssueColor { return [System.Drawing.Color]::FromArgb(255, 217, 179) }

# Значения и цвета ячеек одной строки (общий код полной перерисовки и точечного обновления)
function Set-DistRowCells($GridRow, $row, $columns, [bool]$WithName) {
    for ($k = 0; $k -lt $columns.Count; $k++) {
        $col = $columns[$k]
        $cell = $GridRow.Cells[$k]
        switch ($col.Kind) {
            "name" {
                if ($WithName) {
                    $cell.Value = $row.Title
                    # отступ слева: место под значок раскрытия (у группы) и вложенность (у спецификации); одиночная группа выровнена с остальными
                    $cell.Style.Padding = New-Object System.Windows.Forms.Padding($(if ($row.Kind -eq "spec") { 40 } else { 20 }), 0, 0, 0)
                }
                if ($row.Issue) { $cell.Style.BackColor = Get-DistIssueColor; $cell.ToolTipText = $row.IssueText }
                else { $cell.Style.BackColor = [System.Drawing.Color]::Empty; $cell.ToolTipText = "" }
            }
            "sep" { }
            default {
                $cc = Get-ProjectGridCell $script:Project $row $col
                $cell.Value = $cc.Text
                if ($cc.Color -ne "none") { $cell.Style.BackColor = Get-DistColor $cc.Color }
                elseif ($col.Tint -ge 0 -and $row.In -ne "Unchecked") { $cell.Style.BackColor = Get-DistTint $col.Tint 2 }
                else { $cell.Style.BackColor = [System.Drawing.Color]::Empty }
            }
        }
    }
}

# Точечное обновление после клика: пересчитываются только строки одной группы (сама группа и её спецификации), без перестройки столбцов и всей таблицы.
# Группа, ставшая «не в проекте» при включённом «Скрыть не участвующие», убирается из таблицы.
function Update-DistGridRows([string]$Group) {
    $g = $ui.DistGrid
    $grp = @(Get-ProjectGroups $script:Project | Where-Object { $_.Name -eq $Group })
    if ($grp.Count -eq 0) { Show-DistGrid; return }
    $columns = @($g.Columns | ForEach-Object { $_.Tag })
    $gone = New-Object System.Collections.Generic.List[int]
    $script:DistBusy = $true
    try {
        $g.SuspendLayout()
        for ($i = 0; $i -lt $g.Rows.Count; $i++) {
            $row = $g.Rows[$i].Tag
            if ($null -eq $row -or $row.Group -ne $Group) { continue }
            Update-ProjectGridRowState $script:Project $row          # у группы и у каждой её спецификации — своё состояние
            if ($script:DistExpand.Hide -and $row.In -eq "Unchecked") { $gone.Add($i); continue }
            Set-DistRowCells $g.Rows[$i] $row $columns $false
        }
        for ($i = $gone.Count - 1; $i -ge 0; $i--) { $g.Rows.RemoveAt($gone[$i]) }
        $g.ResumeLayout()
    }
    finally { $script:DistBusy = $false }
    Update-DistColumnWidths
    Update-ActionState
}

# Тонкий значок раскрытия, как в дереве проводника: уголок «›» (свёрнуто) / «⌄» (раскрыто), линия 1,2 px серого цвета; рисуется поверх ячейки
function Write-DistChevron($Graphics, [System.Drawing.Rectangle]$Bounds, [bool]$Open, [System.Drawing.Color]$Color) {
    $old = $Graphics.SmoothingMode
    $Graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $pen = New-Object System.Drawing.Pen($Color, 1.3)
    try {
        $cx = $Bounds.Left + 9; $cy = $Bounds.Top + [int]($Bounds.Height / 2)
        if ($Open) { $Graphics.DrawLines($pen, [System.Drawing.Point[]]@((New-Object System.Drawing.Point(($cx - 4), ($cy - 2))), (New-Object System.Drawing.Point($cx, ($cy + 2))), (New-Object System.Drawing.Point(($cx + 4), ($cy - 2))))) }
        else { $Graphics.DrawLines($pen, [System.Drawing.Point[]]@((New-Object System.Drawing.Point(($cx - 2), ($cy - 4))), (New-Object System.Drawing.Point(($cx + 2), $cy)), (New-Object System.Drawing.Point(($cx - 2), ($cy + 4))))) }
    }
    finally { $pen.Dispose(); $Graphics.SmoothingMode = $old }
}

# Значок рисуется у групп из нескольких спецификаций (в названии столбца «Группа / спецификация») и у заголовков разделов
function Invoke-DistCellPainting($e) {
    $g = $ui.DistGrid
    if ($e.ColumnIndex -lt 0) { return }
    $col = $g.Columns[$e.ColumnIndex].Tag
    if ($null -eq $col) { return }
    $open = $null
    if ($e.RowIndex -ge 0) {
        $row = $g.Rows[$e.RowIndex].Tag
        if ($col.Kind -eq "name" -and $null -ne $row -and $row.Kind -eq "group" -and $row.Expandable) { $open = [bool]$row.Open }
    }
    elseif ($col.Kind -eq "section") { $open = [bool]$col.Open }
    if ($null -eq $open) { return }
    $e.Paint($e.CellBounds, [System.Windows.Forms.DataGridViewPaintParts]::All)
    Write-DistChevron $e.Graphics $e.CellBounds $open ([System.Drawing.Color]::FromArgb(110, 110, 110))
    $e.Handled = $true
}

# Ширина столбцов: «Группа / спецификация» и «В проекте» — фиксированная; «Сводная», разделы и модели делят остаток поровну, но не уже минимума;
# окно уже суммы минимумов — горизонтальная прокрутка вместо сжатия
$script:DistFixedWidths = @{ name = 360; inproject = 78 }
$script:DistMinWidths = @{ summary = 70; section = 84; model = 84; ph = 84 }
$script:DistWidthBusy = $false
function Get-DistColumnWidths([int]$Available, $Kinds) {
    $flex = @($Kinds | Where-Object { -not $script:DistFixedWidths.ContainsKey([string]$_) })
    $fixed = 0
    foreach ($k in @($Kinds)) { if ($script:DistFixedWidths.ContainsKey([string]$k)) { $fixed += [int]$script:DistFixedWidths[[string]$k] } }
    $share = 0
    if ($flex.Count -gt 0) { $share = [int][Math]::Floor(([Math]::Max(0, $Available - $fixed)) / $flex.Count) }
    $out = @()
    foreach ($k in @($Kinds)) {
        if ($script:DistFixedWidths.ContainsKey([string]$k)) { $out += [int]$script:DistFixedWidths[[string]$k] }
        else { $out += [Math]::Max([int]$script:DistMinWidths[[string]$k], $share) }
    }
    return $out
}

function Update-DistColumnWidths {
    $g = $ui.DistGrid
    if ($script:DistWidthBusy -or $null -eq $g -or $g.Columns.Count -eq 0) { return }
    $script:DistWidthBusy = $true
    try {
        $vs = 0
        foreach ($c in @($g.Controls)) { if ($c -is [System.Windows.Forms.VScrollBar] -and $c.Visible) { $vs = $c.Width } }
        $kinds = @($g.Columns | ForEach-Object { [string]$_.Tag.Kind })
        $w = @(Get-DistColumnWidths ([int]($g.ClientSize.Width - $vs - 2)) $kinds)
        for ($i = 0; $i -lt $g.Columns.Count; $i++) {
            $col = $g.Columns[$i]
            $k = [string]$col.Tag.Kind
            $min = if ($script:DistFixedWidths.ContainsKey($k)) { [int]$script:DistFixedWidths[$k] } else { [int]$script:DistMinWidths[$k] }
            $col.MinimumWidth = $min
            if ($col.Width -ne $w[$i]) { $col.Width = $w[$i] }
        }
    }
    finally { $script:DistWidthBusy = $false }
}

# Подсказка по центру пустой таблицы
function Update-DistHintPosition {
    $g = $ui.DistGrid
    $h = $ui.DistHint
    if ($null -eq $g -or $null -eq $h) { return }
    $h.Left = [Math]::Max(0, [int](($g.ClientSize.Width - $h.Width) / 2))
    $h.Top = [Math]::Max(0, [int](($g.ClientSize.Height - $h.Height) / 2))
}

# Пустая таблица: шапка (Группа / В проекте / Сводная / разделы), серые пустые строки и подсказка по центру; без белого поля
function Show-DistPlaceholder {
    $g = $ui.DistGrid
    $gray = [System.Drawing.Color]::FromArgb(243, 243, 243)
    $script:DistBusy = $true
    try {
        $g.SuspendLayout()
        $g.Columns.Clear(); $g.Rows.Clear()
        $g.BackgroundColor = $gray
        $defs = New-Object System.Collections.Generic.List[object]
        $defs.Add(@{ Kind = "name"; Title = "Группа / спецификация"; Tint = -1 })
        $defs.Add(@{ Kind = "inproject"; Title = "В проекте"; Tint = -1 })
        $defs.Add(@{ Kind = "summary"; Title = "Сводная"; Tint = -1 })
        $t = 0
        foreach ($s in $script:PcSections) { $defs.Add(@{ Kind = "ph"; Title = [string]$s.Title; Tint = $t }); $t++ }
        foreach ($d in $defs) {
            $c = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
            $c.HeaderText = $d.Title; $c.SortMode = "NotSortable"
            $c.Tag = [pscustomobject]@{ Kind = $d.Kind; Title = $d.Title; Tint = $d.Tint; Models = @(); Section = ""; Model = ""; Open = $false; Issue = "" }
            $c.DefaultCellStyle.BackColor = $gray
            $c.DefaultCellStyle.Alignment = "MiddleCenter"
            if ($d.Kind -eq "name") { $c.Frozen = $true; $c.DefaultCellStyle.Alignment = "MiddleLeft" }
            elseif ($d.Kind -eq "inproject") { $c.Frozen = $true }
            elseif ($d.Kind -eq "summary") { $c.HeaderCell.Style.BackColor = [System.Drawing.Color]::FromArgb(200, 214, 229) }
            else { $c.HeaderCell.Style.BackColor = Get-DistTint $d.Tint 0 }
            [void]$g.Columns.Add($c)
        }
        for ($i = 0; $i -lt 40; $i++) { [void]$g.Rows.Add() }
        $g.ResumeLayout()
    }
    finally { $script:DistBusy = $false }
    $script:DistPlaceholder = $true
    if ($ui.DistHint) { $ui.DistHint.Visible = $true }
    Update-DistColumnWidths
    Update-DistHintPosition
    Update-ActionState
}

# Вид таблицы из настройки и состояния «раскрыто/свёрнуто»
function Show-DistGrid {
    $g = $ui.DistGrid
    if ($null -eq $script:Project -or $script:Project.Specs.Count -eq 0) {
        Show-DistPlaceholder
        return
    }
    $script:DistPlaceholder = $false
    if ($ui.DistHint) { $ui.DistHint.Visible = $false }
    $g.BackgroundColor = [System.Drawing.Color]::White
    $first = -1
    try { $first = $g.FirstDisplayedScrollingRowIndex } catch { $first = -1 }
    $view = Get-ProjectGridModel $script:Project $script:DistExpand
    $columns = @(Get-DistViewColumns $view)
    $script:DistBusy = $true
    try {
        $g.SuspendLayout()
        $g.Columns.Clear()
        $g.Rows.Clear()
        foreach ($col in $columns) {
            $c = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
            $c.HeaderText = $col.Title; $c.Tag = $col; $c.SortMode = "NotSortable"
            $c.DefaultCellStyle.Alignment = "MiddleCenter"
            switch ($col.Kind) {
                "name" { $c.Width = 360; $c.Frozen = $true; $c.DefaultCellStyle.Alignment = "MiddleLeft" }
                "inproject" { $c.Width = 78; $c.Frozen = $true }
                "summary" { $c.Width = 76; $c.Frozen = $true; $c.HeaderCell.Style.BackColor = [System.Drawing.Color]::FromArgb(200, 214, 229) }
                "section" {
                    $c.Width = 96; $c.HeaderCell.Style.BackColor = Get-DistTint $col.Tint 0
                    $c.HeaderCell.Style.Padding = New-Object System.Windows.Forms.Padding(16, 0, 0, 0)
                }
                "model" { $c.Width = 92; $c.HeaderCell.Style.BackColor = Get-DistTint $col.Tint 1 }
            }
            if (($col.Kind -eq "section" -or $col.Kind -eq "model") -and $col.Issue) {
                $c.HeaderCell.Style.BackColor = Get-DistIssueColor
                $c.HeaderCell.ToolTipText = "Раздел модели не найден в справочнике разделов — спецификации не отмечаются"
            }
            [void]$g.Columns.Add($c)
        }
        foreach ($row in $view.Rows) {
            $i = $g.Rows.Add()
            $r = $g.Rows[$i]
            $r.Tag = $row
            Set-DistRowCells $r $row $columns $true
        }
        if ($first -ge 0 -and $first -lt $g.Rows.Count) { try { $g.FirstDisplayedScrollingRowIndex = $first } catch { } }
        $g.ResumeLayout()
    }
    finally { $script:DistBusy = $false }
    Update-DistColumnWidths
    Update-ActionState
}

# Запись файла проекта. Пока есть несохранённые правки пользователя, молча не пишется (явное сохранение ставит DistDirtyBypass).
function Save-DistProject {
    if ($script:DistSaveTimer) { $script:DistSaveTimer.Stop() }
    if ($script:DistDirty -and -not $script:DistDirtyBypass) { return $false }
    if ($null -eq $script:Project -or -not $script:ProjectPath) { return $false }
    $script:Project.Stage = $ui.DistStage.Text.Trim()
    $script:Project.Template = $ui.DistTemplate.Text.Trim()
    $script:Project.ModelsDir = $ui.DistModelsDir.Text.Trim()
    $script:Project.ModelsRsn = New-Object System.Collections.Generic.List[string]
    if (Test-RsnModelsSource $script:Project.ModelsDir) { foreach ($m in $script:DistRsnModels) { $script:Project.ModelsRsn.Add([string]$m) } }
    # всё, что относится к проекту, — в его настройку: сводная, ссылка на таблицу, «В сводную выполнена», режим вставки, «Выборочно»
    $script:Project.SummaryModel = $ui.Summary.Text.Trim()
    if ($ui.ContainsKey("Google") -and $ui.Google) { $script:Project.GoogleLink = $ui.Google.Text.Trim() }
    $script:Project.SummaryDone = [bool]$script:SummaryDone
    if ($ui.ContainsKey("InsertModeBox") -and $ui.InsertModeBox) { $script:Project.InsertMode = Get-InsertModeValue }
    if ($ui.ContainsKey("AsgPart") -and $ui.AsgPart) {
        $script:Project.AsgScope = $(if ($ui.AsgPart.Checked) { "partial" } else { "all" })
        $script:Project.AsgModels = $script:AsgSel.Models
        $script:Project.AsgPairs = $script:AsgSel.Pairs
    }
    try { Write-ProjectConfig $script:Project $script:ProjectPath }
    catch { $ui.Status.Text = ("Не удалось сохранить настройку проекта: {0}" -f $_.Exception.Message); return $false }
    return $true
}

# Галочки «В проекте» — только в [inproject] настройки проекта: пишутся вместе с ней (Save-DistProject) при явном «Сохранить».

# Правка пользователя: снимок до первой правки, файлы не пишутся. Прежнее имя оставлено — им помечены поля окна.
function Request-DistSave {
    if ($script:DistLoading -or $script:DistRestoring) { return }
    Enter-DistDirty
}

function Flush-DistSave {
    if ($script:DistSaveTimer) { $script:DistSaveTimer.Stop() }
}

# «В сводную» выполнена: флаг в память и одна строка файла. Несохранённые правки пользователя в файл не попадают.
function Update-SummaryDoneInFile([string]$Path, [bool]$Done) {
    if (-not $Path -or -not (Test-Path -LiteralPath $Path)) { return $false }
    $enc = New-Object System.Text.UTF8Encoding $true
    $text = [System.IO.File]::ReadAllText($Path, $enc)
    $line = "summary_done=" + $(if ($Done) { "1" } else { "0" })
    if ($text -match '(?m)^summary_done=') {
        $text = [regex]::Replace($text, '(?m)^summary_done=[^\r\n]*', $line)
    }
    else {
        $m = [regex]::Match($text, '(?m)^\[project\][^\r\n]*')
        if (-not $m.Success) { return $false }
        $text = $text.Remove($m.Index, $m.Length).Insert($m.Index, ($m.Value + "`r`n" + $line))
    }
    $tmp = $Path + ".tmp"
    [System.IO.File]::WriteAllText($tmp, $text, $enc)
    if (Test-Path -LiteralPath $Path) { [System.IO.File]::Delete($Path) }
    [System.IO.File]::Move($tmp, $Path)
    return $true
}

function Publish-SummaryDone([bool]$Done) {
    $script:SummaryDone = [bool]$Done
    if ($null -ne $script:Project) { $script:Project.SummaryDone = [bool]$Done }
    if ($script:DistDirty -and $null -ne $script:DistDirtySnapshot) { $script:DistDirtySnapshot.SummaryDone = [bool]$Done }
    if ($script:ProjectPath) { [void](Update-SummaryDoneInFile $script:ProjectPath ([bool]$Done)) }
}

# Надпись и кнопки «Сохранить» / «Сохранить как новый проект» / «Отменить» — только пока правки не записаны
function Update-DistDirtyChrome {
    if ($null -eq $ui) { return }
    $on = [bool]$script:DistDirty
    foreach ($key in @("DistDirtyNote", "DistCommit", "DistSaveAs", "DistRevert")) {
        if ($ui.ContainsKey($key) -and $ui[$key]) { $ui[$key].Visible = $on }
    }
}

function Exit-DistDirty {
    $script:DistDirty = $false
    $script:DistDirtySnapshot = $null
    $script:DistDirtyPaths = $null
    if ($script:DistSaveTimer) { $script:DistSaveTimer.Stop() }
    Update-DistDirtyChrome
}

# Первая правка: снимок текущего состояния, дальше файлы не трогать. Повторная правка снимок не сдвигает.
function Enter-DistDirty {
    if ($script:DistLoading -or $script:DistRestoring -or $null -eq $script:Project) { return }
    if ($script:DistDirty) { Update-DistDirtyChrome; return }
    if ($script:DistSaveTimer) { $script:DistSaveTimer.Stop() }
    $script:DistDirtySnapshot = Copy-ProjectConfig $script:Project
    $script:DistDirtyPaths = Copy-DistModelPaths $script:DistModelPaths
    $script:DistDirty = $true
    Update-DistDirtyChrome
    if ($null -ne $ui -and $ui.ContainsKey("Status") -and $ui.Status) { $ui.Status.Text = "Есть несохранённые изменения" }
}

# Явная запись: живой выбор «Выборочно», без спрятанных чужих моделей. Неудача запись не оставляет дыркой в памяти.
function Invoke-DistExplicitWrite([scriptblock]$Write) {
    $heldModels = New-Object System.Collections.Generic.List[object]
    $heldCells = New-Object System.Collections.Generic.List[string]
    if ($script:Project.HeldModels) { foreach ($m in $script:Project.HeldModels) { [void]$heldModels.Add([pscustomobject]@{ Name = [string]$m.Name; Section = [string]$m.Section }) } }
    if ($script:Project.HeldCells) { foreach ($c in $script:Project.HeldCells) { if ($null -ne $c) { [void]$heldCells.Add([string]$c) } } }
    if (-not $script:DistRestoring -and $script:PartialUiReady -and $null -ne $ui -and $ui.ContainsKey("AsgTree") -and $ui.AsgTree) { Save-AsgSelection }
    Clear-ProjectHeld $script:Project
    $script:DistDirtyBypass = $true
    $ok = $false
    try { $ok = [bool](& $Write) }
    finally { $script:DistDirtyBypass = $false }
    if (-not $ok) {
        $script:Project.HeldModels = $heldModels
        $script:Project.HeldCells = $heldCells
    }
    return $ok
}

# Записать текущую настройку и выбор в файлы открытого проекта
function Commit-DistDirty {
    if ($null -eq $script:Project -or -not $script:ProjectPath) { return $false }
    $ok = Invoke-DistExplicitWrite { return [bool](Save-DistProject) }
    if (-not $ok) { return $false }
    Exit-DistDirty
    return $true
}

function Restore-DistDirty {
    $paths = $script:DistDirtyPaths
    $script:DistRestoring = $true
    try {
        if ($null -ne $script:DistDirtySnapshot) { $script:Project = $script:DistDirtySnapshot }
        if ($null -ne $ui -and $ui.ContainsKey("DistStage") -and $ui.DistStage) { Set-ProjectFields $script:Project }
        if ($null -ne $paths) { $script:DistModelPaths = $paths }
        if ($null -ne $ui -and $ui.ContainsKey("DistGrid") -and $ui.DistGrid) { Show-DistGrid }
        if (Get-Command Refresh-PartialLists -ErrorAction SilentlyContinue) { Refresh-PartialLists }
    }
    finally { $script:DistRestoring = $false }
    Exit-DistDirty
}

function Invoke-DistCommit {
    if (-not $script:DistDirty) { return }
    if (Commit-DistDirty) {
        if ($null -ne $ui -and $ui.ContainsKey("Status") -and $ui.Status) { $ui.Status.Text = "Изменения сохранены." }
    }
}

function Invoke-DistRevert {
    if (-not $script:DistDirty) { return }
    Restore-DistDirty
    if ($null -ne $ui -and $ui.ContainsKey("Status") -and $ui.Status) { $ui.Status.Text = "Изменения отменены." }
}

# Текущая настройка — в новый проект. Файлы открытого проекта не пишутся. Окно переходит на новый проект.
function Save-DistAsNewProject([string]$Name) {
    if ($null -eq $script:Project) { return $false }
    $name = ([string]$Name).Trim()
    $problem = Test-NewProjectName $ReportDir $name
    if ($problem) {
        if ($form) { [System.Windows.Forms.MessageBox]::Show($form, $problem, "Сохранить как новый проект") | Out-Null }
        else { [System.Windows.Forms.MessageBox]::Show($problem, "Сохранить как новый проект") | Out-Null }
        return $false
    }
    $newPath = Get-ProjectConfigPath $ReportDir $name
    $oldPath = $script:ProjectPath
    $script:ProjectPath = $newPath
    $ok = Invoke-DistExplicitWrite { return [bool](Save-DistProject) }
    if (-not $ok) {
        $script:ProjectPath = $oldPath
        return $false
    }
    Exit-DistDirty
    if ($null -ne $ui -and $ui.ContainsKey("ProjectBox") -and $ui.ProjectBox) { Update-ProjectList $name }
    if (Get-Command Save-Settings -ErrorAction SilentlyContinue) { Save-Settings }
    if ($null -ne $ui -and $ui.ContainsKey("Status") -and $ui.Status) { $ui.Status.Text = ("Проект {0} сохранён. Прежний проект не изменён." -f $name) }
    return $true
}

function Invoke-DistSaveAs {
    if (-not $script:DistDirty -or $null -eq $script:Project) { return }
    Add-Type -AssemblyName Microsoft.VisualBasic
    $name = [Microsoft.VisualBasic.Interaction]::InputBox("Имя нового проекта (например, префикс моделей: PROJ):", "Сохранить как новый проект", "")
    if ([string]::IsNullOrWhiteSpace($name)) { return }
    [void](Save-DistAsNewProject $name.Trim())
}

# Вопрос при уходе с несохранёнными массовыми правками. $true — можно продолжать (уже сохранено или отменено).
function Show-DistDirtyDialog {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Несохранённые изменения"
    $dlg.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $dlg.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
    $dlg.MinimizeBox = $false
    $dlg.MaximizeBox = $false
    $dlg.ShowInTaskbar = $false
    $dlg.ClientSize = New-Object System.Drawing.Size(520, 128)
    $dlg.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = "Есть несохранённые изменения. Сохранить их в текущий проект?"
    $lbl.Location = New-Object System.Drawing.Point(12, 16)
    $lbl.Size = New-Object System.Drawing.Size(496, 40)
    [void]$dlg.Controls.Add($lbl)
    $saveBtn = New-Object System.Windows.Forms.Button
    $saveBtn.Text = "Сохранить"
    $saveBtn.DialogResult = [System.Windows.Forms.DialogResult]::Yes
    $saveBtn.SetBounds(12, 78, 140, 32)
    $dropBtn = New-Object System.Windows.Forms.Button
    $dropBtn.Text = "Не сохранять"
    $dropBtn.DialogResult = [System.Windows.Forms.DialogResult]::No
    $dropBtn.SetBounds(164, 78, 160, 32)
    $cancelBtn = New-Object System.Windows.Forms.Button
    $cancelBtn.Text = "Отмена"
    $cancelBtn.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $cancelBtn.SetBounds(336, 78, 140, 32)
    [void]$dlg.Controls.Add($saveBtn)
    [void]$dlg.Controls.Add($dropBtn)
    [void]$dlg.Controls.Add($cancelBtn)
    $dlg.AcceptButton = $saveBtn
    $dlg.CancelButton = $cancelBtn
    $dr = [System.Windows.Forms.DialogResult]::Cancel
    if ($form -is [System.Windows.Forms.Form]) { $dr = $dlg.ShowDialog($form) } else { $dr = $dlg.ShowDialog() }
    if ($dr -eq [System.Windows.Forms.DialogResult]::Yes) { return "save" }
    if ($dr -eq [System.Windows.Forms.DialogResult]::No) { return "discard" }
    return "cancel"
}

function Get-DistDirtyChoice {
    $c = ""
    if ($script:DistDirtyChoice) { $c = ([string]$script:DistDirtyChoice).Trim().ToLowerInvariant() }
    if ($c -eq "save" -or $c -eq "discard" -or $c -eq "cancel") { return $c }
    return (Show-DistDirtyDialog)
}

function Confirm-DistDirtyLeave {
    if (-not $script:DistDirty) { return $true }
    switch (Get-DistDirtyChoice) {
        "save" { return (Commit-DistDirty) }
        "discard" { Restore-DistDirty; return $true }
        default { return $false }
    }
}

# Настройка проекта при запуске: по сохранённому пути (projectfile= в настройках окна). Нет файла — таблица пустая до «Загрузить…»
# Поля окна из настройки проекта: общий файл, папка моделей (RSN://), стадия, сводная, ссылка на таблицу, режим вставки, «В сводную выполнена»,
# «Выборочно». Подставляются все сразу (смена проекта не смешивает значения разных проектов); автосохранение на время подстановки выключено
function Set-ProjectFields($Cfg) {
    $script:DistLoading = $true
    try {
        $ui.DistStage.Text = $(if ($Cfg.Stage) { [string]$Cfg.Stage } else { "П" })
        $ui.DistTemplate.Text = [string]$Cfg.Template
        $ui.DistModelsDir.Text = [string]$Cfg.ModelsDir
        $script:DistRsnModels = @($Cfg.ModelsRsn)
        $script:DistModelPaths = @{}
        $ui.Summary.Text = [string]$Cfg.SummaryModel
        if ($ui.ContainsKey("Google") -and $ui.Google) { $ui.Google.Text = [string]$Cfg.GoogleLink }
        $script:SummaryDone = [bool]$Cfg.SummaryDone
        if ($ui.ContainsKey("InsertModeBox") -and $ui.InsertModeBox) { Set-InsertModeBox $ui.InsertModeBox ([string]$Cfg.InsertMode) }
        $script:AsgSel = @{ Models = $Cfg.AsgModels; Specs = $null; Pairs = $Cfg.AsgPairs }
        if ($ui.ContainsKey("AsgPart") -and $ui.AsgPart) { $ui.AsgPart.Checked = ($Cfg.AsgScope -eq "partial" -and $script:PartialUiReady) }
    }
    finally { $script:DistLoading = $false }
}

# Пустые поля — когда проекта нет (список проектов пуст)
function Clear-ProjectFields {
    Set-ProjectFields (New-ProjectConfig "П")
}

# Прежний выбор_<проект>.txt -> [inproject] (один раз; до записи — резервная копия отчеты\проект\_backup_<время>). Повторно ничего не делает.
# Ошибка переноса не мешает открыть проект: файл выбора тогда остаётся как есть. Возвращает строку для статуса (пусто — переносить было нечего).
function Invoke-DistStateMigration([string]$ReportRoot) {
    if (-not $ReportRoot) { return "" }
    $r = $null
    try { $r = Invoke-ProjectStateMigration $ReportRoot }
    catch { return ("Перенос служебных файлов не выполнен: {0}. Прежние файлы оставлены как есть." -f $_.Exception.Message) }
    $parts = New-Object System.Collections.Generic.List[string]
    foreach ($n in @($r.Notes)) { if ($n) { $parts.Add([string]$n) } }
    if (@($r.Errors).Count -gt 0) { $parts.Add(("Перенос служебных файлов не выполнен: {0}. Прежние файлы оставлены как есть." -f (@($r.Errors) -join "; "))) }
    if ($r.Backup -and $parts.Count -gt 0) { $parts.Add(("Резервная копия: {0}." -f $r.Backup)) }
    return ($parts -join " ")
}

# Одна строка о переносе — в конец статуса, один раз
function Add-DistStateMigrationNote {
    $note = [string]$script:StateMigrationNote
    $script:StateMigrationNote = ""
    if (-not $note -or $null -eq $ui -or -not $ui.ContainsKey("Status") -or -not $ui.Status) { return }
    $cur = ([string]$ui.Status.Text).Trim()
    if ($cur) { $ui.Status.Text = $cur + " " + $note } else { $ui.Status.Text = $note }
}

function Load-DistProject([string]$Path) {
    $script:Project = $null
    $script:ProjectPath = ""
    if ($Path) {
        $note = Invoke-DistStateMigration (Split-Path -Parent (Split-Path -Parent $Path))
        if ($note) { $script:StateMigrationNote = $note }
    }
    if ($Path -and (Test-Path -LiteralPath $Path)) {
        $cfg = Read-ProjectConfig $Path
        if ($null -ne $cfg) {
            $script:Project = $cfg
            $script:ProjectPath = $Path
            Set-ProjectFields $cfg
        }
    }
    else { Clear-ProjectFields }
    Show-DistGrid
    if (Get-Command Show-ProjectGoogle -ErrorAction SilentlyContinue) { Show-ProjectGoogle }      # адрес Google этого проекта на вкладке «Задание»
    $heldN = 0
    if ($script:Project -and $script:Project.HeldModels) { $heldN = @($script:Project.HeldModels).Count }
    $heldNote = Format-SkippedModelsStatus $heldN
    if ($heldNote -and $null -ne $ui -and $ui.ContainsKey("Status") -and $ui.Status) {
        $cur = [string]$ui.Status.Text
        if ($cur -notlike ("*" + $heldNote + "*")) {
            if ($cur.Trim()) { $ui.Status.Text = $cur.Trim() + " " + $heldNote } else { $ui.Status.Text = $heldNote }
        }
    }
}

# ---------- проекты: список вверху окна, «Новый проект», смена проекта ----------
function Update-ProjectList([string]$Select = "") {
    $names = @(Get-ProjectNames $ReportDir)
    $box = $ui.ProjectBox
    $script:ProjectListBusy = $true
    try {
        $box.Items.Clear()
        foreach ($n in $names) { [void]$box.Items.Add($n) }
        $want = $Select
        if (-not $want) { $want = Get-CurrentProjectName }
        $i = [Array]::IndexOf([string[]]$names, $want)
        if ($i -ge 0) { $box.SelectedIndex = $i } else { $box.SelectedIndex = -1 }
    }
    finally { $script:ProjectListBusy = $false }
}

# Смена проекта: текущий сохраняется целиком (поля окна ещё его), затем подставляется выбранный — пути, выбор «В проекте», Google, журнал, «Выборочно»
function Switch-Project([string]$Name) {
    if ($script:ProjectListBusy) { return }
    if ($Name -eq (Get-CurrentProjectName)) { return }
    if (-not (Confirm-DistDirtyLeave)) {
        if ($null -ne $ui -and $ui.ContainsKey("ProjectBox") -and $ui.ProjectBox) { Update-ProjectList }
        return
    }
    if ($script:ProjectPath) {
        if ($script:DistSaveTimer) { $script:DistSaveTimer.Stop() }
        [void](Save-DistProject)
        if ($script:PartialUiReady -and $ui.AsgTree) { Save-AsgSelection; [void](Save-DistProject) }
    }
    Load-DistProject $(if ($Name) { Get-ProjectConfigPath $ReportDir $Name } else { "" })
    Save-Settings
    if (Get-Command Reset-PartialLists -ErrorAction SilentlyContinue) { Reset-PartialLists; Refresh-PartialLists }
    Update-ProjectList
    Update-ActionState
    if ($script:ProjectPath) {
        $ui.Status.Text = ("Проект {0}." -f (Get-CurrentProjectName))
        $heldN = 0
        if ($script:Project -and $script:Project.HeldModels) { $heldN = @($script:Project.HeldModels).Count }
        $heldNote = Format-SkippedModelsStatus $heldN
        if ($heldNote) { $ui.Status.Text += " " + $heldNote }
    } else { $ui.Status.Text = "Нет проекта: нажмите «Новый проект»." }
    Add-DistStateMigrationNote
}

# «Новый проект»: запрашивает имя (обычно префикс моделей, например PROJ), создаёт пустую настройку и открывает её
function Invoke-NewProject {
    Add-Type -AssemblyName Microsoft.VisualBasic
    $name = [Microsoft.VisualBasic.Interaction]::InputBox("Имя проекта (например, префикс моделей: PROJ):", "Новый проект", "")
    if ([string]::IsNullOrWhiteSpace($name)) { return }
    New-ProjectFile $name.Trim()
}

function New-ProjectFile([string]$Name) {
    $problem = Test-NewProjectName $ReportDir $Name
    if ($problem) { [System.Windows.Forms.MessageBox]::Show($form, $problem, "Новый проект") | Out-Null; return $false }
    if (-not (Confirm-DistDirtyLeave)) { return $false }
    $cfg = New-ProjectConfig $(if ($ui.DistStage.Text.Trim()) { $ui.DistStage.Text.Trim() } else { "П" })
    $cfg.InProject = $null; $cfg.Summary = $null
    Write-ProjectConfig $cfg (Get-ProjectConfigPath $ReportDir $Name)
    Switch-Project $Name
    return $true
}

# Запуск RBP и ожидание с живым окном; возвращает @{ ExitCode; Seconds }
function Invoke-RbpAndWait([string]$TaskScript, [string[]]$Files, [string]$Text, [string]$Year, [string]$Worksets = "close_all", [string[]]$ExtraPaths = @(), [string]$FileList = "") {
    if (-not $Year) { throw "Не определена версия Revit для запуска." }
    $exe = Find-BatchRvt
    if (-not $exe) { throw "Не найден Revit Batch Processor (BatchRvt.exe). Установите RBP и откройте его один раз." }
    $batch = Join-Path (Split-Path -Parent $exe) "BatchRvt.exe"
    if (Test-Path -LiteralPath $batch) { $exe = $batch }
    if (@(@($Files) + @($ExtraPaths) | Where-Object { $_ -match '^(?i)RSN://' }).Count -gt 0) {          # RSN:// у моделей или только у общего файла
        if (-not (Enable-RsnForRbp $exe)) { throw "Поддержка RSN:// не подключилась (rbp_rsn_patch)." }
    }
    # «Посчитать» уже записал тот же список в rvt_list_run.txt (его читает insert_schedules.py). Второй файл не пишем.
    if ($FileList) { $list = $FileList }
    else {
        $list = Join-Path $AppDir "rvt_list_dist.txt"
        [System.IO.File]::WriteAllLines($list, [string[]]$Files, (New-Object System.Text.UTF8Encoding $false))
    }
    # PS 5.1 Start-Process не кавычит элементы массива с пробелами — Quote-BvsProcessArg только для путей с пробелами
    $logDir = $script:Layout.Logs
    if ($logDir -and -not (Test-Path -LiteralPath $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
    $args = @("--task_script", (Quote-BvsProcessArg $TaskScript), "--file_list", (Quote-BvsProcessArg $list), "--create_new_local", "--worksets", $Worksets, "--revit_version", $Year, "--log_folder", (Quote-BvsProcessArg $logDir))
    $live = Join-Path $AppDir "live_status.txt"
    if (Test-Path -LiteralPath $live) { Remove-Item -LiteralPath $live -Force -ErrorAction SilentlyContinue }
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $p = Start-Process -FilePath $exe -ArgumentList $args -WorkingDirectory (Split-Path -Parent $exe) -PassThru -WindowStyle Hidden
    while (-not $p.HasExited) {
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 300
        $ui.Status.Text = ("{0} Прошло {1:mm\:ss}. Окно не закрывайте." -f $Text, $sw.Elapsed)
    }
    $sw.Stop()
    return @{ ExitCode = $p.ExitCode; Seconds = $sw.Elapsed.TotalSeconds }
}

# Все файлы моделей из поля «Модели» (папка или список RSN), ещё без отбора по разделу
function Get-DistModelCandidates {
    $dir = $ui.DistModelsDir.Text.Trim()
    if (Test-RsnModelsSource $dir) { return @($script:DistRsnModels) }
    if ((Get-ModelsDirProblem $dir) -or -not (Test-Path -LiteralPath $dir)) { return @() }
    return @(Get-IosModelsFromFolder $dir)
}

# Инженерные модели: только целевой раздел (_ОВ_, _VK_ …). Сводная, общий файл и чужой код не входят.
function Get-DistModelFiles {
    return @(Select-DistModelPaths @(Get-DistModelCandidates) $ui.DistTemplate.Text.Trim() $ui.Summary.Text.Trim())
}

# Проверка папки моделей: файл вместо папки не принимается (сообщение, поле очищается)
function Confirm-DistModelsDir {
    $problem = Get-ModelsDirProblem $ui.DistModelsDir.Text
    if ($ui.DistModelsDir.Text.Trim() -and $problem) {
        [System.Windows.Forms.MessageBox]::Show($form, $problem, "Распределение") | Out-Null
        $script:DistLoading = $true
        try { $ui.DistModelsDir.Text = "" } finally { $script:DistLoading = $false }
        return $false
    }
    return $true
}

# Файлы моделей настройки проекта (источник для «Посчитать элементы» и «В разделы»): путь из папки моделей
# (кэш загрузки или повторный поиск по имени). После «Взять выбор из сводной» в настройке только связи сводной —
# модели не из этого списка в Paths не попадают. Нет файла в папке — имя в Missing.
# Возвращает @{ Paths = пути; Missing = имена моделей, файл которых не найден }
function Get-DistProjectModelPaths {
    $paths = New-Object System.Collections.Generic.List[string]
    $missing = New-Object System.Collections.Generic.List[string]
    if ($null -eq $script:Project) { return @{ Paths = @(); Missing = @() } }
    $all = $null
    foreach ($m in $script:Project.Models) {
        if (-not (Test-TargetSectionCode ([string]$m.Section))) { continue }
        if ($script:DistModelPaths.ContainsKey($m.Name)) { $paths.Add($script:DistModelPaths[$m.Name]); continue }
        if ($null -eq $all) { $all = @(Get-DistModelFiles) }
        $want = Get-LinkMatchKey $m.Name
        $found = @()
        if ($want) { $found = @($all | Where-Object { (Get-LinkMatchKey $_) -eq $want }) }
        if ($found.Count -gt 0) { $script:DistModelPaths[$m.Name] = $found[0]; $paths.Add($found[0]) } else { $missing.Add($m.Name) }
    }
    return @{ Paths = $paths.ToArray(); Missing = $missing.ToArray() }
}

# Подставляет галочки «В проекте» по списку имён из [inproject] настройки (снятого до перечитывания общего файла); возвращает текст для статуса.
# Раздела [inproject] нет ($null) — подсказка, отметки не меняются. Имена, которых нет в общем файле, — списком; в [inproject] они остаются.
function Restore-DistSelection($Names) {
    if ($null -eq $Names) { return "Выбор не найден. Отметьте спеки или нажмите " + [char]0xAB + "Взять выбор из сводной" + [char]0xBB + "." }
    $r = Set-ProjectSelectionFromNames $script:Project @($Names)
    $text = ("Выбор восстановлен из настройки проекта: {0} спек." -f $r.Marked)
    if (@($r.Unknown).Count -gt 0) { $text += (" нет в общем файле: {0}." -f (@($r.Unknown) -join "; ")) }
    return $text
}

# «Взять выбор из сводной»: один прогон Revit читает спецификации и связанные модели (сводная не меняется).
# Галочки «В проекте» — по спецификациям (список спек остаётся из общего файла). Колонки разделов — по связям.
# Связей нет или сводная их не отдала — колонки моделей остаются из папки, без ошибки. Запись файлов — по «Сохранить».
function Invoke-DistTakeFromSummary {
    if ($null -eq $script:Project -or $script:Project.Specs.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show($form, "Сначала нажмите «Загрузить спецификации и модели».", "Выбор из сводной") | Out-Null
        return
    }
    $summary = $ui.Summary.Text.Trim()
    if (-not (Test-ModelReady $summary)) {
        [System.Windows.Forms.MessageBox]::Show($form, "Укажите сводную модель.", "Выбор из сводной") | Out-Null
        return
    }
    if (Test-ProjectHasSelection $script:Project) {
        $askText = "Заменить текущий выбор выбором из сводной?`r`n`r`nКолонки моделей пересоберутся по связям сводной. Настройки (галочки и числа) моделей, которых нет среди связей, будут удалены при «Сохранить». «Отменить» вернёт всё. После чтения связей будет показано, какие модели уйдут."
        $ask = [System.Windows.Forms.MessageBox]::Show($form, $askText, "Выбор из сводной", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question, [System.Windows.Forms.MessageBoxDefaultButton]::Button2)
        if ($ask -ne [System.Windows.Forms.DialogResult]::Yes) { $ui.Status.Text = "Выбор не изменён."; return }
    }
    $ver = Resolve-RevitRun @($summary) "" (Get-BvsServerYears $ServersCfgPath)
    if ($ver.Problem) {
        [System.Windows.Forms.MessageBox]::Show($form, $ver.Problem, "Выбор из сводной") | Out-Null
        $ui.Status.Text = $ver.Problem
        return
    }
    try { [void](Invoke-BvsHousekeeping -ReportDir $ReportDir -KeepDays $script:AuxKeepDays -AppDir $AppDir) } catch { }
    $aux = New-RunAuxDir $script:Layout.Aux (Get-Date -Format "yyyyMMdd_HHmmss")
    $out = Join-Path $aux "summary_specs.tsv"
    $linksOut = Join-Path $aux "summary_links.txt"
    $cfgLines = @("[template]", ("out_path={0}" -f $out), "only_numbered=1", ("links_path={0}" -f $linksOut))
    if ($script:Layout -and $script:Layout.Logs) { $cfgLines += @("log_folder={0}" -f $script:Layout.Logs) }
    $cfg = $cfgLines -join "`r`n"
    [System.IO.File]::WriteAllText((Join-Path $AppDir "run.cfg"), $cfg + "`r`n", (New-Object System.Text.UTF8Encoding $true))
    Set-ActionsEnabled $false
    try { $res = Invoke-RbpAndWait (Join-Path $ToolDir "read_template_specs.py") @($summary) "Читаю спецификации и связи сводной модели через Revit." ([string]$ver.Year) }
    catch { Set-ActionsEnabled $true; $ui.Status.Text = $_.Exception.Message; return }
    Set-ActionsEnabled $true
    $names = @(Read-TemplateSpecsFile $out)
    if ($names.Count -eq 0) {
        $ui.Status.Text = ("Спецификации сводной не прочитаны (код RBP {0}). Выбор не изменён." -f $res.ExitCode)
        return
    }
    $linkRaw = @(Read-SummaryLinksFile $linksOut)
    $files = @()
    try { $files = @(Get-DistModelFiles) } catch { $files = @() }
    $plan = Get-SummaryLinkPlan $script:Project $linkRaw $files
    $skipModels = (@($plan.Names).Count -eq 0)
    if (-not $skipModels -and (Test-SummaryLinksThin $script:Project.Models.Count @($plan.Names).Count)) {
        $thinText = Get-DistThinLinksPrompt @($plan.Names).Count $script:Project.Models.Count $plan.Leaving
        if (-not (Confirm-DistYes $script:DistThinChoice $thinText "Выбор из сводной")) { $skipModels = $true }
    }
    elseif (-not $skipModels) {
        if (-not (Confirm-DistYes $script:DistColumnChoice (Get-DistColumnRebuildPrompt $plan.Leaving) "Выбор из сводной")) { $skipModels = $true }
    }
    Enter-DistDirty
    $r = Set-ProjectSelectionFromNames $script:Project $names
    $text = ("Отмечено по сводной: {0}; в сводной есть, но нет в общем файле: {1}" -f $r.Marked, @($r.Unknown).Count)
    if (-not $skipModels) {
        $applied = Set-ProjectModelsFromLinkNames $script:Project $linkRaw $files
        Update-AsgSelForProjectModels $script:Project
        $text += (" Моделей по связям сводной: {0}." -f @($applied.Names).Count)
        $missingNote = Format-DistMissingFiles @(Update-DistModelPathsForNames $applied.Names)
        if ($missingNote) { $text += " " + $missingNote }
    }
    elseif (@($plan.Names).Count -eq 0) { $text += " Связи сводной не найдены, колонки моделей оставлены из папки." }
    else { $text += " Колонки моделей оставлены без изменений." }
    $skipNote = Format-SkippedModelsStatus (Measure-NonTargetModels $linkRaw)
    if ($skipNote) { $text += " " + $skipNote }
    Show-DistGrid
    $ui.Status.Text = $text
}

# «Загрузить спецификации и модели»: спецификации — из общего файла (RBP), модели — из папки, только целевые разделы.
# Чужой код в таблицу не входит. Результат остаётся несохранённым, пока не нажато «Сохранить».
function Invoke-DistLoad {
    if (-not $script:ProjectPath) {
        [System.Windows.Forms.MessageBox]::Show($form, "Выберите проект в списке вверху окна или нажмите «Новый проект».", "Распределение") | Out-Null
        return
    }
    if (-not (Confirm-DistDirtyLeave)) { return }
    $template = $ui.DistTemplate.Text.Trim()
    if (-not $template -or ($template -notmatch '^(?i)RSN://' -and -not (Test-Path -LiteralPath $template -PathType Leaf))) {
        [System.Windows.Forms.MessageBox]::Show($form, "Укажите общий файл спецификаций.", "Распределение") | Out-Null
        return
    }
    $problem = Get-ModelsDirProblem $ui.DistModelsDir.Text
    if ($problem) { [System.Windows.Forms.MessageBox]::Show($form, $problem, "Распределение") | Out-Null; return }
    if (-not $ui.DistStage.Text.Trim()) { [System.Windows.Forms.MessageBox]::Show($form, "Укажите стадию (П, РД …).", "Распределение") | Out-Null; return }
    $cands = @(Get-DistModelCandidates)
    $files = @(Select-DistModelPaths $cands $template $ui.Summary.Text.Trim())
    $skippedModels = Measure-NonTargetModels $cands
    if ($files.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show($form, "В папке нет инженерных моделей: в имени файла должен быть код раздела в виде «_КОД_» (ОВ/OV, ПД/PD, ВК/VK, ПТ/PT, ОТ/OT, ХС/HS, ИТП/ITP, ЭОМ/EOM, СС/SS, АВТ/AVT, ХЦ/HC). Сводная, модели без кода и с чужим кодом не используются.", "Распределение") | Out-Null
        $skipNote = Format-SkippedModelsStatus $skippedModels
        if ($skipNote) { $ui.Status.Text = $skipNote }
        return
    }
    # версия Revit: модели одной версии, общий файл не новее моделей; общий файл читается в его собственной версии
    $ver = Resolve-RevitRun @($files) $template (Get-BvsServerYears $ServersCfgPath)
    if ($ver.Problem) {
        [System.Windows.Forms.MessageBox]::Show($form, $ver.Problem, "Распределение") | Out-Null
        $ui.Status.Text = $ver.Problem
        return
    }
    $tplYear = (Get-ModelRevitYear $template (Get-BvsServerYears $ServersCfgPath)).Year
    $runId = Get-Date -Format "yyyyMMdd_HHmmss"
    try { [void](Invoke-BvsHousekeeping -ReportDir $ReportDir -KeepDays $script:AuxKeepDays -AppDir $AppDir) } catch { }
    $aux = New-RunAuxDir $script:Layout.Aux $runId
    $out = Join-Path $aux "template_specs.tsv"
    $cfgLines = @("[template]", ("out_path={0}" -f $out), "only_numbered=1")
    if ($script:Layout -and $script:Layout.Logs) { $cfgLines += @("log_folder={0}" -f $script:Layout.Logs) }
    $cfg = $cfgLines -join "`r`n"
    [System.IO.File]::WriteAllText((Join-Path $AppDir "run.cfg"), $cfg + "`r`n", (New-Object System.Text.UTF8Encoding $true))
    Set-ActionsEnabled $false
    try {
        $res = Invoke-RbpAndWait (Join-Path $ToolDir "read_template_specs.py") @($template) "Читаю спецификации общего файла через Revit." ([string]$tplYear)
    }
    catch { Set-ActionsEnabled $true; $ui.Status.Text = $_.Exception.Message; return }
    Set-ActionsEnabled $true
    $names = @(Read-TemplateSpecsFile $out)
    if ($names.Count -eq 0) {
        $ui.Status.Text = ("Спецификации общего файла не прочитаны (код RBP {0}). Подробности: отчеты\логи\read_template_specs_error.txt и журнал BatchRvt в отчеты\логи." -f $res.ExitCode)
        return
    }
    # результат загрузки — правка в памяти: снимок до неё, файл проекта не пишется
    $path = $script:ProjectPath
    if ($null -eq $script:Project) { $script:Project = New-ProjectConfig $ui.DistStage.Text.Trim() }
    Enter-DistDirty
    $cfgObj = $script:Project
    $prevIn = Get-ProjectInProjectNames $cfgObj                    # «В проекте» из [inproject] — до перечитывания списка спецификаций
    Update-ProjectSpecs $cfgObj $names
    $cfgObj.Categories = Read-TemplateSpecCategories $out          # категории из общего файла — кэш для «Распределить»
    $script:SummaryDone = $false                                   # спецификации перечитаны — сбросить запомненный статус «В сводную»; «В разделы» от этого не зависит. В файл попадёт с «Сохранить»
    $cfgObj.SummaryDone = $false
    $cfgObj.Models.Clear()
    Clear-ProjectHeld $cfgObj                                      # чужие модели уйдут из файла только когда эту загрузку сохранят
    $script:DistModelPaths = @{}
    foreach ($f in $files) { $script:DistModelPaths[(Get-ProjectModelName $f)] = $f }
    Update-ProjectModels $cfgObj $files
    $script:ProjectPath = $path
    # галочки «В проекте» — из [inproject] настройки проекта (файл при загрузке не пишется)
    $selText = Restore-DistSelection $prevIn
    [void](Update-SectionCatalog (Get-SectionCatalogPath $ReportDir))
    Show-DistGrid
    if (Get-Command Show-ProjectGoogle -ErrorAction SilentlyContinue) { Show-ProjectGoogle }
    $noCat = @($names | Where-Object { -not (Get-ProjectSpecCategory $cfgObj $_) }).Count
    $ui.Status.Text = ("Загружено: спецификаций {0} (без категории {1}), моделей {2}. {3}" -f $names.Count, $noCat, $files.Count, $selText)
    $skipNote = Format-SkippedModelsStatus $skippedModels
    if ($skipNote) { $ui.Status.Text += " " + $skipNote }
}

# «Распределить»: галочки по справочнику разделов (отчеты\проект\справочник_разделов.txt). Модели НЕ открываются, Revit не запускается:
# спека отмечена в модели, если она «В проекте» и её номер или категория (категории читаются из общего файла при «Загрузить…») входят в раздел модели.
function Invoke-DistAssign {
    if ($null -eq $script:Project -or $script:Project.Specs.Count -eq 0 -or $script:Project.Models.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show($form, "Сначала нажмите «Загрузить спецификации и модели».", "Распределение") | Out-Null
        return
    }
    if ($script:Project.Categories.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show($form, "Категории спецификаций не прочитаны. Нажмите «Загрузить спецификации и модели».", "Распределение") | Out-Null
        return
    }
    if (@($script:Project.Specs | Where-Object { Test-ProjectSpecIn $script:Project $_ }).Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show($form, "Отметьте в столбце «В проекте» спецификации, которые участвуют в проекте.", "Распределение") | Out-Null
        return
    }
    $catalog = Update-SectionCatalog (Get-SectionCatalogPath $ReportDir)
    if (-not $catalog.Ok) {
        Show-DistGrid
        [System.Windows.Forms.MessageBox]::Show($form, $catalog.Error, "Справочник разделов") | Out-Null
        $ui.Status.Text = $catalog.Error
        return
    }
    Enter-DistDirty
    $r = Set-ProjectDistribution $script:Project $catalog
    Show-DistGrid
    $text = ("Распределено: отметок {0}, моделей {1}." -f $r.Ticked, $r.Models)
    if (@($r.NoCategory).Count -gt 0) { $text += (" Без категории (не отмечены): {0}." -f @($r.NoCategory).Count) }
    if (@($r.NoSection).Count -gt 0) { $text += (" Категория не входит ни в один раздел (не отмечены): {0}." -f @($r.NoSection).Count) }
    if (@($r.UnknownSections).Count -gt 0) { $text += (" Раздела нет в справочнике: {0}." -f (@($r.UnknownSections | ForEach-Object { Get-ProjectSectionTitle $script:Project $_ }) -join ", ")) }
    $ui.Status.Text = $text
}

# «Посчитать элементы»: пробная вставка спецификаций «В проекте» в каждую инженерную модель (фон RBP, одно открытие и одна пачка на модель), подсчёт, откат.
# Результат — числа в ячейках для информации; галочки не меняются.
function Invoke-DistCount {
    if ($null -eq $script:Project -or $script:Project.Specs.Count -eq 0 -or $script:Project.Models.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show($form, "Сначала нажмите «Загрузить спецификации и модели».", "Распределение") | Out-Null
        return
    }
    # скрипт читает файл проекта: сначала решить судьбу несохранённых правок, потом брать пути и галочки
    if (-not (Confirm-DistDirtyLeave)) { return }
    if (@($script:Project.Specs | Where-Object { Test-ProjectSpecIn $script:Project $_ }).Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show($form, "Отметьте в столбце «В проекте» группы спецификаций, которые участвуют в проекте.", "Распределение") | Out-Null
        return
    }
    $template = $ui.DistTemplate.Text.Trim()
    $mp = Get-DistProjectModelPaths
    $paths = New-Object System.Collections.Generic.List[string]
    foreach ($x in $mp.Paths) { $paths.Add([string]$x) }
    $missing = New-Object System.Collections.Generic.List[string]
    foreach ($x in $mp.Missing) { $missing.Add([string]$x) }
    if ($paths.Count -eq 0) {
        $ui.Status.Text = "Файлы моделей не найдены: проверьте поле «Модели»."
        $missingNote = Format-DistMissingFiles $missing
        if ($missingNote) { $ui.Status.Text += " " + $missingNote }
        return
    }
    [void](Save-DistProject)          # в файле уже подтверждённое состояние; спрятанные чужие модели не выкидываются
    $blocker = Get-ProjectRunBlocker $script:ProjectPath
    if ($blocker) { [System.Windows.Forms.MessageBox]::Show($form, $blocker, "Распределение") | Out-Null; return }
    $ver = Resolve-RevitRun @($paths) $template (Get-BvsServerYears $ServersCfgPath)
    if ($ver.Problem) {
        [System.Windows.Forms.MessageBox]::Show($form, $ver.Problem, "Распределение") | Out-Null
        $ui.Status.Text = $ver.Problem
        return
    }
    $runId = Get-Date -Format "yyyyMMdd_HHmmss"
    try { [void](Invoke-BvsHousekeeping -ReportDir $ReportDir -KeepDays $script:AuxKeepDays -AppDir $AppDir) } catch { }
    $aux = New-RunAuxDir $script:Layout.Aux $runId
    $out = Join-Path $aux "prefill.tsv"
    $files = Write-RunFiles -Models @($paths) -Source $template -ReportDir $aux -RunId $runId -Mode "prefill" -ProjectPath $script:ProjectPath -PrefixFrom "" -PrefixTo "" -InsertMode "initial" -Extra @(("prefill_out={0}" -f $out))
    Set-ActionsEnabled $false
    $res = $null
    try {
        $res = Invoke-RbpAndWait (Join-Path $ToolDir "insert_schedules.py") @($paths) ("Считаю элементы спецификаций в моделях ({0}), модели не сохраняются." -f $paths.Count) ([string]$ver.Year) "open_all" @($template) -FileList $files.List          # open_all: подсчёт элементов — по всем рабочим наборам, как в «Задании»; список один — rvt_list_run.txt
    }
    catch { Set-ActionsEnabled $true; $ui.Status.Text = $_.Exception.Message; return }
    Set-ActionsEnabled $true
    $result = Read-PrefillResults $out
    if (@($result.Counts).Count -eq 0) {
        $why = ""
        if (@($result.Errors).Count -gt 0) { $why = " Причина: " + ($result.Errors -join "; ") }
        $ui.Status.Text = ("Подсчёт ничего не вернул (код RBP {0}).{1} Журнал: {2}" -f $res.ExitCode, $why, $aux)
        return
    }
    $text = Merge-PrefillIntoProject $script:Project $result $res.Seconds
    if ($script:DistDirty -and $null -ne $script:DistDirtySnapshot) { [void](Merge-PrefillIntoProject $script:DistDirtySnapshot $result $res.Seconds) }
    $missingNote = Format-DistMissingFiles $missing
    if ($missingNote) { $text += " " + $missingNote }
    if (@($result.Errors).Count -gt 0) { $text += (" Ошибки: {0}." -f ($result.Errors -join "; ")) }
    if (-not $script:DistDirty) { [void](Save-DistProject) }   # числа — служебная запись; несохранённые правки этим сохранением не пишутся
    try {
        $logLines = New-Object System.Collections.Generic.List[string]
        foreach ($x in @(Get-DistLogLines $result $res.Seconds)) { $logLines.Add([string]$x) }
        if ($missingNote) { $logLines.Add($missingNote.Trim()) }
        [System.IO.File]::WriteAllLines((Join-Path $aux "распределение.log"), $logLines.ToArray(), (New-Object System.Text.UTF8Encoding $true))
    } catch { }
    Show-DistGrid
    $ui.Status.Text = $text
}

# Щелчок по таблице: название группы — раскрыть/свернуть; «В проекте», «Сводная» и ячейки — отметить/снять. В файл — только по «Сохранить».
function Invoke-DistCellClick([int]$RowIndex, [int]$ColIndex) {
    if ($script:DistBusy -or $null -eq $script:Project -or $RowIndex -lt 0 -or $ColIndex -lt 0) { return }
    $g = $ui.DistGrid
    $row = $g.Rows[$RowIndex].Tag
    $col = $g.Columns[$ColIndex].Tag
    if ($null -eq $row -or $null -eq $col) { return }
    if ($col.Kind -eq "name") { return }          # раскрытие — значок слева или двойной щелчок по строке (Invoke-DistToggleGroup)
    if ($col.Kind -eq "sep") { return }
    if (($col.Kind -eq "section" -or $col.Kind -eq "model" -or $col.Kind -eq "summary") -and @($row.Active).Count -eq 0) {
        $ui.Status.Text = "Сначала отметьте «В проекте»."
        return
    }
    Enter-DistDirty
    Invoke-ProjectGridClick $script:Project $row $col
    Update-DistGridRows $row.Group
}

# «Выделить всё» / «Снять всё»: видимые группы (с учётом «Скрыть не участвующие»).
# Правка остаётся в памяти; «В проекте» ([inproject]) и настройка пишутся только по «Сохранить».
function Invoke-DistBulk([bool]$On) {
    if ($null -eq $script:Project) { return }
    Enter-DistDirty
    $specs = New-Object System.Collections.Generic.List[string]
    foreach ($grp in @(Get-ProjectGroups $script:Project)) {
        if ($script:DistExpand.Hide -and (Get-ProjectGroupInState $script:Project $grp.Specs) -eq "Unchecked") { continue }
        foreach ($sp in $grp.Specs) { [void]$specs.Add([string]$sp) }
    }
    Set-ProjectInProjectBulk $script:Project $specs.ToArray() $On
    Show-DistGrid
}

# Раскрыть / свернуть группу из нескольких спецификаций: щелчок по значку слева от названия или двойной щелчок по названию
function Invoke-DistToggleGroup([int]$RowIndex) {
    if ($script:DistBusy -or $null -eq $script:Project -or $RowIndex -lt 0) { return }
    $row = $ui.DistGrid.Rows[$RowIndex].Tag
    if ($null -eq $row -or $row.Kind -ne "group" -or -not $row.Expandable) { return }
    $script:DistExpand.Groups[$row.Group] = -not ($script:DistExpand.Groups.ContainsKey($row.Group) -and $script:DistExpand.Groups[$row.Group])
    Show-DistGrid
}

function Invoke-DistHeaderClick([int]$ColIndex) {
    if ($null -eq $script:Project -or $ColIndex -lt 0) { return }
    $col = $ui.DistGrid.Columns[$ColIndex].Tag
    if ($null -eq $col -or $col.Kind -ne "section") { return }
    $script:DistExpand.Sections[$col.Section] = -not ($script:DistExpand.Sections.ContainsKey($col.Section) -and $script:DistExpand.Sections[$col.Section])
    Show-DistGrid
}

function Invoke-DistRowRightClick([int]$RowIndex) {
    if ($null -eq $script:Project -or $RowIndex -lt 0) { return }
    $row = $ui.DistGrid.Rows[$RowIndex].Tag
    if ($null -eq $row -or $row.Kind -ne "spec") { return }
    Add-Type -AssemblyName Microsoft.VisualBasic
    $answer = [Microsoft.VisualBasic.Interaction]::InputBox(("Группа для спецификации «{0}» (пусто — по названию):" -f $row.Spec), "Группа спецификации", [string]$row.Group)
    Enter-DistDirty
    Set-ProjectSpecGroup $script:Project $row.Spec $answer
    Show-DistGrid
}

# Строка пути: подпись | поле | кнопки (до двух). Возвращает поле ввода
function Add-PathRow($Table, [int]$Row, [string]$Caption, $Box, $Buttons) {
    $l = New-Object System.Windows.Forms.Label
    $l.Text = $Caption; $l.AutoSize = $true; $l.Anchor = "Left"; $l.Margin = New-Object System.Windows.Forms.Padding(3, 0, 8, 0)
    $Box.Anchor = "Left, Right"; $Box.Margin = New-Object System.Windows.Forms.Padding(3, 4, 6, 4)
    $Table.Controls.Add($l, 0, $Row); $Table.Controls.Add($Box, 1, $Row)
    $col = 2
    foreach ($b in @($Buttons)) {
        $b.Size = New-Object System.Drawing.Size(104, 28); $b.Anchor = "Left, Right"; $b.Margin = New-Object System.Windows.Forms.Padding(3, 2, 3, 2)
        $Table.Controls.Add($b, $col, $Row); $col++
    }
}

# Обычная кнопка Windows: одна высота у всех, ширина по тексту
function New-DistButton([string]$Text, [int]$Width = 0) {
    $b = New-Object System.Windows.Forms.Button
    $b.Text = $Text
    $b.FlatStyle = [System.Windows.Forms.FlatStyle]::Standard
    $b.UseVisualStyleBackColor = $true
    if ($Width -gt 0) { $b.Size = New-Object System.Drawing.Size($Width, 30) }
    else { $b.AutoSize = $true; $b.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowOnly; $b.MinimumSize = New-Object System.Drawing.Size(0, 30); $b.Padding = New-Object System.Windows.Forms.Padding(6, 0, 6, 0) }
    $b.Margin = New-Object System.Windows.Forms.Padding(3)
    return $b
}

# Вкладка «Спецификации», сверху вниз: общий файл, модели, сводная модель, стадия, управление таблицей, таблица, кнопки шагов
function New-DistPage {
    $page = New-Object System.Windows.Forms.Panel
    $page.Dock = "Fill"
    $page.BackColor = [System.Drawing.SystemColors]::Window
    $t = New-LayoutTable
    Add-LayoutColumn $t "percent" 100
    # 1) пути: общий файл [Обзор…]; модели [Обзор…][С сервера…]; сводная [Обзор…][С сервера…]
    $r = Add-LayoutRow $t
    $paths = New-Object System.Windows.Forms.TableLayoutPanel
    $paths.Dock = "Fill"; $paths.AutoSize = $true; $paths.ColumnCount = 4; $paths.RowCount = 3
    [void]$paths.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
    [void]$paths.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    [void]$paths.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 112)))
    [void]$paths.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 112)))
    for ($i = 0; $i -lt 3; $i++) { [void]$paths.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize))) }
    $ui.DistTemplate = New-Object System.Windows.Forms.TextBox
    $ui.Source = $ui.DistTemplate                         # одно поле: общий файл (source= в настройках окна)
    $ui.DistModelsDir = New-Object System.Windows.Forms.TextBox
    $ui.Summary = New-Object System.Windows.Forms.TextBox
    $bTpl = New-Object System.Windows.Forms.Button; $bTpl.Text = "Обзор…"
    $bDir = New-Object System.Windows.Forms.Button; $bDir.Text = "Обзор…"
    $bDirSrv = New-Object System.Windows.Forms.Button; $bDirSrv.Text = "С сервера…"
    $bSum = New-Object System.Windows.Forms.Button; $bSum.Text = "Обзор…"
    $bSumSrv = New-Object System.Windows.Forms.Button; $bSumSrv.Text = "С сервера…"
    Add-PathRow $paths 0 "Общий файл" $ui.DistTemplate @($bTpl)
    Add-PathRow $paths 1 "Модели" $ui.DistModelsDir @($bDir, $bDirSrv)
    Add-PathRow $paths 2 "Сводная модель" $ui.Summary @($bSum, $bSumSrv)
    Set-UiTip $ui.DistTemplate "Файл Revit со всеми спецификациями (хранится в ранней версии, «только чтение»). Можно путь RSN://."
    Set-UiTip $ui.DistModelsDir "Папка с моделями разделов (внутри папки RVT) или папка RSN:// после «С сервера…». Берутся модели с кодом раздела в имени (_OV_, _ВК_ …)."
    Set-UiTip $ui.Summary "Сводная модель (SVOD/СВОД). Можно путь RSN://."
    Set-UiTip $bDirSrv "Выбрать модели с Revit Server."
    Set-UiTip $bSumSrv "Выбрать сводную модель с Revit Server."
    Add-LayoutControl $t $paths 0 $r
    $bTpl.Add_Click({
        $dlg = New-Object System.Windows.Forms.OpenFileDialog
        $dlg.Filter = "Revit (*.rvt)|*.rvt"; $dlg.Title = "Общий файл спецификаций"
        $start = Get-DialogStartDir $ui.DistTemplate.Text; if ($start) { $dlg.InitialDirectory = $start }
        if ($dlg.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::OK) { $ui.DistTemplate.Text = $dlg.FileName }
    })
    # папка — только Show-FolderPathDialog (адресная строка, серверные пути); системный диалог выбора папки не используется
    $bDir.Add_Click({
        $picked = Show-FolderPathDialog -StartPath $ui.DistModelsDir.Text
        if ($picked) { $script:DistRsnModels = @(); $ui.DistModelsDir.Text = $picked }
    })
    $bDirSrv.Add_Click({
        $picked = @(Get-ServerModels)
        if ($picked.Count -eq 0) { return }
        $script:DistRsnModels = $picked
        $label = Get-RsnSourceLabel $picked
        if ($ui.DistModelsDir.Text -eq $label) { Request-DistSave } else { $ui.DistModelsDir.Text = $label }
        $n = @(Get-DistModelFiles).Count
        $ui.Status.Text = ("С сервера выбрано моделей: {0}, инженерных (по коду раздела): {1}." -f $picked.Count, $n)
    })
    $bSum.Add_Click({
        $dlg = New-Object System.Windows.Forms.OpenFileDialog
        $dlg.Filter = "Revit (*.rvt)|*.rvt"; $dlg.Title = "Сводная модель"
        $start = Get-DialogStartDir $ui.Summary.Text; if ($start) { $dlg.InitialDirectory = $start }
        if ($dlg.ShowDialog($form) -eq [System.Windows.Forms.DialogResult]::OK) { $ui.Summary.Text = $dlg.FileName }
    })
    $bSumSrv.Add_Click({ Set-OneServerModel $ui.Summary })
    $ui.DistTemplate.Add_TextChanged({ Request-DistSave })
    $ui.Summary.Add_TextChanged({ Update-ActionState; Request-DistSave })
    $ui.DistModelsDir.Add_TextChanged({ if (-not (Get-ModelsDirProblem $ui.DistModelsDir.Text)) { Request-DistSave } })
    $ui.DistModelsDir.Add_Leave({ if (-not $script:DistLoading) { [void](Confirm-DistModelsDir) } })
    # 2) стадия
    $r = Add-LayoutRow $t
    $flow = New-Object System.Windows.Forms.FlowLayoutPanel
    $flow.AutoSize = $true; $flow.Dock = "Fill"; $flow.WrapContents = $true
    $lbl = New-Object System.Windows.Forms.Label; $lbl.Text = "Стадия"; $lbl.AutoSize = $true; $lbl.Margin = New-Object System.Windows.Forms.Padding(3, 8, 8, 0)
    $ui.DistStage = New-Object System.Windows.Forms.ComboBox
    $ui.DistStage.DropDownStyle = "DropDown"; $ui.DistStage.Width = 70
    [void]$ui.DistStage.Items.Add("П"); [void]$ui.DistStage.Items.Add("РД")
    $ui.DistStage.Text = "П"
    $ui.DistStage.Add_TextChanged({ Request-DistSave })
    Set-UiTip $ui.DistStage "В фильтрах {СТАДИЯ} подставляется при вставке; имя вида = стадия + 8 пробелов + название."
    $flow.Controls.AddRange(@($lbl, $ui.DistStage))
    Add-LayoutControl $t $flow 0 $r
    # 3) управление таблицей
    $r = Add-LayoutRow $t
    $flow3 = New-Object System.Windows.Forms.FlowLayoutPanel
    $flow3.AutoSize = $true; $flow3.Dock = "Fill"; $flow3.WrapContents = $true
    $ui.DistHide = New-Object System.Windows.Forms.CheckBox; $ui.DistHide.Text = "Скрыть неучаствующие"; $ui.DistHide.AutoSize = $true; $ui.DistHide.Margin = New-Object System.Windows.Forms.Padding(3, 8, 24, 0)
    $ui.DistHide.Add_CheckedChanged({ $script:DistExpand.Hide = [bool]$ui.DistHide.Checked; Show-DistGrid })
    $ui.DistAll = New-DistButton "Выделить всё" 110
    $ui.DistNone = New-DistButton "Снять всё" 110
    $ui.DistFromSummary = New-DistButton "Взять выбор из сводной"
    $ui.DistFromSummary.Add_Click({ Invoke-DistTakeFromSummary })
    Set-UiTip $ui.DistFromSummary "Один раз прочитать в сводной модели (Revit) спецификации и связанные модели: отметить «В проекте» и построить колонки разделов по связям. Список спецификаций остаётся из общего файла. Файлы проекта не меняются, пока не нажато «Сохранить»."
    $ui.DistAll.Add_Click({ Invoke-DistBulk $true })
    $ui.DistNone.Add_Click({ Invoke-DistBulk $false })
    Set-UiTip $ui.DistAll "Отметить «В проекте» у всех видимых групп. Запись в файлы — кнопкой «Сохранить»."
    Set-UiTip $ui.DistNone "Снять «В проекте» у всех видимых групп. Запись в файлы — кнопкой «Сохранить»."
    $ui.DistDirtyNote = New-Object System.Windows.Forms.Label
    $ui.DistDirtyNote.Text = "Есть несохранённые изменения"
    $ui.DistDirtyNote.AutoSize = $true
    $ui.DistDirtyNote.Visible = $false
    $ui.DistDirtyNote.ForeColor = [System.Drawing.Color]::FromArgb(160, 60, 0)
    $ui.DistDirtyNote.Margin = New-Object System.Windows.Forms.Padding(16, 8, 8, 0)
    $ui.DistCommit = New-DistButton "Сохранить"
    $ui.DistSaveAs = New-DistButton "Сохранить как новый проект"
    $ui.DistRevert = New-DistButton "Отменить"
    $ui.DistCommit.Visible = $false
    $ui.DistSaveAs.Visible = $false
    $ui.DistRevert.Visible = $false
    $ui.DistCommit.Add_Click({ Invoke-DistCommit })
    $ui.DistSaveAs.Add_Click({ Invoke-DistSaveAs })
    $ui.DistRevert.Add_Click({ Invoke-DistRevert })
    Set-UiTip $ui.DistCommit "Записать распределение и выбор «В проекте» в файлы открытого проекта."
    Set-UiTip $ui.DistSaveAs "Записать текущее распределение в новый проект. Файлы открытого проекта не меняются."
    Set-UiTip $ui.DistRevert "Вернуть распределение к состоянию до первой несохранённой правки."
    $flow3.Controls.AddRange(@($ui.DistHide, $ui.DistAll, $ui.DistNone, $ui.DistFromSummary))
    $flow3.Controls.AddRange(@($ui.DistDirtyNote, $ui.DistCommit, $ui.DistSaveAs, $ui.DistRevert))
    Add-LayoutControl $t $flow3 0 $r
    # 4) таблица
    $r = Add-LayoutRow $t $true
    $g = New-Object System.Windows.Forms.DataGridView
    # двойная буферизация: без неё каждая перерисовка ячеек мерцает и тормозит
    try { [System.Windows.Forms.DataGridView].GetProperty("DoubleBuffered", [System.Reflection.BindingFlags]"Instance,NonPublic").SetValue($g, $true, $null) } catch { }
    $g.Dock = "Fill"; $g.ReadOnly = $true; $g.AllowUserToAddRows = $false; $g.AllowUserToDeleteRows = $false; $g.AllowUserToResizeRows = $false
    $g.RowHeadersVisible = $false; $g.SelectionMode = "CellSelect"; $g.MultiSelect = $false; $g.BackgroundColor = [System.Drawing.Color]::White
    $g.ColumnHeadersHeightSizeMode = "AutoSize"; $g.EnableHeadersVisualStyles = $false
    $g.ColumnHeadersDefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(232, 237, 245)
    $g.ColumnHeadersDefaultCellStyle.Alignment = "MiddleCenter"
    $g.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::Black
    $g.AutoSizeColumnsMode = "None"
    # внешний вид: тонкая светлая сетка одного цвета (как в Excel), без объёмных границ заголовков; шрифт Segoe UI 9 без жирного
    $grid = [System.Drawing.Color]::FromArgb(218, 220, 224)
    $g.GridColor = $grid
    $g.BorderStyle = [System.Windows.Forms.BorderStyle]::None
    $g.CellBorderStyle = [System.Windows.Forms.DataGridViewCellBorderStyle]::Single
    $g.ColumnHeadersBorderStyle = [System.Windows.Forms.DataGridViewHeaderBorderStyle]::Single
    $g.AdvancedCellBorderStyle.All = [System.Windows.Forms.DataGridViewAdvancedCellBorderStyle]::Single
    $g.AdvancedColumnHeadersBorderStyle.All = [System.Windows.Forms.DataGridViewAdvancedCellBorderStyle]::Single
    $plain = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Regular)
    $g.Font = $plain
    $g.DefaultCellStyle.Font = $plain
    $g.ColumnHeadersDefaultCellStyle.Font = $plain
    $g.RowTemplate.Height = 22
    $ui.DistGrid = $g
    Set-UiTip $g "Цвета: зелёный — авто, жёлтый — вручную. Оранжевое название — спецификация без категории / вне справочника разделов (не отмечается). Правый щелчок по названию — другая группа. Двойной щелчок — раскрыть группу."
    $g.Add_CellClick({ param($s, $e) Invoke-DistCellClick $e.RowIndex $e.ColumnIndex })
    $g.Add_ColumnHeaderMouseClick({ param($s, $e) if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Left) { Invoke-DistHeaderClick $e.ColumnIndex } })
    $g.Add_CellMouseClick({
        param($s, $e)
        if ($e.ColumnIndex -ne 0) { return }
        if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Right) { Invoke-DistRowRightClick $e.RowIndex; return }
        if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Left -and $e.Clicks -eq 1 -and $e.X -lt 24) { Invoke-DistToggleGroup $e.RowIndex }
    })
    $g.Add_CellMouseDoubleClick({
        param($s, $e)
        if ($e.ColumnIndex -eq 0 -and $e.Button -eq [System.Windows.Forms.MouseButtons]::Left -and $e.X -ge 24) { Invoke-DistToggleGroup $e.RowIndex }
    })
    $g.Add_CellPainting({ param($s, $e) Invoke-DistCellPainting $e })
    $g.Add_Resize({ Update-DistColumnWidths; Update-DistHintPosition })
    $g.ScrollBars = [System.Windows.Forms.ScrollBars]::Both
    $g.AllowUserToResizeColumns = $false
    $hint = New-Object System.Windows.Forms.Label
    $hint.AutoSize = $true
    $hint.Text = "Укажите общий файл и нажмите " + [char]0x201E + "Загрузить спецификации и модели" + [char]0x201C
    $hint.BackColor = [System.Drawing.Color]::FromArgb(243, 243, 243)
    $hint.ForeColor = [System.Drawing.SystemColors]::GrayText
    $hint.Padding = New-Object System.Windows.Forms.Padding(12, 8, 12, 8)
    $ui.DistHint = $hint
    $g.Controls.Add($hint)
    Add-LayoutControl $t $g 0 $r
    # 5) кнопки шагов слева направо; «Посчитать элементы» — отдельно, справа
    $r = Add-LayoutRow $t
    $bar = New-Object System.Windows.Forms.TableLayoutPanel
    $bar.Dock = "Fill"; $bar.AutoSize = $true; $bar.ColumnCount = 2; $bar.RowCount = 1
    [void]$bar.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    [void]$bar.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
    [void]$bar.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    $ribbon = New-Object System.Windows.Forms.FlowLayoutPanel
    $ribbon.AutoSize = $true; $ribbon.Dock = "Fill"; $ribbon.WrapContents = $true; $ribbon.Margin = New-Object System.Windows.Forms.Padding(0)
    $ui.DistLoad = New-DistButton "Загрузить спецификации и модели"
    $ui.DistAssign = New-DistButton "Распределить"
    $ui.RunSummary = New-DistButton "В сводную"
    $ui.RunSections = New-DistButton "В разделы"
    $ui.RunClassifier = New-DistButton "Классификатор"
    $ui.DistCount = New-DistButton "Посчитать элементы"
    $ui.DistLoad.Add_Click({ Invoke-DistLoad })
    $ui.DistAssign.Add_Click({ Invoke-DistAssign })
    $ui.RunSummary.Add_Click({ Start-Batch -Mode "summary" })
    $ui.RunSections.Add_Click({ Start-Batch -Mode "sections" })
    $ui.RunClassifier.Add_Click({ Start-Batch -Mode "classifier" })
    $ui.DistCount.Add_Click({ Invoke-DistCount })
    Set-UiTip $ui.DistLoad "Прочитать спецификации и их категории из общего файла (Revit) и найти модели целевых разделов. В файлы попадёт по кнопке «Сохранить»."
    Set-UiTip $ui.DistAssign "Отметить спецификации «В проекте» в моделях по справочнику разделов (отчеты\проект\справочник_разделов.txt). Модели не открываются. Файлы проекта не меняются, пока не нажато «Сохранить»."
    Set-UiTip $ui.RunSummary "Вставить в сводную спецификации, отмеченные в «В проекте» и «Сводная»: со своим фильтром, с учётом связанных моделей."
    Set-UiTip $ui.RunSections "Вставить в модели разделов спецификации по распределению."
    Set-UiTip $ui.RunClassifier "Абсолютный UNC кода по классификатору в моделях разделов (тот же список, что у «В разделы»). Сводная не открывается. Если путь уже такой — модель не сохраняется."
    Set-UiTip $ui.DistCount "Пробная вставка в модели (Revit, модели не сохраняются): число элементов в ячейках для информации, галочки не меняются."
    $combo = New-InsertModeBox
    $ui.InsertModeBox = $combo
    $combo.Margin = New-Object System.Windows.Forms.Padding(3, 5, 3, 3)
    $combo.Add_SelectedIndexChanged({ if (-not $script:DistLoading -and -not $script:DistRestoring) { Enter-DistDirty } })
    $ribbon.Controls.AddRange(@($ui.DistLoad, $ui.DistAssign, $ui.RunSummary, $ui.RunSections, $ui.RunClassifier, $combo))
    $ui.DistCount.Anchor = "Right"
    $bar.Controls.Add($ribbon, 0, 0)
    $bar.Controls.Add($ui.DistCount, 1, 0)
    Add-LayoutControl $t $bar 0 $r
    [void]$page.Controls.Add($t)
    return $page
}
