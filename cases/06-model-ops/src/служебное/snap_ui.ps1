# Name: snap_ui.ps1
# Version: 3.1.0
# What it does: See the case README.
# Notes: Paths and project names in this portfolio copy are placeholders.

# UTF-8 WITH BOM required (PowerShell 5.1)
# Снимки окна операций и проверка раскладки. Окно строится, но диалог не показывается
# пользователю: форма рисуется за пределами экрана, Revit / RBP не запускаются,
# в папку инструмента ничего не пишется.
#
#   powershell -NoProfile -ExecutionPolicy Bypass -File snap_ui.ps1 -OutDir $env:TEMP\bmo_snaps
#
# Для каждой вкладки (и вложенных вкладок 4 и 5) при минимальном размере окна и при 1920x1040:
#   - PNG в -OutDir (01_perenos_min.png ...);
#   - проверка: видимые соседние элементы не перекрываются, каждый лежит внутри
#     ClientRectangle родителя, нет нулевых размеров, текст кнопок и надписей не обрезан.
# Нарушения печатаются; код выхода 1, если они есть, иначе 0.
#   -Demo           подставить пример моделей вместо rvt_list.txt
#   -CaptureScreen  снимать экран (форма в точке 0,0), а не DrawToBitmap
#   -NoImages       только проверка, без PNG
param(
    [string]$OutDir = "",
    [switch]$Demo,
    [switch]$CaptureScreen,
    [switch]$NoImages
)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$snapCodeDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$snapToolDir = Split-Path -Parent $snapCodeDir
if ([string]::IsNullOrWhiteSpace($OutDir)) { $OutDir = Join-Path ([System.IO.Path]::GetTempPath()) "bmo_snaps" }
$OutDir = [System.IO.Path]::GetFullPath($OutDir)
$toolFull = [System.IO.Path]::GetFullPath($snapToolDir).TrimEnd("\", "/")
if ($OutDir.StartsWith($toolFull, [StringComparison]::OrdinalIgnoreCase)) {
    Write-Output ("OutDir must be outside the tool folder: {0}" -f $OutDir)
    exit 2
}
if (-not $NoImages -and -not (Test-Path -LiteralPath $OutDir)) { [void](New-Item -ItemType Directory -Path $OutDir -Force) }

$script:OpsBuildOnly = $true
if ($Demo) {
    $script:OpsDemoPaths = @(
        "\\server\share\Sample\АР\PROJ_К1_АР_R22.rvt",
        "\\server\share\Sample\АР\PROJ_К2_АР_R22.rvt",
        "\\server\share\Sample\КР\PROJ_К1_КР_R22.rvt",
        "\\server\share\Sample\ОВ\PROJ_К1_ОВ_R22.rvt"
    )
}
$ErrorActionPreference = "Continue"
. (Join-Path $snapCodeDir "presets.ps1")
$ErrorActionPreference = "Continue"

$form = $script:Ui.Form
if ($Demo) {
    [void](Add-PlainTabModels "relink" $script:OpsDemoPaths)
    [void](Add-PlainTabModels "upgrade" $script:OpsDemoPaths)
    [void](Add-PlainTabModels "addlinks" $script:OpsDemoPaths)
    # tab 4: ОВ is the new model, so the «Связи» table shows pairs (К1) and a row without a pair (К2)
    Set-WorksetRole @($script:OpsDemoPaths[3]) "new"
    Update-WorksetRoleColumn
    Update-WorksetViews
    [void](Add-PlainTabModels "worksets" $script:OpsDemoPaths)
    $script:Ui.RuleNew.Text = "PROJ_К1_АР_R24"
    $script:Ui.FolderNew.Text = "\\server\share\Sample 2024"
    $script:Ui.RelOldBox.Text = "\\server\share\Sample"
    $script:Ui.RelNewBox.Text = "\\server\share\Sample 2024"
}

$form.StartPosition = "Manual"
$form.ShowInTaskbar = $false
if ($CaptureScreen) { $form.Location = New-Object System.Drawing.Point(0, 0) }
else { $form.Location = New-Object System.Drawing.Point(-30000, -30000) }
$form.Show()

function Wait-SnapUi([int]$Rounds = 6) {
    for ($i = 0; $i -lt $Rounds; $i++) {
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 40
    }
}
if ($script:Ui.FieldTimer) { $script:Ui.FieldTimer.Stop() }
Update-MoveFields
Wait-SnapUi 10

function Get-SnapName($C) {
    $t = ""
    try { $t = [string]$C.Text } catch { }
    if ($t.Length -gt 40) { $t = $t.Substring(0, 40) + "…" }
    $n = $C.GetType().Name
    if ($t) { return ("{0}«{1}»" -f $n, $t) }
    return $n
}

function Get-SnapPath($C) {
    $parts = New-Object System.Collections.Generic.List[string]
    $x = $C
    while ($x -and -not ($x -is [System.Windows.Forms.Form])) {
        $parts.Insert(0, (Get-SnapName $x))
        $x = $x.Parent
    }
    return ($parts.ToArray() -join " / ")
}

# text that must fit: buttons, check boxes, radio buttons, labels (with or without ellipsis)
function Test-SnapTextFits($C) {
    $t = [string]$C.Text
    if (-not $t) { return $true }
    # a label that cuts its text with "…" is clipped too: the user sees «Не удалось определить…»
    if ($C -is [System.Windows.Forms.Label] -and $C.AutoEllipsis) {
        $flags = [System.Windows.Forms.TextFormatFlags]::SingleLine -bor [System.Windows.Forms.TextFormatFlags]::NoPadding
        $sz = [System.Windows.Forms.TextRenderer]::MeasureText($t, $C.Font, [System.Drawing.Size]::Empty, $flags)
        return ($sz.Width -le ($C.ClientSize.Width - $C.Padding.Horizontal + 1))
    }
    # AutoSize controls take the size of their text; checked against the parent bounds instead
    if ($C.AutoSize -and -not ($C -is [System.Windows.Forms.Button])) { return $true }
    $sz = [System.Windows.Forms.TextRenderer]::MeasureText($t, $C.Font)
    if ($C -is [System.Windows.Forms.Button]) { return ($sz.Width + 8 -le $C.Width -and $sz.Height -le $C.Height) }
    if ($C -is [System.Windows.Forms.CheckBox] -or $C -is [System.Windows.Forms.RadioButton]) { return ($sz.Width + 14 -le $C.Width) }
    if ($C -is [System.Windows.Forms.Label]) { return ($sz.Width -le $C.Width + 1) }
    return $true
}

function Test-SnapControls($Parent, $Out) {
    $kids = New-Object System.Collections.Generic.List[object]
    foreach ($c in $Parent.Controls) {
        if (-not $c.Visible) { continue }
        [void]$kids.Add($c)
    }
    $skipKids = ($Parent -is [System.Windows.Forms.ListView]) -or ($Parent -is [System.Windows.Forms.ComboBox]) -or
        ($Parent -is [System.Windows.Forms.TextBox]) -or ($Parent -is [System.Windows.Forms.NumericUpDown])
    if ($skipKids) { return }
    $client = $Parent.ClientRectangle
    foreach ($c in $kids) {
        $b = $c.Bounds
        if ($b.Width -le 0 -or $b.Height -le 0) {
            [void]$Out.Add(("нулевой размер {0}x{1}: {2}" -f $b.Width, $b.Height, (Get-SnapPath $c)))
            continue
        }
        if (-not ($Parent -is [System.Windows.Forms.TabControl]) -and -not $client.Contains($b)) {
            [void]$Out.Add(("за границей родителя {0} (родитель {1}x{2}): {3}" -f $b, $client.Width, $client.Height, (Get-SnapPath $c)))
        }
        if (-not (Test-SnapTextFits $c)) {
            [void]$Out.Add(("текст не помещается ({0} px): {1}" -f $b.Width, (Get-SnapPath $c)))
        }
    }
    if (-not ($Parent -is [System.Windows.Forms.TabControl]) -and -not ($Parent -is [System.Windows.Forms.SplitContainer])) {
        for ($i = 0; $i -lt $kids.Count; $i++) {
            for ($j = $i + 1; $j -lt $kids.Count; $j++) {
                $a = $kids[$i].Bounds
                $b = $kids[$j].Bounds
                if ($a.Width -le 0 -or $b.Width -le 0) { continue }
                if ($a.IntersectsWith($b)) {
                    [void]$Out.Add(("перекрытие: {0}  <>  {1}" -f (Get-SnapPath $kids[$i]), (Get-SnapPath $kids[$j])))
                }
            }
        }
    }
    foreach ($c in $kids) { Test-SnapControls $c $Out }
}

# ListView cells: a text wider than its column is cut with «…» (the layout check above does not
# look inside a list). $Overflow also flags columns wider than the list (a horizontal scroll bar).
function Test-SnapListView($Lv, [string]$What, $Out, [bool]$Overflow = $true) {
    if (-not $Lv -or -not $Lv.Visible -or $Lv.Columns.Count -eq 0) { return }
    $cut = 0
    $first = New-Object System.Collections.Generic.List[string]
    for ($i = 0; $i -lt $Lv.Columns.Count; $i++) {
        $cw = $Lv.Columns[$i].Width
        if ($cw -le 0) { continue }
        $extra = 12
        if ($i -eq 0 -and $Lv.CheckBoxes) { $extra += 18 }
        $hw = [System.Windows.Forms.TextRenderer]::MeasureText([string]$Lv.Columns[$i].Text, $Lv.Font).Width + 12
        if ($hw -gt $cw) { $cut++; if ($first.Count -lt 3) { [void]$first.Add(("заголовок «{0}» ({1} > {2} px)" -f $Lv.Columns[$i].Text, $hw, $cw)) } }
        foreach ($it in $Lv.Items) {
            if ($i -ge $it.SubItems.Count) { continue }
            $t = [string]$it.SubItems[$i].Text
            if (-not $t) { continue }
            $tw = [System.Windows.Forms.TextRenderer]::MeasureText($t, $Lv.Font).Width + $extra
            if ($tw -gt $cw) {
                $cut++
                if ($first.Count -lt 3) { [void]$first.Add(("«{0}» в колонке «{1}» ({2} > {3} px)" -f $t, $Lv.Columns[$i].Text, $tw, $cw)) }
            }
        }
    }
    if ($cut -gt 0) { [void]$Out.Add(("текст в ячейках обрезан ({0}): {1}: {2}" -f $cut, $What, ($first.ToArray() -join "; "))) }
    if ($Overflow) {
        $sum = 0
        foreach ($c in $Lv.Columns) { $sum += [Math]::Max(0, $c.Width) }
        if ($sum -gt $Lv.ClientSize.Width) { [void]$Out.Add(("колонки шире списка ({0} > {1} px, горизонтальная прокрутка): {2}" -f $sum, $Lv.ClientSize.Width, $What)) }
    }
}

# the lists of tab 4 that must show whole texts: the model list and the «Связи» table
function Test-SnapTab4Lists($Out) {
    if ($script:Ui.ContainsKey("PlainLists") -and $script:Ui.PlainLists -and $script:Ui.PlainLists.ContainsKey("addlinks")) {
        $ml = $script:Ui.PlainLists["addlinks"]
        if ($ml -and $ml.Visible) {
            $sum = 0
            foreach ($c in $ml.Columns) { $sum += [Math]::Max(0, $c.Width) }
            if ($sum -gt $ml.ClientSize.Width) { [void]$Out.Add(("колонки шире списка ({0} > {1} px, горизонтальная прокрутка): список моделей вкладки 4" -f $sum, $ml.ClientSize.Width)) }
        }
    }
    if ($script:Ui.ContainsKey("LinkList")) { Test-SnapListView $script:Ui.LinkList "таблица «Связи»" $Out }
}

function Save-SnapImage([string]$Name, $Of = $null) {
    if ($NoImages) { return }
    $form = $(if ($Of) { $Of } else { $script:form })
    $file = Join-Path $OutDir ($Name + ".png")
    $bmp = New-Object System.Drawing.Bitmap($form.Width, $form.Height)
    try {
        if ($CaptureScreen) {
            $g = [System.Drawing.Graphics]::FromImage($bmp)
            try { $g.CopyFromScreen($form.Location, [System.Drawing.Point]::Empty, $form.Size) } finally { $g.Dispose() }
        }
        else {
            $form.DrawToBitmap($bmp, (New-Object System.Drawing.Rectangle(0, 0, $form.Width, $form.Height)))
        }
        $bmp.Save($file, [System.Drawing.Imaging.ImageFormat]::Png)
    }
    finally { $bmp.Dispose() }
}

$total = 0
$sizes = @(
    @("min", $form.MinimumSize.Width, $form.MinimumSize.Height),
    @("max", 1920, 1040)
)
$pages = @(
    @(0, "01_perenos", $null, @()),
    @(1, "02_perepriv", $null, @()),
    @(2, "03_versiya", $null, @()),
    @(3, "04_svyazi", "Tabs", @("svyazi", "kody", "maska")),
    @(4, "05_nabory", "SetTabs", @("nabory", "svyazi_v_modelyah"))
)
foreach ($sz in $sizes) {
    $form.Size = New-Object System.Drawing.Size([int]$sz[1], [int]$sz[2])
    Wait-SnapUi
    Restore-OpsSplitters
    Wait-SnapUi
    Write-Output ("== {0}: {1}x{2}" -f $sz[0], $form.Width, $form.Height)
    foreach ($pg in $pages) {
        $script:Ui.TopTabs.SelectedIndex = [int]$pg[0]
        Wait-SnapUi
        $subs = @($null)
        $sub = $null
        if ($pg[2]) { $sub = $script:Ui[[string]$pg[2]]; $subs = @($pg[3]) }
        for ($k = 0; $k -lt $subs.Count; $k++) {
            $name = [string]$pg[1]
            if ($sub) {
                $sub.SelectedIndex = $k
                $name = ("{0}_{1}" -f $pg[1], $subs[$k])
                Wait-SnapUi
            }
            $name = ("{0}_{1}" -f $name, $sz[0])
            $found = New-Object System.Collections.Generic.List[string]
            Test-SnapControls $form $found
            Test-SnapTab4Lists $found
            if ($found.Count -eq 0) { Write-Output ("  {0}: ok" -f $name) }
            else {
                Write-Output ("  {0}: нарушений {1}" -f $name, $found.Count)
                foreach ($f in $found) { Write-Output ("    " + $f) }
            }
            $total += $found.Count
            Save-SnapImage $name
        }
    }
}
# ---- tab 4 with no new model: the «Связи» table stays empty (no red «нет пары» per target)
$form.Size = New-Object System.Drawing.Size($form.MinimumSize.Width, $form.MinimumSize.Height)
Wait-SnapUi
$script:Ui.TopTabs.SelectedIndex = 3
$script:Ui.Tabs.SelectedIndex = 0
if ($Demo) {
    Set-WorksetRole @($script:OpsDemoPaths) "target"
    Update-WorksetRoleColumn
}
Update-WorksetViews
Wait-SnapUi
$found = New-Object System.Collections.Generic.List[string]
Test-SnapControls $form $found
Test-SnapTab4Lists $found
if ($script:Ui.LinkList.Items.Count -ne 0) { [void]$found.Add(("без новых моделей в «Связях» строк: {0}, должно быть 0" -f $script:Ui.LinkList.Items.Count)) }
$name = "04_svyazi_bez_novyh_min"
if ($found.Count -eq 0) { Write-Output ("  {0}: ok" -f $name) }
else { Write-Output ("  {0}: нарушений {1}" -f $name, $found.Count); foreach ($x in $found) { Write-Output ("    " + $x) } }
$total += $found.Count
Save-SnapImage $name

# ---- confirm forms: «Вставить связи», its «Пробный прогон», and the generic plan dialog,
# at their minimum size and at a 1024x550 remote-desktop window. The buttons must be visible
# and lie inside the form.
$demoHost = "RSN://revit-server-2022.example.local/proj SampleSite/Стадия П/04_ЭОМ/SITE_ORG_EOM_K1_R22.rvt"
$demoActs = @(
    (New-WorksetAction "create" $demoHost "" "01_Связь_RVT_HC_K1" "" "" "" $true),
    (New-WorksetAction "create" $demoHost "" "01_Связь_RVT_AVT_K1" "" "" "" $true),
    (New-WorksetAction "insert" $demoHost "k1" "RSN://revit-server-2022.example.local/proj SampleSite/Стадия П/06_HC/SITE_ORG_HC_K1_R22.rvt" "01_Связь_RVT_HC_K1" "HC" "" $false),
    (New-WorksetAction "insert" $demoHost "k2" "RSN://revit-server-2022.example.local/proj SampleSite/Стадия П/07_AVT/SITE_ORG_AVT_K1_R22.rvt" "01_Связь_RVT_AVT_K1" "AVT" "" $false),
    (New-WorksetAction "skip" $demoHost "k3" "RSN://revit-server-2022.example.local/proj SampleSite/Стадия П/07_AVT/SITE_ORG_AVT_PRK_R22.rvt" "" "" "SKIP_EMPTY_WORKSET" $false)
)
$demoPairs = @([pscustomobject]@{ State = "no_pair"; TargetPath = ""; LinkPath = "RSN://revit-server-2022.example.local/proj SampleSite/Стадия П/06_HC/SITE_ORG_HC_K2_R22.rvt"; LinkName = "SITE_ORG_HC_K2_R22" })
$demoView = Get-WorksetInsertView $demoActs $demoPairs
$dialogs = @(
    @("06_vstavka", { New-WorksetInsertForm $demoView "Вставка связей" $false }),
    @("07_probnyi_progon", { New-WorksetInsertForm $demoView "Пробный прогон" $true }),
    @("08_plan", { New-OpsPlanForm "Рабочие наборы" ("Переименовать:" + [Environment]::NewLine + "  (нет)") $true "Да, записать" $false })
)
foreach ($d in $dialogs) {
    foreach ($sz in @(@("min", 0, 0), @("1024x550", 1024, 550))) {
        $parts = & $d[1]
        $df = $parts.Form
        $df.StartPosition = "Manual"
        if ($CaptureScreen) { $df.Location = New-Object System.Drawing.Point(0, 0) }
        else { $df.Location = New-Object System.Drawing.Point(-30000, -30000) }
        $df.Show()
        if ([int]$sz[1] -gt 0) { $df.Size = New-Object System.Drawing.Size([int]$sz[1], [int]$sz[2]) }
        else { $df.Size = $df.MinimumSize }
        Wait-SnapUi
        $name = ("{0}_{1}" -f $d[0], $sz[0])
        $found = New-Object System.Collections.Generic.List[string]
        Test-SnapControls $df $found
        if ($parts.ContainsKey("List")) { Test-SnapListView $parts.List "список вставки" $found }
        if ($parts.ContainsKey("Skipped")) { Test-SnapListView $parts.Skipped "«Пропущено»" $found }
        foreach ($b in @($parts.Yes, $parts.No)) {
            if (-not $b -or -not $b.Visible) { continue }
            $p = $df.PointToClient($b.PointToScreen([System.Drawing.Point]::Empty))
            $r = New-Object System.Drawing.Rectangle($p.X, $p.Y, $b.Width, $b.Height)
            if (-not $df.ClientRectangle.Contains($r)) { [void]$found.Add(("кнопка «{0}» вне окна: {1}" -f $b.Text, $r)) }
        }
        if (-not $parts.No.Visible) { [void]$found.Add("нет кнопки «Отмена» / «Закрыть»") }
        if ($found.Count -eq 0) { Write-Output ("  {0}: ok" -f $name) }
        else { Write-Output ("  {0}: нарушений {1}" -f $name, $found.Count); foreach ($x in $found) { Write-Output ("    " + $x) } }
        $total += $found.Count
        Save-SnapImage $name $df
        $df.Hide()
        $df.Dispose()
    }
}

$form.Hide()
$form.Dispose()
if (-not $NoImages) { Write-Output ("Снимки: {0}" -f $OutDir) }
if ($total -gt 0) {
    Write-Output ("Нарушений раскладки: {0}" -f $total)
    exit 1
}
Write-Output "Раскладка без нарушений."
exit 0
