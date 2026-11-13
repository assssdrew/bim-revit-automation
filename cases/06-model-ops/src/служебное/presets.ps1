# Name: presets.ps1
# Version: 3.1.0
# What it does: Single WinForms window to plan rename/move/year/relink passes and export job files.
# Inputs: Operator inputs; rvt_list paths.
# Outputs: job_paths.csv, mapping.csv, rbp_checklist.txt.
# How to run: Операции с моделями.vbs or presets.cmd.
# Notes: UTF-8 BOM; re-execs self in STA if needed.
# UTF-8 WITH BOM required (PowerShell 5.1)
# Вариант Б: одно окно операций с моделями. Двойной клик по VBS - без консоли.

if ([Threading.Thread]::CurrentThread.GetApartmentState() -ne "STA") {
    $arg = "-NoProfile -ExecutionPolicy Bypass -STA -File `"$PSCommandPath`""
    Start-Process -FilePath "powershell.exe" -ArgumentList $arg -Wait -NoNewWindow
    exit $LASTEXITCODE
}

$ErrorActionPreference = "Continue"
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# Code lives in служебное\ ($CodeDir); the tool root ($ToolDir) keeps the user data:
# rvt_list.txt, job_paths.csv, mapping.csv, servers.cfg, models_path.cfg ...
$CodeDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ToolDir = Split-Path -Parent $CodeDir
$script:ModelOpsUiLibrary = $true
$ModelOpsUiLibrary = $true
$script:ModelOpsLaunchLibrary = $true

try {
    . (Join-Path $CodeDir "job_lib.ps1")
    . (Join-Path $CodeDir "ui_logic.ps1")
    . (Join-Path $CodeDir "ui_grid.ps1")
    Set-StrictMode -Off
    . (Join-Path $CodeDir "choose_models.ps1")
    . (Join-Path $CodeDir "start_ops.ps1")
    . (Join-Path $CodeDir "workset_logic.ps1")
    . (Join-Path $CodeDir "workset_ui.ps1")
    $ErrorActionPreference = "Continue"
}
catch {
    [System.Windows.Forms.MessageBox]::Show(
        $_.Exception.Message,
        "Операции с моделями",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
    throw
}

$ListPath = Join-Path $ToolDir "rvt_list.txt"
$ServersCfg = Join-Path $ToolDir "servers.cfg"
$script:Years = @("2021", "2022", "2023", "2024", "2025", "2026")
$script:modelPaths = New-Object System.Collections.Generic.List[string]
if ($script:OpsBuildOnly -and $null -ne $script:OpsDemoPaths) {
    foreach ($p in @($script:OpsDemoPaths)) { [void]$script:modelPaths.Add([string]$p) }
}
else {
    foreach ($p in @(Read-ModelList $ListPath)) { [void]$script:modelPaths.Add($p) }
}
$script:ops = @()
$script:nameOverrides = @{}
$script:manualRows = @{}      # old path key -> @{ name; path }: hand edits, beat the rule
$script:RowCache = @{}        # old path -> preview row currently shown
$script:ManualCount = 0
$script:Edit = $null          # cell being edited in place
$script:TipText = ""
$script:autoSeg = ""          # last folder text we put into "было" ourselves
$script:autoRule = ""         # last name we put into the "Имя" example ourselves
$script:UpgradeSrc = ""       # detected from the models (tab «Смена версии Revit»)
$script:UpgradeDst = ""
$script:DryRunOnly = $false
$script:MoveState = $null
$script:ManualColor = [System.Drawing.Color]::FromArgb(255, 245, 190)
$script:servers = @(Read-ServerTable $ServersCfg)
$script:locMode = "same"
$script:oldSeg = ""
$script:newSeg = ""
$script:oldPref = ""
$script:newPref = ""
$script:bufferDir = Get-DefaultBufferDir
$script:copyLinks = "new"
$script:Ui = @{}
$script:Run = @{
    List      = $null
    Active    = $false
    Queue     = @()
    Index     = 0
    Process   = $null
    ExitCodes = @()
}

function Get-OpLabelRu($Op) {
    switch ($Op.kind) {
        "delete" { return ("удалить кусок «{0}»" -f $Op.token) }
        "replace" { return ("заменить «{0}» на «{1}»" -f $Op.from, $Op.to) }
        "insert" {
            if ($Op.after) { return ("вставить «{0}» после «{1}»" -f $Op.token, $Op.after) }
            return ("добавить «{0}» в конец" -f $Op.token)
        }
    }
    return [string]$Op.kind
}

function Get-PassTitle([string]$Name) {
    switch ($Name) {
        "detach" { return "Снимаю модели с Revit Server в буфер на этом ПК" }
        "saveas" { return "Сохраняю новые хранилища" }
        "relink" { return "Перепривязываю связи" }
        "move" { return "Переношу файлы" }
        "delete" { return "Убираю старые модели в _to_delete" }
        "workset-read" { return "Читаю рабочие наборы и связи" }
        "workset-apply" { return "Записываю переименование и новые наборы" }
        default { return $Name }
    }
}

function Test-OpsPassTracked([string]$Name) {
    return ($Name -eq "relink" -or $Name -eq "workset-read" -or $Name -eq "workset-apply")
}

function Get-ModelSource([string]$Path) {
    if (Test-IsRsnPath $Path) { return "Revit Server" }
    if ($Path.StartsWith("\\")) { return "Сеть" }
    return "Этот компьютер"
}

# Source year: always detected from the models (tab «Смена версии Revit»), never chosen by hand.
function Get-SelectedYears {
    if (-not $script:Ui.CbUpgrade.Checked) { return @{ Src = ""; Dst = "" } }
    $src = ""
    $dst = ""
    if ($script:UpgradeSrc) { $src = [string]$script:UpgradeSrc }
    if ($script:UpgradeDst) { $dst = [string]$script:UpgradeDst }
    return @{ Src = $src; Dst = $dst }
}

function Get-PreviewRows([hashtable]$Manual = $null) {
    if ($null -eq $Manual) { $Manual = $script:manualRows }
    # a hand edit counts by itself: a typed name renames, a typed path moves
    $doRename = [bool]$script:Ui.CbRename.Checked
    $doMove = [bool]$script:Ui.CbMove.Checked
    foreach ($k in @($Manual.Keys)) {
        if (Get-ManualField $Manual[$k] "path") { $doMove = $true }
        elseif (Get-ManualField $Manual[$k] "name") { $doRename = $true }
    }
    $effective = Get-EffectiveManualRows $Manual $doRename $doMove
    $years = Get-SelectedYears
    return @(New-JobPreviewRows `
            -OldPaths @($script:modelPaths) `
            -Ops $script:ops `
            -DoRename $doRename `
            -LocationMode $script:locMode `
            -OldSegment $script:oldSeg `
            -NewSegment $script:newSeg `
            -OldPrefix $script:oldPref `
            -NewPrefix $script:newPref `
            -SourceYear $years.Src `
            -TargetYear $years.Dst `
            -BufferDir $script:bufferDir `
            -NameOverrides $script:nameOverrides `
            -ManualRows $effective)
}

function Get-JobPlanText($Job, $Rows) {
    $passes = @(New-OpsPassQueue -Job $Job -ToolDir $ToolDir -CodeDir $CodeDir -Rows $Rows)
    if ($passes.Count -eq 0) { return "Пока нечего делать: отметьте хотя бы одно действие." }
    $i = 1
    $lines = New-Object System.Collections.Generic.List[string]
    foreach ($p in $passes) {
        $t = Get-PassTitle $p.Name
        if ($p.Year) { $t = ("Revit {0}: {1}" -f $p.Year, $t) }
        [void]$lines.Add(("{0}. {1}" -f $i, $t))
        $i++
    }
    return ($lines -join [Environment]::NewLine)
}

function Test-UiRelinkOnly {
    if (-not $script:Ui -or -not $script:Ui.ContainsKey("TxtRelOld")) { return $false }
    return ($script:Ui.CbRelink.Checked -and -not $script:Ui.CbRename.Checked -and -not $script:Ui.CbMove.Checked -and
        -not $script:Ui.CbUpgrade.Checked -and -not $script:Ui.CbCopy.Checked -and -not $script:Ui.CbDelete.Checked)
}

function Build-JobObject($Rows) {
    $years = Get-SelectedYears
    $upgrade = [bool]$script:Ui.CbUpgrade.Checked
    $loc = $script:locMode
    if (-not $script:Ui.CbMove.Checked) { $loc = "same" }
    $uncMove = $false
    if (-not $script:Ui.CbRename.Checked) {
        $uncMove = Test-UncMoveCandidate -Rows $Rows -Upgrade $upgrade -LocationMode $loc
    }
    $needSaveAs = $true
    if ($uncMove -and -not $upgrade -and -not $script:Ui.CbRename.Checked) { $needSaveAs = $false }
    if ($script:Ui.CbRelink.Checked -and -not $script:Ui.CbRename.Checked -and $loc -eq "same" -and -not $upgrade) {
        $needSaveAs = $false
        $uncMove = $false
    }
    $opsJson = @($script:ops | ForEach-Object {
            @{ kind = $_.kind; token = $_.token; from = $_.from; to = $_.to; after = $_.after }
        })
    return [pscustomobject]@{
        version       = "1.0.0"
        rename        = [bool]$script:Ui.CbRename.Checked
        relink        = [bool]$script:Ui.CbRelink.Checked
        copy          = [bool]$script:Ui.CbCopy.Checked
        delete_old    = [bool]$script:Ui.CbDelete.Checked
        allow_overwrite = ([bool]$script:Ui.CbOverwrite.Checked -and -not (Test-UiRelinkOnly))
        upgrade       = $upgrade
        source_year   = $years.Src
        target_year   = $years.Dst
        location_mode = $loc
        old_segment   = $script:oldSeg
        new_segment   = $script:newSeg
        old_prefix    = $script:oldPref
        new_prefix    = $script:newPref
        buffer_dir    = $script:bufferDir
        rsn_buffer    = "detach"
        copy_links    = $script:copyLinks
        saveas_needed = $needSaveAs
        unc_move      = $uncMove
        relink_old    = $(if (Test-UiRelinkOnly) { $script:Ui.TxtRelOld.Text.Trim() } else { "" })
        relink_new    = $(if (Test-UiRelinkOnly) { $script:Ui.TxtRelNew.Text.Trim() } else { "" })
        reports_dir   = (Initialize-OpsReportsDir $ToolDir)
        ops           = $opsJson
        saved_utc     = [DateTime]::UtcNow.ToString("o")
    }
}

# The status line is one short line. A longer text ($Tip: paths, notes) is its tool tip.
function Set-UiStatusTip([string]$Text, [string]$Tip) {
    if (-not $script:Ui.Status -or -not $script:Ui.ContainsKey("StatusTip") -or -not $script:Ui.StatusTip) { return }
    if (-not $Tip) { $Tip = $Text }
    $script:Ui.StatusTip.SetToolTip($script:Ui.Status, $Tip)
}

function Set-UiStatus([string]$Text, [string]$Tip = "") {
    if ($script:Ui.Status) {
        $script:Ui.Status.ForeColor = [System.Drawing.SystemColors]::ControlText
        $script:Ui.Status.Text = $Text
        Set-UiStatusTip $Text $Tip
    }
    [System.Windows.Forms.Application]::DoEvents()
}

# errors of the fields go to the status line, in red (no hint labels on the tabs)
function Set-UiError([string]$Text, [string]$Tip = "") {
    if ($script:Ui.Status) {
        $script:Ui.Status.ForeColor = [System.Drawing.Color]::Firebrick
        $script:Ui.Status.Text = $Text
        Set-UiStatusTip $Text $Tip
    }
}

# ticked models whose new path is not a full path (the run is blocked for them)
function Get-CheckedPathErrorCount {
    $n = 0
    foreach ($it in @($script:Ui.List.Items)) {
        if (-not $it.Checked) { continue }
        $r = $script:RowCache[[string]$it.Tag]
        if ($r -and $r.path_error) { $n++ }
    }
    return $n
}

function Update-ListCount {
    $n = $script:Ui.List.Items.Count
    $c = 0
    foreach ($it in @($script:Ui.List.Items)) { if ($it.Checked) { $c++ } }
    $text = ("В списке: {0} · отмечено: {1}" -f $n, $c)
    if ($n -gt 0) {
        if ($script:ManualCount -gt 0) { $text += (" · вручную: {0}" -f $script:ManualCount) }
        $bad = Get-CheckedPathErrorCount
        if ($bad -gt 0) { $text += (" · ошибок пути: {0}" -f $bad) }
    }
    $script:Ui.CountLabel.Text = $text
    Update-RunButton
}

# The run button is shared by all tabs: off while a run is going or the tab has no ticked model.
function Update-RunButton {
    if (-not $script:Ui.BtnRun) { return }
    $busy = [bool]$script:Run.Active
    $key = Get-ActiveTabKey
    $any = $false
    if ($key -eq "move") {
        foreach ($it in @($script:Ui.List.Items)) { if ($it.Checked) { $any = $true; break } }
    }
    elseif ($script:Ui.PlainLists -and $script:Ui.PlainLists.ContainsKey($key)) {
        foreach ($it in @($script:Ui.PlainLists[$key].Items)) { if ($it.Checked) { $any = $true; break } }
    }
    $script:Ui.BtnRun.Enabled = $any -and (-not $busy)
    if ($script:Ui.BtnDry) { $script:Ui.BtnDry.Enabled = $any -and (-not $busy) }
}

function Refresh-ModelList {
    $lv = $script:Ui.List
    $prev = @{}
    $prevState = @{}
    foreach ($it in @($lv.Items)) {
        $key = [string]$it.Tag
        $prev[$key] = [bool]$it.Checked
        if ($it.SubItems.Count -gt 2 -and [string]$it.SubItems[2].Text -ne "Ошибка пути") { $prevState[$key] = [string]$it.SubItems[2].Text }
    }
    $rows = @(Get-PreviewRows)
    $script:RowCache = @{}
    $script:ManualCount = 0
    foreach ($r in $rows) {
        $script:RowCache[[string]$r.old_path] = $r
        if ($r.manual) { $script:ManualCount++ }
    }
    $lv.BeginUpdate()
    try {
        $lv.Items.Clear()
        foreach ($r in $rows) {
            # subitems: 0 Модель, 1 Станет, 2 Состояние, 3 Откуда (hidden), 4 Старый путь, 5 Новый путь
            $it = New-Object System.Windows.Forms.ListViewItem ($r.old_name)
            $it.UseItemStyleForSubItems = $false
            [void]$it.SubItems.Add($r.new_name)
            $st = ""
            if ($prevState.ContainsKey([string]$r.old_path)) { $st = $prevState[[string]$r.old_path] }
            if ($r.path_error) { $st = "Ошибка пути" }
            [void]$it.SubItems.Add($st)
            [void]$it.SubItems.Add((Get-ModelSource $r.old_path))
            [void]$it.SubItems.Add($r.old_path)
            [void]$it.SubItems.Add($r.new_path)
            if ($r.path_error) {
                $it.SubItems[2].ForeColor = [System.Drawing.Color]::Firebrick
                $it.SubItems[5].ForeColor = [System.Drawing.Color]::Firebrick
            }
            if ($r.manual) {
                # hand edits are marked: name cell, and the path cell when the path itself was typed
                $it.SubItems[1].BackColor = $script:ManualColor
                if ($r.manual -eq "path") { $it.SubItems[5].BackColor = $script:ManualColor }
            }
            $it.Tag = $r.old_path
            if ($prev.ContainsKey([string]$r.old_path)) { $it.Checked = $prev[[string]$r.old_path] }
            else { $it.Checked = $true }
            [void]$lv.Items.Add($it)
        }
    }
    finally { $lv.EndUpdate() }
    Update-ListCount
}

function Add-UiModels([string[]]$Paths) {
    $seen = @{}
    foreach ($old in @($script:modelPaths)) { $seen[$old] = $true }
    $n = 0
    foreach ($raw in @($Paths)) {
        if ([string]::IsNullOrWhiteSpace($raw)) { continue }
        $p = $raw.Trim()
        if ($seen.ContainsKey($p)) { continue }
        if ($p -notmatch '^(?i)RSN://' -and (Get-Command Test-ExcludedModelPath -ErrorAction SilentlyContinue)) {
            if (Test-ExcludedModelPath $p) { continue }
        }
        $seen[$p] = $true
        [void]$script:modelPaths.Add($p)
        $n++
    }
    if (Get-Command Set-MoveSuggestions -ErrorAction SilentlyContinue) { Set-MoveSuggestions }
    if (Get-Command Sync-MoveFlags -ErrorAction SilentlyContinue) { [void](Sync-MoveFlags) }
    Refresh-ModelList
    return $n
}

function Get-CheckedPaths {
    if ($script:IsolatePaths) {
        $out = New-Object System.Collections.Generic.List[string]
        foreach ($p in @($script:IsolatePaths)) {
            if (-not [string]::IsNullOrWhiteSpace([string]$p)) { [void]$out.Add([string]$p) }
        }
        return @($out)
    }
    $out = New-Object System.Collections.Generic.List[string]
    foreach ($it in @($script:Ui.List.Items)) {
        if ($it.Checked -and $it.Tag) { [void]$out.Add([string]$it.Tag) }
    }
    return @($out)
}

function Set-AllChecks([bool]$On) {
    foreach ($it in @($script:Ui.List.Items)) { $it.Checked = $On }
    Update-ListCount
}

# List of the tab that is running now (states and progress go there, not always to tab 1).
function Get-RunList {
    if ($script:Run -and $script:Run.List) { return $script:Run.List }
    return $script:Ui.List
}

function Set-RowState([string]$Path, [string]$Text) {
    foreach ($it in @((Get-RunList).Items)) {
        if ([string]$it.Tag -eq $Path) {
            if ($it.SubItems.Count -gt 2) { $it.SubItems[2].Text = $Text }
            break
        }
    }
}

function Set-CheckedState([string]$Text) {
    foreach ($it in @((Get-RunList).Items)) {
        if ($it.Checked) { Set-RowState ([string]$it.Tag) $Text }
    }
}

function Get-ModelsFromFolder([string]$Folder) {
    if ([string]::IsNullOrWhiteSpace($Folder)) { return @() }
    $found = New-Object System.Collections.Generic.List[string]
    $direct = @(Get-ChildItem -LiteralPath $Folder -File -Filter *.rvt -Force -ErrorAction SilentlyContinue)
    foreach ($f in $direct) {
        if ((Get-Command Test-IsTargetRvtName -ErrorAction SilentlyContinue) -and -not (Test-IsTargetRvtName $f.Name)) { continue }
        $native = Get-NativePath $f.FullName
        if (-not $native) { continue }
        if (Test-ExcludedModelPath $native) { continue }
        [void]$found.Add($native)
    }
    foreach ($p in @(Get-TargetRvtsFromDisciplineFolders -Folders @($Folder))) {
        if ($p) { [void]$found.Add($p) }
    }
    return @($found | Sort-Object -Unique)
}

function Show-ServerPick {
    Ensure-ServersFile
    $servers = @(Get-Servers)
    if ($servers.Count -eq 0) {
        [void][System.Windows.Forms.MessageBox]::Show($script:Ui.Form, "В servers.cfg нет серверов.", "Revit Server")
        return $null
    }
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Revit Server"
    $form.Size = New-Object System.Drawing.Size(460, 360)
    $form.StartPosition = "CenterParent"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.ShowInTaskbar = $false
    $list = New-Object System.Windows.Forms.ListBox
    $list.Location = New-Object System.Drawing.Point(12, 16)
    $list.Size = New-Object System.Drawing.Size(420, 250)
    foreach ($s in $servers) { [void]$list.Items.Add(("{0}    {1}" -f $s.Label, $s.Host)) }
    if ($list.Items.Count -gt 0) { $list.SelectedIndex = 0 }
    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = "Открыть"
    $ok.Location = New-Object System.Drawing.Point(236, 280)
    $ok.Size = New-Object System.Drawing.Size(96, 28)
    $ok.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $cancel = New-Object System.Windows.Forms.Button
    $cancel.Text = "Отмена"
    $cancel.Location = New-Object System.Drawing.Point(336, 280)
    $cancel.Size = New-Object System.Drawing.Size(96, 28)
    $cancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.AcceptButton = $ok
    $form.CancelButton = $cancel
    $form.Controls.AddRange(@($list, $ok, $cancel))
    if ($form.ShowDialog($script:Ui.Form) -ne [System.Windows.Forms.DialogResult]::OK) { return $null }
    if ($list.SelectedIndex -lt 0) { return $null }
    return $servers[$list.SelectedIndex]
}

# ---- tab «Перенос и переименование»: the rule and the folders are typed right on the tab ----

# An error inside an event handler used to vanish (hidden window, ErrorActionPreference=Continue).
# Now it is written to отчёты\_window_errors.log and shown in the status line.
function Write-WindowError([string]$Where, $Err) {
    $text = ""
    try { $text = [string]$Err.Exception.Message } catch { $text = [string]$Err }
    try {
        $dir = Initialize-OpsReportsDir $ToolDir
        $trace = ([string]$Err.ScriptStackTrace) -replace "\r?\n", " | "
        $line = "{0}`t{1}`t{2}`t{3}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Where, $text, $trace
        Add-Content -LiteralPath (Join-Path $dir "_window_errors.log") -Encoding UTF8 -Value $line
    }
    catch { }
    return $text
}

function Show-WindowError([string]$Where, $Err) {
    $msg = Write-WindowError $Where $Err
    $full = "Внутренняя ошибка окна ({0}): {1}. Журнал: отчёты\_window_errors.log" -f $Where, $msg
    Set-UiError ("Внутренняя ошибка окна ({0}). Журнал: отчёты\_window_errors.log" -f $Where) $full
}

# Hand edits that are remembered: name edits / path edits.
function Get-ManualCounts {
    $names = 0
    $paths = 0
    foreach ($k in @($script:manualRows.Keys)) {
        $e = $script:manualRows[$k]
        if (Get-ManualField $e "path") { $paths++ }
        elseif (Get-ManualField $e "name") { $names++ }
    }
    return @{ Names = $names; Paths = $paths }
}

# The fields of the tab decide what the run does: an empty field = that step is not done.
# Sets the job switches (CbRename / CbMove), the rule ($script:ops) and the folder replace.
function Sync-MoveFlags {
    if (-not $script:Ui -or -not $script:Ui.ContainsKey("RuleOld")) { return $null }
    $st = Get-MoveFieldState -RuleOld $script:Ui.RuleOld.Text -RuleNew $script:Ui.RuleNew.Text `
        -FolderOld $script:Ui.FolderOld.Text -FolderNew $script:Ui.FolderNew.Text -Paths @($script:modelPaths)
    $script:ops = @($st.Ops)
    $script:MoveState = $st
    $manual = Get-ManualCounts
    $script:Ui.CbRename.Checked = ($st.Rename -or $manual.Names -gt 0)
    $script:Ui.CbMove.Checked = ($st.Move -or $manual.Paths -gt 0)
    $script:locMode = "same"
    $script:oldSeg = ""; $script:newSeg = ""; $script:oldPref = ""; $script:newPref = ""
    if ($st.Move) {
        $script:locMode = $st.Mode
        if ($st.Mode -eq "prefix") { $script:oldPref = $st.Was; $script:newPref = $st.Became }
        else { $script:oldSeg = $st.Was; $script:newSeg = $st.Became }
    }
    elseif ($script:Ui.CbMove.Checked) { $script:locMode = "segment" }
    if (-not $script:Ui.CbCopy.Checked -and -not $script:Ui.CbDelete.Checked) { $script:Ui.CbCopy.Checked = $true }
    return $st
}

# Status line text for the fields: an error in red, nothing when all is fine.
function Show-MoveFieldState($St) {
    if (-not $St) { return }
    if ($St.RuleMessage) { Set-UiError ("Имя: " + $St.RuleMessage); return }
    if ($St.MoveMessage) { Set-UiError $St.MoveMessage; return }
    if ($script:Ui.Status -and $script:Ui.Status.ForeColor -eq [System.Drawing.Color]::Firebrick) { Set-UiStatus "" }
}

function Update-MoveFields {
    try {
        $st = Sync-MoveFlags
        Refresh-ModelList
        Show-MoveFieldState $st
    }
    catch { Show-WindowError "Перенос и переименование" $_ }
}

# "было" of the name = the first model; "было" of the folder = the common folder of the models.
# Text typed by hand is never overwritten.
function Set-MoveSuggestions {
    if (-not $script:Ui -or -not $script:Ui.ContainsKey("RuleOld")) { return }
    $first = ""
    if ($script:modelPaths.Count -gt 0) { $first = Get-BasenameNoExt $script:modelPaths[0] }
    $cur = $script:Ui.RuleOld.Text.Trim()
    if ($cur -eq "" -or $cur -ceq $script:autoRule) {
        $script:autoRule = $first
        if ($script:Ui.RuleOld.Text -cne $first) { $script:Ui.RuleOld.Text = $first }
    }
    $sug = Get-SegmentSuggestion @(Get-CheckedPaths) @($script:modelPaths)
    $cur = $script:Ui.FolderOld.Text.Trim()
    if ($cur -eq "" -or $cur -ceq $script:autoSeg) {
        $script:autoSeg = [string]$sug
        if ($script:Ui.FolderOld.Text -cne [string]$sug) { $script:Ui.FolderOld.Text = [string]$sug }
    }
}

# ---- copy, tooltips, in-place editing of the table ----

function Get-SelectedRowObjects {
    $out = New-Object System.Collections.Generic.List[object]
    foreach ($it in @($script:Ui.List.SelectedItems | Sort-Object Index)) {
        $r = $script:RowCache[[string]$it.Tag]
        if ($r) { [void]$out.Add($r) }
    }
    return $out.ToArray()      # not @($out): that throws on a List[object] made by New-Object
}

# Mode: both (old + new path, tab-separated) | old | new. One selected row per line.
function Copy-SelectedRows([string]$Mode) {
    $rows = @(Get-SelectedRowObjects)
    if ($rows.Count -eq 0) { Set-UiStatus "Выделите строки, которые нужно скопировать"; return }
    if ($Mode -eq "both" -and $script:Ui.List.Columns[5].Width -le 0) { $Mode = "old" }
    $text = Format-RowsForClipboard $rows $Mode
    try { [System.Windows.Forms.Clipboard]::SetText($text) }
    catch {
        try { [System.Windows.Forms.Clipboard]::SetDataObject($text, $true, 5, 100) }
        catch {
            [void][System.Windows.Forms.MessageBox]::Show($script:Ui.Form, ("Не удалось записать в буфер обмена: {0}" -f $_.Exception.Message), "Копирование")
            return
        }
    }
    Set-UiStatus ("Скопировано строк: {0}" -f $rows.Count)
}

function Close-CellEditor {
    $script:Edit = $null
    $ed = $script:Ui.Editor
    if ($ed) { $ed.Visible = $false }
}

function Show-CellEditor($Item, [int]$Col) {
    if ($script:Run.Active) { return }
    $lv = $script:Ui.List
    if ($lv.Columns[$Col].Width -le 0) { return }
    $r = $script:RowCache[[string]$Item.Tag]
    if (-not $r) { return }
    if ($script:Edit) { Close-CellEditor }
    if (-not $script:Ui.Editor) {
        $ed = New-Object System.Windows.Forms.TextBox
        $ed.BorderStyle = "FixedSingle"
        $ed.Visible = $false
        $ed.Add_KeyDown({
                if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
                    $_.Handled = $true
                    $_.SuppressKeyPress = $true
                    Complete-CellEdit $true
                }
                elseif ($_.KeyCode -eq [System.Windows.Forms.Keys]::Escape) {
                    $_.Handled = $true
                    $_.SuppressKeyPress = $true
                    Complete-CellEdit $false
                }
            })
        $ed.Add_Leave({
                $st = $script:Edit
                if ($st -and (([DateTime]::UtcNow - $st.Opened).TotalMilliseconds -lt 300)) {
                    # the list view can grab the focus right after the double click: take it back later
                    [void]$script:Ui.Form.BeginInvoke([Action] { if ($script:Edit) { $script:Ui.Editor.Focus() } })
                    return
                }
                Complete-CellEdit $true
            })
        $lv.Controls.Add($ed)
        $script:Ui.Editor = $ed
    }
    $ed = $script:Ui.Editor
    $b = $Item.SubItems[$Col].Bounds
    $w = [Math]::Max($b.Width, [Math]::Min(520, $lv.ClientSize.Width - $b.X - 4))
    $ed.SetBounds($b.X, $b.Y, $w, $b.Height)
    $ed.Font = $lv.Font
    $ed.Text = $(if ($Col -eq 1) { [string]$r.new_name } else { [string]$r.new_path })
    $script:Edit = @{ Path = [string]$Item.Tag; Col = $Col; Original = $ed.Text; Busy = $false; Opened = [DateTime]::UtcNow }
    $ed.Visible = $true
    $ed.BringToFront()
    $ed.Focus()
    $ed.SelectAll()
}

# Validates and stores a hand edit. Empty text = back to the rule.
function Set-ManualEdit([string]$OldPath, [int]$Col, [string]$Text) {
    $key = Get-ManualRowKey $OldPath
    if ($Text -eq "") {
        if ($script:manualRows.ContainsKey($key)) { $script:manualRows.Remove($key) }
        return @{ Ok = $true; Message = "" }
    }
    $existing = $null
    if ($script:manualRows.ContainsKey($key)) { $existing = $script:manualRows[$key] }
    if ($Col -eq 1) {
        $chk = Test-ModelFileName $Text
        if (-not $chk.Ok) { return @{ Ok = $false; Message = $chk.Message } }
        $entry = Set-ManualName $existing $chk.Value
    }
    else {
        $chk = Test-NewModelPath $Text
        if (-not $chk.Ok) { return @{ Ok = $false; Message = $chk.Message } }
        $entry = Set-ManualPath $existing $chk.Value
    }
    $trial = $script:manualRows.Clone()
    $trial[$key] = $entry
    $rows = @(Get-PreviewRows $trial)
    $self = @($rows | Where-Object { [string]$_.old_path -ceq $OldPath })
    if ($self.Count -eq 0) { return @{ Ok = $false; Message = "Модель не найдена в списке." } }
    $others = @($rows | Where-Object { [string]$_.old_path -cne $OldPath })
    $conf = Test-PathConflict $self[0].new_path $OldPath $others
    if (-not $conf.Ok) { return @{ Ok = $false; Message = $conf.Message } }
    # typed exactly what the rule gives: that is not a hand edit
    if (([string]$self[0].new_name -ceq [string]$self[0].rule_name) -and ([string]$self[0].new_path -ceq [string]$self[0].rule_path)) {
        $trial.Remove($key)
    }
    $script:manualRows = $trial
    return @{ Ok = $true; Message = "" }
}

function Complete-CellEdit([bool]$Commit) {
    $st = $script:Edit
    if (-not $st -or $st.Busy) { return }
    if (-not $Commit) { Close-CellEditor; return }
    $text = $script:Ui.Editor.Text.Trim()
    if ($text -ceq $st.Original.Trim()) { Close-CellEditor; return }
    $st.Busy = $true
    $res = Set-ManualEdit $st.Path $st.Col $text
    if (-not $res.Ok) {
        # message box takes the focus (Leave fires): Busy keeps that from committing again
        [void][System.Windows.Forms.MessageBox]::Show($script:Ui.Form, $res.Message, "Правка не принята",
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        $st.Busy = $false
        $script:Ui.Editor.Focus()
        return
    }
    $path = $st.Path
    Close-CellEditor
    Update-MoveFields
    foreach ($it in @($script:Ui.List.Items)) {
        if ([string]$it.Tag -ceq $path) { $it.Selected = $true; $it.EnsureVisible(); break }
    }
    Set-UiStatus $(if ($text -eq "") { "Возвращено по правилу" } else { "Изменено вручную" })
}

# "Вернуть по правилу" for the selected rows
function Reset-SelectedManual {
    $n = 0
    foreach ($r in @(Get-SelectedRowObjects)) {
        $key = Get-ManualRowKey ([string]$r.old_path)
        if ($script:manualRows.ContainsKey($key)) { $script:manualRows.Remove($key); $n++ }
    }
    if ($n -gt 0) { Update-MoveFields }
    Set-UiStatus ("Возвращено по правилу: {0}" -f $n)
}

function Get-TabModelPaths([string]$Key) {
    $out = New-Object System.Collections.Generic.List[string]
    $lv = $null
    if ($Key -eq "move") { $lv = $script:Ui.List }
    elseif ($script:Ui.ContainsKey("PlainLists") -and $script:Ui.PlainLists.ContainsKey($Key)) { $lv = $script:Ui.PlainLists[$Key] }
    if ($lv) {
        $any = $false
        foreach ($it in @($lv.Items)) { if ($it.Checked) { $any = $true; break } }
        foreach ($it in @($lv.Items)) {
            if ($any -and -not $it.Checked) { continue }
            if ($it.Tag) { [void]$out.Add([string]$it.Tag) }
        }
        return $out.ToArray()
    }
    if ($Key -eq "move") {
        foreach ($p in @($script:modelPaths)) { [void]$out.Add([string]$p) }
    }
    elseif ($script:TabPaths -and $script:TabPaths.ContainsKey($Key)) {
        foreach ($p in @($script:TabPaths[$Key])) { [void]$out.Add([string]$p) }
    }
    return $out.ToArray()
}

function Add-PlainTabModels([string]$Key, $Paths) {
    if (-not $script:TabPaths -or -not $script:TabPaths.ContainsKey($Key)) { return 0 }
    $before = @($script:TabPaths[$Key]).Count
    $merged = @(Merge-TabModelPaths $script:TabPaths[$Key] $Paths)
    $store = New-Object System.Collections.Generic.List[string]
    foreach ($p in $merged) { [void]$store.Add([string]$p) }
    $script:TabPaths[$Key] = $store
    Update-PlainList $Key
    return ($store.Count - $before)
}

function Get-PathLeaf([string]$Path) {
    $name = [string]$Path
    $slash = [Math]::Max($name.LastIndexOf("\"), $name.LastIndexOf("/"))
    if ($slash -ge 0 -and $slash -lt ($name.Length - 1)) { $name = $name.Substring($slash + 1) }
    return $name
}

# Plain lists: 0 Модель, 1 Путь, 2 Состояние, 3 Версия сейчас / Роль (shown second)
function Update-PlainList([string]$Key) {
    $lv = $null
    if ($script:Ui.ContainsKey("PlainLists") -and $script:Ui.PlainLists.ContainsKey($Key)) { $lv = $script:Ui.PlainLists[$Key] }
    if (-not $lv) { return }
    $prev = @{}
    $prevState = @{}
    foreach ($it in @($lv.Items)) {
        $k = [string]$it.Tag
        $prev[$k] = [bool]$it.Checked
        if ($it.SubItems.Count -gt 2 -and -not $it.SubItems[2].Tag) { $prevState[$k] = [string]$it.SubItems[2].Text }
    }
    $script:PlainRendering = $true
    $lv.BeginUpdate()
    try {
        $lv.Items.Clear()
        foreach ($p in @($script:TabPaths[$Key])) {
            $it = New-Object System.Windows.Forms.ListViewItem (Get-PathLeaf $p)
            $it.UseItemStyleForSubItems = $false
            [void]$it.SubItems.Add([string]$p)
            $st = ""
            if ($prevState.ContainsKey([string]$p)) { $st = $prevState[[string]$p] }
            [void]$it.SubItems.Add($st)
            while ($it.SubItems.Count -lt $lv.Columns.Count) { [void]$it.SubItems.Add("") }
            $it.Tag = [string]$p
            if ($prev.ContainsKey([string]$p)) { $it.Checked = $prev[[string]$p] } else { $it.Checked = $true }
            [void]$lv.Items.Add($it)
        }
    }
    finally {
        $lv.EndUpdate()
        $script:PlainRendering = $false
    }
    if ($Key -eq "upgrade") { Update-UpgradeTab }
    if ($Key -eq "addlinks" -and (Get-Command Update-WorksetRoleColumn -ErrorAction SilentlyContinue)) { Update-WorksetRoleColumn }
    Update-PlainCount $Key
}

function Get-PlainKey($List) {
    foreach ($k in @($script:Ui.PlainLists.Keys)) {
        if ([object]::ReferenceEquals($script:Ui.PlainLists[$k], $List)) { return [string]$k }
    }
    return ""
}

function Update-PlainCount([string]$Key) {
    if (-not $Key -or -not $script:Ui.PlainLists.ContainsKey($Key)) { return }
    $lv = $script:Ui.PlainLists[$Key]
    $c = 0
    foreach ($it in @($lv.Items)) { if ($it.Checked) { $c++ } }
    if ($script:Ui.PlainCounts.ContainsKey($Key)) {
        $script:Ui.PlainCounts[$Key].Text = ("В списке: {0} · отмечено: {1}" -f $lv.Items.Count, $c)
    }
    if ($Key -eq "upgrade" -and -not $script:PlainRendering) { Update-UpgradeTab }
    if (($Key -eq "addlinks" -or $Key -eq "worksets") -and (Get-Command Update-WorksetYearLabel -ErrorAction SilentlyContinue)) { Update-WorksetYearLabel }
    # tabs 4/5: a tick, a new model or a removed one re-proposes the pairs by itself
    if (($Key -eq "addlinks" -or $Key -eq "worksets") -and (Get-Command Request-WorksetRecompute -ErrorAction SilentlyContinue)) { Request-WorksetRecompute }
    Update-RunButton
}

function Set-PlainChecks([string]$Key, [bool]$On) {
    if (-not $script:Ui.PlainLists.ContainsKey($Key)) { return }
    $script:PlainRendering = $true
    try { foreach ($it in @($script:Ui.PlainLists[$Key].Items)) { $it.Checked = $On } }
    finally { $script:PlainRendering = $false }
    Update-PlainCount $Key
}

function Remove-PlainSelected([string]$Key) {
    if ($script:Run.Active -or -not $script:Ui.PlainLists.ContainsKey($Key)) { return }
    $drop = @{}
    foreach ($it in @($script:Ui.PlainLists[$Key].SelectedItems)) { $drop[[string]$it.Tag] = $true }
    if ($drop.Count -eq 0) { Set-UiStatus "Выделите строки, которые убрать из списка."; return }
    $keep = New-Object System.Collections.Generic.List[string]
    foreach ($p in @($script:TabPaths[$Key])) { if (-not $drop.ContainsKey([string]$p)) { [void]$keep.Add([string]$p) } }
    $script:TabPaths[$Key] = $keep
    Update-PlainList $Key
    Set-UiStatus ("Убрано из списка: {0}" -f $drop.Count)
}

# Tab «Смена версии Revit»: what the models are now and what they may become.
function Update-UpgradeTab {
    if (-not $script:Ui.ContainsKey("UpgradeYearLabel") -or -not $script:Ui.PlainLists.ContainsKey("upgrade")) { return }
    $lv = $script:Ui.PlainLists["upgrade"]
    $paths = @(Get-TabModelPaths "upgrade")
    $info = Get-UpgradeYearInfo $paths $script:servers
    $byPath = @{}
    foreach ($r in @($info.Rows)) { $byPath[[string]$r.Path] = $r }
    foreach ($it in @($lv.Items)) {
        # while the list creates its window, ItemChecked fires with items not yet in place
        if ($null -eq $it) { continue }
        $p = [string]$it.Tag
        $ver = ""
        $reason = ""
        if ($byPath.ContainsKey($p)) {
            $r = $byPath[$p]
            if ($r.Year) { $ver = ("Revit {0}" -f $r.Year) }
            $reason = [string]$r.Reason
        }
        if (-not $ver -and -not (Test-IsRsnPath $p)) { $ver = "сетевой файл" }
        elseif ($ver -and $byPath.ContainsKey($p) -and $byPath[$p].Source -eq "name") { $ver = $ver + " (по имени)" }
        if ($it.SubItems.Count -gt 3 -and $it.SubItems[3].Text -ne $ver) { $it.SubItems[3].Text = $ver }
        $cell = $it.SubItems[2]
        if ($reason) {
            $cell.Text = $reason
            $cell.Tag = "reason"
            $cell.ForeColor = [System.Drawing.Color]::Firebrick
        }
        elseif ($cell.Tag) {
            $cell.Text = ""
            $cell.Tag = $null
            $cell.ForeColor = [System.Drawing.SystemColors]::WindowText
        }
    }
    $lbl = $script:Ui.UpgradeYearLabel
    if ($paths.Count -eq 0) {
        $lbl.Text = "модели не выбраны"
        $lbl.ForeColor = [System.Drawing.SystemColors]::GrayText
    }
    elseif (-not $info.Ok) {
        $lbl.Text = [string]$info.Message
        $lbl.ForeColor = [System.Drawing.Color]::Firebrick
    }
    elseif ($info.Year) {
        $lbl.Text = ("Revit {0} (по моделям)" -f $info.Year)
        $lbl.ForeColor = [System.Drawing.SystemColors]::ControlText
    }
    else {
        $lbl.Text = "не определена: сетевые файлы откроет выбранная версия"
        $lbl.ForeColor = [System.Drawing.SystemColors]::ControlText
    }
    $script:UpgradeDetected = $info
    $cmb = $script:Ui.UpgradeYear
    $want = @(Get-UpgradeTargetYears $script:Years ([string]$info.Year))
    $cur = @($cmb.Items | ForEach-Object { [string]$_ })
    if (($cur -join ",") -ne ($want -join ",")) {
        $sel = [string]$cmb.SelectedItem
        $cmb.BeginUpdate()
        try {
            $cmb.Items.Clear()
            foreach ($y in $want) { [void]$cmb.Items.Add($y) }
        }
        finally { $cmb.EndUpdate() }
        if ($sel -and $cmb.Items.Contains($sel)) { $cmb.SelectedItem = $sel }
        elseif ($cmb.Items.Contains("2024")) { $cmb.SelectedItem = "2024" }
        elseif ($cmb.Items.Count -gt 0) { $cmb.SelectedIndex = 0 }
    }
    if ($cmb.SelectedItem) { $script:UpgradeDst = [string]$cmb.SelectedItem } else { $script:UpgradeDst = "" }
}

function Copy-ModelsFromTab([string]$SourceKey, [string]$TargetKey) {
    $paths = @(Get-TabModelPaths $SourceKey)
    if ($paths.Count -eq 0) { Set-UiStatus "На выбранной вкладке нет моделей."; return }
    if ($TargetKey -eq "move") {
        $n = Add-UiModels $paths
        Set-UiStatus ("Добавлено моделей: {0}" -f $n)
        return
    }
    $n = Add-PlainTabModels $TargetKey $paths
    Set-UiStatus ("Добавлено моделей: {0}" -f $n)
}

function Add-TakeMenuItem($Menu, [string]$SourceKey, [string]$TargetKey, [string]$Title) {
    $item = $Menu.Items.Add($Title)
    $item.Add_Click({ Copy-ModelsFromTab $SourceKey $TargetKey }.GetNewClosure())
}

function Show-TakeModelsMenu($Anchor, [string]$TargetKey) {
    $menu = New-Object System.Windows.Forms.ContextMenuStrip
    foreach ($pair in @(
            @("move", "Перенос и переименование"),
            @("relink", "Перепривязка связей"),
            @("upgrade", "Смена версии Revit"),
            @("addlinks", "Связи: добавить"),
            @("worksets", "Рабочие наборы")
        )) {
        if ($pair[0] -eq $TargetKey) { continue }
        Add-TakeMenuItem $menu ([string]$pair[0]) $TargetKey ([string]$pair[1])
    }
    $menu.Show($Anchor, 0, $Anchor.Height)
}

# Isolated runs of tabs 2 and 3: the job switches are set for this run only, and the
# list of tab 1 ($script:modelPaths) is given back afterwards - it is not replaced by the tab's list.
function Invoke-IsolatedTabRun([string]$Kind, [string[]]$Paths, [hashtable]$Flags) {
    $names = @("CbRelink", "CbRename", "CbMove", "CbUpgrade", "CbCopy", "CbDelete")
    $saved = @{}
    foreach ($n in $names) { $saved[$n] = [bool]$script:Ui[$n].Checked }
    $savedPaths = New-Object System.Collections.Generic.List[string]
    foreach ($p in @($script:modelPaths)) { [void]$savedPaths.Add([string]$p) }
    $savedLoc = @($script:locMode, $script:oldSeg, $script:newSeg, $script:oldPref, $script:newPref)
    $savedOps = $script:ops
    try {
        foreach ($n in $names) { $script:Ui[$n].Checked = [bool]$Flags[$n] }
        $script:locMode = "same"
        $script:oldSeg = ""; $script:newSeg = ""; $script:oldPref = ""; $script:newPref = ""
        $script:ops = @()
        $script:IsolatePaths = $Paths
        $script:IsolateKind = $Kind
        $script:RunListKey = $Kind
        Start-UiRun
    }
    finally {
        $script:IsolatePaths = $null
        $script:IsolateKind = $null
        $script:RunListKey = $null
        foreach ($n in $names) { $script:Ui[$n].Checked = [bool]$saved[$n] }
        $script:modelPaths = $savedPaths
        $script:locMode = $savedLoc[0]; $script:oldSeg = $savedLoc[1]; $script:newSeg = $savedLoc[2]
        $script:oldPref = $savedLoc[3]; $script:newPref = $savedLoc[4]
        $script:ops = $savedOps
    }
}

function Start-RelinkTabRun {
    $paths = @(Get-TabModelPaths "relink")
    if ($paths.Count -eq 0) {
        [void][System.Windows.Forms.MessageBox]::Show($script:Ui.Form, "Отметьте модели, в которых перепривязать связи.", "Перепривязка")
        return
    }
    Invoke-IsolatedTabRun "relink" $paths @{ CbRelink = $true; CbRename = $false; CbMove = $false; CbUpgrade = $false; CbCopy = $false; CbDelete = $false }
}

function Start-UpgradeTabRun {
    $paths = @(Get-TabModelPaths "upgrade")
    if ($paths.Count -eq 0) {
        [void][System.Windows.Forms.MessageBox]::Show($script:Ui.Form, "Отметьте модели для смены версии.", "Смена версии")
        return
    }
    Update-UpgradeTab
    $info = Get-UpgradeYearInfo $paths $script:servers
    if (-not $info.Ok) {
        [void][System.Windows.Forms.MessageBox]::Show($script:Ui.Form,
            ([string]$info.Message + [Environment]::NewLine + "Причина - в колонке «Состояние»."), "Смена версии")
        return
    }
    $target = ""
    if ($script:Ui.UpgradeYear.SelectedItem) { $target = [string]$script:Ui.UpgradeYear.SelectedItem }
    if (-not $target) {
        [void][System.Windows.Forms.MessageBox]::Show($script:Ui.Form, "Выберите, какой станет версия Revit.", "Смена версии")
        return
    }
    if ($info.Year -and ([int]$target -le [int]$info.Year)) {
        [void][System.Windows.Forms.MessageBox]::Show($script:Ui.Form, ("Модели уже Revit {0}: выберите версию новее." -f $info.Year), "Смена версии")
        return
    }
    $script:UpgradeSrc = [string]$info.Year
    $script:UpgradeDst = $target
    Invoke-IsolatedTabRun "upgrade" $paths @{ CbRelink = $false; CbRename = $false; CbMove = $false; CbUpgrade = $true; CbCopy = $false; CbDelete = $false }
}

function Fill-RelinkFoldersFromMove {
    $path = Join-Path $ToolDir "job_paths.csv"
    if (-not (Test-Path -LiteralPath $path)) {
        Set-UiStatus "Отчёта переноса ещё нет: сначала запустите вкладку переноса."
        return
    }
    $lines = @(Get-ModelListLines $path)
    $olds = New-Object System.Collections.Generic.List[string]
    $news = New-Object System.Collections.Generic.List[string]
    $first = $true
    foreach ($line in $lines) {
        if ($first) { $first = $false; continue }
        if ([string]::IsNullOrWhiteSpace([string]$line)) { continue }
        $cells = @(([string]$line).Split(";"))
        if ($cells.Count -lt 2) { continue }
        $oldFolder = Get-PathFolder $cells[0].Trim('"')
        $newFolder = Get-PathFolder $cells[1].Trim('"')
        if ($oldFolder) { [void]$olds.Add($oldFolder) }
        if ($newFolder) { [void]$news.Add($newFolder) }
    }
    $oldCommon = Get-CommonPathPrefix $olds.ToArray()
    $newCommon = Get-CommonPathPrefix $news.ToArray()
    if (-not $oldCommon -or -not $newCommon) {
        Set-UiStatus "В job_paths.csv нет общей папки для перепривязки."
        return
    }
    $script:Ui.RelOldBox.Text = $oldCommon
    $script:Ui.RelNewBox.Text = $newCommon
    Set-UiStatus ("Папки из переноса: {0} -> {1}" -f $oldCommon, $newCommon)
}

function Start-MoveTabRun {
    $script:Ui.CbRelink.Checked = $false
    $script:Ui.CbUpgrade.Checked = $false
    $script:RunListKey = "move"
    Start-UiRun
}

function Get-ActiveTabKey {
    if (-not $script:Ui.TopTabs -or -not $script:Ui.TopTabs.SelectedTab) { return "move" }
    $k = [string]$script:Ui.TopTabs.SelectedTab.Tag
    if (-not $k) { return "move" }
    return $k
}

function Get-RunButtonText([string]$Key) {
    switch ($Key) {
        "addlinks" { return "Вставить связи" }
        "worksets" { return "Записать наборы" }
        default { return "Запустить" }
    }
}

# The bottom bar is the same for every tab; only the text of the run button changes.
function Update-ActiveTab {
    if (-not $script:Ui.BtnRun) { return }
    $key = Get-ActiveTabKey
    $script:Ui.BtnRun.Text = Get-RunButtonText $key
    if (Get-Command Set-WorksetActiveTab -ErrorAction SilentlyContinue) { Set-WorksetActiveTab $key }
    if ($key -eq "upgrade") { Update-UpgradeTab }
    Update-RunButton
}

# «Пробный прогон» ($DryRun): the same checks and the same plan, nothing is started or written.
function Invoke-ActiveTabRun([bool]$DryRun) {
    if ($script:Run.Active) { return }
    $key = Get-ActiveTabKey
    $script:DryRunOnly = $DryRun
    try {
        switch ($key) {
            "move" { Start-MoveTabRun }
            "relink" { Start-RelinkTabRun }
            "upgrade" { Start-UpgradeTabRun }
            "addlinks" {
                $script:WsApplyPart = "insert"
                Start-WorksetApply
            }
            "worksets" {
                $script:WsApplyPart = "sets"
                Start-WorksetApply
            }
        }
    }
    catch { Show-WindowError (Get-RunButtonText $key) $_ }
    finally { $script:DryRunOnly = $false }
}

function Set-UiBusy([bool]$Busy) {
    $script:Run.Active = $Busy
    foreach ($c in @($script:Ui.BusyControls)) {
        if ($c) { $c.Enabled = -not $Busy }
    }
    if (-not $Busy) { $script:Run.List = $null }
    Update-ListCount
    $script:Ui.Progress.Visible = $Busy
}

# Modal list of what move/delete would do. "Да" only if there is something to
# apply; "Нет" is the default (Enter / Esc / closing the window).
function Show-OpsPlanDialog([string]$Title, [string]$Text, [bool]$CanApply, [string]$YesText, [bool]$InfoOnly = $false) {
    $parts = New-OpsPlanForm $Title $Text $CanApply $YesText $InfoOnly
    $f = $parts.Form
    try { $res = $f.ShowDialog($script:Ui.Form) }
    finally { $f.Dispose() }
    return ($res -eq [System.Windows.Forms.DialogResult]::Yes)
}

# The form of Show-OpsPlanDialog, built but not shown (snap_ui.ps1 checks it).
function New-OpsPlanForm([string]$Title, [string]$Text, [bool]$CanApply, [string]$YesText, [bool]$InfoOnly = $false) {
    $f = New-Object System.Windows.Forms.Form
    $f.Text = $Title
    $f.StartPosition = "CenterParent"
    $f.Font = $script:Ui.Form.Font
    $wa = [System.Windows.Forms.Screen]::FromControl($script:Ui.Form).WorkingArea
    $f.Size = New-Object System.Drawing.Size([Math]::Max(520, [Math]::Min(1000, $wa.Width - 40)), [Math]::Max(300, [Math]::Min(620, $wa.Height - 40)))
    if ($Text -and $Text.Contains("ВМЕСТЕ С ЭЛЕМЕНТАМИ")) {
        $tb = New-Object System.Windows.Forms.RichTextBox
        $tb.ReadOnly = $true
        $tb.Font = New-Object System.Drawing.Font("Consolas", 9.5)
        $tb.Dock = "Fill"
        $tb.Text = $Text
        $pos = 0
        foreach ($line in @($Text -split "`n")) {
            $len = $line.TrimEnd("`r").Length
            if ($line.StartsWith("ВМЕСТЕ С ЭЛЕМЕНТАМИ")) {
                $tb.Select($pos, $len)
                $tb.SelectionColor = [System.Drawing.Color]::Firebrick
            }
            $pos += $line.Length + 1
        }
        $tb.Select(0, 0)
    }
    else {
        $tb = New-Object System.Windows.Forms.TextBox
        $tb.Multiline = $true
        $tb.ReadOnly = $true
        $tb.ScrollBars = "Both"
        $tb.WordWrap = $false
        $tb.Font = New-Object System.Drawing.Font("Consolas", 9.5)
        $tb.Dock = "Fill"
        $tb.Text = $Text
    }
    # Buttons sit in a bar row of the common grid. They used to be placed by absolute
    # coordinates computed from the FORM width inside a 200 px panel that was not docked yet,
    # with Anchor Right: when the panel stretched, the buttons slid past its right edge and the
    # dialog showed no buttons at all. The grid keeps them inside the window at any size.
    $btnYes = New-OpsButton $(if ($YesText) { $YesText } else { "Да" })
    $btnYes.DialogResult = [System.Windows.Forms.DialogResult]::Yes
    $btnYes.Enabled = $CanApply
    $btnNo = New-OpsButton "Нет"
    $btnNo.DialogResult = [System.Windows.Forms.DialogResult]::No
    $right = @($btnYes, $btnNo)
    if ($InfoOnly) {
        # «Пробный прогон»: only a list, one button
        $btnYes.Visible = $false
        $btnNo.Text = "Закрыть"
        $right = @($btnNo)
    }
    $f.MinimumSize = New-Object System.Drawing.Size(520, 300)
    $tb.Margin = New-Object System.Windows.Forms.Padding(0)
    $grid = New-OpsColumnGrid @(0, $script:OpsRowH)
    $bar = New-OpsBar @() $right
    $bar.Margin = New-Object System.Windows.Forms.Padding(0, 8, 0, 0)
    $grid.Controls.Add($tb, 0, 0)
    $grid.Controls.Add($bar, 0, 1)
    $f.Controls.Add($grid)
    $f.AcceptButton = $btnNo
    $f.CancelButton = $btnNo
    # the read-only text is not selected on open (a focused TextBox selects all of it)
    $f.ActiveControl = $btnNo
    $tb.TabStop = $false
    return @{ Form = $f; Yes = $btnYes; No = $btnNo; Text = $tb }
}

# ---- per-model progress (_progress_ops.txt written by update_rvt_links.py) ----
$script:FinalStatesRu = @("Готово", "Частично", "Ошибка", "Без изменений")

function Get-OpsProgressFile {
    return (Join-Path (Initialize-OpsReportsDir $ToolDir) "_progress_ops.txt")
}

function Get-OpsStateRu([string]$Code) {
    switch ($Code.Trim().ToUpperInvariant()) {
        "STARTED" { return "Идёт" }
        "OK" { return "Готово" }
        "PARTIAL" { return "Частично" }
        "ERROR" { return "Ошибка" }
        "NO_CHANGES" { return "Без изменений" }
        default { return "" }
    }
}

function Get-OpsPathKey([string]$Path) {
    if ([string]::IsNullOrWhiteSpace($Path)) { return "" }
    return $Path.Trim().Replace("\", "/").TrimEnd("/").ToUpperInvariant()
}

function Reset-OpsProgress {
    $f = Get-OpsProgressFile
    try { if (Test-Path -LiteralPath $f) { Remove-Item -LiteralPath $f -Force -ErrorAction Stop } } catch { }
    $script:Run.ProgressText = ""
}

function Read-OpsProgressText([string]$File) {
    if (-not (Test-Path -LiteralPath $File)) { return "" }
    $fs = $null
    try {
        $fs = New-Object System.IO.FileStream($File, [System.IO.FileMode]::Open,
            [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
        $sr = New-Object System.IO.StreamReader($fs, [System.Text.Encoding]::UTF8)
        return $sr.ReadToEnd()
    }
    catch { return $null }
    finally { if ($fs) { $fs.Dispose() } }
}

function Set-RowStateIfChanged($Item, [string]$Text) {
    if ($Item.SubItems.Count -gt 2 -and $Item.SubItems[2].Text -ne $Text) { $Item.SubItems[2].Text = $Text }
}

function Set-UnfinishedState([string]$Text) {
    foreach ($it in @((Get-RunList).Items)) {
        if (-not $it.Checked -or $it.SubItems.Count -le 2) { continue }
        if ($script:FinalStatesRu -notcontains [string]$it.SubItems[2].Text) { Set-RowStateIfChanged $it $Text }
    }
}

function Update-OpsProgress([string]$File = "") {
    if (-not $File) { $File = Get-OpsProgressFile }
    $text = Read-OpsProgressText $File
    if ($null -eq $text) { return }
    $last = @{}
    $order = New-Object System.Collections.Generic.List[string]
    foreach ($ln in ($text -split "`r?`n")) {
        $parts = $ln.Split("`t")
        if ($parts.Count -lt 3) { continue }
        $ru = Get-OpsStateRu $parts[0]
        if (-not $ru) { continue }
        $key = Get-OpsPathKey ($parts[2..($parts.Count - 1)] -join "`t")
        if (-not $key) { continue }
        $last[$key] = $ru
        [void]$order.Add($key)
    }
    $total = 0; $done = 0; $curName = ""; $curKey = ""
    for ($i = $order.Count - 1; $i -ge 0; $i--) {
        if ($last[$order[$i]] -eq "Идёт") { $curKey = $order[$i]; break }
    }
    foreach ($it in @((Get-RunList).Items)) {
        if (-not $it.Checked) { continue }
        $total++
        $keys = @((Get-OpsPathKey ([string]$it.Tag)))
        if ($it.SubItems.Count -gt 5) {
            $keys += (Get-OpsPathKey $it.SubItems[4].Text)
            $keys += (Get-OpsPathKey $it.SubItems[5].Text)
        }
        $st = $null
        foreach ($k in $keys) { if ($k -and $last.ContainsKey($k)) { $st = $last[$k]; if ($k -eq $curKey) { $curName = $it.Text }; break } }
        if (-not $st) { $st = "В очереди" }
        Set-RowStateIfChanged $it $st
        if ($script:FinalStatesRu -contains $st) { $done++ }
    }
    $script:Ui.Progress.Maximum = [Math]::Max(1, $total)
    $script:Ui.Progress.Value = [Math]::Min($done, [Math]::Max(1, $total))
    $n = $done
    if ($curName) { $n = [Math]::Min($done + 1, $total) }
    $label = ("{0} — модель {1} из {2}" -f $script:Run.PassLabel, $n, $total)
    if ($curName) { $label += (" ({0})" -f $curName) }
    if ($script:Ui.Status -and $script:Ui.Status.Text -ne $label) { $script:Ui.Status.Text = $label }
}

function Start-NextPass {
    if ($script:Run.Index -ge $script:Run.Queue.Count) {
        Complete-OpsRun $true
        return
    }
    $pass = $script:Run.Queue[$script:Run.Index]
    $extra = @()
    if ($pass.Kind -eq "move" -or $pass.Kind -eq "delete") {
        # dry-run first: the script changes nothing without -Execute
        Set-UiStatus "Проверяю, что будет сделано…"
        $dry = Invoke-OpsPsDryRun -Pass $pass
        if (-not $dry.Ok) {
            $msg = [string]$dry.Text
            if ($msg.Length -gt 1500) { $msg = $msg.Substring(0, 1500) + "…" }
            Set-UiStatus "Предпросмотр не удался"
            [void][System.Windows.Forms.MessageBox]::Show($script:Ui.Form,
                ("Предпросмотр не удался (код {0}).{1}{1}{2}" -f $dry.Code, [Environment]::NewLine, $msg),
                "Запуск", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            Complete-OpsRun $false
            return
        }
        $isMove = ($pass.Kind -eq "move")
        $title = $(if ($isMove) { "Перенос файлов" } else { "Старые модели: убрать в _to_delete" })
        $yesText = $(if ($isMove) { "Да, перенести" } else { "Да, убрать в _to_delete" })
        $head = $(if ($isMove) {
                "Файлы будут перенесены. Существующие файлы назначения не перезаписываются" + $(if ($pass.Overwrite) { " (включена перезапись: старый файл уйдёт в _to_delete)" } else { "" }) + "."
            }
            else {
                "Старые модели не удаляются: переносятся в папку _to_delete рядом с файлом, откат - вернуть файл. Переносятся только модели с OK в отчётах и непустым новым файлом."
            })
        $planText = (([string]$dry.Text -split "\r?\n") | Where-Object { $_ -notmatch '^(DRY-RUN: nothing|Log: )' }) -join [Environment]::NewLine
        $apply = Show-OpsPlanDialog -Title $title -Text ($head + [Environment]::NewLine + [Environment]::NewLine + $planText) `
            -CanApply ([int]$dry.Count -gt 0) -YesText $yesText
        if (-not $apply) {
            if ($isMove) {
                # relink/delete need the moved files: do not continue
                Set-UiStatus "Перенос не подтверждён - прогон остановлен"
                Complete-OpsRun $false
                return
            }
            Set-UiStatus "Очистка пропущена"
            $script:Run.Index++
            Start-NextPass
            return
        }
        $extra = @("-Execute", "-Yes")
        if ($isMove -and $pass.Overwrite) { $extra += "-Overwrite" }
    }
    $script:Run.PassLabel = ("Проход {0} из {1}: {2}" -f ($script:Run.Index + 1), $script:Run.Queue.Count, (Get-PassTitle $pass.Name))
    Set-UiStatus $script:Run.PassLabel
    $script:Run.Track = (Test-OpsPassTracked ([string]$pass.Name))
    if ($script:Run.Track) {
        Reset-OpsProgress
        Set-CheckedState "В очереди"
        Update-OpsProgress
    }
    else {
        Set-CheckedState "Идёт"
        $script:Ui.Progress.Maximum = [Math]::Max(1, $script:Run.Queue.Count)
        $script:Ui.Progress.Value = [Math]::Min($script:Run.Index, $script:Ui.Progress.Maximum)
    }
    if ($pass.Kind -eq "rbp") {
        $started = Start-OpsRbpPass -Pass $pass -Exe $script:Run.Exe -ExeDir $script:Run.ExeDir
    }
    else {
        $started = Start-OpsPsPass -Pass $pass -ExtraArgs $extra
    }
    if (-not $started.Ok -or -not $started.Process) {
        Set-UiStatus $started.Message
        [void][System.Windows.Forms.MessageBox]::Show($script:Ui.Form, $started.Message, "Запуск",
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        Complete-OpsRun $false
        return
    }
    $script:Run.Process = $started.Process
}

function Complete-OpsRun([bool]$Ok) {
    if ($script:JobBackup) {
        Restore-OpsJobFiles $script:JobBackup $ToolDir
        $script:JobBackup = $null
    }
    if ($script:Ui.Timer) { $script:Ui.Timer.Stop() }
    $script:Run.Process = $null
    if ($Ok) { Set-UnfinishedState "Готово" }
    Set-UiBusy $false
    $xlsx = Find-LatestOpsReport
    $script:Run.LastReport = $xlsx
    if ($Ok) {
        if ($xlsx) { Set-UiStatus ("Готово. Отчёт: {0}" -f [System.IO.Path]::GetFileName([string]$xlsx)) ("Готово. Отчёт: {0}" -f $xlsx) } else { Set-UiStatus "Готово." }
    }
    $script:Ui.Progress.Value = $script:Ui.Progress.Maximum
    if ($Ok -and $script:Run.ReadImport) {
        $script:Run.ReadImport = $false
        if (Get-Command Import-WorksetReadReport -ErrorAction SilentlyContinue) { Import-WorksetReadReport }
    }
    if ($Ok -and $script:Run.ApplyImport) {
        $script:Run.ApplyImport = $false
        if (Get-Command Import-WorksetApplyReport -ErrorAction SilentlyContinue) { Import-WorksetApplyReport }
    }
}

function Tick-OpsRun {
    if (-not $script:Run.Active) { return }
    $proc = $script:Run.Process
    if (-not $proc) { return }
    try { $proc.Refresh() } catch { }
    if ($script:Run.Track) { try { Update-OpsProgress } catch { } }
    if (-not $proc.HasExited) { return }
    $code = 0
    try { $code = [int]$proc.ExitCode } catch { }
    $script:Run.ExitCodes += $code
    $script:Run.Process = $null
    if ($code -ne 0) {
        Set-UnfinishedState "Ошибка"
        Set-UiStatus ("Проход остановился, код {0}." -f $code)
        Complete-OpsRun $false
        return
    }
    Set-UnfinishedState "Готово"
    $script:Run.Index++
    Start-NextPass
}

# Text of «Пробный прогон»: the passes and every model old -> new. Nothing is started.
function Get-DryRunText($Job, $Rows, [string]$Plan) {
    $nl = [Environment]::NewLine
    $lines = New-Object System.Collections.Generic.List[string]
    [void]$lines.Add("Пробный прогон: ничего не запускается и не записывается.")
    [void]$lines.Add("")
    [void]$lines.Add($Plan)
    [void]$lines.Add("")
    foreach ($r in @($Rows)) {
        if ($Job.relink -and -not $Job.rename -and [string]$Job.location_mode -eq "same" -and -not $Job.upgrade) {
            [void]$lines.Add(("{0}{1}   связи: {2} -> {3}" -f [string]$r.old_path, $nl, [string]$Job.relink_old, [string]$Job.relink_new))
        }
        elseif ([string]$r.old_path -ne [string]$r.new_path) {
            [void]$lines.Add(("{0}{1}   -> {2}" -f [string]$r.old_path, $nl, [string]$r.new_path))
        }
        else {
            [void]$lines.Add(("{0}{1}   (путь не меняется)" -f [string]$r.old_path, $nl))
        }
    }
    if ($Job.delete_old) { [void]$lines.Add(""); [void]$lines.Add("Старые модели после прогона уйдут в _to_delete (их можно вернуть).") }
    elseif ($Job.rename -or [string]$Job.location_mode -ne "same") { [void]$lines.Add(""); [void]$lines.Add("Старые модели остаются на месте.") }
    return ($lines -join $nl)
}

function Start-UiRun {
    $script:Run.ReadImport = $false
    $script:Run.ApplyImport = $false
    $isolated = [bool]$script:IsolateKind
    if (-not $isolated) {
        $st = Sync-MoveFlags
        if ($st -and $st.RuleMessage) {
            [void][System.Windows.Forms.MessageBox]::Show($script:Ui.Form, ("Имя: {0}" -f $st.RuleMessage), "Запуск")
            return
        }
        if ($st -and $st.MoveMessage) {
            [void][System.Windows.Forms.MessageBox]::Show($script:Ui.Form, $st.MoveMessage, "Запуск")
            return
        }
    }
    $checked = @(Get-CheckedPaths)
    if ($checked.Count -eq 0) {
        [void][System.Windows.Forms.MessageBox]::Show($script:Ui.Form, "Отметьте модели галочками.", "Запуск")
        return
    }
    if (-not $isolated -and -not $script:Ui.CbRename.Checked -and -not $script:Ui.CbMove.Checked) {
        [void][System.Windows.Forms.MessageBox]::Show($script:Ui.Form,
            "Делать нечего: заполните «Имя» (как сейчас → как станет) и/или «Папка» (было → стало).", "Запуск")
        return
    }
    if ($script:Ui.CbMove.Checked -and $script:locMode -ne "same" -and
        (Test-MoveSettingsMissing ([bool]$script:Ui.CbMove.Checked) $script:locMode $script:oldSeg $script:oldPref)) {
        [void][System.Windows.Forms.MessageBox]::Show($script:Ui.Form, "Папка: укажите «было» и «стало».", "Запуск")
        return
    }
    $keep = New-Object System.Collections.Generic.List[string]
    foreach ($p in $checked) { [void]$keep.Add($p) }
    $script:modelPaths = $keep

    $rows = @(Get-PreviewRows)
    $badRows = @($rows | Where-Object { $_.path_error })
    if ($badRows.Count -gt 0) {
        $lines = @(Get-PathErrorLines $rows 5)
        $more = $(if ($badRows.Count -gt $lines.Count) { [Environment]::NewLine + ("...и ещё {0}" -f ($badRows.Count - $lines.Count)) } else { "" })
        [void][System.Windows.Forms.MessageBox]::Show($script:Ui.Form,
            ("У моделей с ошибкой пути запуск невозможен ({0}):" -f $badRows.Count) + [Environment]::NewLine + [Environment]::NewLine +
            ($lines -join ([Environment]::NewLine + [Environment]::NewLine)) + $more,
            "Запуск", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    $dup = @($rows | Where-Object { $_.note -match "duplicate" })
    if ($dup.Count -gt 0) {
        [void][System.Windows.Forms.MessageBox]::Show($script:Ui.Form, "В списке повторяются новые пути. Поправьте имена или папку.", "Запуск")
        return
    }
    $job = Build-JobObject $rows
    $plan = Get-JobPlanText $job $rows
    if (Test-UiRelinkOnly) {
        $ro = [string]$job.relink_old; $rn = [string]$job.relink_new
        if (-not $ro -or -not $rn) {
            [void][System.Windows.Forms.MessageBox]::Show($script:Ui.Form, "Укажите «Старая папка» и «Новая папка» для связей.", "Запуск")
            return
        }
        if ((ConvertTo-PrefixKey $ro) -eq (ConvertTo-PrefixKey $rn)) {
            [void][System.Windows.Forms.MessageBox]::Show($script:Ui.Form, "Старая и новая папка совпадают.", "Запуск")
            return
        }
        $plan = ("Открываю отмеченные модели: {0}." -f $checked.Count) + [Environment]::NewLine +
        ("Связи из {0}" -f $ro) + [Environment]::NewLine + ("перевожу в {0}" -f $rn) + [Environment]::NewLine +
        "Связи вне старой папки не трогаю." + [Environment]::NewLine + $plan
    }
    if ($script:DryRunOnly) {
        [void](Show-OpsPlanDialog -Title "Пробный прогон" -Text (Get-DryRunText $job $rows $plan) -CanApply $false -YesText "" -InfoOnly $true)
        Set-UiStatus "Пробный прогон: ничего не запускалось."
        return
    }
    $ask = [System.Windows.Forms.MessageBox]::Show(
        $script:Ui.Form,
        ($plan + [Environment]::NewLine + [Environment]::NewLine + "Все вышли из моделей. Продолжить?"),
        "Запуск",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning,
        [System.Windows.Forms.MessageBoxDefaultButton]::Button2
    )
    if ($ask -ne [System.Windows.Forms.DialogResult]::Yes) {
        Set-UiStatus "Запуск отменён"
        return
    }

    $exe = Find-BatchRvt
    $passes = @(New-OpsPassQueue -Job $job -ToolDir $ToolDir -CodeDir $CodeDir -Rows $rows)
    $needRbp = @($passes | Where-Object { $_.Kind -eq "rbp" }).Count -gt 0
    if ($needRbp -and -not $exe) {
        [void][System.Windows.Forms.MessageBox]::Show(
            $script:Ui.Form,
            "На этом ПК не найден Revit Batch Processor (BatchRvt.exe).",
            "Запуск",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return
    }

    Reset-OpsReportSession
    if ($script:IsolateKind) { $script:JobBackup = Backup-OpsJobFiles $ToolDir }
    [void](Save-JobFiles -ToolDir $ToolDir -Rows $rows -Job $job)
    if ($script:IsolateKind -eq "relink") {
        Copy-Item -LiteralPath (Join-Path $ToolDir "job.json") -Destination (Join-Path $ToolDir "job_relink.json") -Force
        $srcList = Join-Path $ToolDir "rvt_list_new.txt"
        if (Test-Path -LiteralPath $srcList) {
            Copy-Item -LiteralPath $srcList -Destination (Join-Path $ToolDir "rvt_list_relink.txt") -Force
        }
    }
    elseif ($script:IsolateKind -eq "upgrade") {
        Copy-Item -LiteralPath (Join-Path $ToolDir "job.json") -Destination (Join-Path $ToolDir "job_upgrade.json") -Force
    }
    $script:Run.List = $script:Ui.List
    if ($script:RunListKey -and $script:RunListKey -ne "move" -and $script:Ui.PlainLists.ContainsKey($script:RunListKey)) {
        $script:Run.List = $script:Ui.PlainLists[$script:RunListKey]
    }
    if ($script:IsolateKind) {
        Set-UiStatus "Запуск отдельной вкладки. Список переноса не менялся."
        Set-CheckedState "Ожидание"
    }
    else {
        Refresh-ModelList
        Set-CheckedState "Ожидание"
    }

    $script:Run.Queue = $passes
    $script:Run.Index = 0
    $script:Run.ExitCodes = @()
    if ($exe) {
        $script:Run.Exe = $exe
        $script:Run.ExeDir = Split-Path -Parent $exe
    }
    $script:Ui.Progress.Maximum = [Math]::Max(1, $passes.Count)
    $script:Ui.Progress.Value = 0
    Set-UiBusy $true
    $script:Ui.Timer.Start()
    Start-NextPass
}

# ---- form ----
# One grid for every tab (see ui_grid.ps1): top toolbar, the table, the row under the table,
# the settings of the tab, and one bottom bar for the whole window (status | Пробный прогон | run).
$form = New-Object System.Windows.Forms.Form
$form.Text = "Операции с моделями Revit"
$form.Font = $script:OpsFont
$form.Size = New-Object System.Drawing.Size(1320, 860)
# smallest size at which every control is fully visible (checked by служебное\snap_ui.ps1)
$form.MinimumSize = New-Object System.Drawing.Size(1180, 760)
$form.StartPosition = "CenterScreen"
$form.Padding = New-Object System.Windows.Forms.Padding(8, 8, 8, 0)

# job switches read by Build-JobObject / New-OpsPassQueue: set from the fields, not shown
$cbRename = New-OpsFlag
$cbMove = New-OpsFlag
$cbUpgrade = New-OpsFlag
$cbRelink = New-OpsFlag
$cbCopy = New-OpsFlag $true
$cbDelete = New-OpsFlag

$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 1000
$fieldTimer = New-Object System.Windows.Forms.Timer
$fieldTimer.Interval = 400

# ---- bottom bar (whole window): status line | Пробный прогон | run button of the tab
$status = New-Object System.Windows.Forms.Label
$status.AutoSize = $false
$status.AutoEllipsis = $true
$status.Dock = "Fill"
$status.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$status.Margin = New-Object System.Windows.Forms.Padding(0)
# the whole text of the status line (with paths and notes) is its tool tip
$statusTip = New-Object System.Windows.Forms.ToolTip
$statusTip.AutoPopDelay = 30000
$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Dock = "Bottom"
$progress.Height = 6
$progress.Visible = $false
$statusPanel = New-Object System.Windows.Forms.Panel
$statusPanel.Margin = New-Object System.Windows.Forms.Padding(0, 0, 12, 0)
$statusPanel.Controls.Add($status)
$statusPanel.Controls.Add($progress)
$status.BringToFront()
$btnDry = New-OpsButton "Пробный прогон"
$btnRun = New-OpsButton "Запустить"
$btnRun.Font = New-Object System.Drawing.Font($script:OpsFont, [System.Drawing.FontStyle]::Bold)
$runW = 0
foreach ($t in @("Запустить", "Вставить связи", "Записать наборы")) {
    $w = (Get-OpsTextWidth $t $btnRun.Font) + 40
    if ($w -gt $runW) { $runW = $w }
}
$btnRun.Width = [Math]::Max(170, $runW)     # the button never moves when its text changes
$bottomBar = New-OpsBar @() @($btnDry, $btnRun) $statusPanel
$bottomHost = New-Object System.Windows.Forms.Panel
$bottomHost.Dock = "Bottom"
$bottomHost.Height = $script:OpsRowH + 12
$bottomHost.Padding = New-Object System.Windows.Forms.Padding(16, 4, 16, 8)
$bottomHost.Controls.Add($bottomBar)

$tabs = New-Object System.Windows.Forms.TabControl
$tabs.Dock = "Fill"
$tabs.Padding = New-Object System.Drawing.Point(14, 6)
$form.Controls.Add($tabs)
$form.Controls.Add($bottomHost)
$tabs.BringToFront()

$script:Ui = @{
    Form         = $form
    Status       = $status
    StatusTip    = $statusTip
    Progress     = $progress
    Timer        = $timer
    FieldTimer   = $fieldTimer
    BtnRun       = $btnRun
    BtnDry       = $btnDry
    CbRename     = $cbRename
    CbMove       = $cbMove
    CbUpgrade    = $cbUpgrade
    CbRelink     = $cbRelink
    CbCopy       = $cbCopy
    CbDelete     = $cbDelete
    TopTabs      = $tabs
    PlainLists   = @{}
    PlainCounts  = @{}
    Pages        = @{}
    BusyControls = (New-Object System.Collections.ArrayList)
}
$script:UiOwner = $form
$script:TabPaths = @{
    relink   = (New-Object System.Collections.Generic.List[string])
    upgrade  = (New-Object System.Collections.Generic.List[string])
    addlinks = (New-Object System.Collections.Generic.List[string])
    worksets = (New-Object System.Collections.Generic.List[string])
}
[void]$script:Ui.BusyControls.Add($btnRun)
[void]$script:Ui.BusyControls.Add($btnDry)

function Add-OpsBusy($Items) {
    foreach ($c in @($Items)) { if ($c) { [void]$script:Ui.BusyControls.Add($c) } }
}

# Model list + its two rows of buttons. The same on every tab.
function New-OpsModelPane([string]$Key, $List, [bool]$WithYear = $false) {
    $p = @{
        Files   = (New-OpsButton "Файлы…")
        Folders = (New-OpsButton "Папки…")
        Server  = (New-OpsButton "Revit Server…")
        Take    = (New-OpsButton "Взять с вкладки…")
        Open    = (New-OpsButton "Открыть список…")
        Save    = (New-OpsButton "Сохранить список…")
        All     = (New-OpsButton "Выделить все")
        None    = (New-OpsButton "Снять")
        Del     = (New-OpsButton "Убрать")
        Count   = (New-OpsCountLabel)
        List    = $List
    }
    $top = New-OpsBar @($p.Files, $p.Folders, $p.Server, $p.Take) @($p.Open, $p.Save)
    # «Связи: добавить» / «Рабочие наборы»: the Revit year of the ticked models, next to the counter
    if ($WithYear) { $p.Year = New-OpsInfoLabel "Год Revit не определён" }
    $under = New-OpsBar @($p.All, $p.None, $p.Del) @($p.Year, $p.Count)
    $grid = New-OpsColumnGrid @($script:OpsRowH, 0, $script:OpsRowH) 0
    $grid.Controls.Add($top, 0, 0)
    $grid.Controls.Add($List, 0, 1)
    $grid.Controls.Add($under, 0, 2)
    $p.Grid = $grid
    Add-OpsBusy @($p.Files, $p.Folders, $p.Server, $p.Take, $p.Open, $p.Save, $p.All, $p.None, $p.Del)
    $p.Take.Add_Click({ param($s, $e) Show-TakeModelsMenu $s ([string]$s.Tag) })
    $p.Take.Tag = $Key
    return $p
}

function New-OpsPage([string]$Title, [string]$Key) {
    $page = New-Object System.Windows.Forms.TabPage
    $page.Text = $Title
    $page.Tag = $Key
    $page.UseVisualStyleBackColor = $true
    $page.Padding = New-Object System.Windows.Forms.Padding(0)
    $script:Ui.Pages[$Key] = $page
    return $page
}

# Page = [model pane] over [settings grid]
function Set-OpsPageContent($Page, $Pane, $Settings) {
    $h = (Get-OpsSettingsHeight $Settings) + $script:OpsGap
    $grid = New-OpsColumnGrid @(0, $h)
    $grid.Controls.Add($Pane.Grid, 0, 0)
    $Settings.Margin = New-Object System.Windows.Forms.Padding(0, $script:OpsGap, 0, 0)
    $grid.Controls.Add($Settings, 0, 1)
    $Page.Controls.Add($grid)
}

# Page = split (mouse-resizable): [model pane] / [bottom host filled by workset_ui.ps1]
function New-OpsSplitPage($Page, $Pane, [string]$Key) {
    $split = New-Object System.Windows.Forms.SplitContainer
    $split.Size = New-Object System.Drawing.Size(1200, 800)
    $split.Orientation = "Horizontal"
    $split.SplitterWidth = 6
    $split.Panel1MinSize = 200
    $split.Panel2MinSize = 240
    $split.SplitterDistance = 340
    $split.Dock = "Fill"
    $split.Tag = $Key
    $top = New-OpsColumnGrid @(0) 0
    $top.Padding = New-Object System.Windows.Forms.Padding($script:OpsGap, $script:OpsGap, $script:OpsGap, 2)
    $top.Controls.Add($Pane.Grid, 0, 0)
    $split.Panel1.Controls.Add($top)
    $bottom = New-Object System.Windows.Forms.Panel
    $bottom.Dock = "Fill"
    $bottom.Padding = New-Object System.Windows.Forms.Padding($script:OpsGap, 4, $script:OpsGap, $script:OpsGap)
    $split.Panel2.Controls.Add($bottom)
    $Page.Controls.Add($split)
    $split.Add_SplitterMoved({ param($s, $e) Save-OpsSplitter $s })
    return @{ Split = $split; Bottom = $bottom }
}

# ui_layout.cfg: "addlinks=<px>" / "worksets=<px>" (an old single number = addlinks)
function Read-OpsLayoutCfg {
    $out = @{}
    $cfg = Join-Path $ToolDir "ui_layout.cfg"
    if (-not (Test-Path -LiteralPath $cfg)) { return $out }
    foreach ($line in @(Get-ModelListLines $cfg)) {
        $t = ([string]$line).Trim().TrimStart([char]0xFEFF)
        if (-not $t) { continue }
        $k = "addlinks"; $v = $t
        $eq = $t.IndexOf("=")
        if ($eq -gt 0) { $k = $t.Substring(0, $eq).Trim(); $v = $t.Substring($eq + 1).Trim() }
        $n = 0
        if ([int]::TryParse($v, [ref]$n)) { $out[$k] = $n }
    }
    return $out
}

function Save-OpsSplitter($Split) {
    if ($script:OpsBuildOnly -or $script:OpsLayoutLoading) { return }
    try {
        $all = Read-OpsLayoutCfg
        $all[[string]$Split.Tag] = [int]$Split.SplitterDistance
        $lines = New-Object System.Collections.Generic.List[string]
        foreach ($k in @($all.Keys | Sort-Object)) { [void]$lines.Add(("{0}={1}" -f $k, $all[$k])) }
        $utf8 = New-Object System.Text.UTF8Encoding $true
        [System.IO.File]::WriteAllLines((Join-Path $ToolDir "ui_layout.cfg"), $lines.ToArray(), $utf8)
    }
    catch { }
}

function Restore-OpsSplitters {
    $saved = @{}
    if (-not $script:OpsBuildOnly) { $saved = Read-OpsLayoutCfg }
    $script:OpsLayoutLoading = $true
    try {
        foreach ($sp in @($script:Ui.Splits)) {
            $h = $sp.Height
            $max = $h - $sp.Panel2MinSize - $sp.SplitterWidth
            $want = [int]($h * 0.42)
            $k = [string]$sp.Tag
            if ($saved.ContainsKey($k)) { $want = [int]$saved[$k] }
            if ($want -gt $max) { $want = $max }
            if ($want -lt $sp.Panel1MinSize) { $want = $sp.Panel1MinSize }
            if ($want -gt 0 -and $max -ge $sp.Panel1MinSize) { $sp.SplitterDistance = $want }
        }
    }
    catch { }
    finally { $script:OpsLayoutLoading = $false }
}

# ---- tab 1: Перенос и переименование
$pageMove = New-OpsPage "Перенос и переименование" "move"
$list = New-OpsListView
$list.CheckBoxes = $true
[void]$list.Columns.Add("Модель", 200)          # 0
[void]$list.Columns.Add("Станет", 200)          # 1
[void]$list.Columns.Add("Состояние", 160)       # 2 (shown last, takes the rest)
[void]$list.Columns.Add("Откуда", 0)            # 3 (hidden)
[void]$list.Columns.Add("Старый путь", 280)     # 4
[void]$list.Columns.Add("Новый путь", 280)      # 5
# shown as: Модель | Станет | Старый путь | Новый путь | Состояние
$list.Columns[4].DisplayIndex = 2
$list.Columns[5].DisplayIndex = 3
$list.Columns[2].DisplayIndex = 4
$list.Columns[3].DisplayIndex = 5
Hide-OpsColumn $list.Columns[3]
Register-OpsFillColumn $list 2 @(0, 1, 4, 5)
$movePane = New-OpsModelPane "move" $list
$lblCount = $movePane.Count

$ruleOld = New-OpsTextBox "имя сейчас (пример)"
$ruleNew = New-OpsTextBox "как должно называться"
$folderOld = New-OpsTextBox "что заменить в пути"
$folderNew = New-OpsTextBox "на что заменить"
$btnBrowseMove = New-OpsButton "Обзор…"
$rbKeep = New-OpsRadio "оставить" $true
$rbDelete = New-OpsRadio "удалить после прогона (перенос в _to_delete)"
$cbOverwrite = New-OpsCheck "Разрешить перезапись существующих файлов"
$moveSettings = New-OpsSettingsGrid
[void](Add-OpsSettingRow $moveSettings "Имя" (New-OpsPairField $ruleOld $ruleNew @((New-OpsSpacer $btnBrowseMove.Width))))
[void](Add-OpsSettingRow $moveSettings "Папка" (New-OpsPairField $folderOld $folderNew @($btnBrowseMove)))
[void](Add-OpsSettingRow $moveSettings "Старые модели" (New-OpsInlineField @($rbKeep, $rbDelete)))
[void](Add-OpsSettingRow $moveSettings "Перезапись" (New-OpsInlineField @($cbOverwrite)))
Set-OpsPageContent $pageMove $movePane $moveSettings
[void]$tabs.TabPages.Add($pageMove)
Add-OpsBusy @($ruleOld, $ruleNew, $folderOld, $folderNew, $btnBrowseMove, $rbKeep, $rbDelete, $cbOverwrite)

$script:Ui.List = $list
$script:Ui.CountLabel = $lblCount
$script:Ui.MovePane = $movePane
$script:Ui.RuleOld = $ruleOld
$script:Ui.RuleNew = $ruleNew
$script:Ui.FolderOld = $folderOld
$script:Ui.FolderNew = $folderNew
$script:Ui.CbOverwrite = $cbOverwrite
$script:Ui.RbKeep = $rbKeep
$script:Ui.RbDelete = $rbDelete
$btnFiles = $movePane.Files
$btnFolders = $movePane.Folders
$btnServer = $movePane.Server
$btnAll = $movePane.All
$btnNone = $movePane.None
$btnDel = $movePane.Del
$btnExport = $movePane.Save
$btnImport = $movePane.Open

# ---- plain tabs: own model list, own settings
function New-PlainOpsList([string]$Key) {
    $lv = New-OpsListView
    $lv.CheckBoxes = $true
    [void]$lv.Columns.Add("Модель", 260)        # 0
    [void]$lv.Columns.Add("Путь", 400)          # 1
    [void]$lv.Columns.Add("Состояние", 170)     # 2
    if ($Key -eq "upgrade") {
        [void]$lv.Columns.Add("Версия сейчас", 190)     # 3
        $lv.Columns[3].DisplayIndex = 1
    }
    elseif ($Key -eq "addlinks") {
        [void]$lv.Columns.Add("Роль", 110)              # 3: новая / целевая
        $lv.Columns[3].DisplayIndex = 1
    }
    # Путь has a fixed width, Состояние (shown last) takes the rest
    Register-OpsFillColumn $lv 2 @(0, 1)
    $script:Ui.PlainLists[$Key] = $lv
    return $lv
}

function Connect-PlainPane($Pane, [string]$Key) {
    $script:Ui.PlainCounts[$Key] = $Pane.Count
    foreach ($b in @($Pane.Files, $Pane.Folders, $Pane.Server, $Pane.Open, $Pane.Save, $Pane.All, $Pane.None, $Pane.Del)) { $b.Tag = $Key }
    $Pane.Files.Add_Click({
            param($s, $e)
            $picked = @(Select-LocalFiles)
            $n = Add-PlainTabModels ([string]$s.Tag) $picked
            Set-UiStatus ("Добавлено файлов: {0}" -f $n)
        })
    $Pane.Folders.Add_Click({
            param($s, $e)
            $folder = Show-FolderPathDialog
            if ([string]::IsNullOrWhiteSpace($folder)) { return }
            $picked = @(Get-ModelsFromFolder $folder)
            $n = Add-PlainTabModels ([string]$s.Tag) $picked
            Set-UiStatus ("Из папки добавлено: {0}" -f $n)
        })
    $Pane.Server.Add_Click({
            param($s, $e)
            $serverObj = Show-ServerPick
            if (-not $serverObj) { return }
            $picked = @(Show-RevitServerBrowser -ServerObj $serverObj)
            $n = Add-PlainTabModels ([string]$s.Tag) $picked
            Set-UiStatus ("С сервера добавлено: {0}" -f $n)
        })
    $Pane.Save.Add_Click({
            param($s, $e)
            $key = [string]$s.Tag
            $paths = @(Get-TabModelPaths $key)
            if ($paths.Count -eq 0) { Set-UiStatus "Список пуст: сохранять нечего."; return }
            $dlg = New-Object System.Windows.Forms.SaveFileDialog
            $dlg.Filter = "Список моделей (*.txt)|*.txt"
            $dlg.InitialDirectory = Get-ModelListStore $ToolDir
            $dlg.FileName = ("список_моделей_{0}.txt" -f (Get-Date -Format "yyyy-MM-dd"))
            if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
            $utf8 = New-Object System.Text.UTF8Encoding $true
            [System.IO.File]::WriteAllLines($dlg.FileName, $paths, $utf8)
            Save-ModelListFolder $ToolDir $dlg.FileName
            Set-UiStatus ("Список сохранён: {0}" -f $dlg.FileName)
        })
    $Pane.Open.Add_Click({
            param($s, $e)
            $dlg = New-Object System.Windows.Forms.OpenFileDialog
            $dlg.Filter = "Список моделей (*.txt)|*.txt|Все файлы (*.*)|*.*"
            $dlg.InitialDirectory = Get-ModelListStore $ToolDir
            if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
            Save-ModelListFolder $ToolDir $dlg.FileName
            $n = Add-PlainTabModels ([string]$s.Tag) (Read-ModelListPaths $dlg.FileName)
            Set-UiStatus ("Добавлено из списка: {0}" -f $n)
        })
    $Pane.All.Add_Click({ param($s, $e) Set-PlainChecks ([string]$s.Tag) $true })
    $Pane.None.Add_Click({ param($s, $e) Set-PlainChecks ([string]$s.Tag) $false })
    $Pane.Del.Add_Click({ param($s, $e) Remove-PlainSelected ([string]$s.Tag) })
    $Pane.List.Add_ItemChecked({
            param($s, $e)
            if ($script:PlainRendering) { return }
            Update-PlainCount (Get-PlainKey $s)
        })
    $Pane.List.Add_KeyDown({
            param($s, $e)
            if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Delete) { Remove-PlainSelected (Get-PlainKey $s) }
        })
}

# ---- tab 2: Перепривязка связей
$pageRelink = New-OpsPage "Перепривязка связей" "relink"
$relinkPane = New-OpsModelPane "relink" (New-PlainOpsList "relink")
Connect-PlainPane $relinkPane "relink"
$relOld = New-OpsTextBox "папка связей сейчас"
$relNew = New-OpsTextBox "куда перевести связи"
$btnRelOldBrowse = New-OpsButton "Обзор…"
$btnRelNewBrowse = New-OpsButton "Обзор…"
$btnFromMove = New-OpsButton "Из переноса"
$relSettings = New-OpsSettingsGrid
[void](Add-OpsSettingRow $relSettings "Старая папка" (New-OpsBoxField $relOld @($btnRelOldBrowse, (New-OpsSpacer $btnFromMove.Width))))
[void](Add-OpsSettingRow $relSettings "Новая папка" (New-OpsBoxField $relNew @($btnRelNewBrowse, $btnFromMove)))
Set-OpsPageContent $pageRelink $relinkPane $relSettings
[void]$tabs.TabPages.Add($pageRelink)
Add-OpsBusy @($relOld, $relNew, $btnRelOldBrowse, $btnRelNewBrowse, $btnFromMove)
$script:Ui.RelOldBox = $relOld
$script:Ui.RelNewBox = $relNew
$script:Ui.TxtRelOld = $relOld          # Build-JobObject reads the relink folders from here
$script:Ui.TxtRelNew = $relNew

# ---- tab 3: Смена версии Revit (the source version is always taken from the models)
$pageUpgrade = New-OpsPage "Смена версии Revit" "upgrade"
$upgradePane = New-OpsModelPane "upgrade" (New-PlainOpsList "upgrade")
Connect-PlainPane $upgradePane "upgrade"
$lblUpNow = New-OpsLabel "модели не выбраны"
$cmbUp = New-Object System.Windows.Forms.ComboBox
$cmbUp.DropDownStyle = "DropDownList"
$cmbUp.Width = 140
$cmbUp.Anchor = "Left"
$upSettings = New-OpsSettingsGrid
[void](Add-OpsSettingRow $upSettings "Сейчас" (New-OpsInlineField @($lblUpNow)))
[void](Add-OpsSettingRow $upSettings "Станет" (New-OpsInlineField @($cmbUp)))
Set-OpsPageContent $pageUpgrade $upgradePane $upSettings
[void]$tabs.TabPages.Add($pageUpgrade)
Add-OpsBusy @($cmbUp)
$script:Ui.UpgradeYearLabel = $lblUpNow
$script:Ui.UpgradeYear = $cmbUp

# ---- tab 4: Связи: добавить
$pageLinks = New-OpsPage "Связи: добавить" "addlinks"
$linksPane = New-OpsModelPane "addlinks" (New-PlainOpsList "addlinks") $true
$script:Ui.WsYearLabel = $linksPane.Year
Connect-PlainPane $linksPane "addlinks"
$linksSplit = New-OpsSplitPage $pageLinks $linksPane "addlinks"
[void]$tabs.TabPages.Add($pageLinks)

# ---- tab 5: Рабочие наборы
$pageSets = New-OpsPage "Рабочие наборы" "worksets"
$setsPane = New-OpsModelPane "worksets" (New-PlainOpsList "worksets") $true
$script:Ui.WsYearLabel2 = $setsPane.Year
Connect-PlainPane $setsPane "worksets"
$setsSplit = New-OpsSplitPage $pageSets $setsPane "worksets"
[void]$tabs.TabPages.Add($pageSets)
$script:Ui.Splits = @($linksSplit.Split, $setsSplit.Split)
$script:Ui.TabWorksets = $pageSets

# ---- events of tab 1
$timer.Add_Tick({ Tick-OpsRun })
$fieldTimer.Add_Tick({ $script:Ui.FieldTimer.Stop(); Update-MoveFields })
$onField = { $script:Ui.FieldTimer.Stop(); $script:Ui.FieldTimer.Start() }
foreach ($c in @($ruleOld, $ruleNew, $folderOld, $folderNew)) { $c.Add_TextChanged($onField) }
$rbKeep.Add_CheckedChanged({
        $script:Ui.CbCopy.Checked = [bool]$script:Ui.RbKeep.Checked
        $script:Ui.CbDelete.Checked = -not [bool]$script:Ui.RbKeep.Checked
    })
$btnBrowseMove.Add_Click({
        $d = New-Object System.Windows.Forms.FolderBrowserDialog
        $d.Description = "Новая папка моделей (заменит «было» в пути)"
        if ($d.ShowDialog() -eq "OK") { $script:Ui.FolderNew.Text = $d.SelectedPath }
    })
$btnRelOldBrowse.Add_Click({
        $d = New-Object System.Windows.Forms.FolderBrowserDialog
        $d.Description = "Папка, в которой связи лежат сейчас"
        if ($d.ShowDialog() -eq "OK") { $script:Ui.RelOldBox.Text = $d.SelectedPath }
    })
$btnRelNewBrowse.Add_Click({
        $d = New-Object System.Windows.Forms.FolderBrowserDialog
        $d.Description = "Папка, куда перевести связи"
        if ($d.ShowDialog() -eq "OK") { $script:Ui.RelNewBox.Text = $d.SelectedPath }
    })
$btnFromMove.Add_Click({ Fill-RelinkFoldersFromMove })
$cmbUp.Add_SelectedIndexChanged({
        if ($script:Ui.UpgradeYear.SelectedItem) { $script:UpgradeDst = [string]$script:Ui.UpgradeYear.SelectedItem }
    })

$btnAll.Add_Click({ Set-AllChecks $true })
$btnNone.Add_Click({ Set-AllChecks $false })
$btnDel.Add_Click({
        foreach ($it in @($list.SelectedItems)) {
            $drop = [string]$it.Tag
            $keep = New-Object System.Collections.Generic.List[string]
            foreach ($p in @($script:modelPaths)) {
                if ($p -ne $drop) { [void]$keep.Add($p) }
            }
            $script:modelPaths = $keep
        }
        Update-MoveFields
    })
$list.Add_ItemChecked({ Update-ListCount })
$list.Add_KeyDown({
        if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Delete) {
            $script:Ui.MovePane.Del.PerformClick()
        }
        elseif ($_.Control -and $_.KeyCode -eq [System.Windows.Forms.Keys]::C) {
            $_.Handled = $true
            $_.SuppressKeyPress = $true
            Copy-SelectedRows "both"
        }
    })

# full path tooltip for the name / path cells
$script:Ui.Tip = New-Object System.Windows.Forms.ToolTip
$script:Ui.Tip.AutoPopDelay = 20000
$script:Ui.Tip.InitialDelay = 400
$list.Add_MouseMove({
        $text = ""
        $h = $script:Ui.List.HitTest($_.Location)
        if ($h.Item -and $h.SubItem) {
            $col = $h.Item.SubItems.IndexOf($h.SubItem)
            $text = Get-CellTipText $script:RowCache[[string]$h.Item.Tag] $col
        }
        if ($text -cne $script:TipText) {
            $script:TipText = $text
            $script:Ui.Tip.SetToolTip($script:Ui.List, $text)
        }
    })

# double click: "Станет" and "Новый путь" are edited in place, the rest is read-only
$list.Add_MouseDoubleClick({
        $h = $script:Ui.List.HitTest($_.Location)
        if (-not $h.Item -or -not $h.SubItem) { return }
        if ($h.Location -eq [System.Windows.Forms.ListViewHitTestLocations]::StateImage) { return }
        $col = $h.Item.SubItems.IndexOf($h.SubItem)
        if ($col -eq 1 -or $col -eq 5) { Show-CellEditor $h.Item $col }
        else { Set-UiStatus "Правятся «Станет» и «Новый путь» (двойной щелчок)." }
    })
$list.Add_MouseWheel({ Complete-CellEdit $false })
$list.Add_Resize({ Complete-CellEdit $false })
$list.Add_ColumnWidthChanged({ Complete-CellEdit $false })

$ctx = New-Object System.Windows.Forms.ContextMenuStrip
$miCopy = $ctx.Items.Add("Копировать")
$miCopy.ShortcutKeyDisplayString = "Ctrl+C"
$miCopyOld = $ctx.Items.Add("Копировать только старые пути")
$miCopyNew = $ctx.Items.Add("Копировать только новые пути")
[void]$ctx.Items.Add("-")
$miReset = $ctx.Items.Add("Вернуть по правилу")
$miCopy.Add_Click({ Copy-SelectedRows "both" })
$miCopyOld.Add_Click({ Copy-SelectedRows "old" })
$miCopyNew.Add_Click({ Copy-SelectedRows "new" })
$miReset.Add_Click({ Reset-SelectedManual })
$ctx.Add_Opening({
        $sel = @(Get-SelectedRowObjects)
        $has = ($sel.Count -gt 0)
        $miCopy.Enabled = $has
        $miCopyOld.Enabled = $has
        $miCopyNew.Enabled = $has
        $miReset.Enabled = (@($sel | Where-Object { $_.manual }).Count -gt 0) -and (-not $script:Run.Active)
    })
$list.ContextMenuStrip = $ctx

$btnFiles.Add_Click({
        $picked = @(Select-LocalFiles)
        $n = Add-UiModels $picked
        if ($n -gt 0) { Set-UiStatus ("Добавлено файлов: {0}" -f $n) }
        elseif ($picked.Count -eq 0) { Set-UiStatus "Файлы не выбраны" }
        else { Set-UiStatus "Эти модели уже в списке" }
    })
$btnFolders.Add_Click({
        try {
            $folder = Show-FolderPathDialog
            if ([string]::IsNullOrWhiteSpace($folder)) { Set-UiStatus "Папка не выбрана"; return }
            $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            Set-UiStatus ("Ищу модели в: {0}" -f $folder)
            $picked = @(Get-ModelsFromFolder $folder)
            $n = Add-UiModels $picked
            if ($n -gt 0) { Set-UiStatus ("Из папки добавлено: {0}" -f $n) }
            else {
                Set-UiStatus ("В папке нет рабочих .rvt: {0}" -f $folder)
                [void][System.Windows.Forms.MessageBox]::Show($form, "В этой папке не нашлось рабочих моделей.", "Папки")
            }
        }
        catch {
            Set-UiStatus "Папка недоступна"
            [void][System.Windows.Forms.MessageBox]::Show($form, $_.Exception.Message, "Папки")
        }
        finally { $form.Cursor = [System.Windows.Forms.Cursors]::Default }
    })
$btnServer.Add_Click({
        $serverObj = Show-ServerPick
        if (-not $serverObj) { Set-UiStatus "Сервер не выбран"; return }
        $picked = @(Show-RevitServerBrowser -ServerObj $serverObj)
        $n = Add-UiModels $picked
        if ($n -gt 0) { Set-UiStatus ("С сервера добавлено: {0}" -f $n) }
        elseif ($picked.Count -eq 0) { Set-UiStatus "С сервера ничего не добавлено" }
        else { Set-UiStatus "Эти модели уже в списке" }
    })

$btnExport.Add_Click({
        if ($script:modelPaths.Count -eq 0) { Set-UiStatus "Список пуст: сохранять нечего."; return }
        $dlg = New-Object System.Windows.Forms.SaveFileDialog
        $dlg.Filter = "Список моделей (*.txt)|*.txt"
        $dlg.FileName = ("список_моделей_{0}.txt" -f (Get-Date -Format "yyyy-MM-dd"))
        $dlg.InitialDirectory = Get-ModelListStore $ToolDir
        if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
        $utf8 = New-Object System.Text.UTF8Encoding $true
        [System.IO.File]::WriteAllLines($dlg.FileName, @($script:modelPaths), $utf8)
        Save-ModelListFolder $ToolDir $dlg.FileName
        Set-UiStatus ("Список сохранён: {0}" -f $dlg.FileName)
    })
$btnImport.Add_Click({
        $dlg = New-Object System.Windows.Forms.OpenFileDialog
        $dlg.Filter = "Список моделей (*.txt)|*.txt|Все файлы (*.*)|*.*"
        $dlg.InitialDirectory = Get-ModelListStore $ToolDir
        if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
        Save-ModelListFolder $ToolDir $dlg.FileName
        $paths = @(Read-ModelListPaths $dlg.FileName)
        if ($script:modelPaths.Count -gt 0) {
            $ask = [System.Windows.Forms.MessageBox]::Show($form, "Добавить к текущему списку? (Нет - заменить)", "Открыть список",
                [System.Windows.Forms.MessageBoxButtons]::YesNoCancel)
            if ($ask -eq [System.Windows.Forms.DialogResult]::Cancel) { return }
            if ($ask -eq [System.Windows.Forms.DialogResult]::No) { $script:modelPaths.Clear() }
        }
        $n = Add-UiModels $paths
        Set-UiStatus ("Добавлено из списка: {0}" -f $n)
    })

# ---- the bottom bar follows the selected tab
$btnRun.Add_Click({ Invoke-ActiveTabRun $false })
$btnDry.Add_Click({ Invoke-ActiveTabRun $true })
$tabs.Add_SelectedIndexChanged({ Update-ActiveTab })
$form.Add_FormClosing({
        if ($script:Ui.Timer) { $script:Ui.Timer.Stop() }
        if ($script:Ui.FieldTimer) { $script:Ui.FieldTimer.Stop() }
    })
$form.Add_Shown({
        try {
            $wa = [System.Windows.Forms.Screen]::FromControl($form).WorkingArea
            $maxW = [Math]::Max($form.MinimumSize.Width, $wa.Width - 16)
            $maxH = [Math]::Max($form.MinimumSize.Height, $wa.Height - 16)
            if ($form.Width -gt $maxW) { $form.Width = $maxW }
            if ($form.Height -gt $maxH) { $form.Height = $maxH }
        }
        catch { }
        Restore-OpsSplitters
        Set-MoveSuggestions
        Update-MoveFields
        foreach ($k in @($script:Ui.PlainLists.Keys)) { Update-PlainCount ([string]$k) }
        Update-ActiveTab
    })
if (Get-Command Install-WorksetUi -ErrorAction SilentlyContinue) { Install-WorksetUi $linksSplit.Bottom $setsSplit.Bottom }
Update-ActiveTab

# snap_ui.ps1 builds the window without the message loop
if ($script:OpsBuildOnly) { return }
[void]$form.ShowDialog()
exit 0
