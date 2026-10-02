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

$ToolDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$script:ModelOpsUiLibrary = $true
$ModelOpsUiLibrary = $true
$script:ModelOpsLaunchLibrary = $true

try {
    . (Join-Path $ToolDir "job_lib.ps1")
    Set-StrictMode -Off
    . (Join-Path $ToolDir "choose_models.ps1")
    . (Join-Path $ToolDir "start_ops.ps1")
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
foreach ($p in @(Read-ModelList $ListPath)) { [void]$script:modelPaths.Add($p) }
$script:ops = @()
$script:nameOverrides = @{}
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
        "delete" { return "Удаляю старые модели" }
        default { return $Name }
    }
}

function Get-ModelSource([string]$Path) {
    if (Test-IsRsnPath $Path) { return "Revit Server" }
    if ($Path.StartsWith("\\")) { return "Сеть" }
    return "Этот компьютер"
}

function Get-SelectedYears {
    if (-not $script:Ui.CbUpgrade.Checked) { return @{ Src = ""; Dst = "" } }
    $src = ""
    $dst = ""
    if ($script:Ui.CmbSrc.SelectedItem) { $src = [string]$script:Ui.CmbSrc.SelectedItem }
    if ($script:Ui.CmbDst.SelectedItem) { $dst = [string]$script:Ui.CmbDst.SelectedItem }
    return @{ Src = $src; Dst = $dst }
}

function Get-PreviewRows {
    $years = Get-SelectedYears
    return @(New-JobPreviewRows `
            -OldPaths @($script:modelPaths) `
            -Ops $script:ops `
            -DoRename $script:Ui.CbRename.Checked `
            -LocationMode $script:locMode `
            -OldSegment $script:oldSeg `
            -NewSegment $script:newSeg `
            -OldPrefix $script:oldPref `
            -NewPrefix $script:newPref `
            -SourceYear $years.Src `
            -TargetYear $years.Dst `
            -BufferDir $script:bufferDir `
            -NameOverrides $script:nameOverrides)
}

function Get-JobPlanText($Job, $Rows) {
    $passes = @(New-OpsPassQueue -Job $Job -ToolDir $ToolDir -Rows $Rows)
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
        ops           = $opsJson
        saved_utc     = [DateTime]::UtcNow.ToString("o")
    }
}

function Set-UiStatus([string]$Text) {
    if ($script:Ui.Status) { $script:Ui.Status.Text = $Text }
    [System.Windows.Forms.Application]::DoEvents()
}

function Update-ListCount {
    $n = $script:Ui.List.Items.Count
    $c = 0
    foreach ($it in @($script:Ui.List.Items)) { if ($it.Checked) { $c++ } }
    if ($n -eq 0) { $script:Ui.CountLabel.Text = "Модели не выбраны" }
    else { $script:Ui.CountLabel.Text = ("В списке: {0}  ·  отмечено: {1}" -f $n, $c) }
    $busy = [bool]$script:Run.Active
    $script:Ui.BtnRun.Enabled = ($c -gt 0) -and (-not $busy)
}

function Update-PathColumns {
    $rename = [bool]$script:Ui.CbRename.Checked
    $move = [bool]$script:Ui.CbMove.Checked
    $script:Ui.List.Columns[1].Text = "Станет"
    $script:Ui.List.Columns[1].Width = $(if ($rename) { 190 } else { 0 })
    if ($move) {
        $script:Ui.List.Columns[4].Text = "Старый путь"
        $script:Ui.List.Columns[4].Width = 260
        $script:Ui.List.Columns[5].Text = "Новый путь"
        $script:Ui.List.Columns[5].Width = 260
    }
    else {
        $script:Ui.List.Columns[4].Text = "Путь"
        $script:Ui.List.Columns[4].Width = 420
        $script:Ui.List.Columns[5].Width = 0
    }
}

function Refresh-ModelList {
    $lv = $script:Ui.List
    $prev = @{}
    $prevState = @{}
    foreach ($it in @($lv.Items)) {
        $key = [string]$it.Tag
        $prev[$key] = [bool]$it.Checked
        if ($it.SubItems.Count -gt 2) { $prevState[$key] = [string]$it.SubItems[2].Text }
    }
    $rows = @(Get-PreviewRows)
    $lv.BeginUpdate()
    try {
        $lv.Items.Clear()
        foreach ($r in $rows) {
            $it = New-Object System.Windows.Forms.ListViewItem ($r.old_name)
            [void]$it.SubItems.Add($r.new_name)
            $st = ""
            if ($prevState.ContainsKey([string]$r.old_path)) { $st = $prevState[[string]$r.old_path] }
            [void]$it.SubItems.Add($st)
            [void]$it.SubItems.Add((Get-ModelSource $r.old_path))
            [void]$it.SubItems.Add($r.old_path)
            [void]$it.SubItems.Add($r.new_path)
            $it.Tag = $r.old_path
            if ($prev.ContainsKey([string]$r.old_path)) { $it.Checked = $prev[[string]$r.old_path] }
            else { $it.Checked = $true }
            [void]$lv.Items.Add($it)
        }
    }
    finally { $lv.EndUpdate() }
    Update-PathColumns
    Update-ListCount
    $job = Build-JobObject $rows
    $script:Ui.Summary.Text = Get-JobPlanText $job $rows
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
    Refresh-ModelList
    return $n
}

function Get-CheckedPaths {
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

function Set-RowState([string]$Path, [string]$Text) {
    foreach ($it in @($script:Ui.List.Items)) {
        if ([string]$it.Tag -eq $Path) {
            if ($it.SubItems.Count -gt 2) { $it.SubItems[2].Text = $Text }
            break
        }
    }
}

function Set-CheckedState([string]$Text) {
    foreach ($it in @($script:Ui.List.Items)) {
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

function Show-DetailsDialog {
    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Подробности задания"
    $dlg.Font = $script:Ui.Form.Font
    $dlg.Size = New-Object System.Drawing.Size(780, 560)
    $dlg.StartPosition = "CenterParent"
    $dlg.FormBorderStyle = "FixedDialog"
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false
    $dlg.ShowInTaskbar = $false

    $lblOld = New-Object System.Windows.Forms.Label
    $lblOld.Text = "Имя сейчас (пример):"
    $lblOld.Location = New-Object System.Drawing.Point(16, 16)
    $lblOld.AutoSize = $true
    $txtOld = New-Object System.Windows.Forms.TextBox
    $txtOld.Location = New-Object System.Drawing.Point(16, 40)
    $txtOld.Width = 460
    if ($script:modelPaths.Count -gt 0) { $txtOld.Text = Get-BasenameNoExt $script:modelPaths[0] }
    $lblNew = New-Object System.Windows.Forms.Label
    $lblNew.Text = "Должно стать:"
    $lblNew.Location = New-Object System.Drawing.Point(16, 72)
    $lblNew.AutoSize = $true
    $txtNew = New-Object System.Windows.Forms.TextBox
    $txtNew.Location = New-Object System.Drawing.Point(16, 96)
    $txtNew.Width = 460
    $btnParse = New-Object System.Windows.Forms.Button
    $btnParse.Text = "Разобрать"
    $btnParse.Location = New-Object System.Drawing.Point(490, 94)
    $btnParse.Size = New-Object System.Drawing.Size(140, 28)

    $lstOps = New-Object System.Windows.Forms.ListBox
    $lstOps.Location = New-Object System.Drawing.Point(16, 136)
    $lstOps.Size = New-Object System.Drawing.Size(460, 110)
    foreach ($op in @($script:ops)) { [void]$lstOps.Items.Add((Get-OpLabelRu $op)) }
    $btnDrop = New-Object System.Windows.Forms.Button
    $btnDrop.Text = "Убрать правило"
    $btnDrop.Location = New-Object System.Drawing.Point(490, 136)
    $btnDrop.Size = New-Object System.Drawing.Size(140, 28)

    $gb = New-Object System.Windows.Forms.GroupBox
    $gb.Text = "Куда сохранить"
    $gb.Location = New-Object System.Drawing.Point(16, 260)
    $gb.Size = New-Object System.Drawing.Size(732, 180)
    $rbSame = New-Object System.Windows.Forms.RadioButton
    $rbSame.Text = "Там же"
    $rbSame.AutoSize = $true
    $rbSame.Location = New-Object System.Drawing.Point(14, 24)
    $rbSeg = New-Object System.Windows.Forms.RadioButton
    $rbSeg.Text = "Поменять кусок пути"
    $rbSeg.AutoSize = $true
    $rbSeg.Location = New-Object System.Drawing.Point(110, 24)
    $rbPref = New-Object System.Windows.Forms.RadioButton
    $rbPref.Text = "Другой сервер или корень шары"
    $rbPref.AutoSize = $true
    $rbPref.Location = New-Object System.Drawing.Point(300, 24)
    $rbSame.Checked = ($script:locMode -eq "same")
    $rbSeg.Checked = ($script:locMode -eq "segment")
    $rbPref.Checked = ($script:locMode -eq "prefix")
    $lbl1 = New-Object System.Windows.Forms.Label
    $lbl1.Text = "было:"
    $lbl1.Location = New-Object System.Drawing.Point(14, 58)
    $lbl1.AutoSize = $true
    $t1 = New-Object System.Windows.Forms.TextBox
    $t1.Location = New-Object System.Drawing.Point(78, 54)
    $t1.Width = 630
    $lbl2 = New-Object System.Windows.Forms.Label
    $lbl2.Text = "стало:"
    $lbl2.Location = New-Object System.Drawing.Point(14, 90)
    $lbl2.AutoSize = $true
    $t2 = New-Object System.Windows.Forms.TextBox
    $t2.Location = New-Object System.Drawing.Point(78, 86)
    $t2.Width = 630
    if ($script:locMode -eq "prefix") { $t1.Text = $script:oldPref; $t2.Text = $script:newPref }
    else { $t1.Text = $script:oldSeg; $t2.Text = $script:newSeg }
    $lblBuf = New-Object System.Windows.Forms.Label
    $lblBuf.Text = "Буфер на этом ПК:"
    $lblBuf.Location = New-Object System.Drawing.Point(14, 122)
    $lblBuf.AutoSize = $true
    $txtBuf = New-Object System.Windows.Forms.TextBox
    $txtBuf.Location = New-Object System.Drawing.Point(14, 144)
    $txtBuf.Width = 600
    $txtBuf.Text = $script:bufferDir
    $btnBuf = New-Object System.Windows.Forms.Button
    $btnBuf.Text = "Обзор…"
    $btnBuf.Location = New-Object System.Drawing.Point(624, 142)
    $btnBuf.Size = New-Object System.Drawing.Size(90, 26)
    $gb.Controls.AddRange(@($rbSame, $rbSeg, $rbPref, $lbl1, $t1, $lbl2, $t2, $lblBuf, $txtBuf, $btnBuf))

    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = "Готово"
    $ok.Location = New-Object System.Drawing.Point(488, 460)
    $ok.Size = New-Object System.Drawing.Size(130, 32)
    $ok.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $no = New-Object System.Windows.Forms.Button
    $no.Text = "Отмена"
    $no.Location = New-Object System.Drawing.Point(628, 460)
    $no.Size = New-Object System.Drawing.Size(120, 32)
    $no.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $dlg.AcceptButton = $ok
    $dlg.CancelButton = $no
    $dlg.Controls.AddRange(@($lblOld, $txtOld, $lblNew, $txtNew, $btnParse, $lstOps, $btnDrop, $gb, $ok, $no))

    $script:detailOps = New-Object System.Collections.Generic.List[object]
    foreach ($op in @($script:ops)) { [void]$script:detailOps.Add($op) }

    $btnParse.Add_Click({
            $a = $txtOld.Text.Trim(); $b = $txtNew.Text.Trim()
            if (-not $a -or -not $b) { return }
            $script:detailOps.Clear()
            foreach ($op in @(New-OpsFromSample $a $b)) { [void]$script:detailOps.Add($op) }
            $lstOps.Items.Clear()
            foreach ($op in @($script:detailOps)) { [void]$lstOps.Items.Add((Get-OpLabelRu $op)) }
        })
    $btnDrop.Add_Click({
            $i = $lstOps.SelectedIndex
            if ($i -lt 0) { return }
            $script:detailOps.RemoveAt($i)
            $lstOps.Items.Clear()
            foreach ($op in @($script:detailOps)) { [void]$lstOps.Items.Add((Get-OpLabelRu $op)) }
        })
    $btnBuf.Add_Click({
            $d = New-Object System.Windows.Forms.FolderBrowserDialog
            $d.Description = "Папка буфера на этом компьютере"
            if ($d.ShowDialog() -eq "OK") { $txtBuf.Text = $d.SelectedPath }
        })

    if ($dlg.ShowDialog($script:Ui.Form) -ne [System.Windows.Forms.DialogResult]::OK) { return }
    $script:ops = @($script:detailOps)
    $script:bufferDir = $txtBuf.Text.Trim()
    if ($rbPref.Checked) {
        $script:locMode = "prefix"
        $script:oldPref = $t1.Text.Trim()
        $script:newPref = $t2.Text.Trim()
        $script:Ui.CbMove.Checked = $true
    }
    elseif ($rbSeg.Checked) {
        $script:locMode = "segment"
        $script:oldSeg = $t1.Text.Trim()
        $script:newSeg = $t2.Text.Trim()
        $script:Ui.CbMove.Checked = $true
    }
    else {
        $script:locMode = "same"
        $script:Ui.CbMove.Checked = $false
    }
    if ($script:detailOps.Count -gt 0) { $script:Ui.CbRename.Checked = $true }
    Refresh-ModelList
}

function Sync-MoveMode {
    if (-not $script:Ui.CbMove.Checked) { $script:locMode = "same" }
    elseif ($script:locMode -eq "same") { $script:locMode = "segment" }
}

function Set-UiBusy([bool]$Busy) {
    $script:Run.Active = $Busy
    foreach ($n in @(
            "BtnFiles", "BtnFolders", "BtnServer", "BtnAll", "BtnNone", "BtnDel",
            "BtnExport", "BtnImport", "BtnMore", "CbRename", "CbMove", "CbUpgrade",
            "CbRelink", "CbCopy", "CbDelete", "CmbSrc", "CmbDst"
        )) {
        if ($script:Ui.ContainsKey($n) -and $script:Ui[$n]) {
            $script:Ui[$n].Enabled = -not $Busy
        }
    }
    if (-not $Busy) {
        $script:Ui.CmbSrc.Enabled = $script:Ui.CbUpgrade.Checked
        $script:Ui.CmbDst.Enabled = $script:Ui.CbUpgrade.Checked
    }
    Update-ListCount
    $script:Ui.Progress.Visible = $Busy
}

function Start-NextPass {
    if ($script:Run.Index -ge $script:Run.Queue.Count) {
        Complete-OpsRun $true
        return
    }
    $pass = $script:Run.Queue[$script:Run.Index]
    if ($pass.Name -eq "delete") {
        $ask = [System.Windows.Forms.MessageBox]::Show(
            $script:Ui.Form,
            "Новые модели записаны. Удалить старые?",
            "Удаление",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning,
            [System.Windows.Forms.MessageBoxDefaultButton]::Button2
        )
        if ($ask -ne [System.Windows.Forms.DialogResult]::Yes) {
            Set-UiStatus "Удаление пропущено"
            $script:Run.Index++
            Start-NextPass
            return
        }
    }
    Set-UiStatus ("Проход {0} из {1}: {2}" -f ($script:Run.Index + 1), $script:Run.Queue.Count, (Get-PassTitle $pass.Name))
    Set-CheckedState "Идёт"
    $script:Ui.Progress.Value = $script:Run.Index
    if ($pass.Kind -eq "rbp") {
        $started = Start-OpsRbpPass -Pass $pass -Exe $script:Run.Exe -ExeDir $script:Run.ExeDir
    }
    else {
        $started = Start-OpsPsPass -Pass $pass
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
    if ($script:Ui.Timer) { $script:Ui.Timer.Stop() }
    $script:Run.Process = $null
    if ($Ok) { Set-CheckedState "Готово" }
    Set-UiBusy $false
    $xlsx = Find-LatestOpsReport
    $script:Ui.BtnReport.Enabled = -not [string]::IsNullOrWhiteSpace($xlsx)
    $script:Run.LastReport = $xlsx
    if ($Ok) { Set-UiStatus "Готово." }
    $script:Ui.Progress.Value = $script:Ui.Progress.Maximum
    if ($Ok -and $script:Ui.ChkReport.Checked -and $xlsx) {
        try { Start-Process -FilePath $xlsx | Out-Null } catch { }
    }
}

function Tick-OpsRun {
    if (-not $script:Run.Active) { return }
    $proc = $script:Run.Process
    if (-not $proc) { return }
    try { $proc.Refresh() } catch { }
    if (-not $proc.HasExited) { return }
    $code = 0
    try { $code = [int]$proc.ExitCode } catch { }
    $script:Run.ExitCodes += $code
    $script:Run.Process = $null
    if ($code -ne 0) {
        Set-CheckedState "Ошибка"
        Set-UiStatus ("Проход остановился, код {0}." -f $code)
        Complete-OpsRun $false
        return
    }
    Set-CheckedState "Готово"
    $script:Run.Index++
    Start-NextPass
}

function Start-UiRun {
    $checked = @(Get-CheckedPaths)
    if ($checked.Count -eq 0) {
        [void][System.Windows.Forms.MessageBox]::Show($script:Ui.Form, "Отметьте модели галочками.", "Запуск")
        return
    }
    if ($script:Ui.CbRename.Checked -and @($script:ops).Count -eq 0) {
        [void][System.Windows.Forms.MessageBox]::Show($script:Ui.Form, "Включили переименование - разберите образец в «Подробности».", "Запуск")
        return
    }
    Sync-MoveMode
    if ($script:Ui.CbMove.Checked -and $script:locMode -eq "same") {
        [void][System.Windows.Forms.MessageBox]::Show($script:Ui.Form, "Включили перенос - укажите куда в «Подробности».", "Запуск")
        return
    }
    $keep = New-Object System.Collections.Generic.List[string]
    foreach ($p in $checked) { [void]$keep.Add($p) }
    $script:modelPaths = $keep

    $rows = @(Get-PreviewRows)
    $dup = @($rows | Where-Object { $_.note -match "duplicate" })
    if ($dup.Count -gt 0) {
        [void][System.Windows.Forms.MessageBox]::Show($script:Ui.Form, "В списке повторяются новые пути. Поправьте имена или папку.", "Запуск")
        return
    }
    $job = Build-JobObject $rows
    $plan = Get-JobPlanText $job $rows
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
    $passes = @(New-OpsPassQueue -Job $job -ToolDir $ToolDir -Rows $rows)
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
    [void](Save-JobFiles -ToolDir $ToolDir -Rows $rows -Job $job)
    Refresh-ModelList
    Set-CheckedState "Ожидание"

    $script:Run.Queue = $passes
    $script:Run.Index = 0
    $script:Run.ExitCodes = @()
    if ($exe) {
        $script:Run.Exe = $exe
        $script:Run.ExeDir = Split-Path -Parent $exe
    }
    $script:Ui.Progress.Maximum = [Math]::Max(1, $passes.Count)
    $script:Ui.Progress.Value = 0
    $script:Ui.BtnReport.Enabled = $false
    Set-UiBusy $true
    $script:Ui.Timer.Start()
    Start-NextPass
}

# ---- form ----
$form = New-Object System.Windows.Forms.Form
$form.Text = "Операции с моделями Revit"
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9.75)
$form.Size = New-Object System.Drawing.Size(1120, 760)
$form.MinimumSize = New-Object System.Drawing.Size(1000, 640)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.SystemColors]::Window

$lblAdd = New-Object System.Windows.Forms.Label
$lblAdd.Text = "Добавить модели"
$lblAdd.Location = New-Object System.Drawing.Point(16, 14)
$lblAdd.AutoSize = $true
$lblAdd.Font = New-Object System.Drawing.Font("Segoe UI", 9.75, [System.Drawing.FontStyle]::Bold)

$btnFiles = New-Object System.Windows.Forms.Button
$btnFiles.Text = "Файлы…"
$btnFiles.Location = New-Object System.Drawing.Point(16, 40)
$btnFiles.Size = New-Object System.Drawing.Size(150, 34)
$btnFolders = New-Object System.Windows.Forms.Button
$btnFolders.Text = "Папки…"
$btnFolders.Location = New-Object System.Drawing.Point(176, 40)
$btnFolders.Size = New-Object System.Drawing.Size(150, 34)
$btnServer = New-Object System.Windows.Forms.Button
$btnServer.Text = "Revit Server…"
$btnServer.Location = New-Object System.Drawing.Point(336, 40)
$btnServer.Size = New-Object System.Drawing.Size(170, 34)

$lblCount = New-Object System.Windows.Forms.Label
$lblCount.Text = "Модели не выбраны"
$lblCount.Location = New-Object System.Drawing.Point(760, 40)
$lblCount.Size = New-Object System.Drawing.Size(320, 34)
$lblCount.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
$lblCount.Anchor = "Top,Right"

$list = New-Object System.Windows.Forms.ListView
$list.Location = New-Object System.Drawing.Point(16, 84)
$list.Size = New-Object System.Drawing.Size(1072, 310)
$list.Anchor = "Top,Bottom,Left,Right"
$list.View = "Details"
$list.CheckBoxes = $true
$list.FullRowSelect = $true
$list.GridLines = $true
$list.HideSelection = $false
[void]$list.Columns.Add("Модель", 190)
[void]$list.Columns.Add("Станет", 0)
[void]$list.Columns.Add("Состояние", 90)
[void]$list.Columns.Add("Откуда", 110)
[void]$list.Columns.Add("Путь", 420)
[void]$list.Columns.Add("Новый путь", 0)

$btnAll = New-Object System.Windows.Forms.Button
$btnAll.Text = "Выделить все"
$btnAll.Location = New-Object System.Drawing.Point(16, 402)
$btnAll.Size = New-Object System.Drawing.Size(150, 30)
$btnAll.Anchor = "Bottom,Left"
$btnNone = New-Object System.Windows.Forms.Button
$btnNone.Text = "Снять выделение"
$btnNone.Location = New-Object System.Drawing.Point(176, 402)
$btnNone.Size = New-Object System.Drawing.Size(150, 30)
$btnNone.Anchor = "Bottom,Left"
$btnDel = New-Object System.Windows.Forms.Button
$btnDel.Text = "Убрать из списка"
$btnDel.Location = New-Object System.Drawing.Point(336, 402)
$btnDel.Size = New-Object System.Drawing.Size(150, 30)
$btnDel.Anchor = "Bottom,Left"
$btnExport = New-Object System.Windows.Forms.Button
$btnExport.Text = "Сохранить список…"
$btnExport.Location = New-Object System.Drawing.Point(772, 402)
$btnExport.Size = New-Object System.Drawing.Size(150, 30)
$btnExport.Anchor = "Bottom,Right"
$btnImport = New-Object System.Windows.Forms.Button
$btnImport.Text = "Открыть список…"
$btnImport.Location = New-Object System.Drawing.Point(932, 402)
$btnImport.Size = New-Object System.Drawing.Size(156, 30)
$btnImport.Anchor = "Bottom,Right"

$gb = New-Object System.Windows.Forms.GroupBox
$gb.Text = "Что сделать"
$gb.Location = New-Object System.Drawing.Point(16, 442)
$gb.Size = New-Object System.Drawing.Size(740, 190)
$gb.Anchor = "Bottom,Left,Right"

$cbRename = New-Object System.Windows.Forms.CheckBox
$cbRename.Text = "Переименовать"
$cbRename.AutoSize = $true
$cbRename.Location = New-Object System.Drawing.Point(16, 26)
$cbMove = New-Object System.Windows.Forms.CheckBox
$cbMove.Text = "Перенести в другое место"
$cbMove.AutoSize = $true
$cbMove.Location = New-Object System.Drawing.Point(180, 26)
$cbUpgrade = New-Object System.Windows.Forms.CheckBox
$cbUpgrade.Text = "Сменить версию Revit"
$cbUpgrade.AutoSize = $true
$cbUpgrade.Location = New-Object System.Drawing.Point(400, 26)

$cbRelink = New-Object System.Windows.Forms.CheckBox
$cbRelink.Text = "Перепривязать связи"
$cbRelink.AutoSize = $true
$cbRelink.Location = New-Object System.Drawing.Point(16, 56)
$cbRelink.Checked = $true
$cbCopy = New-Object System.Windows.Forms.CheckBox
$cbCopy.Text = "Копия: старые оставить"
$cbCopy.AutoSize = $true
$cbCopy.Location = New-Object System.Drawing.Point(180, 56)
$cbDelete = New-Object System.Windows.Forms.CheckBox
$cbDelete.Text = "Удалить старые после прогона"
$cbDelete.AutoSize = $true
$cbDelete.Location = New-Object System.Drawing.Point(400, 56)

$lblY1 = New-Object System.Windows.Forms.Label
$lblY1.Text = "из версии"
$lblY1.Location = New-Object System.Drawing.Point(16, 90)
$lblY1.AutoSize = $true
$cmbSrc = New-Object System.Windows.Forms.ComboBox
$cmbSrc.DropDownStyle = "DropDownList"
$cmbSrc.Location = New-Object System.Drawing.Point(92, 86)
$cmbSrc.Width = 80
$lblY2 = New-Object System.Windows.Forms.Label
$lblY2.Text = "в версию"
$lblY2.Location = New-Object System.Drawing.Point(186, 90)
$lblY2.AutoSize = $true
$cmbDst = New-Object System.Windows.Forms.ComboBox
$cmbDst.DropDownStyle = "DropDownList"
$cmbDst.Location = New-Object System.Drawing.Point(256, 86)
$cmbDst.Width = 80
foreach ($y in $script:Years) { [void]$cmbSrc.Items.Add($y); [void]$cmbDst.Items.Add($y) }
$cmbSrc.SelectedItem = "2022"
$cmbDst.SelectedItem = "2024"
$cmbSrc.Enabled = $false
$cmbDst.Enabled = $false

$btnMore = New-Object System.Windows.Forms.Button
$btnMore.Text = "Подробности…"
$btnMore.Location = New-Object System.Drawing.Point(400, 85)
$btnMore.Size = New-Object System.Drawing.Size(180, 30)

$lblSummary = New-Object System.Windows.Forms.Label
$lblSummary.Location = New-Object System.Drawing.Point(16, 122)
$lblSummary.Size = New-Object System.Drawing.Size(700, 56)
$lblSummary.ForeColor = [System.Drawing.SystemColors]::GrayText
$gb.Controls.AddRange(@(
        $cbRename, $cbMove, $cbUpgrade, $cbRelink, $cbCopy, $cbDelete,
        $lblY1, $cmbSrc, $lblY2, $cmbDst, $btnMore, $lblSummary
    ))

$chkReport = New-Object System.Windows.Forms.CheckBox
$chkReport.Text = "Открыть отчёт после прогона"
$chkReport.Location = New-Object System.Drawing.Point(772, 442)
$chkReport.Size = New-Object System.Drawing.Size(316, 28)
$chkReport.Anchor = "Bottom,Right"
$btnReport = New-Object System.Windows.Forms.Button
$btnReport.Text = "Открыть отчёт"
$btnReport.Location = New-Object System.Drawing.Point(772, 474)
$btnReport.Size = New-Object System.Drawing.Size(316, 30)
$btnReport.Anchor = "Bottom,Right"
$btnReport.Enabled = $false
$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = "Запустить"
$btnRun.Location = New-Object System.Drawing.Point(772, 512)
$btnRun.Size = New-Object System.Drawing.Size(316, 56)
$btnRun.Anchor = "Bottom,Right"
$btnRun.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)

$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Location = New-Object System.Drawing.Point(16, 644)
$progress.Size = New-Object System.Drawing.Size(1072, 14)
$progress.Anchor = "Bottom,Left,Right"
$progress.Visible = $false
$status = New-Object System.Windows.Forms.Label
$status.Text = "Наберите модели, отметьте действия и нажмите «Запустить»."
$status.Location = New-Object System.Drawing.Point(16, 664)
$status.Size = New-Object System.Drawing.Size(1072, 22)
$status.Anchor = "Bottom,Left,Right"
$status.ForeColor = [System.Drawing.SystemColors]::GrayText

$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 1000

$form.Controls.AddRange(@(
        $lblAdd, $btnFiles, $btnFolders, $btnServer, $lblCount, $list,
        $btnAll, $btnNone, $btnDel, $btnExport, $btnImport,
        $gb, $chkReport, $btnReport, $btnRun, $progress, $status
    ))

$script:Ui = @{
    Form       = $form
    List       = $list
    CountLabel = $lblCount
    Status     = $status
    Progress   = $progress
    Timer      = $timer
    BtnRun     = $btnRun
    BtnFiles   = $btnFiles
    BtnFolders = $btnFolders
    BtnServer  = $btnServer
    BtnAll     = $btnAll
    BtnNone    = $btnNone
    BtnDel     = $btnDel
    BtnExport  = $btnExport
    BtnImport  = $btnImport
    BtnMore    = $btnMore
    BtnReport  = $btnReport
    CbRename   = $cbRename
    CbMove     = $cbMove
    CbUpgrade  = $cbUpgrade
    CbRelink   = $cbRelink
    CbCopy     = $cbCopy
    CbDelete   = $cbDelete
    CmbSrc     = $cmbSrc
    CmbDst     = $cmbDst
    Summary    = $lblSummary
    ChkReport  = $chkReport
}
$script:UiOwner = $form

$timer.Add_Tick({ Tick-OpsRun })
$onChange = {
    Sync-MoveMode
    $script:Ui.CmbSrc.Enabled = $script:Ui.CbUpgrade.Checked
    $script:Ui.CmbDst.Enabled = $script:Ui.CbUpgrade.Checked
    Refresh-ModelList
}
foreach ($c in @($cbRename, $cbMove, $cbUpgrade, $cbRelink, $cbCopy, $cbDelete)) { $c.Add_CheckedChanged($onChange) }
foreach ($c in @($cmbSrc, $cmbDst)) { $c.Add_SelectedIndexChanged($onChange) }
$cbCopy.Add_CheckedChanged({ if ($cbCopy.Checked) { $cbDelete.Checked = $false } })
$cbDelete.Add_CheckedChanged({ if ($cbDelete.Checked) { $cbCopy.Checked = $false } })

$btnMore.Add_Click({ Show-DetailsDialog })
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
        Refresh-ModelList
    })
$list.Add_ItemChecked({ Update-ListCount })
$list.Add_KeyDown({
        if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Delete) {
            $script:Ui.BtnDel.PerformClick()
        }
    })

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
        if ($script:modelPaths.Count -eq 0) { return }
        $dlg = New-Object System.Windows.Forms.SaveFileDialog
        $dlg.Filter = "Список моделей (*.txt)|*.txt"
        $dlg.FileName = ("список_моделей_{0}.txt" -f (Get-Date -Format "yyyy-MM-dd"))
        if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
        $utf8 = New-Object System.Text.UTF8Encoding $true
        [System.IO.File]::WriteAllLines($dlg.FileName, @($script:modelPaths), $utf8)
        Set-UiStatus ("Список сохранён: {0}" -f $dlg.FileName)
    })
$btnImport.Add_Click({
        $dlg = New-Object System.Windows.Forms.OpenFileDialog
        $dlg.Filter = "Список моделей (*.txt)|*.txt|Все файлы (*.*)|*.*"
        if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
        $paths = @(
            Get-Content -LiteralPath $dlg.FileName -ErrorAction SilentlyContinue |
                ForEach-Object { $_.Trim() } |
                Where-Object { $_ -and -not $_.StartsWith("#") -and ($_.ToLower().EndsWith(".rvt") -or $_ -match '^(?i)RSN://') }
        )
        if ($script:modelPaths.Count -gt 0) {
            $ask = [System.Windows.Forms.MessageBox]::Show($form, "Добавить к текущему списку? (Нет - заменить)", "Импорт",
                [System.Windows.Forms.MessageBoxButtons]::YesNoCancel)
            if ($ask -eq [System.Windows.Forms.DialogResult]::Cancel) { return }
            if ($ask -eq [System.Windows.Forms.DialogResult]::No) { $script:modelPaths.Clear() }
        }
        $n = Add-UiModels $paths
        Set-UiStatus ("Импортировано: {0}" -f $n)
    })

$btnReport.Add_Click({
        $p = [string]$script:Run.LastReport
        if ([string]::IsNullOrWhiteSpace($p) -or -not (Test-Path -LiteralPath $p)) { $p = Find-LatestOpsReport }
        if ([string]::IsNullOrWhiteSpace($p) -or -not (Test-Path -LiteralPath $p)) {
            [void][System.Windows.Forms.MessageBox]::Show($form, "Отчёта этого прогона ещё нет.", "Отчёт")
            return
        }
        try { Start-Process -FilePath $p | Out-Null } catch {
            [void][System.Windows.Forms.MessageBox]::Show($form, $_.Exception.Message, "Отчёт")
        }
    })
$btnRun.Add_Click({ Start-UiRun })
$form.Add_FormClosing({ if ($script:Ui.Timer) { $script:Ui.Timer.Stop() } })
$form.Add_Shown({
        Refresh-ModelList
        if ($script:modelPaths.Count -gt 0) {
            Set-UiStatus "Подставлен прошлый список - можно править и запускать."
        }
    })

[void]$form.ShowDialog()
exit 0
