# Unified presets window for batch_model_ops.
# STA WinForms. Fit to screen WorkingArea. UTF-8 source (BOM added on save).

if ([Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
    $arg = "-NoProfile -ExecutionPolicy Bypass -STA -File `"$PSCommandPath`""
    Start-Process -FilePath "powershell.exe" -ArgumentList $arg -Wait -NoNewWindow
    exit $LASTEXITCODE
}

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$ToolDir = Split-Path -Parent $MyInvocation.MyCommand.Path
. (Join-Path $ToolDir "job_lib.ps1")

$ListPath = Join-Path $ToolDir "rvt_list.txt"
$ServersCfg = Join-Path $ToolDir "servers.cfg"
$ChoosePs1 = Join-Path $ToolDir "choose_models.ps1"

$script:modelPaths = @(Read-ModelList $ListPath)
$script:ops = @()
$script:nameOverrides = @{}
$script:servers = @(Read-ServerTable $ServersCfg)

function New-Label([string]$Text, [int]$X, [int]$Y, [int]$W, [int]$H) {
    $c = New-Object System.Windows.Forms.Label
    $c.Text = $Text
    $c.Location = New-Object System.Drawing.Point ($X, $Y)
    $c.Size = New-Object System.Drawing.Size ($W, $H)
    $c.AutoSize = $false
    return $c
}

function Get-ScreenBox {
    $wa = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
    $w = [Math]::Min(1280, [Math]::Max(900, $wa.Width - 24))
    $h = [Math]::Min(860, [Math]::Max(640, $wa.Height - 24))
    if ($w -gt $wa.Width) { $w = $wa.Width }
    if ($h -gt $wa.Height) { $h = $wa.Height }
    $x = $wa.X + [Math]::Max(0, [int](($wa.Width - $w) / 2))
    $y = $wa.Y + [Math]::Max(0, [int](($wa.Height - $h) / 2))
    return @{ X = $x; Y = $y; W = $w; H = $h; Wa = $wa }
}

function Refresh-OpList {
    $lstOps.Items.Clear()
    foreach ($op in @($script:ops)) {
        [void]$lstOps.Items.Add((Get-OpLabel $op))
    }
}

function Get-SelectedYears {
    $src = ""
    $dst = ""
    if ($cbUpgrade.Checked) {
        if ($cmbSrcYear.SelectedItem) { $src = [string]$cmbSrcYear.SelectedItem }
        if ($cmbDstYear.SelectedItem) { $dst = [string]$cmbDstYear.SelectedItem }
    }
    return @{ Src = $src; Dst = $dst }
}

function Get-LocationMode {
    if ($rbSeg.Checked) { return "segment" }
    if ($rbPref.Checked) { return "prefix" }
    return "same"
}

function Get-PrefixValues {
    $oldP = $txtOldPrefix.Text.Trim()
    $newP = $txtNewPrefix.Text.Trim()
    if ($cmbNewRsn.SelectedItem -and ($cmbNewRsn.SelectedIndex -gt 0)) {
        $item = [string]$cmbNewRsn.SelectedItem
        if ($item -match '^(\S+)\s') { $newHost = $Matches[1] }
        else { $newHost = $item }
        if ($script:modelPaths.Count -gt 0 -and (Test-IsRsnPath $script:modelPaths[0])) {
            $oldHost = Get-RsnHost $script:modelPaths[0]
            if ($oldHost) {
                if ([string]::IsNullOrWhiteSpace($oldP)) { $oldP = "RSN://$oldHost/" }
                if ([string]::IsNullOrWhiteSpace($newP)) { $newP = "RSN://$newHost/" }
            }
        }
    }
    return @{ Old = $oldP; New = $newP }
}

function Refresh-Preview {
    $years = Get-SelectedYears
    $pref = Get-PrefixValues
    $rows = @(New-JobPreviewRows `
            -OldPaths $script:modelPaths `
            -Ops $script:ops `
            -DoRename $cbRename.Checked `
            -LocationMode (Get-LocationMode) `
            -OldSegment $txtOldSeg.Text.Trim() `
            -NewSegment $txtNewSeg.Text.Trim() `
            -OldPrefix $pref.Old `
            -NewPrefix $pref.New `
            -SourceYear $years.Src `
            -TargetYear $years.Dst `
            -BufferDir $txtBuffer.Text.Trim() `
            -NameOverrides $script:nameOverrides)
    $grid.Rows.Clear()
    $bufN = 0
    foreach ($r in $rows) {
        $idx = $grid.Rows.Add(
            $r.old_name, $r.new_name, $r.old_path, $r.new_path,
            $r.staging_path, $(if ($r.needs_buffer) { "yes" } else { "" }), $r.note
        )
        $grid.Rows[$idx].Tag = $r
        if ($r.needs_buffer) { $bufN++ }
    }
    $lblCount.Text = ("Models: {0}    Buffer (Detach on this PC): {1}" -f $script:modelPaths.Count, $bufN)
    $lblBufHint.Visible = ($bufN -gt 0)
}

function Guess-YearsFromList {
    foreach ($p in @($script:modelPaths)) {
        $y = Get-YearFromName $p
        if ($y) {
            $cmbSrcYear.SelectedItem = $y
            break
        }
    }
    if ($script:modelPaths.Count -gt 0 -and (Test-IsRsnPath $script:modelPaths[0])) {
        $h = Get-RsnHost $script:modelPaths[0]
        $sy = Get-YearForRsnHost $h $script:servers
        if ($sy) { $cmbSrcYear.SelectedItem = $sy }
    }
}

function Add-OpFromUi([string]$Kind) {
    if ($Kind -eq "delete") {
        $t = $txtOpToken.Text.Trim()
        if (-not $t) { return }
        $script:ops += (New-NameOp -Kind delete -Token $t)
    }
    elseif ($Kind -eq "replace") {
        $a = $txtOpFrom.Text.Trim(); $b = $txtOpTo.Text.Trim()
        if (-not $a -or -not $b) { return }
        $script:ops += (New-NameOp -Kind replace -From $a -To $b)
    }
    elseif ($Kind -eq "insert") {
        $t = $txtOpToken.Text.Trim()
        if (-not $t) { return }
        $script:ops += (New-NameOp -Kind insert -Token $t -After $txtOpAfter.Text.Trim())
    }
    Refresh-OpList
    Refresh-Preview
}

# ---- form ----
$box = Get-ScreenBox
$form = New-Object System.Windows.Forms.Form
$form.Text = "Model ops - presets"
$form.StartPosition = "Manual"
$form.Bounds = New-Object System.Drawing.Rectangle ($box.X, $box.Y, $box.W, $box.H)
$form.MinimumSize = New-Object System.Drawing.Size (900, 640)
$form.MaximumSize = New-Object System.Drawing.Size ($box.Wa.Width, $box.Wa.Height)
$form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Font
$form.Font = New-Object System.Drawing.Font ("Segoe UI", 9)
$form.KeyPreview = $true

$root = New-Object System.Windows.Forms.TableLayoutPanel
$root.Dock = "Fill"
$root.ColumnCount = 1
$root.RowCount = 3
$root.Padding = New-Object System.Windows.Forms.Padding (8)
[void]$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle ([System.Windows.Forms.SizeType]::Absolute, 40)))
[void]$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle ([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle ([System.Windows.Forms.SizeType]::Absolute, 52)))
[void]$root.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle ([System.Windows.Forms.SizeType]::Percent, 100)))
$form.Controls.Add($root)

$top = New-Object System.Windows.Forms.FlowLayoutPanel
$top.Dock = "Fill"
$top.WrapContents = $false
$btnPick = New-Object System.Windows.Forms.Button
$btnPick.Text = "1. Choose models"
$btnPick.AutoSize = $true
$btnPick.Height = 28
$btnReload = New-Object System.Windows.Forms.Button
$btnReload.Text = "Reload list"
$btnReload.AutoSize = $true
$btnReload.Height = 28
$lblCount = New-Object System.Windows.Forms.Label
$lblCount.AutoSize = $true
$lblCount.Padding = New-Object System.Windows.Forms.Padding (12, 6, 0, 0)
$lblCount.Text = "Models: 0"
$top.Controls.AddRange(@($btnPick, $btnReload, $lblCount))
$root.Controls.Add($top, 0, 0)

$tabs = New-Object System.Windows.Forms.TabControl
$tabs.Dock = "Fill"
$tabTask = New-Object System.Windows.Forms.TabPage
$tabTask.Text = "Task"
$tabNames = New-Object System.Windows.Forms.TabPage
$tabNames.Text = "Names"
$tabWhere = New-Object System.Windows.Forms.TabPage
$tabWhere.Text = "Location / version"
$tabPrev = New-Object System.Windows.Forms.TabPage
$tabPrev.Text = "Preview / save"
foreach ($tp in @($tabTask, $tabNames, $tabWhere, $tabPrev)) {
    $tp.AutoScroll = $true
    $tp.Padding = New-Object System.Windows.Forms.Padding (8)
}
[void]$tabs.TabPages.Add($tabTask)
[void]$tabs.TabPages.Add($tabNames)
[void]$tabs.TabPages.Add($tabWhere)
[void]$tabs.TabPages.Add($tabPrev)
$root.Controls.Add($tabs, 0, 1)

$bottom = New-Object System.Windows.Forms.FlowLayoutPanel
$bottom.Dock = "Fill"
$bottom.FlowDirection = "RightToLeft"
$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Text = "Close"
$btnCancel.Width = 110
$btnCancel.Height = 32
$btnSave = New-Object System.Windows.Forms.Button
$btnSave.Text = "Save job"
$btnSave.Width = 140
$btnSave.Height = 32
$btnSave.Name = "SaveJob"
$bottom.Controls.AddRange(@($btnCancel, $btnSave))
$root.Controls.Add($bottom, 0, 2)

# ---- Tab Task ----
$cbRename = New-Object System.Windows.Forms.CheckBox
$cbRename.Text = "Rename files (operations on Names tab)"
$cbRename.AutoSize = $true
$cbRename.Location = New-Object System.Drawing.Point (8, 12)
$cbRename.Checked = $true

$cbRelink = New-Object System.Windows.Forms.CheckBox
$cbRelink.Text = "Update RVT Links after new files exist (Pass 2)"
$cbRelink.AutoSize = $true
$cbRelink.Location = New-Object System.Drawing.Point (8, 40)
$cbRelink.Checked = $true

$cbCopy = New-Object System.Windows.Forms.CheckBox
$cbCopy.Text = "Copy: keep old files (do not delete)"
$cbCopy.AutoSize = $true
$cbCopy.Location = New-Object System.Drawing.Point (8, 68)

$cbDelete = New-Object System.Windows.Forms.CheckBox
$cbDelete.Text = "Delete old files after green Pass 2 (UNC auto / RSN Admin list)"
$cbDelete.AutoSize = $true
$cbDelete.Location = New-Object System.Drawing.Point (8, 96)
$cbDelete.Checked = $true

$cbUpgrade = New-Object System.Windows.Forms.CheckBox
$cbUpgrade.Text = "Upgrade Revit version (open in newer Revit + SaveAs)"
$cbUpgrade.AutoSize = $true
$cbUpgrade.Location = New-Object System.Drawing.Point (8, 132)

$lblSrcY = New-Label "Source Revit year:" 8 164 150 22
$cmbSrcYear = New-Object System.Windows.Forms.ComboBox
$cmbSrcYear.DropDownStyle = "DropDownList"
$cmbSrcYear.Location = New-Object System.Drawing.Point (170, 160)
$cmbSrcYear.Width = 90
@("2021", "2022", "2023", "2024", "2025", "2026") | ForEach-Object { [void]$cmbSrcYear.Items.Add($_) }
$cmbSrcYear.SelectedItem = "2022"

$lblDstY = New-Label "Target Revit year:" 280 164 150 22
$cmbDstYear = New-Object System.Windows.Forms.ComboBox
$cmbDstYear.DropDownStyle = "DropDownList"
$cmbDstYear.Location = New-Object System.Drawing.Point (440, 160)
$cmbDstYear.Width = 90
@("2021", "2022", "2023", "2024", "2025", "2026") | ForEach-Object { [void]$cmbDstYear.Items.Add($_) }
$cmbDstYear.SelectedItem = "2024"

$lblCopyLinks = New-Label "If Copy: links in the new files should point to" 8 200 420 22
$rbLinksNew = New-Object System.Windows.Forms.RadioButton
$rbLinksNew.Text = "the new copies"
$rbLinksNew.AutoSize = $true
$rbLinksNew.Location = New-Object System.Drawing.Point (8, 224)
$rbLinksNew.Checked = $true
$rbLinksOld = New-Object System.Windows.Forms.RadioButton
$rbLinksOld.Text = "the original files"
$rbLinksOld.AutoSize = $true
$rbLinksOld.Location = New-Object System.Drawing.Point (160, 224)

$lblTaskHelp = New-Object System.Windows.Forms.Label
$lblTaskHelp.Location = New-Object System.Drawing.Point (8, 268)
$lblTaskHelp.Size = New-Object System.Drawing.Size (1100, 90)
$lblTaskHelp.Text = "RBP still runs as separate passes. This window only writes mapping.csv, lists and the checklist.`r`nNested RVT links are skipped: open every host that owns the link.`r`nRSN year jump uses Detach to a folder on THIS PC (Location tab), then SaveAs in the new Revit."
$tabTask.Controls.AddRange(@(
        $cbRename, $cbRelink, $cbCopy, $cbDelete, $cbUpgrade,
        $lblSrcY, $cmbSrcYear, $lblDstY, $cmbDstYear,
        $lblCopyLinks, $rbLinksNew, $rbLinksOld, $lblTaskHelp
    ))

# ---- Tab Names ----
$lblOldS = New-Label "Current name (sample):" 8 12 220 22
$txtOldSample = New-Object System.Windows.Forms.TextBox
$txtOldSample.Location = New-Object System.Drawing.Point (8, 36)
$txtOldSample.Width = 520
$lblNewS = New-Label "Required name (sample):" 8 68 220 22
$txtNewSample = New-Object System.Windows.Forms.TextBox
$txtNewSample.Location = New-Object System.Drawing.Point (8, 92)
$txtNewSample.Width = 520
$btnParse = New-Object System.Windows.Forms.Button
$btnParse.Text = "Build operations from sample"
$btnParse.Location = New-Object System.Drawing.Point (540, 90)
$btnParse.AutoSize = $true

$lblOps = New-Label "Operations (in order, applied to every selected name):" 8 128 500 22
$lstOps = New-Object System.Windows.Forms.ListBox
$lstOps.Location = New-Object System.Drawing.Point (8, 154)
$lstOps.Size = New-Object System.Drawing.Size (520, 220)

$lblTok = New-Label "Token / from:" 548 154 120 20
$txtOpToken = New-Object System.Windows.Forms.TextBox
$txtOpToken.Location = New-Object System.Drawing.Point (548, 176)
$txtOpToken.Width = 180
$txtOpFrom = New-Object System.Windows.Forms.TextBox
$txtOpFrom.Location = New-Object System.Drawing.Point (548, 176)
$txtOpFrom.Width = 180
$txtOpFrom.Visible = $false
$lblTo = New-Label "Replace with / insert after:" 548 204 200 20
$txtOpTo = New-Object System.Windows.Forms.TextBox
$txtOpTo.Location = New-Object System.Drawing.Point (548, 226)
$txtOpTo.Width = 180
$txtOpAfter = New-Object System.Windows.Forms.TextBox
$txtOpAfter.Location = New-Object System.Drawing.Point (548, 226)
$txtOpAfter.Width = 180
$txtOpAfter.Visible = $false

$btnDelOp = New-Object System.Windows.Forms.Button
$btnDelOp.Text = "Add DELETE token"
$btnDelOp.Location = New-Object System.Drawing.Point (548, 262)
$btnDelOp.Width = 200
$btnRepOp = New-Object System.Windows.Forms.Button
$btnRepOp.Text = "Add REPLACE"
$btnRepOp.Location = New-Object System.Drawing.Point (548, 294)
$btnRepOp.Width = 200
$btnInsOp = New-Object System.Windows.Forms.Button
$btnInsOp.Text = "Add INSERT token"
$btnInsOp.Location = New-Object System.Drawing.Point (548, 326)
$btnInsOp.Width = 200
$btnRmOp = New-Object System.Windows.Forms.Button
$btnRmOp.Text = "Remove selected"
$btnRmOp.Location = New-Object System.Drawing.Point (548, 358)
$btnRmOp.Width = 200

$lblNameHelp = New-Object System.Windows.Forms.Label
$lblNameHelp.Location = New-Object System.Drawing.Point (8, 384)
$lblNameHelp.Size = New-Object System.Drawing.Size (1100, 70)
$lblNameHelp.Text = "Example: sample  ..._01_K01_R24  ->  ..._K01_R24  builds DELETE 01.`r`nThen add REPLACE STLB -> K03-K05-K07-K10-STLB and DELETE 02. Replace runs only if the token exists."
$tabNames.Controls.AddRange(@(
        $lblOldS, $txtOldSample, $lblNewS, $txtNewSample, $btnParse,
        $lblOps, $lstOps, $lblTok, $txtOpToken, $txtOpFrom, $lblTo, $txtOpTo, $txtOpAfter,
        $btnDelOp, $btnRepOp, $btnInsOp, $btnRmOp, $lblNameHelp
    ))
# keep both token/from visible: use one row of fields
$txtOpFrom.Visible = $true
$txtOpToken.Visible = $true
$txtOpFrom.Location = New-Object System.Drawing.Point (740, 176)
$lblFrom2 = New-Label "Replace FROM:" 740 154 140 20
$tabNames.Controls.Add($lblFrom2)

# ---- Tab Where ----
$gbLoc = New-Object System.Windows.Forms.GroupBox
$gbLoc.Text = "Where the new central will live"
$gbLoc.Location = New-Object System.Drawing.Point (8, 8)
$gbLoc.Size = New-Object System.Drawing.Size (1180, 210)

$rbSame = New-Object System.Windows.Forms.RadioButton
$rbSame.Text = "Same folder (rename only)"
$rbSame.AutoSize = $true
$rbSame.Location = New-Object System.Drawing.Point (12, 24)
$rbSame.Checked = $true
$rbSeg = New-Object System.Windows.Forms.RadioButton
$rbSeg.Text = "Replace a path segment (folder move, same storage type)"
$rbSeg.AutoSize = $true
$rbSeg.Location = New-Object System.Drawing.Point (12, 52)
$lblOldSeg = New-Label "Old segment:" 36 80 100 20
$txtOldSeg = New-Object System.Windows.Forms.TextBox
$txtOldSeg.Location = New-Object System.Drawing.Point (140, 76)
$txtOldSeg.Width = 400
$lblNewSeg = New-Label "New segment:" 560 80 100 20
$txtNewSeg = New-Object System.Windows.Forms.TextBox
$txtNewSeg.Location = New-Object System.Drawing.Point (670, 76)
$txtNewSeg.Width = 400

$rbPref = New-Object System.Windows.Forms.RadioButton
$rbPref.Text = "Replace path prefix (other share root or other Revit Server host)"
$rbPref.AutoSize = $true
$rbPref.Location = New-Object System.Drawing.Point (12, 108)
$lblOldPr = New-Label "Old prefix:" 36 136 100 20
$txtOldPrefix = New-Object System.Windows.Forms.TextBox
$txtOldPrefix.Location = New-Object System.Drawing.Point (140, 132)
$txtOldPrefix.Width = 500
$lblNewPr = New-Label "New prefix:" 36 164 100 20
$txtNewPrefix = New-Object System.Windows.Forms.TextBox
$txtNewPrefix.Location = New-Object System.Drawing.Point (140, 160)
$txtNewPrefix.Width = 500
$lblRsn = New-Label "Or pick new RSN host:" 660 136 160 20
$cmbNewRsn = New-Object System.Windows.Forms.ComboBox
$cmbNewRsn.DropDownStyle = "DropDownList"
$cmbNewRsn.Location = New-Object System.Drawing.Point (660, 160)
$cmbNewRsn.Width = 420
[void]$cmbNewRsn.Items.Add("(do not change host)")
foreach ($s in $script:servers) {
    [void]$cmbNewRsn.Items.Add(("{0}  {1}  ({2})" -f $s.Host, $s.Label, $s.Year))
}
$cmbNewRsn.SelectedIndex = 0
$gbLoc.Controls.AddRange(@(
        $rbSame, $rbSeg, $lblOldSeg, $txtOldSeg, $lblNewSeg, $txtNewSeg,
        $rbPref, $lblOldPr, $txtOldPrefix, $lblNewPr, $txtNewPrefix, $lblRsn, $cmbNewRsn
    ))

$gbBuf = New-Object System.Windows.Forms.GroupBox
$gbBuf.Text = "Buffer on this PC (only when leaving old RSN for a newer Revit year)"
$gbBuf.Location = New-Object System.Drawing.Point (8, 230)
$gbBuf.Size = New-Object System.Drawing.Size (1180, 130)
$lblBuf = New-Label "Folder:" 12 28 80 20
$txtBuffer = New-Object System.Windows.Forms.TextBox
$txtBuffer.Location = New-Object System.Drawing.Point (90, 24)
$txtBuffer.Width = 900
$txtBuffer.Text = Get-DefaultBufferDir
$btnBuf = New-Object System.Windows.Forms.Button
$btnBuf.Text = "Browse"
$btnBuf.Location = New-Object System.Drawing.Point (1000, 22)
$btnBuf.Width = 90
$lblBufHint = New-Object System.Windows.Forms.Label
$lblBufHint.Location = New-Object System.Drawing.Point (12, 56)
$lblBufHint.Size = New-Object System.Drawing.Size (1140, 60)
$lblBufHint.Text = "Unload from old RSN: RBP Detach = ON, task detach_to_buffer.py, list = rvt_list.txt.`r`nThis is NOT a central. Next pass: new Revit SaveAs from rvt_list_staging.txt to the real destination.`r`nNeed enough disk for ALL models of the run at once."
$gbBuf.Controls.AddRange(@($lblBuf, $txtBuffer, $btnBuf, $lblBufHint))

$tabWhere.Controls.AddRange(@($gbLoc, $gbBuf))

# ---- Tab Preview ----
$grid = New-Object System.Windows.Forms.DataGridView
$grid.Dock = "Fill"
$grid.AllowUserToAddRows = $false
$grid.AllowUserToDeleteRows = $false
$grid.AutoSizeColumnsMode = "AllCells"
$grid.ScrollBars = "Both"
$grid.SelectionMode = "FullRowSelect"
$grid.RowHeadersVisible = $false
$grid.MultiSelect = $false
$cols = @(
    @{ N = "old_name"; H = "Old name"; Ro = $true },
    @{ N = "new_name"; H = "New name"; Ro = $false },
    @{ N = "old_path"; H = "Old path"; Ro = $true },
    @{ N = "new_path"; H = "New path"; Ro = $true },
    @{ N = "staging"; H = "Buffer on PC"; Ro = $true },
    @{ N = "buffer"; H = "Detach hop"; Ro = $true },
    @{ N = "note"; H = "Note"; Ro = $true }
)
foreach ($c in $cols) {
    $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $col.Name = $c.N
    $col.HeaderText = $c.H
    $col.ReadOnly = $c.Ro
    [void]$grid.Columns.Add($col)
}
$tabPrev.Controls.Add($grid)

# ---- events ----
$refresh = { Refresh-Preview }
foreach ($c in @(
        $cbRename, $cbRelink, $cbCopy, $cbDelete, $cbUpgrade,
        $rbSame, $rbSeg, $rbPref, $cmbSrcYear, $cmbDstYear
    )) {
    if ($c -is [System.Windows.Forms.CheckBox] -or $c -is [System.Windows.Forms.RadioButton]) {
        $c.Add_CheckedChanged($refresh)
    }
    else {
        $c.Add_SelectedIndexChanged($refresh)
    }
}
$txtOldSeg.Add_TextChanged($refresh)
$txtNewSeg.Add_TextChanged($refresh)
$txtOldPrefix.Add_TextChanged($refresh)
$txtNewPrefix.Add_TextChanged($refresh)
$txtBuffer.Add_TextChanged($refresh)

$cmbNewRsn.Add_SelectedIndexChanged({
        if ($cmbNewRsn.SelectedIndex -gt 0) {
            $rbPref.Checked = $true
            $item = [string]$cmbNewRsn.SelectedItem
            $newHost = $null
            $newYear = $null
            if ($item -match '^(\S+)\s') { $newHost = $Matches[1] }
            if ($item -match '\((\d{4})\)\s*$') { $newYear = $Matches[1] }
            if ($script:modelPaths.Count -gt 0 -and (Test-IsRsnPath $script:modelPaths[0])) {
                $oldHost = Get-RsnHost $script:modelPaths[0]
                if ($oldHost -and $newHost) {
                    $txtOldPrefix.Text = "RSN://$oldHost/"
                    $txtNewPrefix.Text = "RSN://$newHost/"
                }
            }
            if ($newYear) {
                $cmbDstYear.SelectedItem = $newYear
                $srcY = Get-SelectedYears
                if ($srcY.Src -and ($srcY.Src -ne $newYear)) {
                    $cbUpgrade.Checked = $true
                }
            }
        }
        Refresh-Preview
    })

$cbCopy.Add_CheckedChanged({
        if ($cbCopy.Checked) { $cbDelete.Checked = $false }
        Refresh-Preview
    })
$cbDelete.Add_CheckedChanged({
        if ($cbDelete.Checked) { $cbCopy.Checked = $false }
    })

$btnPick.Add_Click({
        if (-not (Test-Path -LiteralPath $ChoosePs1)) {
            [void][System.Windows.Forms.MessageBox]::Show("choose_models.ps1 not found")
            return
        }
        $p = Start-Process -FilePath "powershell.exe" -ArgumentList @(
            "-NoProfile", "-ExecutionPolicy", "Bypass", "-STA", "-File", $ChoosePs1, "-OutFile", "rvt_list.txt"
        ) -WorkingDirectory $ToolDir -Wait -PassThru
        $script:modelPaths = @(Read-ModelList $ListPath)
        if ($script:modelPaths.Count -gt 0 -and [string]::IsNullOrWhiteSpace($txtOldSample.Text)) {
            $txtOldSample.Text = Get-BasenameNoExt $script:modelPaths[0]
        }
        Guess-YearsFromList
        Refresh-Preview
    })

$btnReload.Add_Click({
        $script:modelPaths = @(Read-ModelList $ListPath)
        Refresh-Preview
    })

$btnParse.Add_Click({
        $a = $txtOldSample.Text.Trim()
        $b = $txtNewSample.Text.Trim()
        if (-not $a -or -not $b) {
            [void][System.Windows.Forms.MessageBox]::Show("Fill both sample names.")
            return
        }
        $script:ops = @(New-OpsFromSample $a $b)
        $cbRename.Checked = $true
        Refresh-OpList
        Refresh-Preview
        $tabs.SelectedTab = $tabNames
    })

$btnDelOp.Add_Click({ Add-OpFromUi "delete" })
$btnRepOp.Add_Click({
        $script:ops += (New-NameOp -Kind replace -From $txtOpFrom.Text.Trim() -To $txtOpTo.Text.Trim())
        Refresh-OpList
        Refresh-Preview
    })
$btnInsOp.Add_Click({ Add-OpFromUi "insert" })
$btnRmOp.Add_Click({
        $i = $lstOps.SelectedIndex
        if ($i -lt 0) { return }
        $list = New-Object System.Collections.Generic.List[object]
        foreach ($op in @($script:ops)) { $list.Add($op) | Out-Null }
        $list.RemoveAt($i)
        $script:ops = @($list)
        Refresh-OpList
        Refresh-Preview
    })

$btnBuf.Add_Click({
        $d = New-Object System.Windows.Forms.FolderBrowserDialog
        $d.Description = "Buffer folder on this PC"
        $d.SelectedPath = $txtBuffer.Text
        if ($d.ShowDialog() -eq "OK") { $txtBuffer.Text = $d.SelectedPath }
    })

$grid.Add_CellEndEdit({
        param($sender, $e)
        if ($e.RowIndex -lt 0) { return }
        if ($grid.Columns[$e.ColumnIndex].Name -ne "new_name") { return }
        $row = $grid.Rows[$e.RowIndex]
        $oldN = [string]$row.Cells["old_name"].Value
        $newN = [string]$row.Cells["new_name"].Value
        if ($oldN) { $script:nameOverrides[$oldN.ToUpperInvariant()] = $newN }
        Refresh-Preview
    })

$btnCancel.Add_Click({ $form.Close() })

$btnSave.Add_Click({
        if ($script:modelPaths.Count -lt 1) {
            [void][System.Windows.Forms.MessageBox]::Show("Choose models first.")
            return
        }
        $years = Get-SelectedYears
        $pref = Get-PrefixValues
        $rows = @(New-JobPreviewRows `
                -OldPaths $script:modelPaths `
                -Ops $script:ops `
                -DoRename $cbRename.Checked `
                -LocationMode (Get-LocationMode) `
                -OldSegment $txtOldSeg.Text.Trim() `
                -NewSegment $txtNewSeg.Text.Trim() `
                -OldPrefix $pref.Old `
                -NewPrefix $pref.New `
                -SourceYear $years.Src `
                -TargetYear $years.Dst `
                -BufferDir $txtBuffer.Text.Trim() `
                -NameOverrides $script:nameOverrides)
        $dup = @($rows | Where-Object { $_.note -match "duplicate" })
        if ($dup.Count -gt 0) {
            [void][System.Windows.Forms.MessageBox]::Show("Duplicate new paths in preview. Fix names/location.")
            return
        }
        $upgrade = $cbUpgrade.Checked
        $loc = Get-LocationMode
        $uncMove = $false
        if (-not $cbRename.Checked) {
            $uncMove = Test-UncMoveCandidate -Rows $rows -Upgrade $upgrade -LocationMode $loc
        }
        $needSaveAs = $true
        if ($uncMove -and -not $upgrade -and -not $cbRename.Checked) {
            $needSaveAs = $false
        }
        if ($cbRelink.Checked -and -not $cbRename.Checked -and $loc -eq "same" -and -not $upgrade) {
            # relink-only: files already at destination
            $needSaveAs = $false
            $uncMove = $false
        }
        $opsJson = @($script:ops | ForEach-Object {
                @{ kind = $_.kind; token = $_.token; from = $_.from; to = $_.to; after = $_.after }
            })
        $job = [pscustomobject]@{
            version        = "1.0.0"
            rename         = [bool]$cbRename.Checked
            relink         = [bool]$cbRelink.Checked
            copy           = [bool]$cbCopy.Checked
            delete_old     = [bool]$cbDelete.Checked
            upgrade        = $upgrade
            source_year    = $years.Src
            target_year    = $years.Dst
            location_mode  = $loc
            old_segment    = $txtOldSeg.Text.Trim()
            new_segment    = $txtNewSeg.Text.Trim()
            old_prefix     = $pref.Old
            new_prefix     = $pref.New
            buffer_dir     = $txtBuffer.Text.Trim()
            rsn_buffer     = "detach"
            copy_links     = $(if ($rbLinksNew.Checked) { "new" } else { "old" })
            saveas_needed  = $needSaveAs
            unc_move       = $uncMove
            ops            = $opsJson
            saved_utc      = [DateTime]::UtcNow.ToString("o")
        }
        $saved = Save-JobFiles -ToolDir $ToolDir -Rows $rows -Job $job
        $msg = "Saved:`r`n{0}`r`n{1}`r`n`r`nRBP checklist:`r`n{2}`r`n`r`n{3}" -f `
            $saved.mapping, $saved.paths, $saved.checklist, ((Get-Content -LiteralPath $saved.checklist -Encoding UTF8) -join "`r`n")
        [void][System.Windows.Forms.MessageBox]::Show($msg, "Job saved")
        $tabs.SelectedTab = $tabPrev
    })

$form.Add_Shown({
        if ($script:modelPaths.Count -gt 0) {
            $txtOldSample.Text = Get-BasenameNoExt $script:modelPaths[0]
            Guess-YearsFromList
        }
        Refresh-OpList
        Refresh-Preview
    })

[void]$form.ShowDialog()
exit 0
