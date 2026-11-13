# Name: ui_grid.ps1
# Version: 3.1.0
# What it does: See the case README.
# Notes: Paths and project names in this portfolio copy are placeholders.

# UTF-8 WITH BOM required (PowerShell 5.1).
# Layout helpers of the operations window: one grid for every tab.
# Everything is placed by TableLayoutPanel / FlowLayoutPanel / Dock, never by
# absolute coordinates, so nothing overlaps or leaves the window when it is resized.
# Dot-source after Add-Type System.Windows.Forms. Builds controls only when called.

$script:OpsGap = 12          # outer margins and the gap between blocks
$script:OpsBtnH = 32         # every button has this height
$script:OpsRowH = 44         # a row of buttons: button + 12 px
$script:OpsFieldRowH = 38    # a settings row
$script:OpsLabelW = 140      # label column of the settings grid
if (-not (Get-Variable -Name OpsFont -Scope Script -ErrorAction SilentlyContinue) -or -not $script:OpsFont) {
    $script:OpsFont = New-Object System.Drawing.Font("Segoe UI", 9.75)
}
$script:OpsHiddenCols = New-Object System.Collections.ArrayList

# Width that fits the text in the window font (and in a wider fallback font) plus padding.
function Get-OpsTextWidth([string]$Text, $Font = $null) {
    if (-not $Font) { $Font = $script:OpsFont }
    $sz = [System.Windows.Forms.TextRenderer]::MeasureText($Text, $Font)
    return [int]$sz.Width
}

function New-OpsButton([string]$Text, [int]$MinWidth = 0) {
    $b = New-Object System.Windows.Forms.Button
    $b.Text = $Text
    $b.AutoSize = $false
    $b.UseVisualStyleBackColor = $true
    $w = (Get-OpsTextWidth $Text) + 32
    if ($w -lt $MinWidth) { $w = $MinWidth }
    if ($w -lt 96) { $w = 96 }
    $b.Size = New-Object System.Drawing.Size($w, $script:OpsBtnH)
    $b.Margin = New-Object System.Windows.Forms.Padding(0, 0, 8, 0)
    $b.Anchor = "Left"
    return $b
}

# A horizontal run of controls that never wraps; it is as wide as its content.
function New-OpsFlow {
    $f = New-Object System.Windows.Forms.FlowLayoutPanel
    $f.FlowDirection = "LeftToRight"
    $f.WrapContents = $false
    $f.AutoSize = $true
    $f.AutoSizeMode = "GrowAndShrink"
    $f.Margin = New-Object System.Windows.Forms.Padding(0)
    $f.Padding = New-Object System.Windows.Forms.Padding(0)
    $f.Anchor = "Left"
    return $f
}

function Add-OpsFlowItems($Flow, $Items) {
    foreach ($c in @($Items)) {
        if ($null -eq $c) { continue }
        [void]$Flow.Controls.Add($c)
    }
    # last item: no gap on the right, so a right-aligned group ends exactly at the edge
    $n = $Flow.Controls.Count
    if ($n -gt 0) {
        $last = $Flow.Controls[$n - 1]
        $m = $last.Margin
        $last.Margin = New-Object System.Windows.Forms.Padding($m.Left, $m.Top, 0, $m.Bottom)
    }
}

# Fixed-size gap between button groups inside a flow.
function New-OpsSpacer([int]$Width = 16) {
    $p = New-Object System.Windows.Forms.Panel
    $p.Size = New-Object System.Drawing.Size($Width, $script:OpsBtnH)
    $p.Margin = New-Object System.Windows.Forms.Padding(0)
    return $p
}

# One row: a group on the left, a group on the right, empty space between them.
# $Middle (optional) takes the space between the groups (e.g. the status line).
function New-OpsBar($Left, $Right, $Middle = $null, [int]$Height = 0) {
    if ($Height -le 0) { $Height = $script:OpsRowH }
    $t = New-Object System.Windows.Forms.TableLayoutPanel
    $t.Dock = "Fill"
    $t.Margin = New-Object System.Windows.Forms.Padding(0)
    $t.Padding = New-Object System.Windows.Forms.Padding(0)
    $t.ColumnCount = 3
    $t.RowCount = 1
    [void]$t.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle ([System.Windows.Forms.SizeType]::AutoSize)))
    [void]$t.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle ([System.Windows.Forms.SizeType]::Percent, 100)))
    [void]$t.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle ([System.Windows.Forms.SizeType]::AutoSize)))
    [void]$t.RowStyles.Add((New-Object System.Windows.Forms.RowStyle ([System.Windows.Forms.SizeType]::Percent, 100)))
    $t.Height = $Height
    if (@($Left | Where-Object { $_ }).Count -gt 0) {
        $lf = New-OpsFlow
        Add-OpsFlowItems $lf $Left
        $t.Controls.Add($lf, 0, 0)
    }
    if ($Middle) {
        if ($Middle -is [System.Windows.Forms.Label]) {
            # one line, cut with "…" when narrow (a docked label would wrap)
            $Middle.AutoSize = $false
            $Middle.AutoEllipsis = $true
            $Middle.Height = $script:OpsFont.Height + 6
            $Middle.Anchor = "Left,Right"
        }
        else { $Middle.Dock = "Fill" }
        $Middle.Margin = New-Object System.Windows.Forms.Padding(0)
        $t.Controls.Add($Middle, 1, 0)
    }
    if (@($Right | Where-Object { $_ }).Count -gt 0) {
        $rf = New-OpsFlow
        $rf.Anchor = "Right"
        $rf.Margin = New-Object System.Windows.Forms.Padding(12, 0, 0, 0)
        Add-OpsFlowItems $rf $Right
        $t.Controls.Add($rf, 2, 0)
    }
    return $t
}

# A label that sits vertically centred in its cell.
# one-line text that sizes itself to its text (never cut with "…"), for a cell in a bar flow
function New-OpsInfoLabel([string]$Text) {
    $l = New-OpsLabel $Text
    $l.AutoSize = $true
    $l.AutoEllipsis = $false
    $l.Margin = New-Object System.Windows.Forms.Padding(0, 0, 24, 0)
    return $l
}
function New-OpsLabel([string]$Text, [bool]$Bold = $false) {
    $l = New-Object System.Windows.Forms.Label
    $l.Text = $Text
    $l.AutoSize = $true
    $l.Anchor = "Left"
    $l.Margin = New-Object System.Windows.Forms.Padding(0, 0, 8, 0)
    if ($Bold) { $l.Font = New-Object System.Drawing.Font($script:OpsFont, [System.Drawing.FontStyle]::Bold) }
    return $l
}

# Counter "В списке: N · отмечено: M" at the right end of the row under a table.
function New-OpsCountLabel {
    $l = New-Object System.Windows.Forms.Label
    $l.AutoSize = $true
    $l.Anchor = "Right"
    $l.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
    $l.Margin = New-Object System.Windows.Forms.Padding(0)
    $l.Text = "В списке: 0 · отмечено: 0"
    return $l
}

function New-OpsListView {
    $lv = New-Object System.Windows.Forms.ListView
    $lv.View = "Details"
    $lv.FullRowSelect = $true
    $lv.GridLines = $true
    $lv.HideSelection = $false
    $lv.Dock = "Fill"
    $lv.Margin = New-Object System.Windows.Forms.Padding(0)
    return $lv
}

# The last column takes the rest of the width. Stretch columns keep their start width at the
# minimum window size and grow in proportion when the window is larger (zero-width columns stay hidden).
$script:OpsFillInfo = @{}
function Set-OpsFillColumn($List) {
    if (-not $List -or $List.Columns.Count -eq 0) { return }
    $info = $script:OpsFillInfo[$List]
    if (-not $info) { return }
    $fill = [int]$info.Fill
    $grow = @($info.Stretch) + @($fill)
    $fixed = 0
    for ($i = 0; $i -lt $List.Columns.Count; $i++) {
        if ($grow -notcontains $i) { $fixed += [Math]::Max(0, $List.Columns[$i].Width) }
    }
    $avail = $List.ClientSize.Width - $fixed - 4
    $base = 0
    foreach ($i in $grow) { $base += [int]$info.Base[$i] }
    if ($base -le 0) { return }
    $info.Busy = $true
    try {
        # shrinking: narrow the last column first, so the columns never get wider than the list
        # for a moment (a horizontal scroll bar shown then may stay after the window got smaller)
        $total = 0
        foreach ($c in $List.Columns) { $total += [Math]::Max(0, $c.Width) }
        if ($total -gt $List.ClientSize.Width -and $List.Columns[$fill].Width -gt 80) { $List.Columns[$fill].Width = 80 }
        if ($avail -ge $base) {
            $used = 0
            foreach ($i in @($info.Stretch)) {
                $w = [int][Math]::Floor([double]$info.Base[$i] * $avail / $base)
                if ($List.Columns[$i].Width -ne $w) { $List.Columns[$i].Width = $w }
                $used += $w
            }
            $w = $avail - $used
        }
        else {
            $used = 0
            foreach ($i in @($info.Stretch)) {
                $w = [int]$info.Base[$i]
                if ($List.Columns[$i].Width -ne $w) { $List.Columns[$i].Width = $w }
                $used += $w
            }
            $w = $avail - $used
        }
        # the client width may have changed meanwhile (a scroll bar came or went)
        $w = $w + ($List.ClientSize.Width - $fixed - 4 - $avail)
        if ($w -lt 80) { $w = 80 }
        if ($List.Columns[$fill].Width -ne $w) { $List.Columns[$fill].Width = $w }
    }
    finally { $info.Busy = $false }
}

function Register-OpsFillColumn($List, [int]$FillIndex, [int[]]$Stretch = @()) {
    $base = @{}
    for ($i = 0; $i -lt $List.Columns.Count; $i++) { $base[$i] = [int]$List.Columns[$i].Width }
    $script:OpsFillInfo[$List] = @{ Fill = $FillIndex; Stretch = @($Stretch); Base = $base; Busy = $false }
    $List.Add_Resize({ param($s, $e) Set-OpsFillColumn $s })
    # a hidden column (width 0) stays hidden even if someone drags its divider
    $List.Add_ColumnWidthChanging({
            param($s, $e)
            if ($s.Columns[$e.ColumnIndex].Width -eq 0 -and $script:OpsHiddenCols.Contains($s.Columns[$e.ColumnIndex])) {
                $e.Cancel = $true
                $e.NewWidth = 0
            }
        })
}

# Width a ListView cell needs for its text (same margins for the cells and the header).
function Get-OpsCellWidth([string]$Text, $Font) {
    if (-not $Text) { return 0 }
    return ([System.Windows.Forms.TextRenderer]::MeasureText($Text, $Font).Width + 16)
}

# Sets every column to the width of its longest text (header included) and returns the sum.
# Call before Register-OpsFillColumn: the last column then takes what is left.
function Set-OpsColumnsToContent($List, $Font) {
    $sum = 0
    for ($i = 0; $i -lt $List.Columns.Count; $i++) {
        $w = Get-OpsCellWidth ([string]$List.Columns[$i].Text) $Font
        foreach ($it in $List.Items) {
            if ($i -ge $it.SubItems.Count) { continue }
            $cw = Get-OpsCellWidth ([string]$it.SubItems[$i].Text) $Font
            if ($cw -gt $w) { $w = $cw }
        }
        $w = [Math]::Max(60, $w)
        $List.Columns[$i].Width = $w
        $sum += $w
    }
    return $sum
}

function Hide-OpsColumn($Column) {
    $Column.Width = 0
    [void]$script:OpsHiddenCols.Add($Column)
}

# Main grid of a tab page: a single column; rows are given as heights (0 = take the rest).
function New-OpsColumnGrid([int[]]$Heights, [int]$Pad = -1) {
    if ($Pad -lt 0) { $Pad = $script:OpsGap }
    $t = New-Object System.Windows.Forms.TableLayoutPanel
    $t.Dock = "Fill"
    $t.Margin = New-Object System.Windows.Forms.Padding(0)
    $t.Padding = New-Object System.Windows.Forms.Padding($Pad)
    $t.ColumnCount = 1
    $t.RowCount = $Heights.Count
    [void]$t.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle ([System.Windows.Forms.SizeType]::Percent, 100)))
    foreach ($h in $Heights) {
        if ($h -le 0) { [void]$t.RowStyles.Add((New-Object System.Windows.Forms.RowStyle ([System.Windows.Forms.SizeType]::Percent, 100))) }
        else { [void]$t.RowStyles.Add((New-Object System.Windows.Forms.RowStyle ([System.Windows.Forms.SizeType]::Absolute, $h))) }
    }
    return $t
}

# Settings block: label column of fixed width | field column that stretches. One setting per row.
function New-OpsSettingsGrid {
    $t = New-Object System.Windows.Forms.TableLayoutPanel
    $t.Dock = "Fill"
    $t.Margin = New-Object System.Windows.Forms.Padding(0)
    $t.Padding = New-Object System.Windows.Forms.Padding(0)
    $t.ColumnCount = 2
    $t.RowCount = 0
    [void]$t.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle ([System.Windows.Forms.SizeType]::Absolute, $script:OpsLabelW)))
    [void]$t.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle ([System.Windows.Forms.SizeType]::Percent, 100)))
    return $t
}

function Add-OpsSettingRow($Grid, [string]$Label, $Field) {
    $row = $Grid.RowCount
    $Grid.RowCount = $row + 1
    [void]$Grid.RowStyles.Add((New-Object System.Windows.Forms.RowStyle ([System.Windows.Forms.SizeType]::Absolute, $script:OpsFieldRowH)))
    $l = New-OpsLabel $Label
    $Grid.Controls.Add($l, 0, $row)
    if ($Field) {
        if ($Field -is [System.Windows.Forms.TableLayoutPanel]) { $Field.Dock = "Fill" }
        $Grid.Controls.Add($Field, 1, $row)
    }
    return $l
}

function Get-OpsSettingsHeight($Grid) {
    return ($Grid.RowCount * $script:OpsFieldRowH)
}

# Grey hint inside an empty text box (EM_SETCUEBANNER); silently skipped where unsupported.
function Set-OpsCue($Box, [string]$Text) {
    if (-not ("OpsNativeCue" -as [type])) {
        try {
            Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public static class OpsNativeCue {
    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    public static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, string lParam);
}
"@ -ErrorAction Stop
        }
        catch { return }
    }
    $Box.Tag = $Text
    $apply = {
        param($s, $e)
        try { [void][OpsNativeCue]::SendMessage($s.Handle, 0x1501, [IntPtr]1, [string]$s.Tag) } catch { }
    }
    if ($Box.IsHandleCreated) { & $apply $Box $null }
    $Box.Add_HandleCreated($apply)
}

function New-OpsTextBox([string]$Cue = "") {
    $t = New-Object System.Windows.Forms.TextBox
    $t.Anchor = "Left,Right"
    $t.Margin = New-Object System.Windows.Forms.Padding(0, 0, 8, 0)
    if ($Cue) { Set-OpsCue $t $Cue }
    return $t
}

# Field cell "[было] → [стало] [buttons...]": both boxes share the width equally.
function New-OpsPairField($OldBox, $NewBox, $Buttons = @()) {
    $t = New-Object System.Windows.Forms.TableLayoutPanel
    $t.Dock = "Fill"
    $t.Margin = New-Object System.Windows.Forms.Padding(0)
    $t.Padding = New-Object System.Windows.Forms.Padding(0)
    $btns = @($Buttons | Where-Object { $_ })
    $t.ColumnCount = 3 + $btns.Count
    $t.RowCount = 1
    [void]$t.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle ([System.Windows.Forms.SizeType]::Percent, 50)))
    [void]$t.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle ([System.Windows.Forms.SizeType]::Absolute, 28)))
    [void]$t.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle ([System.Windows.Forms.SizeType]::Percent, 50)))
    foreach ($b in $btns) { [void]$t.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle ([System.Windows.Forms.SizeType]::AutoSize))) }
    [void]$t.RowStyles.Add((New-Object System.Windows.Forms.RowStyle ([System.Windows.Forms.SizeType]::Percent, 100)))
    $OldBox.Anchor = "Left,Right"
    $OldBox.Margin = New-Object System.Windows.Forms.Padding(0)
    $NewBox.Anchor = "Left,Right"
    $NewBox.Margin = New-Object System.Windows.Forms.Padding(0)
    $arrow = New-Object System.Windows.Forms.Label
    $arrow.Text = "→"
    $arrow.AutoSize = $false
    $arrow.Dock = "Fill"
    $arrow.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $arrow.Margin = New-Object System.Windows.Forms.Padding(0)
    $t.Controls.Add($OldBox, 0, 0)
    $t.Controls.Add($arrow, 1, 0)
    $t.Controls.Add($NewBox, 2, 0)
    $col = 3
    foreach ($b in $btns) {
        $b.Margin = New-Object System.Windows.Forms.Padding(8, 0, 0, 0)
        $t.Controls.Add($b, $col, 0)
        $col++
    }
    return $t
}

# Field cell "[box][buttons...]": the box stretches.
function New-OpsBoxField($Box, $Buttons = @()) {
    $t = New-Object System.Windows.Forms.TableLayoutPanel
    $t.Dock = "Fill"
    $t.Margin = New-Object System.Windows.Forms.Padding(0)
    $t.Padding = New-Object System.Windows.Forms.Padding(0)
    $btns = @($Buttons | Where-Object { $_ })
    $t.ColumnCount = 1 + $btns.Count
    $t.RowCount = 1
    [void]$t.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle ([System.Windows.Forms.SizeType]::Percent, 100)))
    foreach ($b in $btns) { [void]$t.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle ([System.Windows.Forms.SizeType]::AutoSize))) }
    [void]$t.RowStyles.Add((New-Object System.Windows.Forms.RowStyle ([System.Windows.Forms.SizeType]::Percent, 100)))
    $Box.Anchor = "Left,Right"
    $Box.Margin = New-Object System.Windows.Forms.Padding(0)
    $t.Controls.Add($Box, 0, 0)
    $col = 1
    foreach ($b in $btns) {
        $b.Margin = New-Object System.Windows.Forms.Padding(8, 0, 0, 0)
        $t.Controls.Add($b, $col, 0)
        $col++
    }
    return $t
}

# Field cell with controls in a row from the left (radio buttons, a combo...).
function New-OpsInlineField($Items) {
    $f = New-OpsFlow
    foreach ($c in @($Items)) {
        if ($null -eq $c) { continue }
        if (-not ($c -is [System.Windows.Forms.Button])) {
            $c.Margin = New-Object System.Windows.Forms.Padding(0, 0, 24, 0)
        }
        [void]$f.Controls.Add($c)
    }
    return $f
}

function New-OpsRadio([string]$Text, [bool]$Checked = $false) {
    $r = New-Object System.Windows.Forms.RadioButton
    $r.Text = $Text
    $r.AutoSize = $true
    $r.Checked = $Checked
    $r.Anchor = "Left"
    return $r
}

function New-OpsCheck([string]$Text, [bool]$Checked = $false) {
    $c = New-Object System.Windows.Forms.CheckBox
    $c.Text = $Text
    $c.AutoSize = $true
    $c.Checked = $Checked
    $c.Anchor = "Left"
    return $c
}

# A state switch the job builder reads (.Checked), not a control on the screen.
function New-OpsFlag([bool]$On = $false) {
    return [pscustomobject]@{ Checked = $On; Enabled = $true; Visible = $false }
}
