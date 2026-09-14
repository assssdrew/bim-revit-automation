# Main window for compact save. UTF-8 WITH BOM required (PowerShell 5.1).
# Double-click Сжатие.cmd — no console menu.

$ErrorActionPreference = "Continue"
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$ToolDir = Split-Path -Parent $MyInvocation.MyCommand.Path

try {
    $CompactUiLibrary = $true
    $CompactLaunchLibrary = $true
    . (Join-Path $ToolDir "choose_models.ps1")
    . (Join-Path $ToolDir "start_compact.ps1")
    # start_compact sets Stop — for the window that makes Test-Path on a dead UNC crash the form.
    $ErrorActionPreference = "Continue"
}
catch {
    $msg = $_.Exception.Message
    if ($_.InvocationInfo -and $_.InvocationInfo.PositionMessage) {
        $msg = $msg + [Environment]::NewLine + $_.InvocationInfo.PositionMessage
    }
    [System.Windows.Forms.MessageBox]::Show(
        $msg,
        "Сжатие моделей Revit",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
    throw
}

$script:SelectedModels = New-Object System.Collections.Generic.List[string]
$script:ModelStatus = @{}
$script:RunState = @{
    Active     = $false
    Process    = $null
    OpenReport = $false
    Total      = 0
    YearQueue  = @()
    YearIndex  = 1
    YearTotal  = 1
    YearLabel  = ""
    Exe        = ""
    ExeDir     = ""
    Mode       = "deep"
    ExitCodes  = @()
}
$script:Ui = @{}

function Get-UiSettingsPath {
    $dir = Join-Path $env:LOCALAPPDATA "BatchRvt\batch_compact_save"
    if (-not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
    return (Join-Path $dir "ui_settings.txt")
}

function Read-UiSettings {
    $p = Get-UiSettingsPath
    $openReport = $false
    if (Test-Path -LiteralPath $p) {
        foreach ($line in @(Get-Content -LiteralPath $p -ErrorAction SilentlyContinue)) {
            if ($line -match '^(?i)open_report=(1|true|yes)$') { $openReport = $true }
        }
    }
    return @{ OpenReport = $openReport }
}

function Save-UiSettings {
    $on = if ($script:Ui.ChkOpenReport -and $script:Ui.ChkOpenReport.Checked) { "1" } else { "0" }
    Set-Content -LiteralPath (Get-UiSettingsPath) -Value ("open_report={0}" -f $on) -Encoding UTF8
}

function Get-ModelKey([string]$Path) {
    if ([string]::IsNullOrWhiteSpace($Path)) { return "" }
    $p = $Path.Trim()
    if ($p -match '^(?i)RSN://') { return $p.TrimEnd("/").ToLower() }
    return $p.TrimEnd("\").ToLower()
}

function Get-StatusText([string]$Path) {
    $key = Get-ModelKey $Path
    if ($script:ModelStatus.ContainsKey($key)) { return [string]$script:ModelStatus[$key] }
    return ""
}

function Get-ModelDisplayName([string]$Path) {
    if ([string]::IsNullOrWhiteSpace($Path)) { return "" }
    if ($Path -match '^(?i)RSN://') {
        $tail = $Path.Substring(6).Trim("/")
        $parts = @($tail -split "/")
        if ($parts.Count -ge 2) {
            return ("{0}  ({1})" -f $parts[-1], $parts[0])
        }
        return $Path
    }
    return [System.IO.Path]::GetFileName($Path)
}

function Get-ModelSource([string]$Path) {
    if ($Path -match '^(?i)RSN://') { return "Revit Server" }
    if ($Path.StartsWith("\\")) { return "Сеть" }
    return "Этот компьютер"
}

function Refresh-ModelList {
    $lv = $script:Ui.List
    $lv.BeginUpdate()
    try {
        $lv.Items.Clear()
        foreach ($p in @($script:SelectedModels)) {
            $item = New-Object System.Windows.Forms.ListViewItem((Get-ModelDisplayName $p))
            [void]$item.SubItems.Add((Get-StatusText $p))
            [void]$item.SubItems.Add((Get-ModelSource $p))
            [void]$item.SubItems.Add($p)
            $item.Tag = $p
            [void]$lv.Items.Add($item)
        }
    }
    finally {
        $lv.EndUpdate()
    }
    $n = $script:SelectedModels.Count
    $script:Ui.CountLabel.Text = if ($n -eq 0) { "Модели не выбраны" } else { ("Выбрано моделей: {0}" -f $n) }
    $busy = [bool]$script:RunState.Active
    $script:Ui.BtnRun.Enabled = ($n -gt 0) -and (-not $busy)
}

function Set-ModelStatus {
    param([string]$Path, [string]$Text)
    if ([string]::IsNullOrWhiteSpace($Path)) { return }
    $script:ModelStatus[(Get-ModelKey $Path)] = $Text
    $lv = $script:Ui.List
    foreach ($item in @($lv.Items)) {
        if ((Get-ModelKey ([string]$item.Tag)) -eq (Get-ModelKey $Path)) {
            if ($item.SubItems.Count -gt 1) { $item.SubItems[1].Text = $Text }
            break
        }
    }
}

function Update-RunProgressBar {
    $done = 0
    foreach ($p in @($script:SelectedModels)) {
        $st = Get-StatusText $p
        if ($st -match 'Готово|Ошибка|Пропуск') { $done++ }
    }
    $total = [Math]::Max(1, [int]$script:RunState.Total)
    if ($script:Ui.Progress) {
        $script:Ui.Progress.Maximum = $total
        $script:Ui.Progress.Value = [Math]::Min($done, $total)
        $script:Ui.Progress.Visible = [bool]$script:RunState.Active -or ($done -gt 0)
    }
    $current = ""
    foreach ($p in @($script:SelectedModels)) {
        if ((Get-StatusText $p) -eq "Идёт") {
            $current = Get-ModelDisplayName $p
            break
        }
    }
    if ($script:RunState.Active) {
        $msg = ("Готово {0} из {1}" -f $done, $script:RunState.Total)
        if ($current) { $msg += (" · сейчас: {0}" -f $current) }
        Set-UiStatus $msg
    }
}

function Set-UiStatus([string]$Text) {
    $script:Ui.Status.Text = $Text
    [System.Windows.Forms.Application]::DoEvents()
}

function Add-UiModels([string[]]$Paths) {
    $added = 0
    $seen = @{}
    foreach ($old in @($script:SelectedModels)) { $seen[$old] = $true }
    foreach ($raw in @($Paths)) {
        if ([string]::IsNullOrWhiteSpace($raw)) { continue }
        $p = $raw.Trim()
        if ($seen.ContainsKey($p)) { continue }
        if ($p -notmatch '^(?i)RSN://' -and (Get-Command Test-ExcludedModelPath -ErrorAction SilentlyContinue)) {
            if (Test-ExcludedModelPath $p) { continue }
        }
        $seen[$p] = $true
        [void]$script:SelectedModels.Add($p)
        $added++
    }
    Refresh-ModelList
    return $added
}

function Get-ModelsFromFolder([string]$Folder) {
    if ([string]::IsNullOrWhiteSpace($Folder)) { return @() }

    $found = New-Object System.Collections.Generic.List[string]
    $direct = @(
        Get-ChildItem -LiteralPath $Folder -File -Filter *.rvt -Force -ErrorAction SilentlyContinue
    )
    foreach ($f in $direct) {
        if (-not (Test-IsTargetRvtName $f.Name)) { continue }
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

function Show-ServerPicker {
    Ensure-ServersFile
    $servers = @(Get-Servers)
    if ($servers.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            $script:Ui.Form,
            "В servers.cfg нет серверов.",
            "Revit Server",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
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
    $form.Font = $script:Ui.Form.Font

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = "Выберите сервер:"
    $lbl.Location = New-Object System.Drawing.Point(12, 12)
    $lbl.AutoSize = $true

    $list = New-Object System.Windows.Forms.ListBox
    $list.Location = New-Object System.Drawing.Point(12, 36)
    $list.Size = New-Object System.Drawing.Size(420, 230)
    $list.IntegralHeight = $false
    foreach ($s in $servers) {
        [void]$list.Items.Add(("{0}    {1}" -f $s.Label, $s.Host))
    }
    if ($list.Items.Count -gt 0) { $list.SelectedIndex = 0 }

    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = "Открыть"
    $btnOk.Location = New-Object System.Drawing.Point(236, 280)
    $btnOk.Size = New-Object System.Drawing.Size(96, 28)
    $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Отмена"
    $btnCancel.Location = New-Object System.Drawing.Point(336, 280)
    $btnCancel.Size = New-Object System.Drawing.Size(96, 28)
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

    $form.AcceptButton = $btnOk
    $form.CancelButton = $btnCancel
    $form.Controls.AddRange(@($lbl, $list, $btnOk, $btnCancel))

    if ($form.ShowDialog($script:Ui.Form) -ne [System.Windows.Forms.DialogResult]::OK) {
        return $null
    }
    if ($list.SelectedIndex -lt 0) { return $null }
    return $servers[$list.SelectedIndex]
}

function Add-FromFiles {
    $picked = @(Select-LocalFiles)
    $n = Add-UiModels $picked
    if ($n -gt 0) { Set-UiStatus ("Добавлено файлов: {0}" -f $n) }
    elseif ($picked.Count -eq 0) { Set-UiStatus "Файлы не выбраны" }
    else { Set-UiStatus "Эти модели уже в списке" }
}

function Add-FromFolders {
    try {
        $folder = Show-FolderPathDialog
        if ([string]::IsNullOrWhiteSpace($folder)) {
            Set-UiStatus "Папка не выбрана"
            return
        }
        $script:Ui.Form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        Set-UiStatus "Ищу модели в папке…"
        $picked = @(Get-ModelsFromFolder $folder)
        $n = Add-UiModels $picked
        if ($n -gt 0) {
            Set-UiStatus ("Из папки добавлено: {0}" -f $n)
        }
        else {
            Set-UiStatus "В папке нет рабочих .rvt (резервы пропускаются)"
        }
    }
    catch {
        Set-UiStatus "Папка недоступна"
        [System.Windows.Forms.MessageBox]::Show(
            $script:Ui.Form,
            "Не удалось прочитать папку. Если это сеть — проверьте путь и доступ." + [Environment]::NewLine + [Environment]::NewLine + $_.Exception.Message,
            "Папки",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
    }
    finally {
        $script:Ui.Form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
}

function Add-FromServer {
    $serverObj = Show-ServerPicker
    if (-not $serverObj) {
        Set-UiStatus "Сервер не выбран"
        return
    }
    Set-UiStatus ("Сервер: {0}" -f $serverObj.Label)
    $picked = @(Show-RevitServerBrowser -ServerObj $serverObj)
    $n = Add-UiModels $picked
    if ($n -gt 0) { Set-UiStatus ("С сервера добавлено: {0}" -f $n) }
    elseif ($picked.Count -eq 0) { Set-UiStatus "С сервера ничего не добавлено" }
    else { Set-UiStatus "Эти модели уже в списке" }
}

function Remove-SelectedUiModels {
    $lv = $script:Ui.List
    if ($lv.SelectedItems.Count -eq 0) { return }
    $drop = @{}
    foreach ($it in @($lv.SelectedItems)) { $drop[[string]$it.Tag] = $true }
    $keep = New-Object System.Collections.Generic.List[string]
    foreach ($p in @($script:SelectedModels)) {
        if (-not $drop.ContainsKey($p)) { [void]$keep.Add($p) }
    }
    $script:SelectedModels = $keep
    Refresh-ModelList
    Set-UiStatus "Убраны выделенные строки"
}

function Clear-UiModels {
    $script:SelectedModels.Clear()
    $script:ModelStatus = @{}
    Refresh-ModelList
    Set-UiStatus "Список очищен"
}

function Get-SelectedMode {
    if ($script:Ui.RadioQuick.Checked) { return "quick" }
    return "deep"
}

function Update-ModeUi {
    $deep = -not $script:Ui.RadioQuick.Checked
    $script:Ui.Hint.Visible = $deep
}

function Read-LiveStatusFile {
    $p = $null
    if (Get-Command Get-LiveStatusPath -ErrorAction SilentlyContinue) {
        $p = Get-LiveStatusPath
    }
    else {
        $p = Join-Path $env:LOCALAPPDATA "BatchRvt\batch_compact_save\live_status.txt"
    }
    if (-not $p -or -not (Test-Path -LiteralPath $p)) { return $null }
    $map = @{}
    foreach ($line in @(Get-Content -LiteralPath $p -Encoding UTF8 -ErrorAction SilentlyContinue)) {
        if ($line -match '^(path|state|message)=(.*)$') {
            $map[$Matches[1]] = $Matches[2]
        }
    }
    if (-not $map.ContainsKey("path")) { return $null }
    return $map
}

function Apply-LiveStatus {
    $live = Read-LiveStatusFile
    if (-not $live) { return }
    $state = [string]$live["state"]
    $path = [string]$live["path"]
    $text = "Идёт"
    if ($state -eq "ok") { $text = "Готово" }
    elseif ($state -eq "warn") { $text = "Готово" }
    elseif ($state -eq "error") { $text = "Ошибка" }
    elseif ($state -eq "start") { $text = "Идёт" }
    Set-ModelStatus -Path $path -Text $text
    Update-RunProgressBar
}

function Set-UiBusy([bool]$Busy) {
    $script:RunState.Active = $Busy
    foreach ($name in @("BtnFiles", "BtnFolders", "BtnServer", "BtnRemove", "BtnClear", "BtnImport", "RadioDeep", "RadioQuick")) {
        if ($script:Ui.ContainsKey($name) -and $script:Ui[$name]) {
            $script:Ui[$name].Enabled = -not $Busy
        }
    }
    $n = $script:SelectedModels.Count
    $script:Ui.BtnRun.Enabled = ($n -gt 0) -and (-not $Busy)
    if ($script:Ui.Progress) { $script:Ui.Progress.Visible = $Busy }
}

function Find-LatestReportXlsx {
    $dirs = @(
        (Join-Path $ToolDir "отчеты")
        (Join-Path $env:LOCALAPPDATA "BatchRvt\batch_compact_save\отчеты")
    )
    $files = @()
    foreach ($d in $dirs) {
        if (Test-Path -LiteralPath $d) {
            $files += @(Get-ChildItem -LiteralPath $d -Filter *.xlsx -File -ErrorAction SilentlyContinue)
        }
    }
    if ($files.Count -eq 0) { return $null }
    return ($files | Sort-Object LastWriteTime -Descending | Select-Object -First 1).FullName
}

function Complete-CompactRun {
    if ($script:Ui.Timer) { $script:Ui.Timer.Stop() }
    Apply-LiveStatus
    foreach ($p in @($script:SelectedModels)) {
        $st = Get-StatusText $p
        if ($st -eq "Ожидание" -or $st -eq "Идёт") {
            Set-ModelStatus -Path $p -Text "Не обработана"
        }
    }
    $ok = 0; $err = 0
    foreach ($p in @($script:SelectedModels)) {
        $st = Get-StatusText $p
        if ($st -eq "Готово") { $ok++ }
        elseif ($st -eq "Ошибка") { $err++ }
    }
    Set-UiBusy $false
    Update-RunProgressBar
    $code = 0
    try {
        $codes = @($script:RunState.ExitCodes)
        $bad = @($codes | Where-Object { $_ -ne 0 })
        if ($bad.Count -gt 0) { $code = [int]$bad[-1] }
        elseif ($script:RunState.Process) {
            $proc = $script:RunState.Process
            $proc.Refresh()
            $code = [int]$proc.ExitCode
        }
    }
    catch { }
    $script:RunState.Process = $null
    $script:RunState.YearQueue = @()
    $msg = ("Готово: {0} успешно, ошибок: {1}" -f $ok, $err)
    if ($code -ne 0) { $msg += (" · RBP код {0}" -f $code) }
    Set-UiStatus $msg

    $xlsx = Find-LatestReportXlsx
    if ($script:RunState.OpenReport -and $xlsx) {
        try { Start-Process -FilePath $xlsx | Out-Null } catch { }
    }
}

function Start-NextYearIfNeeded {
    $next = @($script:RunState.YearQueue)
    if ($next.Count -eq 0) { return $false }
    $g = $next[0]
    $script:RunState.YearQueue = @($next | Select-Object -Skip 1)
    $script:RunState.YearIndex = [int]$script:RunState.YearIndex + 1
    $script:RunState.YearLabel = [string]$g.Label
    try {
        $started = Start-CompactYearBatch -Group $g -Mode $script:RunState.Mode -Exe $script:RunState.Exe -ExeDir $script:RunState.ExeDir -NoWait
    }
    catch {
        Set-UiStatus $_.Exception.Message
        return $false
    }
    if (-not $started.Ok -or -not $started.Process) {
        Set-UiStatus $started.Message
        return $false
    }
    $script:RunState.Process = $started.Process
    Set-UiStatus ("Revit {0} ({1} из {2})" -f $script:RunState.YearLabel, $script:RunState.YearIndex, $script:RunState.YearTotal)
    return $true
}

function Tick-CompactRun {
    if (-not $script:RunState.Active) { return }
    Apply-LiveStatus
    $proc = $script:RunState.Process
    if ($proc) {
        try { $proc.Refresh() } catch { }
        if ($proc.HasExited) {
            $code = 0
            try { $code = [int]$proc.ExitCode } catch { }
            $script:RunState.ExitCodes += $code
            $script:RunState.Process = $null
            if (Start-NextYearIfNeeded) { return }
            Complete-CompactRun
        }
    }
}

function Export-UiList {
    if ($script:SelectedModels.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            $script:Ui.Form, "Список пуст.", "Экспорт",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
        return
    }
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Title = "Сохранить список моделей"
    $dlg.Filter = "Список моделей (*.txt)|*.txt|Все файлы (*.*)|*.*"
    $dlg.FileName = ("список_моделей_{0}.txt" -f (Get-Date -Format "yyyy-MM-dd"))
    $dlg.OverwritePrompt = $true
    if ((Show-OwnedDialog $dlg) -ne [System.Windows.Forms.DialogResult]::OK) { return }
    try {
        Write-TextFileLines -Path $dlg.FileName -Lines @($script:SelectedModels)
        Set-UiStatus ("Список сохранён: {0}" -f $dlg.FileName)
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            $script:Ui.Form, $_.Exception.Message, "Экспорт",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
    }
}

function Import-UiList {
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Title = "Открыть список моделей"
    $dlg.Filter = "Список моделей (*.txt)|*.txt|Все файлы (*.*)|*.*"
    $dlg.Multiselect = $false
    if ((Show-OwnedDialog $dlg) -ne [System.Windows.Forms.DialogResult]::OK) { return }
    $paths = @(
        Get-Content -LiteralPath $dlg.FileName -ErrorAction SilentlyContinue |
            ForEach-Object { $_.Trim() } |
            Where-Object {
                $_ -and -not $_.StartsWith("#") -and (
                    $_.ToLower().EndsWith(".rvt") -or $_ -match '^(?i)RSN://'
                )
            }
    )
    if ($paths.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            $script:Ui.Form, "В файле нет путей к моделям.", "Импорт",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
        return
    }
    $append = $false
    if ($script:SelectedModels.Count -gt 0) {
        $ask = [System.Windows.Forms.MessageBox]::Show(
            $script:Ui.Form,
            "В списке уже есть модели. Добавить к ним? (Нет — заменить список)",
            "Импорт",
            [System.Windows.Forms.MessageBoxButtons]::YesNoCancel,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )
        if ($ask -eq [System.Windows.Forms.DialogResult]::Cancel) { return }
        $append = ($ask -eq [System.Windows.Forms.DialogResult]::Yes)
    }
    if (-not $append) {
        $script:SelectedModels.Clear()
        $script:ModelStatus = @{}
    }
    $n = Add-UiModels $paths
    Set-UiStatus ("Импортировано: {0}" -f $n)
}

function Start-UiCompact {
    $models = @($script:SelectedModels)
    if ($models.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            $script:Ui.Form,
            "Сначала добавьте модели.",
            "Сжатие",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        ) | Out-Null
        return
    }

    $mode = Get-SelectedMode
    if ($mode -eq "deep") {
        $ask = [System.Windows.Forms.MessageBox]::Show(
            $script:Ui.Form,
            "Глубокое сжатие перезапишет хранилище (ФХ)." + [Environment]::NewLine + [Environment]::NewLine +
            "Если кто-то ещё в модели — его работу можно повредить." + [Environment]::NewLine +
            "После прогона локали у остальных нужно создать заново." + [Environment]::NewLine + [Environment]::NewLine +
            "Все вышли из файлов. Продолжить?",
            "Глубокое сжатие",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning,
            [System.Windows.Forms.MessageBoxDefaultButton]::Button2
        )
        if ($ask -ne [System.Windows.Forms.DialogResult]::Yes) {
            Set-UiStatus "Запуск отменён"
            return
        }
    }

    $script:Ui.Form.UseWaitCursor = $false
    $script:Ui.Form.Cursor = [System.Windows.Forms.Cursors]::Default
    $script:ModelStatus = @{}
    foreach ($p in $models) { Set-ModelStatus -Path $p -Text "Ожидание" }
    Refresh-ModelList
    $script:RunState.Total = $models.Count
    $script:RunState.OpenReport = [bool]$script:Ui.ChkOpenReport.Checked
    Save-UiSettings
    try { Remove-Item -LiteralPath (Get-LiveStatusPath) -Force -ErrorAction SilentlyContinue } catch { }

    Set-UiBusy $true
    Set-UiStatus "Запускаю Revit Batch Processor…"
    $script:RunState.YearQueue = @()
    $script:RunState.ExitCodes = @()
    $script:RunState.Mode = $mode
    try {
        $result = Invoke-CompactLaunch -Mode $mode -Models $models -NoWait
    }
    catch {
        $result = @{ Ok = $false; Message = $_.Exception.Message; Process = $null }
    }

    if (-not $result.Ok -or -not $result.Process) {
        Set-UiBusy $false
        Set-UiStatus $result.Message
        [System.Windows.Forms.MessageBox]::Show(
            $script:Ui.Form,
            $result.Message,
            "Сжатие",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        ) | Out-Null
        return
    }

    $script:RunState.Process = $result.Process
    $script:RunState.YearQueue = @($result.YearQueue)
    $script:RunState.YearIndex = [int]$(if ($result.YearIndex) { $result.YearIndex } else { 1 })
    $script:RunState.YearTotal = [int]$(if ($result.YearTotal) { $result.YearTotal } else { 1 })
    $script:RunState.YearLabel = [string]$result.YearLabel
    $script:RunState.Exe = [string]$result.Exe
    $script:RunState.ExeDir = [string]$result.ExeDir
    if ($script:Ui.Progress) {
        $script:Ui.Progress.Value = 0
        $script:Ui.Progress.Maximum = [Math]::Max(1, $models.Count)
        $script:Ui.Progress.Visible = $true
    }
    $script:Ui.Timer.Start()
    Set-UiStatus $result.Message
}

function New-MainForm {
    $font = New-Object System.Drawing.Font("Segoe UI", 9.75)

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Сжатие моделей Revit"
    $form.Size = New-Object System.Drawing.Size(840, 650)
    $form.MinimumSize = New-Object System.Drawing.Size(720, 520)
    $form.StartPosition = "CenterScreen"
    $form.Font = $font
    $form.BackColor = [System.Drawing.SystemColors]::Window

    $lblAdd = New-Object System.Windows.Forms.Label
    $lblAdd.Text = "Добавить модели"
    $lblAdd.Location = New-Object System.Drawing.Point(16, 14)
    $lblAdd.AutoSize = $true
    $lblAdd.Font = New-Object System.Drawing.Font("Segoe UI", 9.75, [System.Drawing.FontStyle]::Bold)

    $btnFiles = New-Object System.Windows.Forms.Button
    $btnFiles.Text = "Файлы…"
    $btnFiles.Location = New-Object System.Drawing.Point(16, 40)
    $btnFiles.Size = New-Object System.Drawing.Size(150, 36)

    $btnFolders = New-Object System.Windows.Forms.Button
    $btnFolders.Text = "Папки…"
    $btnFolders.Location = New-Object System.Drawing.Point(176, 40)
    $btnFolders.Size = New-Object System.Drawing.Size(150, 36)

    $btnServer = New-Object System.Windows.Forms.Button
    $btnServer.Text = "Revit Server…"
    $btnServer.Location = New-Object System.Drawing.Point(336, 40)
    $btnServer.Size = New-Object System.Drawing.Size(170, 36)

    $lblCount = New-Object System.Windows.Forms.Label
    $lblCount.Text = "Модели не выбраны"
    $lblCount.Location = New-Object System.Drawing.Point(16, 88)
    $lblCount.AutoSize = $true

    $list = New-Object System.Windows.Forms.ListView
    $list.Location = New-Object System.Drawing.Point(16, 112)
    $list.Size = New-Object System.Drawing.Size(792, 250)
    $list.Anchor = "Top,Bottom,Left,Right"
    $list.View = [System.Windows.Forms.View]::Details
    $list.FullRowSelect = $true
    $list.MultiSelect = $true
    $list.HideSelection = $false
    $list.GridLines = $true
    [void]$list.Columns.Add("Модель", 220)
    [void]$list.Columns.Add("Состояние", 110)
    [void]$list.Columns.Add("Откуда", 100)
    [void]$list.Columns.Add("Путь", 340)

    $btnRemove = New-Object System.Windows.Forms.Button
    $btnRemove.Text = "Убрать выбранные"
    $btnRemove.Location = New-Object System.Drawing.Point(16, 372)
    $btnRemove.Size = New-Object System.Drawing.Size(160, 30)
    $btnRemove.Anchor = "Bottom,Left"

    $btnClear = New-Object System.Windows.Forms.Button
    $btnClear.Text = "Очистить список"
    $btnClear.Location = New-Object System.Drawing.Point(184, 372)
    $btnClear.Size = New-Object System.Drawing.Size(150, 30)
    $btnClear.Anchor = "Bottom,Left"

    $btnExport = New-Object System.Windows.Forms.Button
    $btnExport.Text = "Экспорт списка…"
    $btnExport.Location = New-Object System.Drawing.Point(342, 372)
    $btnExport.Size = New-Object System.Drawing.Size(150, 30)
    $btnExport.Anchor = "Bottom,Left"

    $btnImport = New-Object System.Windows.Forms.Button
    $btnImport.Text = "Импорт списка…"
    $btnImport.Location = New-Object System.Drawing.Point(500, 372)
    $btnImport.Size = New-Object System.Drawing.Size(150, 30)
    $btnImport.Anchor = "Bottom,Left"

    $grpMode = New-Object System.Windows.Forms.GroupBox
    $grpMode.Text = "Режим"
    $grpMode.Location = New-Object System.Drawing.Point(16, 414)
    $grpMode.Size = New-Object System.Drawing.Size(600, 128)
    $grpMode.Anchor = "Bottom,Left,Right"

    $radioDeep = New-Object System.Windows.Forms.RadioButton
    $radioDeep.Text = "Глубокое сжатие (хранилище, проверка ФХ)"
    $radioDeep.Location = New-Object System.Drawing.Point(14, 24)
    $radioDeep.AutoSize = $true
    $radioDeep.Checked = $true

    $radioQuick = New-Object System.Windows.Forms.RadioButton
    $radioQuick.Text = "Быстрое сжатие (локаль + синхронизация)"
    $radioQuick.Location = New-Object System.Drawing.Point(14, 48)
    $radioQuick.AutoSize = $true

    $hint = New-Object System.Windows.Forms.Label
    $hint.Text = "Перезаписывает хранилище. Перед запуском все должны выйти из моделей. Локали после прогона создают заново."
    $hint.Location = New-Object System.Drawing.Point(14, 76)
    $hint.Size = New-Object System.Drawing.Size(570, 42)
    $hint.Anchor = "Top,Left,Right"
    $hint.ForeColor = [System.Drawing.SystemColors]::GrayText

    $grpMode.Controls.AddRange(@($radioDeep, $radioQuick, $hint))

    $chkOpenReport = New-Object System.Windows.Forms.CheckBox
    $chkOpenReport.Text = "Открыть отчёт после прогона"
    $chkOpenReport.Location = New-Object System.Drawing.Point(640, 414)
    $chkOpenReport.Size = New-Object System.Drawing.Size(168, 36)
    $chkOpenReport.Anchor = "Bottom,Right"

    $btnRun = New-Object System.Windows.Forms.Button
    $btnRun.Text = "Сжать"
    $btnRun.Location = New-Object System.Drawing.Point(640, 448)
    $btnRun.Size = New-Object System.Drawing.Size(168, 48)
    $btnRun.Anchor = "Bottom,Right"
    $btnRun.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)

    $progress = New-Object System.Windows.Forms.ProgressBar
    $progress.Location = New-Object System.Drawing.Point(16, 546)
    $progress.Size = New-Object System.Drawing.Size(792, 14)
    $progress.Anchor = "Bottom,Left,Right"
    $progress.Visible = $false

    $status = New-Object System.Windows.Forms.Label
    $status.Text = "Добавьте модели и нажмите «Сжать»."
    $status.Location = New-Object System.Drawing.Point(16, 564)
    $status.Size = New-Object System.Drawing.Size(792, 22)
    $status.Anchor = "Bottom,Left,Right"
    $status.ForeColor = [System.Drawing.SystemColors]::GrayText

    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 1000
    $timer.Add_Tick({ Tick-CompactRun })

    $form.Controls.AddRange(@(
        $lblAdd, $btnFiles, $btnFolders, $btnServer,
        $lblCount, $list, $btnRemove, $btnClear, $btnExport, $btnImport,
        $grpMode, $chkOpenReport, $btnRun, $progress, $status
    ))

    $script:Ui = @{
        Form         = $form
        List         = $list
        CountLabel   = $lblCount
        Status       = $status
        Progress     = $progress
        Timer        = $timer
        BtnRun       = $btnRun
        BtnFiles     = $btnFiles
        BtnFolders   = $btnFolders
        BtnServer    = $btnServer
        BtnRemove    = $btnRemove
        BtnClear     = $btnClear
        BtnImport    = $btnImport
        RadioQuick   = $radioQuick
        RadioDeep    = $radioDeep
        Hint         = $hint
        ChkOpenReport = $chkOpenReport
    }
    $script:UiOwner = $form

    $btnFiles.Add_Click({ Add-FromFiles })
    $btnFolders.Add_Click({ Add-FromFolders })
    $btnServer.Add_Click({ Add-FromServer })
    $btnRemove.Add_Click({ Remove-SelectedUiModels })
    $btnClear.Add_Click({ Clear-UiModels })
    $btnExport.Add_Click({ Export-UiList })
    $btnImport.Add_Click({ Import-UiList })
    $btnRun.Add_Click({ Start-UiCompact })
    $radioDeep.Add_CheckedChanged({ Update-ModeUi })
    $radioQuick.Add_CheckedChanged({ Update-ModeUi })
    $chkOpenReport.Add_CheckedChanged({ Save-UiSettings })
    $list.Add_KeyDown({
        if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Delete) {
            Remove-SelectedUiModels
        }
    })
    $form.Add_FormClosing({
        Save-UiSettings
        if ($script:Ui.Timer) { $script:Ui.Timer.Stop() }
    })

    $savedUi = Read-UiSettings
    $chkOpenReport.Checked = [bool]$savedUi.OpenReport

    Update-ModeUi
    return $form
}

try {
$form = New-MainForm
[void](Add-UiModels @(Read-ExistingList))
if ($script:SelectedModels.Count -gt 0) {
    Set-UiStatus "Подставлен прошлый список — можно править и сжать."
}
[void]$form.ShowDialog()
}
catch {
    [System.Windows.Forms.MessageBox]::Show(
        $_.Exception.Message,
        "Сжатие моделей Revit",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
    throw
}
