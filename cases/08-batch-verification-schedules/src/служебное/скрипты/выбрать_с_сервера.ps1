# UTF-8 WITH BOM required (PowerShell 5.1).
# Список моделей Revit Server. Папки читаются по HTTP, в файл пишутся пути RSN://.
param(
    [Parameter(Mandatory = $true)][string]$OutFile
)

$ErrorActionPreference = "Continue"
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

if (-not ("VerifConsole" -as [type])) {
    Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public static class VerifConsole {
    [DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow();
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
}
"@
}
$console = [VerifConsole]::GetConsoleWindow()
if ($console -ne [IntPtr]::Zero) { [void][VerifConsole]::ShowWindow($console, 0) }

$here = Split-Path -Parent $MyInvocation.MyCommand.Path          # служебное\скрипты
$toolRoot = Split-Path -Parent (Split-Path -Parent $here)        # корень: рядом .vbs, servers.cfg, отчеты
$compactCandidates = @(
    (Join-Path (Split-Path -Parent $here) "choose_models.ps1"),                    # служебное\choose_models.ps1
    (Join-Path (Split-Path -Parent $toolRoot) "batch_compact_save\choose_models.ps1"),
    (Join-Path (Split-Path -Parent (Split-Path -Parent $toolRoot)) "automation\cases\05-compact-save\src\choose_models.ps1")
)
$compact = $null
foreach ($candidate in $compactCandidates) {
    if (Test-Path -LiteralPath $candidate) { $compact = $candidate; break }
}
if (-not $compact) {
    [System.Windows.Forms.MessageBox]::Show("Не найден список Revit Server (служебное\choose_models.ps1).", "Проверочные спецификации") | Out-Null
    exit 2
}

$script:CompactUiLibrary = $true
$CompactUiLibrary = $true
. $compact

$localCfg = Join-Path $toolRoot "servers.cfg"
if (Test-Path -LiteralPath $localCfg) { $ServersCfg = $localCfg }

function Show-ServerChoice {
    $servers = @(Get-Servers)
    if ($servers.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("В servers.cfg нет серверов.", "Revit Server") | Out-Null
        return $null
    }
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Revit Server"
    $form.Size = New-Object System.Drawing.Size(460, 360)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 9.75)
    $form.TopMost = $true
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = "Сервер"
    $lbl.Location = New-Object System.Drawing.Point(12, 12)
    $lbl.AutoSize = $true
    $list = New-Object System.Windows.Forms.ListBox
    $list.Location = New-Object System.Drawing.Point(12, 36)
    $list.Size = New-Object System.Drawing.Size(420, 230)
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
    if ($form.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return $null }
    if ($list.SelectedIndex -lt 0) { return $null }
    return $servers[$list.SelectedIndex]
}

$server = Show-ServerChoice
if (-not $server) { exit 0 }
$picked = @(Show-RevitServerBrowser -ServerObj $server)
if ($picked.Count -lt 1) { exit 0 }
$utf8 = New-Object System.Text.UTF8Encoding $true
[System.IO.File]::WriteAllLines($OutFile, @($picked), $utf8)
exit 0
