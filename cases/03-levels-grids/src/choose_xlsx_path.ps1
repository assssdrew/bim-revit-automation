# Name: choose_xlsx_path.ps1
# Version: 1.0
# What it does: Pick local Excel report folder (xlsx_out_path.cfg).
# Inputs: Folder dialog.
# Outputs: xlsx_out_path.cfg.
# How to run: reset_report_session.cmd or choose_xlsx_path.cmd.
# Notes: Local folder recommended.
# Choose LOCAL folder for colored XLSX using OpenFileDialog with address bar.
# This avoids FolderBrowserDialog and allows pasting paths.

Add-Type -AssemblyName System.Windows.Forms | Out-Null

$ToolDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$CfgPath = Join-Path $ToolDir "xlsx_out_path.cfg"

function Get-NativePath([string]$Path) {
    if ([string]::IsNullOrWhiteSpace($Path)) { return $null }
    $p = $Path.Trim().Trim('"')
    $marker = "FileSystem::"
    $idx = $p.IndexOf($marker, [StringComparison]::OrdinalIgnoreCase)
    if ($idx -ge 0) { $p = $p.Substring($idx + $marker.Length) }
    if ($p.StartsWith("\\")) { return $p }
    try { return (New-Object System.IO.FileInfo($p)).FullName }
    catch { return $p }
}

function Resolve-FolderFromDialogResult([string]$Picked) {
    if ([string]::IsNullOrWhiteSpace($Picked)) { return $null }
    $p = Get-NativePath $Picked
    if (Test-Path -LiteralPath $p -PathType Leaf) {
        return (Split-Path -Parent $p)
    }
    if (Test-Path -LiteralPath $p -PathType Container) {
        return $p
    }
    # User pasted a folder path into File name + dummy "Select this folder"
    $dir = Split-Path -Parent $p
    if ($dir -and (Test-Path -LiteralPath $dir -PathType Container)) {
        $leaf = Split-Path -Leaf $p
        if ($leaf -eq "Select this folder" -or $leaf -eq "Select_this_folder") {
            return $dir
        }
        # pasted full folder path that doesn't exist yet - still accept parent? no
    }
    # Exact path typed as folder
    if ($p -and -not (Test-Path -LiteralPath $p)) {
        $asDir = $p.TrimEnd('\', '/')
        if (Test-Path -LiteralPath $asDir -PathType Container) { return $asDir }
    }
    return $null
}

$dlg = New-Object System.Windows.Forms.OpenFileDialog
$dlg.Title = "XLSX folder: paste path in address bar or File name"
$dlg.Filter = "Folders|*.nevermatch|All files (*.*)|*.*"
$dlg.FilterIndex = 1
$dlg.CheckFileExists = $false
$dlg.CheckPathExists = $true
$dlg.ValidateNames = $false
$dlg.Multiselect = $false
$dlg.DereferenceLinks = $true
$dlg.FileName = "Select this folder"

if (Test-Path -LiteralPath $CfgPath) {
    $prev = (
        Get-Content -LiteralPath $CfgPath -ErrorAction SilentlyContinue |
            Where-Object { $_ -and -not $_.StartsWith("#") } |
            Select-Object -First 1
    )
    if ($prev) {
        $prev = Get-NativePath $prev.Trim()
        if ($prev -and (Test-Path -LiteralPath $prev -PathType Container)) {
            $dlg.InitialDirectory = $prev
        }
    }
}
if ([string]::IsNullOrWhiteSpace($dlg.InitialDirectory)) {
    $auto = Join-Path $env:USERPROFILE "Documents\doc\script\rbp\reports"
    if (-not (Test-Path -LiteralPath $auto)) {
        New-Item -ItemType Directory -Path $auto -Force | Out-Null
    }
    $dlg.InitialDirectory = $auto
}

if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
    Write-Host "Cancelled: XLSX folder not selected."
    exit 1
}

$folder = Resolve-FolderFromDialogResult $dlg.FileName
if ([string]::IsNullOrWhiteSpace($folder)) {
    Write-Host ("ERROR: could not resolve folder from: {0}" -f $dlg.FileName)
    Write-Host "Tip: open target folder, keep File name as-is, press Open."
    exit 1
}

$folder = $folder.TrimEnd("\")
if ($folder.StartsWith("\\")) {
    $ans = [System.Windows.Forms.MessageBox]::Show(
        "Selected a network folder. Prefer a LOCAL folder (e.g. Desktop\lg_xlsx)." +
        [Environment]::NewLine + [Environment]::NewLine +
        "Keep this path?",
        "XLSX path",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )
    if ($ans -ne [System.Windows.Forms.DialogResult]::Yes) {
        Write-Host "Cancelled."
        exit 1
    }
}

Set-Content -LiteralPath $CfgPath -Value $folder -Encoding UTF8
Write-Host ""
Write-Host ("XLSX folder saved: {0}" -f $folder)
Write-Host ("Config: {0}" -f $CfgPath)
Write-Host ""
exit 0
