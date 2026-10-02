# Choose local folder for Excel report (shared by several users; each saves locally).
Add-Type -AssemblyName System.Windows.Forms | Out-Null

$ToolDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$CfgPath = Join-Path $ToolDir "xlsx_out_path.cfg"

$dlg = New-Object System.Windows.Forms.FolderBrowserDialog
$dlg.Description = "Выберите ЛОКАЛЬНУЮ папку для сохранения отчёта в Excel (не сетевой диск)."
$dlg.ShowNewFolderButton = $true

if (Test-Path -LiteralPath $CfgPath) {
    $prev = (Get-Content -LiteralPath $CfgPath -ErrorAction SilentlyContinue | Select-Object -First 1)
    if ($prev -and (Test-Path -LiteralPath $prev -PathType Container)) {
        $dlg.SelectedPath = $prev
    }
}

if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
    Write-Host "Cancelled: XLSX folder not selected."
    exit 1
}

$folder = $dlg.SelectedPath.TrimEnd("\")
if ([string]::IsNullOrWhiteSpace($folder)) {
    Write-Host "ERROR: empty folder."
    exit 1
}

# Soft warning for UNC (Excel often fails on long share paths)
if ($folder.StartsWith("\\")) {
    $ans = [System.Windows.Forms.MessageBox]::Show(
        "Выбрана сетевая папка. Excel часто не умеет сохранять по длинным UNC-путям." +
        [Environment]::NewLine +
        "Рекомендуется локальная папка (например Desktop\log\xlsx)." +
        [Environment]::NewLine + [Environment]::NewLine +
        "Оставить этот путь?",
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
