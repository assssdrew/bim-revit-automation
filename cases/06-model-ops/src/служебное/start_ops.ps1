# Name: start_ops.ps1
# Version: 3.1.0
# What it does: Launch RBP for a selected model-ops pass script.
# Inputs: ToolDir, pass name, lists from job.
# Outputs: Running RBP with the right script/list.
# How to run: Invoked from presets.ps1.
# Notes: UTF-8 BOM.
# Launch Revit Batch Processor for batch_model_ops passes.
# UTF-8 WITH BOM required (PowerShell 5.1).
# Dotted into presets.ps1. ASCII-friendly comments.

$ErrorActionPreference = "Stop"
if (-not $ToolDir) {
    # this file is in служебное\; the data root is one level up
    $ToolDir = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
}

function Find-BatchRvt {
    $names = @("BatchRvt.exe", "BatchRvtGUI.exe")
    $dirs = New-Object System.Collections.Generic.List[string]

    Get-Process -ErrorAction SilentlyContinue |
        Where-Object { $_.ProcessName -match '^(?i)BatchRvt(GUI)?$|RevitBatchProcessor$' } |
        ForEach-Object {
            try {
                $exeDir = Split-Path -Parent $_.MainModule.FileName
                if ($exeDir) { $dirs.Add($exeDir) }
            }
            catch { }
        }

    foreach ($root in @($env:LOCALAPPDATA, $env:ProgramFiles, ${env:ProgramFiles(x86)}, $env:ProgramW6432)) {
        if ([string]::IsNullOrWhiteSpace($root)) { continue }
        foreach ($rel in @(
                "RevitBatchProcessor",
                "BatchRvt",
                "Revit Batch Processor",
                "Programs\Revit Batch Processor"
            )) {
            $dirs.Add((Join-Path $root $rel))
        }
    }

    if ($env:LOCALAPPDATA -and (Test-Path -LiteralPath $env:LOCALAPPDATA)) {
        Get-ChildItem -LiteralPath $env:LOCALAPPDATA -Filter "BatchRvt.exe" -Recurse -ErrorAction SilentlyContinue -Depth 5 |
            Select-Object -First 8 |
            ForEach-Object { $dirs.Add($_.DirectoryName) }
    }

    $seen = @{}
    foreach ($dir in $dirs) {
        if ([string]::IsNullOrWhiteSpace($dir)) { continue }
        try { $full = [System.IO.Path]::GetFullPath($dir) } catch { continue }
        if ($seen.ContainsKey($full)) { continue }
        $seen[$full] = $true
        foreach ($name in $names) {
            $exe = Join-Path $full $name
            if (Test-Path -LiteralPath $exe) { return $exe }
        }
    }
    return $null
}

function Reset-OpsReportSession {
    foreach ($reports in @(Get-OpsReportsReadDirs $ToolDir)) {
    if (-not (Test-Path -LiteralPath $reports)) { continue }
    foreach ($n in @(
            "_active_update_rvt_links.txt",
            "_active_update_rvt_links_details.txt",
            "_active_saveas_new_name.txt",
            "_active_saveas_central.txt",
            "_active_detach_to_buffer.txt",
            "_active_workset_read.txt",
            "_active_workset_apply.txt"
        )) {
        $p = Join-Path $reports $n
        if (Test-Path -LiteralPath $p) {
            try { Remove-Item -LiteralPath $p -Force -ErrorAction Stop } catch { }
        }
    }
    }
}

function Find-LatestOpsReport {
    foreach ($reports in @(Get-OpsReportsReadDirs $ToolDir)) {
        if (-not (Test-Path -LiteralPath $reports)) { continue }
        $files = @(Get-ChildItem -LiteralPath $reports -File -ErrorAction SilentlyContinue |
                Where-Object { $_.Extension -match '(?i)\.(xlsx|csv)$' } |
                Sort-Object LastWriteTime -Descending)
        if ($files.Count -gt 0) { return $files[0].FullName }
    }
    return $null
}

function Start-OpsHiddenProcess {
    # CreateNoWindow: no console and no Windows Terminal tab (WindowStyle Hidden is not enough there).
    # Logs are written by BatchRvt / scripts to their own files; stdout is not needed.
    param(
        [Parameter(Mandatory = $true)][string]$FilePath,
        [string]$Arguments = "",
        [string]$WorkingDirectory = ""
    )
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $FilePath
    $psi.Arguments = $Arguments
    if ($WorkingDirectory) { $psi.WorkingDirectory = $WorkingDirectory }
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    $psi.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Hidden
    return [System.Diagnostics.Process]::Start($psi)
}

function Start-OpsRbpPass {
    param(
        [Parameter(Mandatory = $true)]$Pass,
        [Parameter(Mandatory = $true)][string]$Exe,
        [Parameter(Mandatory = $true)][string]$ExeDir
    )
    if (-not (Test-Path -LiteralPath $Pass.Task)) {
        return @{ Ok = $false; Message = ("Нет скрипта: {0}" -f $Pass.Task); Process = $null }
    }
    if (-not (Test-Path -LiteralPath $Pass.List)) {
        return @{ Ok = $false; Message = ("Нет списка: {0}" -f $Pass.List); Process = $null }
    }

    $rbpArgs = @(
        "--task_script", [string]$Pass.Task,
        "--file_list", [string]$Pass.List
    )
    # Relink, full workset read and apply set Worksets=open_all.
    # Fast workset read sets close_all. Anything else stays close_all.
    $ws = "close_all"
    if ($Pass.Name -eq "relink") { $ws = "open_all" }
    if ($Pass.PSObject.Properties["Worksets"] -and $Pass.Worksets) { $ws = [string]$Pass.Worksets }
    if ($Pass.Detach) {
        $rbpArgs += @("--detach", "--worksets", $ws)
    }
    else {
        $rbpArgs += @("--create_new_local", "--worksets", $ws)
    }
    if ($Pass.Year) {
        $rbpArgs += @("--revit_version", [string]$Pass.Year)
    }
    # RBP's own log goes next to the other reports of this tool
    $logDir = Initialize-OpsReportsDir $ToolDir
    if ($logDir) { $rbpArgs += @("--log_folder", [string]$logDir) }

    $argLine = New-Object System.Collections.Generic.List[string]
    foreach ($a in $rbpArgs) {
        if ($a -match '\s') { [void]$argLine.Add(('"{0}"' -f $a)) }
        else { [void]$argLine.Add($a) }
    }
    $proc = Start-OpsHiddenProcess -FilePath $Exe -Arguments ($argLine -join " ") -WorkingDirectory $ExeDir
    return @{
        Ok      = $true
        Message = ("Запущен проход {0}." -f $Pass.Name)
        Process = $proc
    }
}

# Dry-run of move/delete: runs the script WITHOUT -Execute and returns the plan
# text (read from -PlanFile, UTF-8) so the window can show it before anything changes.
function Invoke-OpsPsDryRun {
    param(
        [Parameter(Mandatory = $true)]$Pass
    )
    $plan = Join-Path ([System.IO.Path]::GetTempPath()) ("ops_plan_{0}.txt" -f [Guid]::NewGuid().ToString("N"))
    $arg = "-NoProfile -ExecutionPolicy Bypass -File `"$($Pass.Task)`" -ToolDir `"$ToolDir`" -DryRun -PlanFile `"$plan`""
    if ($Pass.Name -eq "move" -and $Pass.Overwrite) { $arg += " -Overwrite" }
    $proc = Start-OpsHiddenProcess -FilePath "powershell.exe" -Arguments $arg -WorkingDirectory $ToolDir
    if (-not $proc.WaitForExit(120000)) {
        try { $proc.Kill() } catch { }
        return @{ Ok = $false; Code = -1; Text = "Предпросмотр не завершился за 2 минуты." }
    }
    $text = ""
    if (Test-Path -LiteralPath $plan) {
        $text = (Get-Content -LiteralPath $plan -Raw -Encoding UTF8)
        Remove-Item -LiteralPath $plan -Force -ErrorAction SilentlyContinue
    }
    $count = 0
    if ($text -match 'To move to _to_delete:\s+(\d+)\s+RSN to list:\s+(\d+)') {
        $count = [int]$Matches[1] + [int]$Matches[2]
    }
    elseif ($text -match 'To move:\s+(\d+)') {
        $count = [int]$Matches[1]
    }
    return @{ Ok = ($proc.ExitCode -eq 0); Code = $proc.ExitCode; Text = $text; Count = $count }
}

function Start-OpsPsPass {
    param(
        [Parameter(Mandatory = $true)]$Pass,
        [string[]]$ExtraArgs = @()
    )
    $arg = "-NoProfile -ExecutionPolicy Bypass -File `"$($Pass.Task)`" -ToolDir `"$ToolDir`""
    foreach ($a in $ExtraArgs) { $arg += " " + $a }
    $proc = Start-OpsHiddenProcess -FilePath "powershell.exe" -Arguments $arg -WorkingDirectory $ToolDir
    return @{
        Ok      = $true
        Message = ("Запущен {0}." -f $Pass.Name)
        Process = $proc
    }
}
