# Name: start_ops.ps1
# Version: 1.0
# What it does: Launch RBP for a selected model-ops pass script.
# Inputs: ToolDir, pass name, lists from job.
# Outputs: Running RBP with the right script/list.
# How to run: Invoked from presets.ps1.
# Notes: UTF-8 BOM.
$ErrorActionPreference = "Stop"
if (-not $ToolDir) {
    $ToolDir = Split-Path -Parent $MyInvocation.MyCommand.Path
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
    $reports = Join-Path (Get-RbpRoot) "reports"
    if (-not (Test-Path -LiteralPath $reports)) { return }
    foreach ($n in @(
            "_active_update_rvt_links.txt",
            "_active_update_rvt_links_details.txt",
            "_active_saveas_new_name.txt",
            "_active_saveas_central.txt",
            "_active_detach_to_buffer.txt"
        )) {
        $p = Join-Path $reports $n
        if (Test-Path -LiteralPath $p) {
            try { Remove-Item -LiteralPath $p -Force -ErrorAction Stop } catch { }
        }
    }
}

function Find-LatestOpsReport {
    $reports = Join-Path (Get-RbpRoot) "reports"
    if (-not (Test-Path -LiteralPath $reports)) { return $null }
    $files = @(Get-ChildItem -LiteralPath $reports -File -ErrorAction SilentlyContinue |
            Where-Object { $_.Extension -match '(?i)\.(xlsx|csv)$' } |
            Sort-Object LastWriteTime -Descending)
    if ($files.Count -eq 0) { return $null }
    return $files[0].FullName
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
    if ($Pass.Detach) {
        $rbpArgs += @("--detach", "--worksets", "close_all")
    }
    else {
        $rbpArgs += @("--create_new_local", "--worksets", "close_all")
    }
    if ($Pass.Year) {
        $rbpArgs += @("--revit_version", [string]$Pass.Year)
    }

    $argLine = New-Object System.Collections.Generic.List[string]
    foreach ($a in $rbpArgs) {
        if ($a -match '\s') { [void]$argLine.Add(('"{0}"' -f $a)) }
        else { [void]$argLine.Add($a) }
    }
    $proc = Start-Process -FilePath $Exe -ArgumentList ($argLine -join " ") -WorkingDirectory $ExeDir -WindowStyle Hidden -PassThru
    return @{
        Ok      = $true
        Message = ("Запущен проход {0}." -f $Pass.Name)
        Process = $proc
    }
}

function Start-OpsPsPass {
    param(
        [Parameter(Mandatory = $true)]$Pass
    )
    $arg = "-NoProfile -ExecutionPolicy Bypass -File `"$($Pass.Task)`""
    if ($Pass.Name -eq "move") { $arg += " -Yes" }
    $proc = Start-Process -FilePath "powershell.exe" -ArgumentList $arg -WorkingDirectory $ToolDir -WindowStyle Hidden -PassThru
    return @{
        Ok      = $true
        Message = ("Запущен {0}." -f $Pass.Name)
        Process = $proc
    }
}
