# Name: delete_old_models.ps1
# Version: 1.0
# What it does: After successful Pass 2: delete old UNC models and list RSN paths for manual removal.
# Inputs: job_paths.csv, optional -ForceDeleteOldWithoutNewCheck.
# Outputs: Deleted UNC files; to_delete_rsn.txt.
# How to run: delete_old_models.cmd after green Pass 2.
# Notes: Confirms destructive deletes.
param(
    [switch]$ForceDeleteOldWithoutNewCheck
)

$ToolDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$PathsCsv = Join-Path $ToolDir "job_paths.csv"
$RsnOut = Join-Path $ToolDir "to_delete_rsn.txt"
$Reports = Join-Path $env:USERPROFILE "Documents\doc\script\rbp\reports"
if (-not (Test-Path -LiteralPath $Reports)) {
    New-Item -ItemType Directory -Path $Reports -Force | Out-Null
}
$Log = Join-Path $Reports ("delete_old_models_{0}.log" -f (Get-Date -Format "yyyy-MM-dd_HHmmss"))

if (-not (Test-Path -LiteralPath $PathsCsv)) {
    Write-Host "ERROR: job_paths.csv missing. Save the job in the presets window first."
    exit 1
}

function Log([string]$msg) {
    $msg | Tee-Object -FilePath $Log -Append
}

$rsn = New-Object System.Collections.Generic.List[string]
$deleted = 0
$skipped = 0

Log "=== delete old models (job_paths) ==="

Get-Content -LiteralPath $PathsCsv -Encoding UTF8 | Select-Object -Skip 1 | ForEach-Object {
    $line = $_.Trim()
    if (-not $line -or $line.StartsWith("#")) { return }
    $parts = $line.Split(";")
    if ($parts.Count -lt 5) { return }
    $old = $parts[0].Trim().Trim('"')
    $new = $parts[1].Trim().Trim('"')
    $stg = $parts[2].Trim().Trim('"')
    if (-not $old) { return }

    if ($old.ToUpper().StartsWith("RSN://")) {
        [void]$rsn.Add($old)
        Log "RSN (manual Admin): $old"
    }
    elseif (Test-Path -LiteralPath $old) {
        $okNew = $ForceDeleteOldWithoutNewCheck
        if (-not $okNew) {
            if ($new -and (Test-Path -LiteralPath $new)) { $okNew = $true }
        }
        if (-not $okNew) {
            Log "SKIP old (new not found): $old"
            $skipped++
            return
        }
        try {
            Remove-Item -LiteralPath $old -Force
            $deleted++
            Log "DELETED old: $old"
        }
        catch {
            Log "ERR delete old: $old | $_"
        }
    }
    else {
        Log "SKIP old missing: $old"
        $skipped++
    }

    if ($stg -and (Test-Path -LiteralPath $stg)) {
        try {
            Remove-Item -LiteralPath $stg -Force
            Log "DELETED buffer: $stg"
        }
        catch {
            Log "ERR delete buffer: $stg | $_"
        }
    }
}

$rsn | Set-Content -LiteralPath $RsnOut -Encoding UTF8
Log ""
Log ("Deleted UNC/old: {0}  skipped: {1}  RSN listed: {2}" -f $deleted, $skipped, $rsn.Count)
Log "RSN list: $RsnOut"
Log "Log: $Log"
Write-Host "Done. See $Log"
if ($rsn.Count -gt 0) {
    Write-Host "Delete these in Revit Server Admin: $RsnOut"
}
