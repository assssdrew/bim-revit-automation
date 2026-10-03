# Name: move_centrals.ps1
# Version: 1.0
# What it does: Move UNC central files per job_paths.csv (same storage, no year upgrade).
# Inputs: job_paths.csv.
# Outputs: Moved files on disk.
# How to run: move_centrals.cmd; confirm unless -Yes.
# Notes: Skips RSN:// paths.
param(
    [switch]$Yes
)

$ToolDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$PathsCsv = Join-Path $ToolDir "job_paths.csv"
$Reports = Join-Path $env:USERPROFILE "Documents\doc\script\rbp\reports"
if (-not (Test-Path -LiteralPath $Reports)) {
    New-Item -ItemType Directory -Path $Reports -Force | Out-Null
}

if (-not (Test-Path -LiteralPath $PathsCsv)) {
    Write-Host "ERROR: job_paths.csv missing. Save the job first."
    exit 1
}

$rows = New-Object System.Collections.Generic.List[object]
Get-Content -LiteralPath $PathsCsv -Encoding UTF8 | Select-Object -Skip 1 | ForEach-Object {
    $line = $_.Trim()
    if (-not $line -or $line.StartsWith("#")) { return }
    $p = $line.Split(";")
    if ($p.Count -lt 2) { return }
    $old = $p[0].Trim().Trim('"')
    $new = $p[1].Trim().Trim('"')
    if (-not $old -or -not $new) { return }
    if ($old.ToUpper().StartsWith("RSN://") -or $new.ToUpper().StartsWith("RSN://")) { return }
    if ($old.Replace("\", "/").ToLower() -eq $new.Replace("\", "/").ToLower()) { return }
    $rows.Add([pscustomobject]@{ Old = $old; New = $new }) | Out-Null
}

if ($rows.Count -eq 0) {
    Write-Host "Nothing to move (no UNC path changes in job_paths.csv)."
    exit 0
}

Write-Host ("Will MOVE {0} file(s):" -f $rows.Count)
$rows | Select-Object -First 15 | ForEach-Object {
    Write-Host ("  {0}" -f $_.Old)
    Write-Host ("  -> {0}" -f $_.New)
}
if ($rows.Count -gt 15) { Write-Host ("  ... +{0} more" -f ($rows.Count - 15)) }
if (-not $Yes) {
    $ans = Read-Host "Type YES to move"
    if ($ans -ne "YES") { Write-Host "Cancelled."; exit 0 }
}

$ok = 0
$err = 0
$csv = Join-Path $Reports ("move_centrals_{0}.csv" -f (Get-Date -Format "yyyy-MM-dd_HHmmss"))
"old_path;new_path;status;message" | Set-Content -LiteralPath $csv -Encoding UTF8

foreach ($r in $rows) {
    try {
        $destDir = Split-Path -Parent $r.New
        if ($destDir -and -not (Test-Path -LiteralPath $destDir)) {
            New-Item -ItemType Directory -Path $destDir -Force | Out-Null
        }
        Move-Item -LiteralPath $r.Old -Destination $r.New -Force
        $ok++
        Add-Content -LiteralPath $csv -Encoding UTF8 -Value ("{0};{1};OK;" -f $r.Old, $r.New)
        Write-Host ("OK  {0}" -f (Split-Path -Leaf $r.Old))
    }
    catch {
        $err++
        Add-Content -LiteralPath $csv -Encoding UTF8 -Value ("{0};{1};ERROR;{2}" -f $r.Old, $r.New, $_.Exception.Message)
        Write-Host ("ERR {0} | {1}" -f $r.Old, $_)
    }
}

Write-Host ("Moved {0}, errors {1}" -f $ok, $err)
Write-Host "CSV: $csv"
if ($err -gt 0) { exit 2 }
exit 0
