# Name: delete_old_models.ps1
# Version: 3.1.0
# What it does: After successful Pass 2: delete old UNC models and list RSN paths for manual removal.
# Inputs: job_paths.csv, optional -ForceDeleteOldWithoutNewCheck.
# Outputs: Deleted UNC files; to_delete_rsn.txt.
# How to run: delete_old_models.cmd after green Pass 2.
# Notes: Confirms destructive deletes.
# Clean-up of OLD models after green Pass 1/2. Nothing is deleted: files are
# MOVED to <folder>\_to_delete\<stamp>\ (manifest.csv inside). RSN -> to_delete_rsn.txt.
#
# DRY-RUN by default (prints the plan, changes nothing).
#   -Execute          apply: shows the plan, asks to type YES
#   -Execute -Yes     apply without the question (caller already confirmed)
#   -DryRun           force dry-run even with -Execute
#   -PlanFile <path>  also write the printed plan there (UTF-8, for the window)
# An old model is moved only when the Pass 1/2 reports say OK for it and the
# new file exists and is not empty (see safety_lib.ps1).

param(
    [switch]$Execute,
    [switch]$Yes,
    [switch]$DryRun,
    [string]$ToolDir,
    [string]$ReportsDir,
    [string]$PlanFile
)

# script in служебное\; job_paths.csv / job.json live in the tool root one level up
if (-not $ToolDir) { $ToolDir = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path) }
# reports of this tool: <tool root>\отчёты (the window and the RBP scripts use the same folder)
if (-not $ReportsDir) { $ReportsDir = Join-Path $ToolDir "отчёты" }
. (Join-Path $PSScriptRoot "safety_lib.ps1")

$Apply = $Execute -and (-not $DryRun)
$PathsCsv = Join-Path $ToolDir "job_paths.csv"
$JobJson = Join-Path $ToolDir "job.json"
$RsnOut = Join-Path $ToolDir "to_delete_rsn.txt"
$Stamp = Get-RunStamp
if (-not (Test-Path -LiteralPath $ReportsDir)) {
    New-Item -ItemType Directory -Path $ReportsDir -Force | Out-Null
}
$mode = $(if ($Apply) { "apply" } else { "dryrun" })
$Log = Join-Path $ReportsDir ("delete_old_models_{0}_{1}.log" -f $mode, $Stamp)

# -PlanFile: same text as UTF-8 for the window (hidden console is OEM-encoded)
if ($PlanFile -and (Test-Path -LiteralPath $PlanFile)) { Remove-Item -LiteralPath $PlanFile -Force }
function Log([string]$msg) {
    Write-Host $msg
    Add-Content -LiteralPath $Log -Encoding UTF8 -Value $msg
    if ($PlanFile) { Add-Content -LiteralPath $PlanFile -Encoding UTF8 -Value $msg }
}

try {
    $rows = @(Read-JobPaths $PathsCsv)
    $job = Read-JobJson $JobJson
}
catch {
    Write-Host ("ERROR: {0}. Save the job in the presets window first." -f $_.Exception.Message)
    exit 1
}

Log ("=== old models clean-up ({0}) ===" -f $(if ($Apply) { "APPLY" } else { "DRY-RUN" }))

if (Get-JobFlag $job "copy") {
    Log "Job is in COPY mode: old models stay. Nothing to do."
    exit 0
}
if ($rows.Count -eq 0) {
    Log "job_paths.csv has no rows. Nothing to do."
    exit 0
}

$jobTime = (Get-Item -LiteralPath $PathsCsv).LastWriteTime
$ctx = New-PassContext -Job $job -Rows $rows -ReportsDir $ReportsDir -JobTime $jobTime

# ---- plan -------------------------------------------------------------
$plan = New-Object System.Collections.Generic.List[object]
foreach ($r in $rows) {
    $v = Test-ModelVerified -Row $r -Ctx $ctx
    if (-not $v.Ok) {
        $plan.Add([pscustomobject]@{ Action = "SKIP"; Path = $r.OldPath; Note = ($v.Reasons -join "; ") }) | Out-Null
        continue
    }
    if ((ConvertTo-PathKey $r.OldPath) -eq (ConvertTo-PathKey $r.NewPath)) {
        $plan.Add([pscustomobject]@{ Action = "SKIP"; Path = $r.OldPath; Note = "old path equals new path (same file)" }) | Out-Null
        continue
    }
    $note = "pass OK"
    if ($v.NewFile.State -eq "ok") { $note = "pass OK, new {0:N0} bytes" -f $v.NewFile.Size }
    elseif ($v.NewFile.State -eq "unverifiable") { $note = "pass OK, new file on RSN (size not verifiable)" }

    if (Test-RsnPath $r.OldPath) {
        $plan.Add([pscustomobject]@{ Action = "RSN"; Path = $r.OldPath; Note = $note }) | Out-Null
    }
    elseif (Test-Path -LiteralPath $r.OldPath -PathType Leaf) {
        $plan.Add([pscustomobject]@{ Action = "MOVE"; Path = $r.OldPath; Note = $note }) | Out-Null
    }
    else {
        $plan.Add([pscustomobject]@{ Action = "SKIP"; Path = $r.OldPath; Note = "old file already missing" }) | Out-Null
    }
    if ($r.StagingPath -and (Test-Path -LiteralPath $r.StagingPath -PathType Leaf)) {
        $plan.Add([pscustomobject]@{ Action = "MOVE"; Path = $r.StagingPath; Note = "buffer copy" }) | Out-Null
    }
}

$toMove = @($plan | Where-Object { $_.Action -eq "MOVE" })
$toRsn = @($plan | Where-Object { $_.Action -eq "RSN" })
$skipped = @($plan | Where-Object { $_.Action -eq "SKIP" })

foreach ($p in $plan) { Log ("{0,-5} {1}  [{2}]" -f $p.Action, $p.Path, $p.Note) }
Log ""
Log ("To move to _to_delete: {0}   RSN to list: {1}   Skipped: {2}" -f $toMove.Count, $toRsn.Count, $skipped.Count)

if (-not $Apply) {
    Log "DRY-RUN: nothing was changed. To apply run with -Execute."
    Log "Log: $Log"
    exit 0
}
if ($toMove.Count -eq 0 -and $toRsn.Count -eq 0) {
    Log "Nothing to do."
    exit 0
}
if (-not $Yes) {
    if (-not (Confirm-Yes ("Type YES to move {0} file(s) to _to_delete" -f $toMove.Count))) {
        Log "Cancelled."
        exit 0
    }
}

# ---- apply ------------------------------------------------------------
$moved = 0
$errors = 0
foreach ($p in $toMove) {
    try {
        $dest = Move-ToQuarantine -Path $p.Path -Stamp $Stamp -Reason ("clean-up: " + $p.Note)
        $moved++
        Log "MOVED $($p.Path) -> $dest"
    }
    catch {
        $errors++
        Log "ERR   $($p.Path) | $($_.Exception.Message)"
    }
}
$toRsn | ForEach-Object { $_.Path } | Set-Content -LiteralPath $RsnOut -Encoding UTF8
Log ""
Log ("Moved: {0}  errors: {1}  skipped: {2}  RSN listed: {3}" -f $moved, $errors, $skipped.Count, $toRsn.Count)
Log "Restore: move the file back from _to_delete\$Stamp (see manifest.csv there)."
Log "RSN list: $RsnOut"
Log "Log: $Log"
if ($toRsn.Count -gt 0) { Write-Host "Delete these in Revit Server Admin: $RsnOut" }
if ($errors -gt 0) { exit 2 }
exit 0
