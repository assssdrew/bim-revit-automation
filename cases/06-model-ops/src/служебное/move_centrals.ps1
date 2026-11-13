# Name: move_centrals.ps1
# Version: 3.1.0
# What it does: Move UNC central files per job_paths.csv (same storage, no year upgrade).
# Inputs: job_paths.csv.
# Outputs: Moved files on disk.
# How to run: move_centrals.cmd; confirm unless -Yes.
# Notes: Skips RSN:// paths.
# Windows Move of UNC centrals using job_paths.csv (same storage, no version upgrade).
# RSN paths are skipped.
#
# DRY-RUN by default (prints the plan, changes nothing).
#   -Execute          apply: shows the plan, asks to type YES
#   -Execute -Yes     apply without the question (caller already confirmed)
#   -DryRun           force dry-run even with -Execute
#   -PlanFile <path>  also write the printed plan there (UTF-8, for the window)
#   -Overwrite        destination that already exists is first moved to
#                     <folder>\_to_delete\<stamp>\ and then replaced. Without it
#                     such rows are skipped (SKIP_EXISTS). Never uses Move-Item -Force.

param(
    [switch]$Execute,
    [switch]$Yes,
    [switch]$DryRun,
    [switch]$Overwrite,
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
if ($PlanFile -and (Test-Path -LiteralPath $PlanFile)) { Remove-Item -LiteralPath $PlanFile -Force }
function Out-Plan([string]$msg) {
    Write-Host $msg
    if ($PlanFile) { Add-Content -LiteralPath $PlanFile -Encoding UTF8 -Value $msg }
}
$PathsCsv = Join-Path $ToolDir "job_paths.csv"
$Stamp = Get-RunStamp

try {
    $all = @(Read-JobPaths $PathsCsv)
}
catch {
    Write-Host ("ERROR: {0}. Save the job first." -f $_.Exception.Message)
    exit 1
}

Out-Plan ("=== move centrals ({0}) ===" -f $(if ($Apply) { "APPLY" } else { "DRY-RUN" }))

# ---- plan -------------------------------------------------------------
$plan = New-Object System.Collections.Generic.List[object]
$seenDest = @{}
foreach ($r in $all) {
    $old = $r.OldPath
    $new = $r.NewPath
    $item = [pscustomobject]@{ Old = $old; New = $new; Action = "MOVE"; Note = "" }
    if (-not $old -or -not $new) { continue }
    if ((Test-RsnPath $old) -or (Test-RsnPath $new)) { continue }
    if ((ConvertTo-PathKey $old) -eq (ConvertTo-PathKey $new)) { continue }

    $dk = ConvertTo-PathKey $new
    if (-not (Test-Path -LiteralPath $old -PathType Leaf)) {
        $item.Action = "SKIP_NO_SOURCE"
    }
    elseif ($seenDest.ContainsKey($dk)) {
        $item.Action = "SKIP_DUPLICATE_DEST"
        $item.Note = "another row in this job moves to the same path"
    }
    elseif (Test-Path -LiteralPath $new -PathType Leaf) {
        if ($Overwrite) {
            $item.Action = "MOVE_OVERWRITE"
            $item.Note = "existing file goes to _to_delete first"
        }
        else {
            $item.Action = "SKIP_EXISTS"
            $item.Note = "destination exists (use -Overwrite to replace)"
        }
    }
    if ($item.Action -like "MOVE*") { $seenDest[$dk] = $true }
    $plan.Add($item) | Out-Null
}

if ($plan.Count -eq 0) {
    Out-Plan "Nothing to move (no UNC path changes in job_paths.csv)."
    exit 0
}

foreach ($p in $plan) {
    Out-Plan ("{0,-20} {1}" -f $p.Action, $p.Old)
    Out-Plan ("{0,-20} -> {1}{2}" -f "", $p.New, $(if ($p.Note) { "  [" + $p.Note + "]" } else { "" }))
}
$todo = @($plan | Where-Object { $_.Action -like "MOVE*" })
Out-Plan ("To move: {0}   Skipped: {1}" -f $todo.Count, ($plan.Count - $todo.Count))

if (-not $Apply) {
    Out-Plan "DRY-RUN: nothing was changed. To apply run with -Execute."
    exit 0
}
if ($todo.Count -eq 0) {
    Write-Host "Nothing to do."
    exit 0
}
if (-not $Yes) {
    if (-not (Confirm-Yes ("Type YES to move {0} file(s)" -f $todo.Count))) {
        Write-Host "Cancelled."
        exit 0
    }
}

# ---- apply ------------------------------------------------------------
if (-not (Test-Path -LiteralPath $ReportsDir)) {
    New-Item -ItemType Directory -Path $ReportsDir -Force | Out-Null
}
$csv = Join-Path $ReportsDir ("move_centrals_{0}.csv" -f $Stamp)
Set-Content -LiteralPath $csv -Encoding UTF8 -Value "old_path;new_path;status;message"
function Write-MoveRow($p, [string]$status, [string]$msg) {
    $line = "{0};{1};{2};{3}" -f (Format-CsvField $p.Old), (Format-CsvField $p.New), $status, (Format-CsvField $msg)
    Add-Content -LiteralPath $csv -Encoding UTF8 -Value $line
}

$ok = 0
$err = 0
foreach ($p in $plan) {
    if ($p.Action -notlike "MOVE*") {
        Write-MoveRow $p $p.Action $p.Note
        continue
    }
    try {
        $srcSize = (Get-Item -LiteralPath $p.Old).Length
        $destDir = Split-Path -Parent $p.New
        if ($destDir -and -not (Test-Path -LiteralPath $destDir)) {
            New-Item -ItemType Directory -Path $destDir -Force | Out-Null
        }
        if ($p.Action -eq "MOVE_OVERWRITE") {
            $q = Move-ToQuarantine -Path $p.New -Stamp $Stamp -Reason "replaced by move_centrals"
            Write-Host ("QUARANTINED existing {0} -> {1}" -f $p.New, $q)
        }
        Move-Item -LiteralPath $p.Old -Destination $p.New
        if ((Test-Path -LiteralPath $p.Old) -or -not (Test-Path -LiteralPath $p.New -PathType Leaf) `
                -or ((Get-Item -LiteralPath $p.New).Length -ne $srcSize)) {
            throw "move not verified (size/presence mismatch)"
        }
        $ok++
        Write-MoveRow $p "OK" ""
        Write-Host ("OK  {0}" -f (Split-Path -Leaf $p.Old))
    }
    catch {
        $err++
        Write-MoveRow $p "ERROR" $_.Exception.Message
        Write-Host ("ERR {0} | {1}" -f $p.Old, $_.Exception.Message)
    }
}

Write-Host ("Moved {0}, errors {1}" -f $ok, $err)
Write-Host "CSV: $csv"
if ($err -gt 0) { exit 2 }
exit 0
