# Name: safety.tests.ps1
# Version: 3.1.0
# What it does: See the case README.
# Notes: Paths and project names in this portfolio copy are placeholders.

# Tests for safety_lib.ps1, delete_old_models.ps1, move_centrals.ps1.
# No dependencies (works on Windows PowerShell 5.1 and PowerShell 7):
#   powershell -NoProfile -ExecutionPolicy Bypass -File tests\safety.tests.ps1
# Everything happens in a temp folder; exit code 1 if any check fails.

$ErrorActionPreference = "Stop"
$ToolSrc = Split-Path -Parent $PSScriptRoot
$Ps = (Get-Process -Id $PID).Path
$Work = Join-Path ([System.IO.Path]::GetTempPath()) ("ops_safety_tests_" + [Guid]::NewGuid().ToString("N").Substring(0, 8))
New-Item -ItemType Directory -Path $Work -Force | Out-Null

$script:Fail = 0
$script:Pass = 0
function Check([bool]$Cond, [string]$Msg) {
    if ($Cond) { $script:Pass++ }
    else { $script:Fail++; Write-Host "  FAIL: $Msg" -ForegroundColor Red }
}

function Csv([string]$v) {
    if ($v.Contains(";") -or $v.Contains('"')) { return '"' + $v.Replace('"', '""') + '"' }
    return $v
}

function New-Case([string]$Name) {
    $root = Join-Path $Work $Name
    $c = [pscustomobject]@{
        Root = $root
        Tool = (Join-Path $root "tool")
        Rep  = (Join-Path $root "reports")
        Data = (Join-Path $root "data")
    }
    foreach ($d in @($c.Tool, $c.Rep, $c.Data)) { New-Item -ItemType Directory -Path $d -Force | Out-Null }
    Write-Host ("[{0}]" -f $Name)
    return $c
}

function New-File([string]$Path, [string]$Content = "model") {
    $d = Split-Path -Parent $Path
    if (-not (Test-Path -LiteralPath $d)) { New-Item -ItemType Directory -Path $d -Force | Out-Null }
    Set-Content -LiteralPath $Path -Value $Content -Encoding UTF8
}

# rows: @{ old; new; stg; buf }
function Write-Job($c, [object[]]$Rows, [hashtable]$Job) {
    $lines = @("old_path;new_path;staging_path;old_name;new_name;needs_buffer")
    foreach ($r in $Rows) {
        $stg = $(if ($r.stg) { $r.stg } else { "" })
        $on = [System.IO.Path]::GetFileNameWithoutExtension(($r.old -replace "\\", "/").Split("/")[-1])
        $nn = [System.IO.Path]::GetFileNameWithoutExtension(($r.new -replace "\\", "/").Split("/")[-1])
        $lines += ("{0};{1};{2};{3};{4};{5}" -f (Csv $r.old), (Csv $r.new), (Csv $stg), (Csv $on), (Csv $nn), $(if ($r.buf) { "1" } else { "0" }))
    }
    $utf8 = New-Object System.Text.UTF8Encoding $true
    $pj = Join-Path $c.Tool "job_paths.csv"
    [System.IO.File]::WriteAllLines($pj, $lines, $utf8)
    (Get-Item -LiteralPath $pj).LastWriteTime = (Get-Date).AddMinutes(-5)
    ($Job | ConvertTo-Json -Depth 4) | Set-Content -LiteralPath (Join-Path $c.Tool "job.json") -Encoding UTF8
}

# kind: saveas | detach | relink ; lines: array of arrays (fields)
function Write-Report($c, [string]$Kind, [object[]]$Rows, [switch]$WithHostPath, [switch]$Stale) {
    $map = @{
        saveas = @{ File = "saveas_central_2026-01-01_1200.csv"; Marker = "_active_saveas_central.txt"; Head = "old_path;new_path;status;message;script_version" }
        detach = @{ File = "detach_to_buffer_2026-01-01_1200.csv"; Marker = "_active_detach_to_buffer.txt"; Head = "old_path;staging_path;status;message;script_version" }
        relink = @{ File = "update_rvt_links_2026-01-01_1200.csv"; Marker = "_active_update_rvt_links.txt"; Head = "host;status;updated;skipped;errors;issues;script_version" }
    }
    $m = $map[$Kind]
    $head = $m.Head
    if ($WithHostPath) { $head += ";host_path" }
    $lines = @($head)
    foreach ($row in $Rows) { $lines += (($row | ForEach-Object { Csv ([string]$_) }) -join ";") }
    $p = Join-Path $c.Rep $m.File
    $utf8 = New-Object System.Text.UTF8Encoding $true
    [System.IO.File]::WriteAllLines($p, $lines, $utf8)
    Set-Content -LiteralPath (Join-Path $c.Rep $m.Marker) -Value $p -Encoding UTF8
    if ($Stale) { (Get-Item -LiteralPath $p).LastWriteTime = (Get-Date).AddDays(-2) }
    else { (Get-Item -LiteralPath $p).LastWriteTime = (Get-Date) }
}

function Invoke-Script($c, [string]$Script, [string[]]$ExtraArgs = @(), [string]$StdIn = $null) {
    $argv = @("-NoProfile", "-ExecutionPolicy", "Bypass", "-File", (Join-Path $ToolSrc $Script),
        "-ToolDir", $c.Tool, "-ReportsDir", $c.Rep) + $ExtraArgs
    if ($null -ne $StdIn) { $out = $StdIn | & $Ps @argv 2>&1 }
    else { $out = & $Ps @argv 2>&1 }
    return [pscustomobject]@{ Code = $LASTEXITCODE; Out = (($out | ForEach-Object { "$_" }) -join "`n") }
}

function Quarantined($c, [string]$Dir, [string]$Leaf) {
    $q = Join-Path $Dir "_to_delete"
    if (-not (Test-Path -LiteralPath $q)) { return $false }
    return (@(Get-ChildItem -LiteralPath $q -Recurse -File -Filter $Leaf).Count -gt 0)
}

$SaveasJob = @{ saveas_needed = $true; relink = $false; copy = $false; delete_old = $true }

# ------------------------------------------------------------------ delete
$c = New-Case "del_dryrun_default"
New-File (Join-Path $c.Data "old\A.rvt"); New-File (Join-Path $c.Data "new\A2.rvt")
$o = Join-Path $c.Data "old\A.rvt"; $n = Join-Path $c.Data "new\A2.rvt"
Write-Job $c @(@{ old = $o; new = $n }) $SaveasJob
Write-Report $c saveas @(, @($o, $n, "OK", "", "3.0.0"))
$r = Invoke-Script $c "delete_old_models.ps1"
Check ($r.Code -eq 0) "dry-run exits 0 ($($r.Code))"
Check (Test-Path -LiteralPath $o) "dry-run keeps old file"
Check (-not (Test-Path -LiteralPath (Join-Path $c.Data "old\_to_delete"))) "dry-run creates no _to_delete"
Check ($r.Out -match "DRY-RUN") "dry-run says DRY-RUN"
$r = Invoke-Script $c "delete_old_models.ps1" @("-Execute", "-DryRun")
Check (Test-Path -LiteralPath $o) "-Execute -DryRun still dry-run"

$c = New-Case "del_execute_ok"
$o = Join-Path $c.Data "old\A.rvt"; $n = Join-Path $c.Data "new\A2.rvt"
New-File $o; New-File $n
Write-Job $c @(@{ old = $o; new = $n }) $SaveasJob
Write-Report $c saveas @(, @($o, $n, "OK", "", "3.0.0"))
$r = Invoke-Script $c "delete_old_models.ps1" @("-Execute", "-Yes")
Check ($r.Code -eq 0) "execute exits 0 ($($r.Code)): $($r.Out)"
Check (-not (Test-Path -LiteralPath $o)) "old file moved away"
Check (Quarantined $c (Join-Path $c.Data "old") "A.rvt") "old file is in _to_delete"
Check (Test-Path -LiteralPath $n) "new file untouched"
Check (@(Get-ChildItem -LiteralPath (Join-Path $c.Data "old\_to_delete") -Recurse -Filter manifest.csv).Count -eq 1) "manifest.csv written"

$c = New-Case "del_prompt_no"
$o = Join-Path $c.Data "old\A.rvt"; $n = Join-Path $c.Data "new\A2.rvt"
New-File $o; New-File $n
Write-Job $c @(@{ old = $o; new = $n }) $SaveasJob
Write-Report $c saveas @(, @($o, $n, "OK", "", "3.0.0"))
$r = Invoke-Script $c "delete_old_models.ps1" @("-Execute") "no"
Check (Test-Path -LiteralPath $o) "answer 'no' keeps old file"
$r = Invoke-Script $c "delete_old_models.ps1" @("-Execute") "yes"
Check (Test-Path -LiteralPath $o) "lowercase 'yes' is not YES"
$r = Invoke-Script $c "delete_old_models.ps1" @("-Execute") "YES"
Check (-not (Test-Path -LiteralPath $o)) "typed YES moves old file"

$c = New-Case "del_blocked_cases"
$cases = @(
    @{ n = "err"; rep = @("ERROR"); newSize = "x"; expectKept = $true },
    @{ n = "nonew"; rep = @("OK"); newSize = $null; expectKept = $true },
    @{ n = "empty"; rep = @("OK"); newSize = ""; expectKept = $true },
    @{ n = "good"; rep = @("OK"); newSize = "x"; expectKept = $false }
)
$rows = @(); $repRows = @(); $olds = @{}
foreach ($k in $cases) {
    $o = Join-Path $c.Data ("old\{0}.rvt" -f $k.n); $n = Join-Path $c.Data ("new\{0}_N.rvt" -f $k.n)
    New-File $o
    if ($null -ne $k.newSize) {
        if ($k.newSize -eq "") { New-Item -ItemType File -Path $n -Force | Out-Null } else { New-File $n }
    }
    $rows += @{ old = $o; new = $n }
    $repRows += , @($o, $n, $k.rep[0], "", "3.0.0")
    $olds[$k.n] = $o
}
Write-Job $c $rows $SaveasJob
Write-Report $c saveas $repRows
$r = Invoke-Script $c "delete_old_models.ps1" @("-Execute", "-Yes")
foreach ($k in $cases) {
    $kept = Test-Path -LiteralPath $olds[$k.n]
    Check ($kept -eq $k.expectKept) ("case {0}: kept={1} expected={2}" -f $k.n, $kept, $k.expectKept)
}
Check ($r.Out -match "SKIP") "skips are reported"

$c = New-Case "del_no_or_stale_report"
$o = Join-Path $c.Data "old\A.rvt"; $n = Join-Path $c.Data "new\A2.rvt"
New-File $o; New-File $n
Write-Job $c @(@{ old = $o; new = $n }) $SaveasJob
$r = Invoke-Script $c "delete_old_models.ps1" @("-Execute", "-Yes")
Check (Test-Path -LiteralPath $o) "no report -> old kept"
Write-Report $c saveas @(, @($o, $n, "OK", "", "3.0.0")) -Stale
$r = Invoke-Script $c "delete_old_models.ps1" @("-Execute", "-Yes")
Check (Test-Path -LiteralPath $o) "stale report (older than job) -> old kept"
Check ($r.Out -match "stale") "stale reason shown"

$c = New-Case "del_semicolon_path"
$o = Join-Path $c.Data "old\A;B.rvt"; $n = Join-Path $c.Data "new\A;B2.rvt"
New-File $o; New-File $n
Write-Job $c @(@{ old = $o; new = $n }) $SaveasJob
Write-Report $c saveas @(, @($o, $n, "OK", "", "3.0.0"))
$r = Invoke-Script $c "delete_old_models.ps1" @("-Execute", "-Yes")
Check (-not (Test-Path -LiteralPath $o)) "path with ';' handled"

$c = New-Case "del_same_path"
$o = Join-Path $c.Data "m\A.rvt"
New-File $o
Write-Job $c @(@{ old = $o; new = $o }) $SaveasJob
Write-Report $c saveas @(, @($o, $o, "OK", "", "3.0.0"))
$r = Invoke-Script $c "delete_old_models.ps1" @("-Execute", "-Yes")
Check (Test-Path -LiteralPath $o) "old == new path is never moved"

$c = New-Case "del_skip_exists"
$o = Join-Path $c.Data "old\A.rvt"; $n = Join-Path $c.Data "new\A2.rvt"; $o2 = Join-Path $c.Data "old\B.rvt"; $n2 = Join-Path $c.Data "new\B2.rvt"
New-File $o; New-File $n; New-File $o2; New-File $n2
Write-Job $c @(@{ old = $o; new = $n }, @{ old = $o2; new = $n2 }) $SaveasJob
Write-Report $c saveas @(@($o, $n, "OK", "", "3.1.0"), @($o, "", "SKIP_EXISTS", "", "3.1.0"), @($o2, "", "SKIP_EXISTS", "", "3.1.0"))
$r = Invoke-Script $c "delete_old_models.ps1" @("-Execute", "-Yes")
Check (-not (Test-Path -LiteralPath $o)) "OK then SKIP_EXISTS counts as OK"
Check (Test-Path -LiteralPath $o2) "SKIP_EXISTS alone does not count"

$c = New-Case "del_relink"
$relinkJob = @{ saveas_needed = $true; relink = $true; copy = $false; delete_old = $true }
$names = "ok", "nochg", "partial", "err", "none"
$rows = @(); $sv = @(); $rl = @(); $olds = @{}
foreach ($nm in $names) {
    $o = Join-Path $c.Data ("old\{0}.rvt" -f $nm); $n = Join-Path $c.Data ("new\{0}_N.rvt" -f $nm)
    New-File $o; New-File $n
    $rows += @{ old = $o; new = $n }; $sv += , @($o, $n, "OK", "", "3.0.0"); $olds[$nm] = $o
}
$rl += , @("ok_N", "OK", 1, 0, 0, "", "3.0.0")
$rl += , @("nochg_N", "NO_CHANGES", 0, 0, 0, "", "3.0.0")
$rl += , @("partial_N", "PARTIAL", 1, 0, 1, "", "3.0.0")
$rl += , @("err_N", "ERROR", 0, 0, 1, "", "3.0.0")
Write-Job $c $rows $relinkJob
Write-Report $c saveas $sv
Write-Report $c relink $rl
$r = Invoke-Script $c "delete_old_models.ps1" @("-Execute", "-Yes")
Check (-not (Test-Path -LiteralPath $olds["ok"])) "relink OK -> moved"
Check (-not (Test-Path -LiteralPath $olds["nochg"])) "relink NO_CHANGES -> moved"
Check (Test-Path -LiteralPath $olds["partial"]) "relink PARTIAL -> kept"
Check (Test-Path -LiteralPath $olds["err"]) "relink ERROR -> kept"
Check (Test-Path -LiteralPath $olds["none"]) "relink no record -> kept"

$c = New-Case "del_relink_missing_report"
$o = Join-Path $c.Data "old\A.rvt"; $n = Join-Path $c.Data "new\A_N.rvt"
New-File $o; New-File $n
Write-Job $c @(@{ old = $o; new = $n }) $relinkJob
Write-Report $c saveas @(, @($o, $n, "OK", "", "3.0.0"))
$r = Invoke-Script $c "delete_old_models.ps1" @("-Execute", "-Yes")
Check (Test-Path -LiteralPath $o) "relink in job but no relink report -> kept"

$c = New-Case "del_relink_duplicate_names"
$o1 = Join-Path $c.Data "old\d1\A.rvt"; $n1 = Join-Path $c.Data "new\d1\A_N.rvt"
$o2 = Join-Path $c.Data "old\d2\B.rvt"; $n2 = Join-Path $c.Data "new\d2\A_N.rvt"
New-File $o1; New-File $n1; New-File $o2; New-File $n2
Write-Job $c @(@{ old = $o1; new = $n1 }, @{ old = $o2; new = $n2 }) $relinkJob
Write-Report $c saveas @(@($o1, $n1, "OK", "", "3"), @($o2, $n2, "OK", "", "3"))
Write-Report $c relink @(, @("A_N", "OK", 1, 0, 0, "", "3.0.0"))
$r = Invoke-Script $c "delete_old_models.ps1" @("-Execute", "-Yes")
Check ((Test-Path -LiteralPath $o1) -and (Test-Path -LiteralPath $o2)) "same new name, report without host_path -> both kept"
Write-Report $c relink @(@("A_N", "OK", 1, 0, 0, "", "3.1.0", $n1), @("A_N", "ERROR", 0, 0, 1, "", "3.1.0", $n2)) -WithHostPath
$r = Invoke-Script $c "delete_old_models.ps1" @("-Execute", "-Yes")
Check (-not (Test-Path -LiteralPath $o1)) "host_path disambiguates: OK one moved"
Check (Test-Path -LiteralPath $o2) "host_path disambiguates: ERROR one kept"

# update_rvt_links.py writes the FULL host path into "host"
Write-Report $c relink @(@($n1, "OK", 1, 0, 0, "", "3.0.0"), @($n2, "ERROR", 0, 0, 1, "", "3.0.0"))
$o1 = Join-Path $c.Data "old\d1\A.rvt"; New-File $o1
$r = Invoke-Script $c "delete_old_models.ps1" @("-Execute", "-Yes")
Check (-not (Test-Path -LiteralPath $o1)) "host holds full path: OK one moved"
Check (Test-Path -LiteralPath $o2) "host holds full path: ERROR one kept"

$c = New-Case "del_buffer"
$bufJob = @{ saveas_needed = $true; relink = $false; copy = $false; delete_old = $true }
$rsnOld = "RSN://revit-server-2022.example.local/Proj/A_R22.rvt"
$stg = Join-Path $c.Data "buffer\A_N.rvt"; $nw = Join-Path $c.Data "new\A_N.rvt"
$rsnOld2 = "RSN://revit-server-2022.example.local/Proj/B_R22.rvt"; $stg2 = Join-Path $c.Data "buffer\B_N.rvt"; $nw2 = Join-Path $c.Data "new\B_N.rvt"
New-File $stg; New-File $nw; New-File $stg2; New-File $nw2
Write-Job $c @(@{ old = $rsnOld; new = $nw; stg = $stg; buf = $true }, @{ old = $rsnOld2; new = $nw2; stg = $stg2; buf = $true }) $bufJob
Write-Report $c detach @(@($rsnOld, $stg, "OK", "", "3"), @($rsnOld2, $stg2, "OK", "", "3"))
Write-Report $c saveas @(@($stg, $nw, "OK", "", "3"), @($stg2, $nw2, "ERROR", "boom", "3"))
$r = Invoke-Script $c "delete_old_models.ps1" @("-Execute", "-Yes")
Check (-not (Test-Path -LiteralPath $stg)) "buffer copy moved when both passes OK"
Check (Test-Path -LiteralPath $stg2) "buffer kept when saveas(buffer) ERROR"
$rsnList = Join-Path $c.Tool "to_delete_rsn.txt"
Check ((Test-Path -LiteralPath $rsnList) -and ((Get-Content -LiteralPath $rsnList -Raw) -match "A_R22") -and ((Get-Content -LiteralPath $rsnList -Raw) -notmatch "B_R22")) "RSN list has only verified models"

$c = New-Case "del_copy_mode"
$o = Join-Path $c.Data "old\A.rvt"; $n = Join-Path $c.Data "new\A2.rvt"
New-File $o; New-File $n
Write-Job $c @(@{ old = $o; new = $n }) @{ saveas_needed = $true; relink = $false; copy = $true; delete_old = $true }
Write-Report $c saveas @(, @($o, $n, "OK", "", "3"))
$r = Invoke-Script $c "delete_old_models.ps1" @("-Execute", "-Yes")
Check (Test-Path -LiteralPath $o) "copy mode keeps old"

$c = New-Case "del_no_job_json"
$r = Invoke-Script $c "delete_old_models.ps1" @("-Execute", "-Yes")
Check ($r.Code -eq 1) "missing job files -> exit 1"

# -------------------------------------------------------------------- move
$c = New-Case "move_dryrun_default"
$o = Join-Path $c.Data "a\M.rvt"; $n = Join-Path $c.Data "b\M.rvt"
New-File $o
Write-Job $c @(@{ old = $o; new = $n }) @{}
$r = Invoke-Script $c "move_centrals.ps1"
Check ((Test-Path -LiteralPath $o) -and -not (Test-Path -LiteralPath $n)) "move dry-run changes nothing"
Check (-not (Test-Path -LiteralPath (Join-Path $c.Data "b"))) "move dry-run creates no folders"
$r = Invoke-Script $c "move_centrals.ps1" @("-Execute") "no"
Check (Test-Path -LiteralPath $o) "move: answer no keeps file"
$r = Invoke-Script $c "move_centrals.ps1" @("-Execute", "-Yes")
Check ((-not (Test-Path -LiteralPath $o)) -and (Test-Path -LiteralPath $n)) "move executes with -Execute -Yes"
Check (@(Get-ChildItem -LiteralPath $c.Rep -Filter "move_centrals_*.csv").Count -eq 1) "move writes report csv"

$c = New-Case "move_dest_exists"
$o = Join-Path $c.Data "a\M.rvt"; $n = Join-Path $c.Data "b\M.rvt"
New-File $o "source-content"; New-File $n "dest-content"
Write-Job $c @(@{ old = $o; new = $n }) @{}
$r = Invoke-Script $c "move_centrals.ps1" @("-Execute", "-Yes")
Check (Test-Path -LiteralPath $o) "dest exists -> source stays"
Check ((Get-Content -LiteralPath $n -Raw).Trim() -eq "dest-content") "dest exists -> dest not overwritten"
Check ($r.Out -match "SKIP_EXISTS") "SKIP_EXISTS reported"
$r = Invoke-Script $c "move_centrals.ps1" @("-Execute", "-Yes", "-Overwrite")
Check (-not (Test-Path -LiteralPath $o)) "-Overwrite moves source"
Check ((Get-Content -LiteralPath $n -Raw).Trim() -eq "source-content") "-Overwrite puts new content in place"
Check (Quarantined $c (Join-Path $c.Data "b") "M.rvt") "-Overwrite: old dest kept in _to_delete"

$c = New-Case "move_misc"
$o1 = Join-Path $c.Data "a\X.rvt"; $o2 = Join-Path $c.Data "a\Y.rvt"; $dst = Join-Path $c.Data "b\Z.rvt"; $gone = Join-Path $c.Data "a\Gone.rvt"
New-File $o1; New-File $o2
Write-Job $c @(@{ old = $o1; new = $dst }, @{ old = $o2; new = $dst }, @{ old = $gone; new = (Join-Path $c.Data "b\G.rvt") }, @{ old = "RSN://h/p/R.rvt"; new = "RSN://h/p/R2.rvt" }) @{}
$r = Invoke-Script $c "move_centrals.ps1" @("-Execute", "-Yes")
Check ((-not (Test-Path -LiteralPath $o1)) -and (Test-Path -LiteralPath $dst)) "first row moves"
Check (Test-Path -LiteralPath $o2) "duplicate destination row skipped"
Check ($r.Out -match "SKIP_NO_SOURCE") "missing source reported"
Check ($r.Code -eq 0) "skips are not errors"

# ------------------------------------------------------- plan file / window
$ru = -join ([char[]](0x41F, 0x440, 0x43E, 0x435, 0x43A, 0x442))   # Cyrillic folder name
$c = New-Case "planfile_delete"
$o = Join-Path $c.Data "$ru\old\A.rvt"; $n = Join-Path $c.Data "$ru\new\A2.rvt"
New-File $o; New-File $n
Write-Job $c @(@{ old = $o; new = $n }) $SaveasJob
Write-Report $c saveas @(, @($o, $n, "OK", "", "3.0.0"))
$plan = Join-Path $c.Root "plan.txt"
$r = Invoke-Script $c "delete_old_models.ps1" @("-PlanFile", $plan)
$pt = Get-Content -LiteralPath $plan -Raw -Encoding UTF8
Check ($pt -match "MOVE" -and $pt.Contains($ru)) "delete -PlanFile has the plan, Cyrillic intact"
Check ($pt -match "To move to _to_delete:\s+1\s+RSN to list:\s+0") "delete -PlanFile has the summary line the window parses"
Check (Test-Path -LiteralPath $o) "-PlanFile run is still a dry-run"

$c = New-Case "planfile_move"
$o = Join-Path $c.Data "$ru\a\M.rvt"; $n = Join-Path $c.Data "$ru\b\M.rvt"
New-File $o
Write-Job $c @(@{ old = $o; new = $n }) @{}
$plan = Join-Path $c.Root "plan.txt"
$r = Invoke-Script $c "move_centrals.ps1" @("-PlanFile", $plan)
$pt = Get-Content -LiteralPath $plan -Raw -Encoding UTF8
Check ($pt -match "To move:\s+1" -and $pt.Contains($ru)) "move -PlanFile has the plan and summary, Cyrillic intact"
Check (Test-Path -LiteralPath $o) "move -PlanFile run is still a dry-run"

# Invoke-OpsPsDryRun (window helper) through a powershell.exe shim on non-Windows
$onWindows = ($env:OS -eq "Windows_NT")
if (-not $onWindows) {
    $shim = Join-Path $Work "shim"
    New-Item -ItemType Directory -Path $shim -Force | Out-Null
    Set-Content -LiteralPath (Join-Path $shim "powershell.exe") -Value ("#!/bin/sh`nexec `"{0}`" `"`$@`"" -f $Ps) -Encoding ASCII
    & chmod +x (Join-Path $shim "powershell.exe")
    $env:PATH = $shim + [System.IO.Path]::PathSeparator + $env:PATH
}
$c = New-Case "window_dryrun"
$home2 = Join-Path $c.Root "home"
$env:USERPROFILE = $home2
$toolCopy = Join-Path $c.Root "toolcopy"
New-Item -ItemType Directory -Path $toolCopy -Force | Out-Null
foreach ($f in @("delete_old_models.ps1", "move_centrals.ps1", "safety_lib.ps1")) { Copy-Item -LiteralPath (Join-Path $ToolSrc $f) -Destination $toolCopy }
$c.Tool = $toolCopy
# the scripts default to the tool's own reports folder <tool root>\отчёты
$c.Rep = Join-Path $toolCopy "отчёты"
New-Item -ItemType Directory -Path $c.Rep -Force | Out-Null
$o = Join-Path $c.Data "$ru\old\A.rvt"; $n = Join-Path $c.Data "$ru\new\A2.rvt"
New-File $o; New-File $n
Write-Job $c @(@{ old = $o; new = $n }) $SaveasJob
Write-Report $c saveas @(, @($o, $n, "OK", "", "3.0.0"))
$ToolDir = $toolCopy
. (Join-Path $ToolSrc "start_ops.ps1")
$d = Invoke-OpsPsDryRun -Pass ([pscustomobject]@{ Name = "delete"; Kind = "delete"; Task = (Join-Path $toolCopy "delete_old_models.ps1") })
Check ($d.Ok -and $d.Count -eq 1) "Invoke-OpsPsDryRun delete: Ok and Count=1 (Ok=$($d.Ok) Count=$($d.Count) Code=$($d.Code))"
Check ($d.Text.Contains($ru)) "Invoke-OpsPsDryRun returns Cyrillic path intact"
Check (Test-Path -LiteralPath $o) "Invoke-OpsPsDryRun changed nothing"
$d = Invoke-OpsPsDryRun -Pass ([pscustomobject]@{ Name = "move"; Kind = "move"; Overwrite = $false; Task = (Join-Path $toolCopy "move_centrals.ps1") })
Check ($d.Ok) "Invoke-OpsPsDryRun move: Ok (nothing to move is not an error)"
Remove-Item -LiteralPath (Join-Path $toolCopy "job_paths.csv") -Force
$d = Invoke-OpsPsDryRun -Pass ([pscustomobject]@{ Name = "delete"; Kind = "delete"; Task = (Join-Path $toolCopy "delete_old_models.ps1") })
Check (-not $d.Ok) "Invoke-OpsPsDryRun reports failure when job files are missing"

# ------------------------------------------------------------------ result
Remove-Item -LiteralPath $Work -Recurse -Force -ErrorAction SilentlyContinue
Write-Host ""
Write-Host ("Checks passed: {0}, failed: {1}" -f $script:Pass, $script:Fail)
if ($script:Fail -gt 0) { exit 1 }
exit 0
