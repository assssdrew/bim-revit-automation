# Name: relink_job.tests.ps1
# Version: 3.1.0
# What it does: See the case README.
# Notes: Paths and project names in this portfolio copy are placeholders.

# Tests for the 20.10 additions to job_lib.ps1: Get-CommonModelFolder,
# ConvertTo-PrefixKey, Get-JobProp, relink_map.csv export, year from servers.cfg,
# and the pass queue (together with the safety/overwrite flag of #9).
#   powershell -NoProfile -ExecutionPolicy Bypass -File tests\relink_job.tests.ps1

$ToolSrc = Split-Path -Parent $PSScriptRoot
. (Join-Path $ToolSrc "job_lib.ps1")

$script:Fail = 0
$script:Pass = 0
function Check([bool]$Cond, [string]$Msg) {
    if ($Cond) { $script:Pass++ }
    else { $script:Fail++; Write-Host "  FAIL: $Msg" -ForegroundColor Red }
}
function Eq($Got, $Want, [string]$Msg) { Check ($Got -ceq $Want) ("{0}: got '{1}', want '{2}'" -f $Msg, $Got, $Want) }

$Work = Join-Path ([System.IO.Path]::GetTempPath()) ("ops_relink_tests_" + [Guid]::NewGuid().ToString("N").Substring(0, 8))
New-Item -ItemType Directory -Path $Work -Force | Out-Null

# ---------------------------------------------------------- ConvertTo-PrefixKey
Eq (ConvertTo-PrefixKey '  \\SRV\Share\Old\ ') '//srv/share/old' "prefix key: case, slashes, trailing"
Eq (ConvertTo-PrefixKey 'RSN://H/Proj/') 'rsn://h/proj' "prefix key: RSN"
Eq (ConvertTo-PrefixKey '') '' "prefix key: empty"
Eq (ConvertTo-PrefixKey $null) '' "prefix key: null"
Check ((ConvertTo-PrefixKey '\\srv\a\B') -eq (ConvertTo-PrefixKey '//SRV/a/b/')) "prefix key: same folder spelled differently"

# ------------------------------------------------------------------ Get-JobProp
$j = [pscustomobject]@{ a = '  x  '; b = $null; c = 5 }
Eq (Get-JobProp $j 'a') 'x' "job prop: trimmed"
Eq (Get-JobProp $j 'b') '' "job prop: null"
Eq (Get-JobProp $j 'zz') '' "job prop: absent (works under strict mode)"
Eq (Get-JobProp $j 'c') '5' "job prop: non-string"
Eq (Get-JobProp $null 'a') '' "job prop: no job"

# ------------------------------------------------------- Get-CommonModelFolder
$cyr = -join ([char[]](0x421, 0x442, 0x440, 0x43E, 0x433, 0x438, 0x43D, 0x43E))   # SampleSite
Eq (Get-CommonModelFolder @('\\srv\proj\obj\A\x.rvt', '\\srv\proj\obj\B\y.rvt')) '\\srv\proj\obj' "UNC siblings keep backslashes"
Eq (Get-CommonModelFolder @("RSN://revit-server-2022.example.local/proj $cyr/Stage/04/x.rvt", "RSN://revit-server-2022.example.local/proj $cyr/Stage/09/y.rvt")) "RSN://revit-server-2022.example.local/proj $cyr/Stage" "RSN siblings keep forward slashes"
Eq (Get-CommonModelFolder @('D:\a\b\c\x.rvt', 'D:\a\b\d\y.rvt')) 'D:\a\b' "drive siblings"
Eq (Get-CommonModelFolder @('\\SRV\Proj\Obj\A\x.rvt', '\\srv\proj\obj\B\y.rvt')) '\\SRV\Proj\Obj' "case-insensitive, first spelling kept"
Eq (Get-CommonModelFolder @('\\srv\proj\obj\A\x.rvt', '\\srv\proj\obj\A\y.rvt')) '' "one folder: nothing to guess"
Eq (Get-CommonModelFolder @('\\srv\proj\A\x.rvt', '\\srv\proj\B\y.rvt')) '' "common part shorter than 3 segments is not guessed"
Eq (Get-CommonModelFolder @('\\s1\p\a\b\x.rvt', '\\s2\p\a\b\y.rvt')) '' "different servers"
Eq (Get-CommonModelFolder @('x.rvt', 'y.rvt')) '' "names without folders"
Eq (Get-CommonModelFolder @()) '' "no paths"
Eq (Get-CommonModelFolder @('', $null)) '' "blank paths"

# ------------------------------------------------- relink_map.csv in Save-JobFiles
function New-TestJob([string]$RelOld, [string]$RelNew) {
    return [pscustomobject]@{
        source_year = ''; target_year = ''; relink = $true; copy = $false; delete_old = $false
        saveas_needed = $false; unc_move = $false; relink_old = $RelOld; relink_new = $RelNew
    }
}
$rows = @([pscustomobject]@{
        old_name = 'A'; new_name = 'A'; old_path = '\\s\p\A.rvt'; new_path = '\\s\p\A.rvt'
        staging_path = ''; needs_buffer = $false; note = ''
    })
$tool = Join-Path $Work 'save'
$map = Join-Path $tool 'relink_map.csv'
[void](Save-JobFiles -ToolDir $tool -Rows $rows -Job (New-TestJob 'RSN://h/Old;Folder' 'RSN://h/New'))
Check (Test-Path -LiteralPath $map) "relink_map.csv written for a relink-only job"
$lines = @(Get-Content -LiteralPath $map -Encoding UTF8)
Eq $lines[0] 'old_prefix;new_prefix' "map header"
Eq $lines[1] '"RSN://h/Old;Folder";RSN://h/New' "map row: ';' inside a path is quoted"
Check ((Get-Content -LiteralPath $map -Raw -Encoding UTF8).Length -gt 0 -and ([System.IO.File]::ReadAllBytes($map)[0] -eq 0xEF)) "map is UTF-8 with BOM (relink_prefix.py reads utf-8-sig)"
Check (Test-Path -LiteralPath (Join-Path $tool 'job.json')) "job.json written as before"
[void](Save-JobFiles -ToolDir $tool -Rows $rows -Job (New-TestJob '\\SRV\a\B\' '//srv/a/b'))
Check (-not (Test-Path -LiteralPath $map)) "same folder spelled differently: map removed"
[void](Save-JobFiles -ToolDir $tool -Rows $rows -Job (New-TestJob 'RSN://h/Old' 'RSN://h/New'))
[void](Save-JobFiles -ToolDir $tool -Rows $rows -Job (New-TestJob '' ''))
Check (-not (Test-Path -LiteralPath $map)) "not a relink-only job: a stale map from the previous job is removed"
[void](Save-JobFiles -ToolDir $tool -Rows $rows -Job (New-TestJob 'RSN://h/Old' ''))
Check (-not (Test-Path -LiteralPath $map)) "only one folder given: no map"

# ------------------------------------------------------------ New-OpsPassQueue
$qd = Join-Path $Work 'queue'
New-Item -ItemType Directory -Path $qd -Force | Out-Null
Set-Content -LiteralPath (Join-Path $qd 'servers.cfg') -Encoding UTF8 -Value @('# host|label|year', 'revit-server-2022.example.local|Revit Server 2022|2022', 'revit-server-2024.example.local|Revit Server 2024|2024')
function New-QJob([hashtable]$Over) {
    $h = @{ source_year = ''; target_year = ''; relink = $true; copy = $false; delete_old = $false; saveas_needed = $false; unc_move = $false }
    foreach ($k in $Over.Keys) { $h[$k] = $Over[$k] }
    return [pscustomobject]$h
}
function New-QRow([string]$Path, [bool]$Buf = $false) { return [pscustomobject]@{ old_path = $Path; needs_buffer = $Buf } }
$rsn22 = New-QRow 'RSN://revit-server-2022.example.local/P/A.rvt'
$rsn24 = New-QRow 'RSN://revit-server-2024.example.local/P/A.rvt'

$q = @(New-OpsPassQueue -Job (New-QJob @{}) -ToolDir $qd -Rows @($rsn22))
Eq $q.Count 1 "relink-only: one pass"
Eq $q[0].Name 'relink' "relink-only: pass name"
Eq $q[0].Year '2022' "no year chosen: taken from servers.cfg by the RSN host"
$q = @(New-OpsPassQueue -Job (New-QJob @{}) -ToolDir $qd -Rows @($rsn24))
Eq $q[0].Year '2024' "year follows the host (2024)"
$q = @(New-OpsPassQueue -Job (New-QJob @{ target_year = '2023' }) -ToolDir $qd -Rows @($rsn22))
Eq $q[0].Year '2023' "chosen target year wins"
$q = @(New-OpsPassQueue -Job (New-QJob @{ source_year = '2021' }) -ToolDir $qd -Rows @($rsn22))
Eq $q[0].Year '2021' "chosen source year beats the server year"
$q = @(New-OpsPassQueue -Job (New-QJob @{}) -ToolDir $qd -Rows @(New-QRow '\\s\p\A.rvt'))
Eq $q[0].Year '' "not an RSN model: no year guessed"
$q = @(New-OpsPassQueue -Job (New-QJob @{}) -ToolDir $qd -Rows @(New-QRow 'RSN://fake.example.local/P/A.rvt'))
Eq $q[0].Year '' "RSN host missing from servers.cfg: no year guessed"
$q = @(New-OpsPassQueue -Job (New-QJob @{ saveas_needed = $true }) -ToolDir $qd -Rows @($rsn22))
Eq $q[0].Name 'saveas' "saveas pass first"
Eq $q[0].Year '2022' "saveas pass: year from servers.cfg too"

# the move pass carries the overwrite flag from #9 (default: no overwrite)
$mv = New-QJob @{ unc_move = $true; relink = $false }
$q = @(New-OpsPassQueue -Job $mv -ToolDir $qd -Rows @(New-QRow '\\s\p\A.rvt'))
Eq $q[0].Name 'move' "move pass"
Check ($q[0].Overwrite -eq $false) "move pass: overwrite is off by default"
$mv2 = New-QJob @{ unc_move = $true; relink = $false; allow_overwrite = $true }
$q = @(New-OpsPassQueue -Job $mv2 -ToolDir $qd -Rows @(New-QRow '\\s\p\A.rvt'))
Check ($q[0].Overwrite -eq $true) "move pass: overwrite on when the job allows it"

# order and delete pass
$full = New-QJob @{ saveas_needed = $true; unc_move = $true; relink = $true; delete_old = $true }
$q = @(New-OpsPassQueue -Job $full -ToolDir $qd -Rows @(New-QRow '\\s\p\A.rvt'))
Eq (($q | ForEach-Object { $_.Name }) -join ',') 'saveas,move,relink,delete' "pass order"
$q = @(New-OpsPassQueue -Job (New-QJob @{ saveas_needed = $true; relink = $true; delete_old = $true; target_year = '2024' }) -ToolDir $qd -Rows @((New-QRow 'RSN://revit-server-2022.example.local/P/A.rvt' $true)))
Eq (($q | ForEach-Object { $_.Name }) -join ',') 'detach,saveas,relink,delete' "buffer job: detach first"
Eq $q[0].Year '' "detach pass runs in the source Revit (no year when none chosen)"
Eq $q[1].Year '2024' "saveas of the buffer runs in the target Revit"

$bakRoot = Join-Path $Work "bak"
New-Item -ItemType Directory -Path $bakRoot -Force | Out-Null
[System.IO.File]::WriteAllText((Join-Path $bakRoot "job.json"), '{"keep":1}')
$bak = Backup-OpsJobFiles $bakRoot
Check (Test-Path -LiteralPath (Join-Path $bak.Dir "job.json")) "backup keeps the previous job.json"
[System.IO.File]::WriteAllText((Join-Path $bakRoot "job.json"), '{"other":1}')
[System.IO.File]::WriteAllText((Join-Path $bakRoot "job_relink.json"), '{"relink":1}')
[System.IO.File]::WriteAllText((Join-Path $bakRoot "mapping.csv"), "old;new")
Restore-OpsJobFiles $bak $bakRoot
Eq ([System.IO.File]::ReadAllText((Join-Path $bakRoot "job.json"))) '{"keep":1}' "restore puts the previous job back"
Check (-not (Test-Path -LiteralPath (Join-Path $bakRoot "mapping.csv"))) "a file created by the other tab is removed"
Check (Test-Path -LiteralPath (Join-Path $bakRoot "job_relink.json")) "the relink copy is not part of the restore set"

Remove-Item -LiteralPath $Work -Recurse -Force -ErrorAction SilentlyContinue
Write-Host ""
Write-Host ("Checks passed: {0}, failed: {1}" -f $script:Pass, $script:Fail)
if ($script:Fail -gt 0) { exit 1 }
exit 0
