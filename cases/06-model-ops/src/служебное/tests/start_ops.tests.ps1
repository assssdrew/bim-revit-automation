# Name: start_ops.tests.ps1
# Version: 3.1.0
# What it does: See the case README.
# Notes: Paths and project names in this portfolio copy are placeholders.

# Tests for start_ops.ps1: how the RBP and PowerShell passes are launched.
# The process start is replaced by a recorder, nothing is started.
#   powershell -NoProfile -ExecutionPolicy Bypass -File tests\start_ops.tests.ps1

$ToolSrc = Split-Path -Parent $PSScriptRoot
$ToolDir = $ToolSrc
. (Join-Path $ToolSrc "job_lib.ps1")
. (Join-Path $ToolSrc "start_ops.ps1")

$script:Fail = 0
$script:Pass = 0
function Check([bool]$Cond, [string]$Msg) {
    if ($Cond) { $script:Pass++ }
    else { $script:Fail++; Write-Host "  FAIL: $Msg" -ForegroundColor Red }
}

$Work = Join-Path ([System.IO.Path]::GetTempPath()) ("ops_start_tests_" + [Guid]::NewGuid().ToString("N").Substring(0, 8))
New-Item -ItemType Directory -Path $Work -Force | Out-Null
$task = Join-Path $Work "task.py"
$list = Join-Path $Work "my list.txt"
Set-Content -LiteralPath $task -Value "x"
Set-Content -LiteralPath $list -Value "y"

$script:Rec = $null
function Start-OpsHiddenProcess {
    param([string]$FilePath, [string]$Arguments = "", [string]$WorkingDirectory = "")
    $script:Rec = @{ File = $FilePath; Args = $Arguments; Dir = $WorkingDirectory }
    return [pscustomobject]@{ Fake = $true }
}
function Run-Rbp([string]$Name, [bool]$Detach, [string]$Year) {
    $script:Rec = $null
    $pass = [pscustomobject]@{ Kind = "rbp"; Name = $Name; Task = $task; List = $list; Year = $Year; Detach = $Detach }
    return (Start-OpsRbpPass -Pass $pass -Exe "C:\BatchRvt\BatchRvt.exe" -ExeDir "C:\BatchRvt")
}

# ---- RBP passes
$r = Run-Rbp "relink" $false "2024"
Check ($r.Ok -and $script:Rec) "relink pass starts"
Check ($script:Rec.Args -match '--worksets open_all') "relink opens all worksets (LoadFrom fails in a closed workset)"
Check ($script:Rec.Args -match '--create_new_local' -and $script:Rec.Args -notmatch '--detach') "relink: Create New Local, no Detach"
Check ($script:Rec.Args -match '--revit_version 2024') "revit version passed"
Check ($script:Rec.Args -match '--task_script') "task script passed"
Check ($script:Rec.Args -match '--file_list "[^"]*my list\.txt"') "path with a space is quoted"
Check ($script:Rec.File -eq "C:\BatchRvt\BatchRvt.exe" -and $script:Rec.Dir -eq "C:\BatchRvt") "runs the found BatchRvt in its folder"

$r = Run-Rbp "saveas" $false "2024"
Check ($script:Rec.Args -match '--worksets close_all') "saveas keeps close_all"
$r = Run-Rbp "detach" $true "2022"
Check ($script:Rec.Args -match '--detach --worksets close_all') "detach: --detach with close_all"
Check ($script:Rec.Args -notmatch '--create_new_local') "detach: no Create New Local"
$r = Run-Rbp "relink" $false ""
Check ($script:Rec.Args -notmatch '--revit_version') "no year chosen: no --revit_version"

$script:Rec = $null
$wsPass = [pscustomobject]@{ Kind = "rbp"; Name = "workset-read"; Task = $task; List = $list; Year = "2024"; Detach = $false; Worksets = "open_all" }
$r = Start-OpsRbpPass -Pass $wsPass -Exe "C:\BatchRvt\BatchRvt.exe" -ExeDir "C:\BatchRvt"
Check ($r.Ok) "workset read pass starts"
Check ($script:Rec.Args -match '--worksets open_all') "full workset read opens all worksets"
Check ($script:Rec.Args -match '--create_new_local' -and $script:Rec.Args -notmatch '--detach') "workset read: Create New Local, Detach off"
$script:Rec = $null
$fastPass = [pscustomobject]@{ Kind = "rbp"; Name = "workset-read"; Task = $task; List = $list; Year = "2024"; Detach = $false; Worksets = "close_all" }
$r = Start-OpsRbpPass -Pass $fastPass -Exe "C:\BatchRvt\BatchRvt.exe" -ExeDir "C:\BatchRvt"
Check ($r.Ok) "fast workset read pass starts"
Check ($script:Rec.Args -match '--worksets close_all') "fast workset read keeps worksets closed"
Check ($script:Rec.Args -match '--create_new_local' -and $script:Rec.Args -notmatch '--detach') "fast workset read still uses a new local"
Check ($script:Rec.Args -match 'read_worksets\.py|task\.py') "workset read task script is passed"
$script:Rec = $null
$applyPass = [pscustomobject]@{ Kind = "rbp"; Name = "workset-apply"; Task = $task; List = $list; Year = "2024"; Detach = $false; Worksets = "open_all" }
$r = Start-OpsRbpPass -Pass $applyPass -Exe "C:\BatchRvt\BatchRvt.exe" -ExeDir "C:\BatchRvt"
Check ($r.Ok) "workset apply pass starts"
Check ($script:Rec.Args -match '--worksets open_all') "workset apply opens all worksets"
Check ($script:Rec.Args -match '--create_new_local' -and $script:Rec.Args -notmatch '--detach') "workset apply: Create New Local, Detach off"

$script:Rec = $null
$bad = Start-OpsRbpPass -Pass ([pscustomobject]@{ Kind = "rbp"; Name = "relink"; Task = (Join-Path $Work "nope.py"); List = $list; Year = ""; Detach = $false }) -Exe "x" -ExeDir "y"
Check ((-not $bad.Ok) -and $null -eq $script:Rec) "missing task script: not started"
$bad = Start-OpsRbpPass -Pass ([pscustomobject]@{ Kind = "rbp"; Name = "relink"; Task = $task; List = (Join-Path $Work "nope.txt"); Year = ""; Detach = $false }) -Exe "x" -ExeDir "y"
Check ((-not $bad.Ok) -and $null -eq $script:Rec) "missing file list: not started"

# ---- PowerShell passes (move / delete): never confirmed or applied on their own
$mv = [pscustomobject]@{ Kind = "move"; Name = "move"; Task = (Join-Path $Work "move_centrals.ps1"); List = ""; Year = ""; Detach = $false; Overwrite = $false }
$script:Rec = $null
$r = Start-OpsPsPass -Pass $mv
Check ($r.Ok -and $script:Rec.File -eq "powershell.exe") "move pass starts via powershell.exe"
Check ($script:Rec.Args -notmatch '-Yes' -and $script:Rec.Args -notmatch '-Execute') "move pass gets no -Yes / -Execute by itself (dry-run is the default)"
$r = Start-OpsPsPass -Pass $mv -ExtraArgs @("-Execute", "-Yes")
Check ($script:Rec.Args -match '-Execute -Yes$') "extra args are appended after the window confirmed"
Check ($script:Rec.Args -match '-File "[^"]*move_centrals\.ps1"') "script path quoted"
$r = Start-OpsPsPass -Pass $mv -ExtraArgs @("-Execute", "-Yes", "-Overwrite")
Check ($script:Rec.Args -match '-Overwrite') "overwrite only when passed explicitly"

$oldProfile = $env:USERPROFILE
try {
    $env:USERPROFILE = $Work
    $rep = Join-Path $Work "Documents\doc\script\rbp\reports"
    New-Item -ItemType Directory -Path $rep -Force | Out-Null
    $pointer = Join-Path $rep "_active_workset_read.txt"
    Set-Content -LiteralPath $pointer -Value "keep" -Encoding ASCII
    $applyPointer = Join-Path $rep "_active_workset_apply.txt"
    Set-Content -LiteralPath $applyPointer -Value "keep" -Encoding ASCII
    $keep = Join-Path $rep "_active_saveas_central.txt"
    Set-Content -LiteralPath $keep -Value "old" -Encoding ASCII
    Reset-OpsReportSession
    Check (-not (Test-Path -LiteralPath $pointer)) "read pointer _active_workset_read.txt is removed"
    Check (-not (Test-Path -LiteralPath $applyPointer)) "apply pointer _active_workset_apply.txt is removed"
    Check (-not (Test-Path -LiteralPath $keep)) "existing report pointers are still removed"
}
finally { $env:USERPROFILE = $oldProfile }

Remove-Item -LiteralPath $Work -Recurse -Force -ErrorAction SilentlyContinue
Write-Host ""
Write-Host ("Checks passed: {0}, failed: {1}" -f $script:Pass, $script:Fail)
if ($script:Fail -gt 0) { exit 1 }
exit 0
