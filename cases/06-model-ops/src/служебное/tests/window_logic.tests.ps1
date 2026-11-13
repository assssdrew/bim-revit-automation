# Name: window_logic.tests.ps1
# Version: 3.1.0
# What it does: See the case README.
# Notes: Paths and project names in this portfolio copy are placeholders.

# Logic of the window that does not need WinForms: relink-only mode and the per-model
# progress (_progress_ops.txt). The real functions are taken out of presets.ps1 with the
# PowerShell parser and run against stub controls, so the file itself is not changed.
#   powershell -NoProfile -ExecutionPolicy Bypass -File tests\window_logic.tests.ps1

$ToolSrc = Split-Path -Parent $PSScriptRoot
. (Join-Path $ToolSrc "job_lib.ps1")
Set-StrictMode -Off     # presets.ps1 switches strict mode off after loading the libraries

$script:Fail = 0
$script:Pass = 0
function Check([bool]$Cond, [string]$Msg) {
    if ($Cond) { $script:Pass++ }
    else { $script:Fail++; Write-Host "  FAIL: $Msg" -ForegroundColor Red }
}
function Eq($Got, $Want, [string]$Msg) { Check ($Got -ceq $Want) ("{0}: got '{1}', want '{2}'" -f $Msg, $Got, $Want) }

$want = @("Test-UiRelinkOnly", "Get-RunList", "Test-OpsPassTracked", "Get-OpsStateRu", "Get-OpsPathKey", "Read-OpsProgressText",
    "Set-RowStateIfChanged", "Set-UnfinishedState", "Update-OpsProgress", "Reset-OpsProgress", "Get-OpsProgressFile",
    "Get-CheckedPathErrorCount")
$errs = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $ToolSrc "presets.ps1"), [ref]$null, [ref]$errs)
$defs = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)
$found = @()
foreach ($d in $defs) { if ($want -contains $d.Name) { Invoke-Expression $d.Extent.Text; $found += $d.Name } }
Check ($found.Count -eq $want.Count) ("all window functions found in presets.ps1 (found {0} of {1})" -f $found.Count, $want.Count)
$script:FinalStatesRu = @("Готово", "Частично", "Ошибка", "Без изменений")   # defined next to the functions in presets.ps1
Check ((Get-Content -LiteralPath (Join-Path $ToolSrc "presets.ps1") -Raw -Encoding UTF8) -match [regex]::Escape('$script:FinalStatesRu = @("Готово", "Частично", "Ошибка", "Без изменений")')) "test copy of FinalStatesRu matches presets.ps1"

# ---------------------------------------------------------------- stubs
function New-Box([bool]$Checked = $false) { return [pscustomobject]@{ Checked = $Checked; Visible = $true; Text = ""; Top = 0; Height = 0 } }
function New-Ui {
    $ui = @{}
    foreach ($n in "CbRelink", "CbRename", "CbMove", "CbUpgrade", "CbCopy", "CbDelete", "CbOverwrite", "CbWorksets",
        "LblRelOld", "TxtRelOld", "LblRelNew", "TxtRelNew", "LblY1", "CmbSrc", "LblY2", "CmbDst", "BtnMore", "Summary") { $ui[$n] = New-Box }
    $ui.Status = New-Box
    $ui.Progress = [pscustomobject]@{ Maximum = 1; Value = 0 }
    $ui.List = [pscustomobject]@{ Items = @() }
    return $ui
}
$script:CheckedPaths = @()
function Get-CheckedPaths { return @($script:CheckedPaths) }
function New-Item2([string]$Name, [string]$Old, [string]$New, [bool]$Checked = $true, [string]$State = "") {
    $subs = @([pscustomobject]@{ Text = $Name }, [pscustomobject]@{ Text = $Name }, [pscustomobject]@{ Text = $State },
        [pscustomobject]@{ Text = "" }, [pscustomobject]@{ Text = $Old }, [pscustomobject]@{ Text = $New })
    return [pscustomobject]@{ Text = $Name; Tag = $Old; Checked = $Checked; SubItems = $subs }
}

# --------------------------------------------------- relink-only mode
$script:Ui = New-Ui
$script:Ui.Remove("TxtRelOld")
Check (-not (Test-UiRelinkOnly)) "no relink controls yet -> not relink-only"
$script:Ui = New-Ui
$script:Ui.CbRelink.Checked = $true
Check (Test-UiRelinkOnly) "only 'relink' ticked -> relink-only"
foreach ($other in "CbRename", "CbMove", "CbUpgrade", "CbCopy", "CbDelete") {
    $script:Ui = New-Ui; $script:Ui.CbRelink.Checked = $true; $script:Ui[$other].Checked = $true
    Check (-not (Test-UiRelinkOnly)) "relink + $other -> not relink-only"
}
$script:Ui = New-Ui
Check (-not (Test-UiRelinkOnly)) "nothing ticked -> not relink-only"
$script:Ui.CbRelink.Checked = $true; $script:Ui.CbOverwrite.Checked = $true
Check (Test-UiRelinkOnly) "the overwrite tick does not change the mode"
$relOnly = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq "Test-UiRelinkOnly" }, $true) | Select-Object -First 1
Check ($relOnly.Extent.Text -notmatch "CbWorksets") "worksets have their own tabs: no worksets switch in the relink-only test"
Check (Test-OpsPassTracked "relink") "relink progress stays tracked"
Check (Test-OpsPassTracked "workset-read") "workset read progress is tracked"
Check (Test-OpsPassTracked "workset-apply") "workset apply progress is tracked"
Check (-not (Test-OpsPassTracked "saveas")) "saveas progress is not the per-model tracker"
Check (-not (Test-OpsPassTracked "detach")) "detach progress is not the per-model tracker"

# tab «Перепривязка связей»: the folders of the job are the two fields of that tab
$presetsText = Get-Content -LiteralPath (Join-Path $ToolSrc "presets.ps1") -Raw -Encoding UTF8
Check ($presetsText -match [regex]::Escape('$script:Ui.TxtRelOld = $relOld')) "relink old folder of the job = «Старая папка» of the tab"
Check ($presetsText -match [regex]::Escape('$script:Ui.TxtRelNew = $relNew')) "relink new folder of the job = «Новая папка» of the tab"
foreach ($gone in @("Sync-RelinkFields", "Show-DetailsDialog", "BtnMore", "CmbSrc", "ChkReport", "Подставлен прошлый список", "Только перепривязка", "Только смена версии", "Вставка связей подключится")) {
    Check (-not $presetsText.Contains($gone)) ("removed from the window: {0}" -f $gone)
}
$moveRun = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq "Start-MoveTabRun" }, $true) | Select-Object -First 1
Check ($null -ne $moveRun) "the move tab has its own run"
Check ($moveRun.Extent.Text -match "CbRelink.Checked = \`$false") "the move tab does not relink"
Check ($moveRun.Extent.Text -match "CbUpgrade.Checked = \`$false") "the move tab does not upgrade"
Check ($moveRun.Extent.Text -match "Start-UiRun") "the move tab still uses the existing run"

# ---------------------------------------------------- path errors (only ticked models block the run)
$script:Ui = New-Ui
$script:RowCache = @{
    'a' = [pscustomobject]@{ path_error = "" }
    'b' = [pscustomobject]@{ path_error = "Новый путь не полный" }
    'c' = [pscustomobject]@{ path_error = "Новый путь не полный" }
}
$script:Ui.List.Items = @(
    [pscustomobject]@{ Tag = 'a'; Checked = $true }, [pscustomobject]@{ Tag = 'b'; Checked = $true },
    [pscustomobject]@{ Tag = 'c'; Checked = $false })
Eq (Get-CheckedPathErrorCount) 1 "path errors: the unticked model with an error is not counted"
$script:Ui.List.Items[2].Checked = $true
Eq (Get-CheckedPathErrorCount) 2 "path errors: ticking it counts it"
$script:Ui.List.Items = @()
Eq (Get-CheckedPathErrorCount) 0 "path errors: empty list"

# ---------------------------------------------------------- progress
Eq (Get-OpsStateRu "STARTED") "Идёт" "state STARTED"
Eq (Get-OpsStateRu "ok") "Готово" "state ok (case-insensitive)"
Eq (Get-OpsStateRu " PARTIAL ") "Частично" "state PARTIAL (trimmed)"
Eq (Get-OpsStateRu "ERROR") "Ошибка" "state ERROR"
Eq (Get-OpsStateRu "NO_CHANGES") "Без изменений" "state NO_CHANGES"
Eq (Get-OpsStateRu "whatever") "" "unknown state"
Eq (Get-OpsPathKey '  \\SRV\a\B.rvt ') '//SRV/A/B.RVT' "path key: slashes and case"
Eq (Get-OpsPathKey '') '' "path key: empty"

$work = Join-Path ([System.IO.Path]::GetTempPath()) ("ops_window_tests_" + [Guid]::NewGuid().ToString("N").Substring(0, 8))
New-Item -ItemType Directory -Path $work -Force | Out-Null
$pf = Join-Path $work "_progress_ops.txt"
$cy = -join ([char[]](0x41C, 0x43E, 0x434, 0x435, 0x43B, 0x44C))      # Модель
$p1 = "RSN://h/P/${cy}_1.rvt"; $p2 = "RSN://h/P/${cy}_2.rvt"; $p3 = "RSN://h/P/${cy}_3.rvt"; $p4 = "RSN://h/P/${cy}_4.rvt"; $p5 = "RSN://h/P/${cy}_5.rvt"
$lines = @(
    "STARTED`t10:00:01`t$p1", "OK`t10:00:09`t$p1",
    "STARTED`t10:00:10`t$p2", "PARTIAL`t10:00:20`t$p2",
    "STARTED`t10:00:21`t$p3", "ERROR`t10:00:25`t$p3",
    "STARTED`t10:00:26`t$p4"
)
[System.IO.File]::WriteAllLines($pf, $lines, (New-Object System.Text.UTF8Encoding $true))

$script:Ui = New-Ui
$script:Run = @{ PassLabel = "Проход 2 из 2: Перепривязываю связи" }
$items = @(
    (New-Item2 "m1" $p1 $p1), (New-Item2 "m2" $p2 $p2), (New-Item2 "m3" $p3 $p3), (New-Item2 "m4" $p4 $p4),
    (New-Item2 "m5" $p5 $p5), (New-Item2 "m6-unchecked" "RSN://h/P/other.rvt" "RSN://h/P/other.rvt" $false "keep")
)
$script:Ui.List.Items = $items
Update-OpsProgress $pf
Eq $items[0].SubItems[2].Text "Готово" "finished OK -> Готово"
Eq $items[1].SubItems[2].Text "Частично" "PARTIAL -> Частично"
Eq $items[2].SubItems[2].Text "Ошибка" "ERROR -> Ошибка"
Eq $items[3].SubItems[2].Text "Идёт" "STARTED, no result yet -> Идёт"
Eq $items[4].SubItems[2].Text "В очереди" "not seen yet -> В очереди"
Eq $items[5].SubItems[2].Text "keep" "unchecked rows are not touched"
Check ($script:Ui.Progress.Maximum -eq 5) "progress maximum = checked models"
Check ($script:Ui.Progress.Value -eq 3) "progress value = finished models (3 of 5)"
Check ($script:Ui.Status.Text -match "Проход 2 из 2" -and $script:Ui.Status.Text -match "модель 4 из 5" -and $script:Ui.Status.Text -match "m4") "status line: pass, 'модель 4 из 5' and the current model"

# the file keeps growing: the next poll moves on
[System.IO.File]::AppendAllText($pf, "NO_CHANGES`t10:00:40`t$p4`r`nSTARTED`t10:00:41`t$p5`r`n", (New-Object System.Text.UTF8Encoding $false))
Update-OpsProgress $pf
Eq $items[3].SubItems[2].Text "Без изменений" "NO_CHANGES -> Без изменений"
Eq $items[4].SubItems[2].Text "Идёт" "next model is running"
Check ($script:Ui.Progress.Value -eq 4 -and $script:Ui.Status.Text -match "модель 5 из 5") "progress moved to 4 done, model 5 of 5"

# a model matched by its NEW path (column 5) or its old path (tag)
$script:Ui.List.Items = @((New-Item2 "x" "RSN://h/P/old.rvt" $p1))
Update-OpsProgress $pf
Eq $script:Ui.List.Items[0].SubItems[2].Text "Готово" "matched by the new path column"

# missing / empty file: nothing breaks, state stays
$script:Ui.List.Items = @((New-Item2 "x" $p1 $p1 $true "Идёт"))
Update-OpsProgress (Join-Path $work "nope.txt")
Eq $script:Ui.List.Items[0].SubItems[2].Text "В очереди" "no progress file: rows are 'В очереди'"
Eq (Read-OpsProgressText (Join-Path $work "nope.txt")) "" "reading a missing file gives an empty string"

# the file can be locked for writing by RBP while the window reads it
$fs = New-Object System.IO.FileStream($pf, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Write, [System.IO.FileShare]::ReadWrite)
$txt = Read-OpsProgressText $pf
$fs.Dispose()
Check ($txt -and $txt.Contains("STARTED")) "a file open for writing by RBP can still be read"

# end of a pass: unfinished rows get the pass result, finished ones keep theirs
$script:Ui.List.Items = @((New-Item2 "a" $p1 $p1 $true "Готово"), (New-Item2 "b" $p2 $p2 $true "Идёт"), (New-Item2 "c" $p3 $p3 $false "x"))
Set-UnfinishedState "Ошибка"
Eq $script:Ui.List.Items[0].SubItems[2].Text "Готово" "finished row keeps its state"
Eq $script:Ui.List.Items[1].SubItems[2].Text "Ошибка" "unfinished row gets the pass result"
Eq $script:Ui.List.Items[2].SubItems[2].Text "x" "unchecked row untouched"

# a run of another tab writes the states into that tab's list, not into tab 1
$tab1 = @((New-Item2 "t1" $p1 $p1 $true "было"))
$other = [pscustomobject]@{ Items = @((New-Item2 "o1" $p1 $p1 $true "Идёт")) }
$script:Ui.List.Items = $tab1
$script:Run = @{ PassLabel = "Проход 1 из 1: Перепривязываю связи"; List = $other }
Update-OpsProgress $pf
Eq $other.Items[0].SubItems[2].Text "Готово" "progress goes to the list of the running tab"
Eq $tab1[0].SubItems[2].Text "было" "tab 1 list is not touched by another tab's run"
Set-UnfinishedState "Ошибка"
Eq $tab1[0].SubItems[2].Text "было" "end of another tab's pass leaves tab 1 alone"
$script:Run = @{ PassLabel = "x" }

# Reset-OpsProgress removes the previous run's file in <tool>\отчёты
$oldProfile = $env:USERPROFILE
$env:USERPROFILE = $work
$ToolDir = $work
$progFile = Get-OpsProgressFile
Eq (Split-Path -Leaf (Split-Path -Parent $progFile)) "отчёты" "progress file lives in the tool's отчёты folder"
New-Item -ItemType Directory -Path (Split-Path -Parent $progFile) -Force | Out-Null
Set-Content -LiteralPath $progFile -Value "STARTED`t10:00:01`tx"
$script:Run = @{ ProgressText = "old" }
Reset-OpsProgress
Check (-not (Test-Path -LiteralPath $progFile)) "Reset-OpsProgress removes the old progress file"
Eq $script:Run.ProgressText "" "Reset-OpsProgress clears the cached text"
Reset-OpsProgress
Check $true "Reset-OpsProgress with no file does not fail"
$env:USERPROFILE = $oldProfile

Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue
Write-Host ""
Write-Host ("Checks passed: {0}, failed: {1}" -f $script:Pass, $script:Fail)
if ($script:Fail -gt 0) { exit 1 }
exit 0
