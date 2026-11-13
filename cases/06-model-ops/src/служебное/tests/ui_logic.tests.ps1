# Name: ui_logic.tests.ps1
# Version: 3.1.0
# What it does: See the case README.
# Notes: Paths and project names in this portfolio copy are placeholders.

# Tests for ui_logic.ps1 (pure logic of the operations window).
# No dependencies, Windows PowerShell 5.1 or 7:
#   powershell -NoProfile -ExecutionPolicy Bypass -File tests\ui_logic.tests.ps1

$ToolSrc = Split-Path -Parent $PSScriptRoot
. (Join-Path $ToolSrc "job_lib.ps1")
. (Join-Path $ToolSrc "ui_logic.ps1")

$script:Fail = 0
$script:Pass = 0
function Check([bool]$Cond, [string]$Msg) {
    if ($Cond) { $script:Pass++ }
    else { $script:Fail++; Write-Host "  FAIL: $Msg" -ForegroundColor Red }
}
function Eq($Got, $Want, [string]$Msg) { Check ($Got -ceq $Want) ("{0}: got '{1}', want '{2}'" -f $Msg, $Got, $Want) }

# ----------------------------------------------------- common path prefix
Eq (Get-CommonPathPrefix @('\\srv\proj\A\RVT', '\\srv\proj\B\RVT')) '\\srv\proj' "UNC siblings"
Eq (Get-CommonPathPrefix @('\\srv\proj\A\RVT')) '\\srv\proj\A\RVT' "single folder"
Eq (Get-CommonPathPrefix @('\\SRV\Proj\A', '\\srv\proj\a')) '\\SRV\Proj\A' "case-insensitive, first spelling kept"
Eq (Get-CommonPathPrefix @('\\srv\proj\Abc', '\\srv\proj\Abd')) '\\srv\proj' "whole segments only"
Eq (Get-CommonPathPrefix @('\\srv\proj\A\', '\\srv\proj\A')) '\\srv\proj\A' "trailing slash ignored"
Eq (Get-CommonPathPrefix @('\\srv1\p\A', '\\srv2\p\A')) '' "different servers share nothing"
Eq (Get-CommonPathPrefix @('RSN://revit-server-2022.example.local/P/A', 'RSN://revit-server-2022.example.local/P/B')) 'RSN://revit-server-2022.example.local/P' "RSN same project"
Eq (Get-CommonPathPrefix @('RSN://revit-server-2022.example.local/P/A', 'RSN://revit-server-2022.example.local/Q/B')) 'RSN://revit-server-2022.example.local' "RSN same host"
Eq (Get-CommonPathPrefix @('RSN://revit-server-2022.example.local/P', 'RSN://revit-server-2023.example.local/P')) '' "RSN different hosts"
Eq (Get-CommonPathPrefix @('RSN://h/P', '\\srv\P')) '' "RSN vs UNC"
Eq (Get-CommonPathPrefix @('C:\a\b', 'C:\a\c')) 'C:\a' "drive siblings"
Eq (Get-CommonPathPrefix @('C:\a', 'D:\a')) '' "different drives"
Eq (Get-CommonPathPrefix @()) '' "empty list"
Eq (Get-CommonPathPrefix @('', $null)) '' "blank entries"
$ru = -join ([char[]](0x41F, 0x440, 0x43E, 0x435, 0x43A, 0x442))
Eq (Get-CommonPathPrefix @("\\srv\$ru\A", "\\srv\$ru\B")) "\\srv\$ru" "Cyrillic segment"

Eq (Get-PathFolder '\\srv\proj\A\x.rvt') '\\srv\proj\A' "folder UNC"
Eq (Get-PathFolder 'RSN://h/P/x.rvt') 'RSN://h/P' "folder RSN"
Eq (Get-PathFolder 'C:\a\x.rvt') 'C:\a' "folder drive"
Eq (Get-PathFolder 'x.rvt') '' "folder of bare name"

# ------------------------------------------------------ segment suggestion
$all = @('\\srv\proj\A\RVT\a.rvt', '\\srv\proj\B\RVT\b.rvt', '\\srv\other\c.rvt')
Eq (Get-SegmentSuggestion @('\\srv\proj\A\RVT\a.rvt', '\\srv\proj\B\RVT\b.rvt') $all) '\\srv\proj' "uses checked models"
Eq (Get-SegmentSuggestion @('\\srv\proj\A\RVT\a.rvt') $all) '\\srv\proj\A\RVT' "one checked model: its folder"
Eq (Get-SegmentSuggestion @() $all) '\\srv' "nothing checked: all models"
Eq (Get-SegmentSuggestion @('RSN://h/P/a.rvt', '\\srv\proj\b.rvt') @()) 'RSN://h/P' "no common part: folder of the first"
Eq (Get-SegmentSuggestion @() @()) '' "no models"

# ---------------------------------------------------------------- names
$r = Test-ModelFileName 'PROJ_IOS_EOM_KLWT_K01_R24'
Check ($r.Ok -and $r.Value -ceq 'PROJ_IOS_EOM_KLWT_K01_R24') "plain name ok"
$r = Test-ModelFileName '  PROJ_IOS_EOM_KLWT_K01_R24.RVT '
Check ($r.Ok -and $r.Value -ceq 'PROJ_IOS_EOM_KLWT_K01_R24') "name with .RVT and spaces normalized"
$r = Test-ModelFileName 'PROJ_IOS_EOM_KLWT_K03-K05-K07-K10-ST_R24'
Check ($r.Ok) "compound name ok"
$r = Test-ModelFileName "${ru}_K01"
Check ($r.Ok) "Cyrillic name ok"
$r = Test-ModelFileName 'Model v1.2_K01'
Check ($r.Ok) "dots inside name ok"
foreach ($c in @('', '   ', '.rvt')) { Eq (Test-ModelFileName $c).Code 'empty' "empty '$c'" }
Eq (Test-ModelFileName $null).Code 'empty' "null name"
foreach ($ch in @('\', '/', ':', '*', '?', '"', '<', '>', '|', "`t")) {
    $x = Test-ModelFileName ("A{0}B" -f $ch)
    Check ((-not $x.Ok) -and $x.Code -eq 'chars') "invalid char [$([int][char]$ch)] rejected"
}
Eq (Test-ModelFileName 'A.').Code 'trailing' "trailing dot"
Eq (Test-ModelFileName 'A .rvt').Code 'trailing' "trailing space before .rvt"
foreach ($n in @('CON', 'nul.rvt', 'Com3', 'LPT9', 'aux')) { Eq (Test-ModelFileName $n).Code 'reserved' "reserved $n" }
Check ((Test-ModelFileName 'CONSOLE').Ok) "CONSOLE is not reserved"
Eq (Test-ModelFileName 'A.rfa').Code 'ext' "foreign extension .rfa"
Eq (Test-ModelFileName 'A.RFA.rvt').Code 'ext' "foreign extension before .rvt"
Check ((Test-ModelFileName ('x' * 251)).Ok) "251 chars + .rvt = 255 ok"
Eq (Test-ModelFileName ('x' * 252)).Code 'length' "252 chars + .rvt = 256 too long"

# ----------------------------------------------------------------- paths
foreach ($p in @('\\srv\p\A\x.rvt', 'C:\a\x.rvt', 'RSN://h/p/x.rvt', 'c:/a/x.RVT', 'RSN://h/x.rvt', "\\srv\$ru\x.rvt")) {
    Check ((Test-NewModelPath $p).Ok) "path ok: $p"
}
$r = Test-NewModelPath '"C:\a\x.rvt"'
Check ($r.Ok -and $r.Value -ceq 'C:\a\x.rvt') "quotes trimmed"
Eq (Test-NewModelPath '').Code 'empty' "empty path"
Eq (Test-NewModelPath 'C:\a\x').Code 'ext' "no extension"
Eq (Test-NewModelPath 'C:\a\x.rfa').Code 'ext' "wrong extension"
foreach ($p in @('x.rvt', 'a\x.rvt', '\\srv\x.rvt', 'RSN://x.rvt', 'C:x.rvt')) {
    Eq (Test-NewModelPath $p).Code 'format' "not a full path: $p"
}
Eq (Test-NewModelPath 'RSN://h/').Code 'ext' "RSN folder without file"
Eq (Test-NewModelPath 'C:\a\\b\x.rvt').Code 'segments' "double slash"
Eq (Test-NewModelPath 'C:\a\b?\x.rvt').Code 'chars' "invalid char in folder"
Eq (Test-NewModelPath 'C:\a:b\x.rvt').Code 'chars' "colon in folder"
Eq (Test-NewModelPath 'C:\a\CON.rvt').Code 'reserved' "reserved file name inside path"
Eq (Test-NewModelPath 'C:\a\ .rvt').Code 'empty' "blank file name inside path"

# ------------------------------------------------------------- conflicts
$rows = @(
    [pscustomobject]@{ old_name = 'A'; old_path = '\\s\p\A.rvt'; new_path = '\\s\p\A2.rvt' },
    [pscustomobject]@{ old_name = 'B'; old_path = '\\s\p\B.rvt'; new_path = '\\s\p\B2.rvt' }
)
Check ((Test-PathConflict '\\s\p\Z.rvt' '\\s\p\A.rvt' $rows).Ok) "free path"
Eq (Test-PathConflict '\\s\p\B2.rvt' '\\s\p\A.rvt' $rows).Code 'duplicate_new' "same as another new path"
Eq (Test-PathConflict '\\S\P\b2.RVT' '\\s\p\A.rvt' $rows).Code 'duplicate_new' "case differs"
Eq (Test-PathConflict '//s/p/B2.rvt' '\\s\p\A.rvt' $rows).Code 'duplicate_new' "slash style differs"
Eq (Test-PathConflict '\\s\p\B.rvt' '\\s\p\A.rvt' $rows).Code 'other_old' "equals another model's OLD path"
Check ((Test-PathConflict '\\s\p\A.rvt' '\\s\p\A.rvt' $rows).Ok) "own old path is fine"
Check ((Test-PathConflict '\\s\p\A2.rvt' '\\s\p\A.rvt' $rows).Ok) "own current new path is fine"
Check ((Test-PathConflict '\\s\p\Z.rvt' '\\s\p\A.rvt' @()).Ok) "no rows"

# ---------------------------------------------------------------- rules
$s = Get-RuleStatus 'PROJ_IOS_EOM_ORG_K01_R24' 'PROJ_IOS_EOM_KLWT_K01_R24'
Check ($s.State -eq 'ok' -and $s.Ops.Count -eq 1 -and $s.Got -ceq 'PROJ_IOS_EOM_KLWT_K01_R24') "example: rule built and reproduces the sample"
$out = Invoke-NameOps -Name 'PROJ_IOS_EOM_ORG_K03-K05-K07-K10-ST_R24' -Ops $s.Ops
Eq $out 'PROJ_IOS_EOM_KLWT_K03-K05-K07-K10-ST_R24' "example rule on the compound name"
$s = Get-RuleStatus 'PROJ_IOS_EOM_ORG_K01_R24.rvt' 'PROJ_IOS_EOM_KLWT_K01_R24.RVT'
Eq $s.State 'ok' ".rvt in the fields is ignored"
Eq (Get-RuleStatus '' '').State 'empty' "both empty"
Eq (Get-RuleStatus $null $null).State 'empty' "both null"
$s = Get-RuleStatus 'A_B' ''
Check ($s.State -eq 'fail' -and $s.Message) "one field empty -> fail with a message"
$s = Get-RuleStatus 'A_B' 'A_B'
Check ($s.State -eq 'same' -and $s.Message) "identical names"
foreach ($pair in @(@('A_B_C', 'C_B_A'), @('A_B_C_D', 'D_A_B_C'), @('X_Y', 'Y_X'))) {
    $s = Get-RuleStatus $pair[0] $pair[1]
    $consistent = (($s.State -eq 'ok') -and ($s.Got -ceq $pair[1])) -or (($s.State -eq 'fail') -and $s.Message)
    Check $consistent ("reorder {0} -> {1} is either reproduced or reported, state {2}" -f $pair[0], $pair[1], $s.State)
}

# ---------------------------------------------------------- manual edits
$e = Set-ManualName $null 'NEW_K01'
Check ($e.name -ceq 'NEW_K01' -and -not $e.path) "name edit without path override"
$e = Set-ManualPath $null '\\s\p2\Other_K01.rvt'
Check ($e.path -ceq '\\s\p2\Other_K01.rvt' -and $e.name -ceq 'Other_K01') "path edit sets name from file name"
$e2 = Set-ManualName $e 'Third_K01'
Eq $e2.path '\\s\p2\Third_K01.rvt' "name edit rewrites the file name inside the path override"
Eq $e2.name 'Third_K01' "name kept"
$e3 = Set-ManualName (Set-ManualPath $null 'RSN://h/P/Old.rvt') 'New'
Eq $e3.path 'RSN://h/P/New.rvt' "RSN path keeps its folder"
Eq $e.path '\\s\p2\Other_K01.rvt' "original entry not mutated"

# ------------------------------------------------- preview rows + manual rows
$ops = @(New-OpsFromSample 'PROJ_IOS_EOM_ORG_K01_R24' 'PROJ_IOS_EOM_KLWT_K01_R24')
$olds = @('\\s\p\PROJ_IOS_EOM_ORG_K01_R24.rvt', '\\s\p\PROJ_IOS_EOM_ORG_K03-K05-K07-K10-ST_R24.rvt')
function Get-Rows($Manual) {
    return @(New-JobPreviewRows -OldPaths $olds -Ops $ops -DoRename $true -LocationMode 'same' `
            -OldSegment '' -NewSegment '' -OldPrefix '' -NewPrefix '' -SourceYear '' -TargetYear '' `
            -BufferDir 'C:\buf' -NameOverrides @{} -ManualRows $Manual)
}
$rows = Get-Rows @{}
Eq $rows[0].new_name 'PROJ_IOS_EOM_KLWT_K01_R24' "rule: example row"
Eq $rows[1].new_name 'PROJ_IOS_EOM_KLWT_K03-K05-K07-K10-ST_R24' "rule: compound row"
Eq $rows[0].new_path '\\s\p\PROJ_IOS_EOM_KLWT_K01_R24.rvt' "rule: path follows the new name"
Check ($rows[0].manual -eq '' -and $rows[1].manual -eq '') "no manual edits -> manual is empty"

$k0 = Get-ManualRowKey $olds[0]
$k1 = Get-ManualRowKey $olds[1]
Eq (Get-ManualRowKey '\\S\p\A.rvt') (Get-ManualRowKey '//s/P/a.RVT') "manual key ignores case and slash style"

$m = @{}
$m[$k0] = Set-ManualName $null 'Hand_Made_01'
$rows = Get-Rows $m
Eq $rows[0].new_name 'Hand_Made_01' "manual name beats the rule"
Eq $rows[0].new_path '\\s\p\Hand_Made_01.rvt' "manual name flows into the path"
Eq $rows[0].manual 'name' "marked as manual name"
Eq $rows[0].rule_name 'PROJ_IOS_EOM_KLWT_K01_R24' "rule value is still available for the tooltip"
Eq $rows[1].new_name 'PROJ_IOS_EOM_KLWT_K03-K05-K07-K10-ST_R24' "other rows keep the rule"

$m = @{}
$m[$k0] = Set-ManualPath $null '\\t\other\Moved_K01.rvt'
$rows = Get-Rows $m
Eq $rows[0].new_path '\\t\other\Moved_K01.rvt' "manual path beats the rule"
Eq $rows[0].new_name 'Moved_K01' "name follows the file name in the manual path"
Eq $rows[0].manual 'path' "marked as manual path"
Eq $rows[0].rule_path '\\s\p\PROJ_IOS_EOM_KLWT_K01_R24.rvt' "rule path kept for the tooltip"

$m = @{}
$m[$k0] = @{ name = 'OnlyName' }          # entry without a path key
$rows = Get-Rows $m
Eq $rows[0].new_name 'OnlyName' "entry with only a name key works"

# staging path (RSN -> new year) uses the manual name
$rsn = @('RSN://revit-server-2022.example.local/P/PROJ_IOS_EOM_ORG_K01_R24.rvt')
$m = @{}
$m[(Get-ManualRowKey $rsn[0])] = Set-ManualName $null 'Hand_Made_01'
$rows = @(New-JobPreviewRows -OldPaths $rsn -Ops $ops -DoRename $true -LocationMode 'same' `
        -OldSegment '' -NewSegment '' -OldPrefix '' -NewPrefix '' -SourceYear '2022' -TargetYear '2024' `
        -BufferDir ([System.IO.Path]::GetTempPath()) -NameOverrides @{} -ManualRows $m)
Eq (Split-Path -Leaf $rows[0].staging_path) 'Hand_Made_01.rvt' "buffer file uses the manual name"

# two rows pointed at the same path are flagged by the preview itself
$m = @{}
$m[$k0] = Set-ManualPath $null '\\s\p\Same.rvt'
$m[$k1] = Set-ManualPath $null '\\s\p\same.RVT'
$rows = Get-Rows $m
Check ($rows[1].note -match 'duplicate') "duplicate manual paths are flagged in the note"

# ------------------------------------------------------------- clipboard
$rows = @(
    [pscustomobject]@{ old_path = '\\s\p\A.rvt'; new_path = '\\s\q\A2.rvt' },
    [pscustomobject]@{ old_path = "\\s\$ru\B.rvt"; new_path = "\\s\q\$ru.rvt" }
)
Eq (Format-RowsForClipboard $rows 'both') ("\\s\p\A.rvt`t\\s\q\A2.rvt`r`n\\s\$ru\B.rvt`t\\s\q\$ru.rvt") "both: tab between, one row per line"
Eq (Format-RowsForClipboard $rows 'old') ("\\s\p\A.rvt`r`n\\s\$ru\B.rvt") "old only"
Eq (Format-RowsForClipboard $rows 'new') ("\\s\q\A2.rvt`r`n\\s\q\$ru.rvt") "new only"
Eq (Format-RowsForClipboard @([pscustomobject]@{ old_path = 'X'; new_path = '' }) 'both') 'X' "no new path column -> old only"
Eq (Format-RowsForClipboard @($rows[0]) 'both') "\\s\p\A.rvt`t\\s\q\A2.rvt" "single row has no trailing newline"
Eq (Format-RowsForClipboard @() 'both') '' "no rows"

# ------------------------------------------- effective manual rows / tooltips
$man = @{}
$man['K1'] = Set-ManualName $null 'NameOnly'
$man['K2'] = Set-ManualPath $null '\\t\q\FromPath.rvt'
$man['K3'] = Set-ManualName (Set-ManualPath $null '\\t\q\Both.rvt') 'Both2'
$eff = Get-EffectiveManualRows $man $true $true
Check ($eff.Count -eq 3 -and $eff['K1'].name -ceq 'NameOnly' -and $eff['K2'].path -ceq '\\t\q\FromPath.rvt') "rename+move: everything applied"
Eq $eff['K3'].path '\\t\q\Both2.rvt' "name edit after a path edit rewrote the file name"
$eff = Get-EffectiveManualRows $man $true $false
Check ($eff.ContainsKey('K1') -and $eff['K1'].path -eq $null) "rename only: name edit applied"
Check ($eff['K2'].name -ceq 'FromPath' -and $eff['K2'].path -eq $null) "rename only: path edit gives only its file name"
$eff = Get-EffectiveManualRows $man $false $true
Check ((-not $eff.ContainsKey('K1')) -and $eff['K2'].path -ceq '\\t\q\FromPath.rvt') "move only: name-only edit is not applied, path edit is"
Eq (Get-EffectiveManualRows $man $false $false).Count 0 "both off: nothing applied"
Eq (Get-EffectiveManualRows $null $true $true).Count 0 "no manual rows"
Check ($man.Count -eq 3) "source hashtable is not modified"

$row = [pscustomobject]@{ manual = ''; old_path = '\\s\p\A.rvt'; new_path = '\\s\p\B.rvt'; new_name = 'B'; rule_name = 'B'; rule_path = '\\s\p\B.rvt' }
Eq (Get-CellTipText $row 4) '\\s\p\A.rvt' "tip: old path in full"
Check ((Get-CellTipText $row 5).StartsWith('\\s\p\B.rvt')) "tip: new path in full"
Check ((Get-CellTipText $row 1) -notmatch 'Изменено вручную') "tip: no 'manual' note for rule rows"
Check ((Get-CellTipText $row 1) -match 'Двойной щелчок - править вручную') "tip: rule rows say how to edit (also proves Cyrillic literals are read correctly)"
Eq (Get-CellTipText $row 0) '' "tip: other columns have none"
Eq (Get-CellTipText $null 4) '' "tip: no row"
$row.manual = 'path'; $row.rule_path = '\\s\p\Rule.rvt'
Check ((Get-CellTipText $row 5) -match 'Изменено вручную' -and (Get-CellTipText $row 5) -match 'Rule\.rvt') "tip: manual path shows the rule value"

# -------------------------------------------------- "Подробности": blocks follow the checkboxes
foreach ($c in @(@($false, $false, $false), @($true, $false, $false), @($false, $true, $false), @($false, $false, $true), @($true, $true, $true))) {
    $b = Get-DetailsBlockState $c[0] $c[1] $c[2]
    $tag = "rename=$($c[0]) move=$($c[1]) upgrade=$($c[2])"
    Check ($b.RuleEnabled -eq $c[0] -and $b.MoveEnabled -eq $c[1] -and $b.BufferEnabled -eq $c[2]) "block state follows the checkboxes ($tag)"
    Check (($b.RuleHint -eq "") -eq $c[0] -and ($b.MoveHint -eq "") -eq $c[1] -and ($b.BufferHint -eq "") -eq $c[2]) "hint exactly when the block is off ($tag)"
}
Eq (Get-DetailsBlockState $false $true $true).RuleHint "Включите «Переименовать» в главном окне" "rule hint text"
Eq (Get-DetailsBlockState $true $false $true).MoveHint "Включите «Перенести» в главном окне" "move hint text"
Eq (Get-DetailsBlockState $true $true $false).BufferHint "Включите «Сменить версию Revit» в главном окне" "buffer hint text"
Check (Test-MoveSettingsMissing $true "segment" "" "x") "move on, segment mode, empty 'было' -> missing"
Check (Test-MoveSettingsMissing $true "prefix" "x" "  ") "move on, prefix mode, blank prefix -> missing"
Check (-not (Test-MoveSettingsMissing $true "segment" "\\a" "")) "move on, segment mode with 'было' -> configured"
Check (-not (Test-MoveSettingsMissing $true "prefix" "" "\\a")) "move on, prefix mode with a prefix -> configured"
Check (-not (Test-MoveSettingsMissing $false "segment" "" "")) "move off -> nothing to configure"

# -------------------------------------------------- copy models between tabs
$merged = @(Merge-TabModelPaths @("RSN://h/A.rvt") @("rsn://h/a.rvt", "\\srv\b.rvt"))
Eq $merged.Count 2 "the same model is not copied twice"
Eq $merged[0] "RSN://h/A.rvt" "the path already on the tab is kept"
Eq $merged[1] "\\srv\b.rvt" "a new path is appended"

# -------------------------------------------------- saved model lists
$storeRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("lists_" + [Guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $storeRoot -Force | Out-Null
try {
    $made = Get-ModelListStore $storeRoot
    Eq $made (Join-Path $storeRoot "списки") "the default list folder is списки under the tool"
    Check (Test-Path -LiteralPath $made -PathType Container) "the default folder is created"
    $other = Join-Path $storeRoot "other"
    New-Item -ItemType Directory -Path $other -Force | Out-Null
    $saved = Join-Path $other "модели.txt"
    $utf8 = New-Object System.Text.UTF8Encoding $true
    [System.IO.File]::WriteAllLines($saved, @("# comment", "RSN://revit-server-2022.example.local/П/А.rvt", "\\srv\share\Б.rvt", "skip.txt"), $utf8)
    Save-ModelListFolder $storeRoot $saved
    $cfgBytes = [System.IO.File]::ReadAllBytes((Join-Path $storeRoot "lists_dir.cfg"))
    Check ($cfgBytes.Length -ge 3 -and $cfgBytes[0] -eq 0xEF -and $cfgBytes[1] -eq 0xBB -and $cfgBytes[2] -eq 0xBF) "lists_dir.cfg is UTF-8 with BOM"
    Eq (Get-ModelListStore $storeRoot) $other "the next dialog opens the last folder"
    $read = @(Read-ModelListPaths $saved)
    Eq $read.Count 2 "comments and non-rvt lines are dropped"
    Eq $read[0] "RSN://revit-server-2022.example.local/П/А.rvt" "a one-line-or-more list keeps the RSN path"
    Eq $read[1] "\\srv\share\Б.rvt" "Cyrillic UNC path survives UTF-8"
}
finally { Remove-Item -LiteralPath $storeRoot -Recurse -Force -ErrorAction SilentlyContinue }

Write-Host ""
Write-Host ("Checks passed: {0}, failed: {1}" -f $script:Pass, $script:Fail)
if ($script:Fail -gt 0) { exit 1 }
exit 0
