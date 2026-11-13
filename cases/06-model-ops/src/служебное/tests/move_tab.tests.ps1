# Name: move_tab.tests.ps1
# Version: 3.1.0
# What it does: See the case README.
# Notes: Paths and project names in this portfolio copy are placeholders.

# Tab «Перенос и переименование» without WinForms: the rule and the folders are typed on the tab,
# an empty field = that step is not done. The REAL functions are taken out of presets.ps1 with
# the PowerShell parser and run against stub controls.
# Also: «Смена версии Revit» (the source version is detected), isolated tab runs keep the list
# of tab 1, «Пробный прогон» text, the run button per tab.
#   powershell -NoProfile -ExecutionPolicy Bypass -File tests\move_tab.tests.ps1

$ToolSrc = Split-Path -Parent $PSScriptRoot
. (Join-Path $ToolSrc "job_lib.ps1")
. (Join-Path $ToolSrc "ui_logic.ps1")
Set-StrictMode -Off     # presets.ps1 switches strict mode off after loading the libraries
Add-Type -AssemblyName System.Drawing

$script:Fail = 0
$script:Pass = 0
function Check([bool]$Cond, [string]$Msg) {
    if ($Cond) { $script:Pass++ }
    else { $script:Fail++; Write-Host "  FAIL: $Msg" -ForegroundColor Red }
}
function Eq($Got, $Want, [string]$Msg) { Check ($Got -ceq $Want) ("{0}: got '{1}', want '{2}'" -f $Msg, $Got, $Want) }

$presets = Join-Path $ToolSrc "presets.ps1"
$ast = [System.Management.Automation.Language.Parser]::ParseFile($presets, [ref]$null, [ref]$null)
$presetsText = [System.IO.File]::ReadAllText($presets)
function Find-Func([string]$Name) {
    return $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $Name }, $true)
}
$need = @("Get-SelectedYears", "Get-PreviewRows", "Get-ManualCounts", "Sync-MoveFlags", "Show-MoveFieldState", "Update-MoveFields",
    "Set-MoveSuggestions", "Write-WindowError", "Show-WindowError", "Set-UiError", "Set-UiStatusTip", "Build-JobObject", "Test-UiRelinkOnly",
    "Invoke-IsolatedTabRun", "Start-RelinkTabRun", "Start-UpgradeTabRun", "Get-DryRunText", "Get-RunButtonText", "Get-ActiveTabKey",
    "Update-UpgradeTab", "Get-TabModelPaths")
$found = 0
foreach ($n in $need) { $f = Find-Func $n; if ($f) { Invoke-Expression $f.Extent.Text; $found++ } }
Eq $found $need.Count "functions found in presets.ps1"

# ---------------------------------------------------------------- stubs
function New-Box([bool]$Checked = $false, [string]$Text = "") {
    return [pscustomobject]@{ Checked = $Checked; Text = $Text; ForeColor = $null; Enabled = $true; Visible = $true }
}
$script:Rows = @()
function Set-UiStatus([string]$Text) { $script:Ui.Status.Text = $Text; $script:Ui.Status.ForeColor = [System.Drawing.SystemColors]::ControlText }   # the real one pumps WinForms messages
function Refresh-ModelList { $script:Rows = @(Get-PreviewRows) }      # the real one needs the ListView
function Get-CheckedPaths { return @($script:modelPaths) }
function Initialize-OpsReportsDir([string]$Root) { return (Join-Path $Root "отчёты") }   # no folders in the test
$ToolDir = Join-Path ([System.IO.Path]::GetTempPath()) "move_tab_tool"

function Reset-World {
    $script:Ui = @{
        CbRename = (New-Box); CbMove = (New-Box); CbUpgrade = (New-Box); CbRelink = (New-Box)
        CbCopy = (New-Box $true); CbDelete = (New-Box); CbOverwrite = (New-Box)
        RuleOld = (New-Box); RuleNew = (New-Box); FolderOld = (New-Box); FolderNew = (New-Box)
        Status = (New-Box); TxtRelOld = (New-Box); TxtRelNew = (New-Box)
    }
    $script:modelPaths = New-Object System.Collections.Generic.List[string]
    [void]$script:modelPaths.Add('\\srv\p\Old\PROJ_IOS_EOM_ORG_K01_R24.rvt')
    [void]$script:modelPaths.Add('\\srv\p\Old\PROJ_IOS_EOM_ORG_K03-K05-K07-K10-ST_R24.rvt')
    $script:ops = @(); $script:nameOverrides = @{}; $script:manualRows = @{}; $script:locMode = "same"
    $script:oldSeg = ""; $script:newSeg = ""; $script:oldPref = ""; $script:newPref = ""
    $script:bufferDir = [System.IO.Path]::GetTempPath()
    $script:autoRule = ""; $script:autoSeg = ""
    $script:UpgradeSrc = ""; $script:UpgradeDst = ""
    $script:Rows = @()
}

# ---------------------------------------------------------------- pure helpers (ui_logic.ps1)
$paths = @('\\srv\p\Old\a.rvt', '\\srv\p\Old\b.rvt')
Eq (Get-AutoLocationMode $paths '\\srv\p\Old') "prefix" "every path starts with «было» -> prefix"
Eq (Get-AutoLocationMode $paths '\\srv\p\Old\') "prefix" "trailing slash of «было» is ignored"
Eq (Get-AutoLocationMode $paths 'Old') "segment" "a piece inside the path -> segment"
Eq (Get-AutoLocationMode $paths '\\srv\p\Ol') "segment" "no folder boundary -> segment, not a half-name prefix"
Eq (Get-AutoLocationMode $paths '') "same" "empty «было» -> nothing to move"
Eq (Get-AutoLocationMode @('RSN://h/P/a.rvt', '\\srv\p\Old\b.rvt') '\\srv\p\Old') "segment" "one path outside -> segment"

$st = Get-MoveFieldState -RuleOld "" -RuleNew "" -FolderOld "" -FolderNew "" -Paths $paths
Check (-not $st.Rename -and -not $st.Move) "all fields empty: nothing is done"
$st = Get-MoveFieldState -RuleOld "PROJ_IOS_EOM_ORG_K01_R24" -RuleNew "" -FolderOld '\\srv\p\Old' -FolderNew "" -Paths $paths
Check (-not $st.Rename -and -not $st.Move -and -not $st.RuleMessage -and -not $st.MoveMessage) "only «было» filled (the suggestions): nothing is done, no error"
$st = Get-MoveFieldState -RuleOld "PROJ_IOS_EOM_ORG_K01_R24" -RuleNew "PROJ_IOS_EOM_KLWT_K01_R24" -FolderOld "" -FolderNew "" -Paths $paths
Check ($st.Rename -and @($st.Ops).Count -eq 1 -and -not $st.Move) "rule filled: rename, no move"
$st = Get-MoveFieldState -RuleOld "" -RuleNew "X_R24" -FolderOld "" -FolderNew "" -Paths $paths
Check (-not $st.Rename -and $st.RuleMessage) "«стало» without an example: an error for the status line"
$st = Get-MoveFieldState -RuleOld "" -RuleNew "" -FolderOld "" -FolderNew '\\srv\new' -Paths $paths
Check (-not $st.Move -and $st.MoveMessage -match "было") "folder «стало» without «было»: an error"
$st = Get-MoveFieldState -RuleOld "" -RuleNew "" -FolderOld '\\srv\p\Old' -FolderNew '\\SRV\p\old\' -Paths $paths
Check (-not $st.Move -and $st.MoveMessage -match "совпадают") "the same folder twice is not a move"
$st = Get-MoveFieldState -RuleOld "" -RuleNew "" -FolderOld '\\srv\p\Old' -FolderNew '\\srv\p\New' -Paths $paths
Check ($st.Move -and $st.Mode -eq "prefix" -and $st.Was -eq '\\srv\p\Old' -and $st.Became -eq '\\srv\p\New') "both folder fields: move, prefix"

# ---------------------------------------------------------------- live preview in «Станет» / «Новый путь»
Reset-World
Update-MoveFields
Eq $script:Rows[0].new_name "PROJ_IOS_EOM_ORG_K01_R24" "empty fields: «Станет» = the name now"
Check (-not $script:Ui.CbRename.Checked -and -not $script:Ui.CbMove.Checked) "empty fields: no rename, no move in the job"
$script:Ui.RuleOld.Text = "PROJ_IOS_EOM_ORG_K01_R24"
$script:Ui.RuleNew.Text = "PROJ_IOS_EOM_KLWT_K01_R24"
Update-MoveFields
Check ($script:Ui.CbRename.Checked) "rule typed: rename is on"
Eq $script:Rows[0].new_name "PROJ_IOS_EOM_KLWT_K01_R24" "«Станет» follows the rule (example row)"
Eq $script:Rows[1].new_name "PROJ_IOS_EOM_KLWT_K03-K05-K07-K10-ST_R24" "«Станет» follows the rule (compound name)"
Eq $script:Rows[0].new_path '\\srv\p\Old\PROJ_IOS_EOM_KLWT_K01_R24.rvt' "no folder fields: the file stays in its folder"
$script:Ui.FolderOld.Text = '\\srv\p\Old'
$script:Ui.FolderNew.Text = '\\srv\p\New'
Update-MoveFields
Check ($script:Ui.CbMove.Checked -and $script:locMode -eq "prefix") "folders typed: move is on, prefix mode"
Eq $script:Rows[0].new_path '\\srv\p\New\PROJ_IOS_EOM_KLWT_K01_R24.rvt' "«Новый путь» follows the folder fields"
$script:Ui.RuleNew.Text = ""
Update-MoveFields
Check (-not $script:Ui.CbRename.Checked -and @($script:ops).Count -eq 0) "«стало» emptied: no rename"
Eq $script:Rows[0].new_path '\\srv\p\New\PROJ_IOS_EOM_ORG_K01_R24.rvt' "move without rename keeps the name"
$script:Ui.FolderNew.Text = ""
Update-MoveFields
Check (-not $script:Ui.CbMove.Checked -and $script:locMode -eq "same") "folder «стало» emptied: no move"
$script:Ui.RuleNew.Text = "X_R24"
$script:Ui.RuleOld.Text = ""
Update-MoveFields
Eq $script:Ui.Status.ForeColor ([System.Drawing.Color]::Firebrick) "a field error is shown in red in the status line"
Check ($script:Ui.Status.Text -match "^Имя:") "the error names the field"
$script:Ui.RuleNew.Text = ""
Update-MoveFields
Eq $script:Ui.Status.Text "" "fixed field: the red error goes away"

# hand edit of a path: it moves the model even with empty folder fields
Reset-World
$key = Get-ManualRowKey $script:modelPaths[0]
$script:manualRows[$key] = @{ name = $null; path = '\\srv\x\Hand_R24.rvt' }
Update-MoveFields
Eq $script:Rows[0].new_path '\\srv\x\Hand_R24.rvt' "a typed path is used"
Check ($script:Ui.CbMove.Checked) "a typed path switches the move on"

# old models: keep (default) or delete after the run
Reset-World
Check ($script:Ui.CbCopy.Checked -and -not $script:Ui.CbDelete.Checked) "old models are kept by default"
Check ($presetsText -match [regex]::Escape('$rbKeep = New-OpsRadio "оставить" $true')) "«оставить» is the default radio"
Check ($presetsText -match [regex]::Escape('$cbOverwrite = New-OpsCheck "Разрешить перезапись существующих файлов"')) "overwrite is a checkbox, off by default"

# suggestions: the first model and the common folder; typed text is never overwritten
Reset-World
Set-MoveSuggestions
Eq $script:Ui.RuleOld.Text "PROJ_IOS_EOM_ORG_K01_R24" "«Имя» example = the first model"
Eq $script:Ui.FolderOld.Text '\\srv\p\Old' "«Папка» было = the common folder"
$script:Ui.RuleOld.Text = "typed"
[void]$script:modelPaths.Insert(0, '\\srv\p\Old\ZZZ_R24.rvt')
Set-MoveSuggestions
Eq $script:Ui.RuleOld.Text "typed" "a typed example stays"

# job.json: the reports folder is passed to the RBP scripts
Reset-World
Update-MoveFields
$job = Build-JobObject $script:Rows
Eq $job.reports_dir (Join-Path $ToolDir "отчёты") "job.json carries reports_dir = <tool>\отчёты"
Check ($job.PSObject.Properties["relink_old"] -and $job.PSObject.Properties["copy"] -and $job.PSObject.Properties["delete_old"]) "job.json keeps its fields"

# ---------------------------------------------------------------- isolated runs keep the list of tab 1
Reset-World
$script:Ui.CbRename.Checked = $true
$script:ops = @([pscustomobject]@{ kind = "replace"; token = ""; from = "ORG"; to = "KLWT"; after = "" })
$script:locMode = "prefix"; $script:oldPref = '\\srv\p\Old'; $script:newPref = '\\srv\p\New'
$script:Seen = $null
function Start-UiRun {
    $script:Seen = @{
        Relink = $script:Ui.CbRelink.Checked; Rename = $script:Ui.CbRename.Checked; Move = $script:Ui.CbMove.Checked
        Copy = $script:Ui.CbCopy.Checked; Loc = $script:locMode; Ops = @($script:ops).Count; Paths = @($script:IsolatePaths); Kind = $script:IsolateKind
    }
    # the real Start-UiRun replaces the list with the run's models
    $script:modelPaths = New-Object System.Collections.Generic.List[string]
    [void]$script:modelPaths.Add('RSN://other/x.rvt')
}
function Get-TabModelPaths([string]$Key) { return @('RSN://h/P/r1.rvt', 'RSN://h/P/r2.rvt') }
Start-RelinkTabRun
Check ($script:Seen.Relink -and -not $script:Seen.Rename -and -not $script:Seen.Move -and -not $script:Seen.Copy) "relink tab: only relink is on during its run"
Check ($script:Seen.Loc -eq "same" -and $script:Seen.Ops -eq 0) "relink tab: no rule, no move of tab 1 leaks into its job"
Eq $script:Seen.Kind "relink" "relink tab: isolated kind"
Eq $script:modelPaths.Count 2 "after the relink run the list of tab 1 is back"
Eq $script:modelPaths[0] '\\srv\p\Old\PROJ_IOS_EOM_ORG_K01_R24.rvt' "tab 1 keeps its models"
Check ($script:Ui.CbRename.Checked -and @($script:ops).Count -eq 1 -and $script:locMode -eq "prefix" -and $script:oldPref -eq '\\srv\p\Old') "tab 1 settings are back"

# ---------------------------------------------------------------- «Смена версии Revit»: source year is detected
$servers = @([pscustomobject]@{ Host = "srv22"; Year = "2022" }, [pscustomobject]@{ Host = "srv24"; Year = "2024" })
$i = Get-UpgradeYearInfo @("RSN://srv22/P/a.rvt", "RSN://srv22/P/b.rvt") $servers
Check ($i.Ok -and $i.Year -eq "2022") "Revit Server models: year from servers.cfg"
$i = Get-UpgradeYearInfo @("RSN://srv22/P/a.rvt", "RSN://srv24/P/b.rvt") $servers
Check (-not $i.Ok -and $i.Message -match "разных лет") "two years in one run: stopped"
Check (@($i.Rows | Where-Object { $_.Reason }).Count -eq 2) "...and every row says why"
$i = Get-UpgradeYearInfo @("RSN://nobody/P/a.rvt", "RSN://srv22/P/b.rvt") $servers
Check (-not $i.Ok) "unknown server: stopped"
Check (@($i.Rows)[0].Reason -match "nobody" -and -not @($i.Rows)[1].Reason) "the reason is on the row of the unknown server only"
$i = Get-UpgradeYearInfo @('\\srv\p\A_R22.rvt') $servers
Check ($i.Ok -and $i.Year -eq "2022" -and @($i.Rows)[0].Source -eq "name") "network file: year from its name"
$i = Get-UpgradeYearInfo @('\\srv\p\A.rvt') $servers
Check ($i.Ok -and $i.Year -eq "") "network file without a year: fine, not needed"
Eq ((Get-UpgradeTargetYears @("2021", "2022", "2023", "2024", "2025") "2022") -join ",") "2023,2024,2025" "only newer versions are offered"
Eq ((Get-UpgradeTargetYears @("2021", "2022") "") -join ",") "2021,2022" "unknown source: all versions"
$upFn = (Find-Func "Start-UpgradeTabRun").Extent.Text
Check ($upFn -notmatch "CmbSrc" -and $upFn -match "Get-UpgradeYearInfo") "the upgrade run takes the source year from the models"
Check (-not $presetsText.Contains("CmbSrc") -and -not $presetsText.Contains("CmbDst")) "no source-version combo in the window"

# ---------------------------------------------------------------- dry run and the run button
$job = [pscustomobject]@{ relink = $false; rename = $true; location_mode = "prefix"; upgrade = $false; delete_old = $false; relink_old = ""; relink_new = "" }
$rows = @([pscustomobject]@{ old_path = '\\a\x.rvt'; new_path = '\\b\y.rvt' })
$t = Get-DryRunText $job $rows "1. Сохраняю новые хранилища"
Check ($t -match "ничего не запускается" -and $t -match [regex]::Escape('\\a\x.rvt') -and $t -match [regex]::Escape('-> \\b\y.rvt')) "dry run lists old -> new"
Check ($t -match "остаются на месте") "dry run says the old models stay"
Eq (Get-RunButtonText "move") "Запустить" "tab 1 button"
Eq (Get-RunButtonText "relink") "Запустить" "tab 2 button"
Eq (Get-RunButtonText "addlinks") "Вставить связи" "tab 4 button"
Eq (Get-RunButtonText "worksets") "Записать наборы" "tab 5 button"
$startText = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq "Start-UiRun" }, $true).Extent.Text
$atDry = $startText.IndexOf('$script:DryRunOnly')
Check ($atDry -gt 0 -and $atDry -lt $startText.IndexOf("Save-JobFiles") -and $atDry -lt $startText.IndexOf("Start-NextPass")) "«Пробный прогон» returns before anything is written or started"

# ---------------------------------------------------------------- the tab has no hints and no old controls
foreach ($gone in @("Подробности", "Переименовать""", "Перенести в другое место", "Открыть отчёт после прогона", "Подставлен прошлый список", "BtnReport")) {
    Check (-not $presetsText.Contains($gone)) ("removed: {0}" -f $gone)
}
Check ($presetsText -match 'New-OpsPairField \$ruleOld \$ruleNew' -and $presetsText -match 'New-OpsPairField \$folderOld \$folderNew') "«Имя» and «Папка» are было → стало rows on the tab"
Check ($presetsText -match 'FieldTimer\.Interval = 400') "typing is parsed after a short pause"
$formPart = $presetsText.Substring($presetsText.LastIndexOf('$form = New-Object System.Windows.Forms.Form'))
Check ($formPart -notmatch '\.Location = New-Object System\.Drawing\.Point') "no absolute control positions in the main window (grids only)"

$snapText = [System.IO.File]::ReadAllText((Join-Path $ToolSrc "snap_ui.ps1"))
$fitFn = [System.Management.Automation.Language.Parser]::ParseInput($snapText, [ref]$null, [ref]$null).Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq "Test-SnapTextFits" }, $true)
Check ($fitFn -and $fitFn.Extent.Text -match 'AutoEllipsis\) \{[\s\S]*?MeasureText') "snap_ui: a label cut with «…» is a violation, not skipped"

Write-Host ""
Write-Host ("Checks passed: {0}, failed: {1}" -f $script:Pass, $script:Fail)
if ($script:Fail -gt 0) { exit 1 }
exit 0
