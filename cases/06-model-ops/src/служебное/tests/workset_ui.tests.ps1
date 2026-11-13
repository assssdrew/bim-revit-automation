# Name: workset_ui.tests.ps1
# Version: 3.1.0
# What it does: See the case README.
# Notes: Paths and project names in this portfolio copy are placeholders.

# Window logic for worksets and links. No WinForms.
#   powershell -NoProfile -ExecutionPolicy Bypass -File tests\workset_ui.tests.ps1

# Windows PowerShell 5.1 does not load System.Drawing by itself (PowerShell 7 does);
# WinForms exists on Windows only, so a failed load is not an error here
foreach ($asm in @("System.Drawing", "System.Windows.Forms")) {
    try { Add-Type -AssemblyName $asm -ErrorAction Stop } catch { }
}
$ToolSrc = Split-Path -Parent $PSScriptRoot
. (Join-Path $ToolSrc "job_lib.ps1")
. (Join-Path $ToolSrc "ui_logic.ps1")
. (Join-Path $ToolSrc "workset_logic.ps1")
. (Join-Path $ToolSrc "workset_ui.ps1")

$script:Fail = 0
$script:Pass = 0

# a statement that throws is a failure, not a silently skipped check (seen as 207 vs 209 on 5.1)
trap {
    $script:Fail++
    Write-Host ("  ERROR: {0} (строка {1})" -f $_.Exception.Message, $_.InvocationInfo.ScriptLineNumber) -ForegroundColor Red
    continue
}
function Check([bool]$Cond, [string]$Msg) {
    if ($Cond) { $script:Pass++ }
    else { $script:Fail++; Write-Host "  FAIL: $Msg" -ForegroundColor Red }
}
function Eq($Got, $Want, [string]$Msg) { Check ($Got -ceq $Want) ("{0}: got '{1}', want '{2}'" -f $Msg, $Got, $Want) }

# ---------------------------------------------------------------- workset run does not touch the old job files
$ast = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $ToolSrc "presets.ps1"), [ref]$null, [ref]$null)
$start = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq "Start-UiRun" }, $true) | Select-Object -First 1
$src = $start.Extent.Text
# worksets have their own tabs: the tab-1 run never starts them and they never reach Save-JobFiles
Check ($src -notmatch "Start-WorksetApply") "the tab-1 run does not start the workset apply"
$dispatch = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq "Invoke-ActiveTabRun" }, $true) | Select-Object -First 1
Check ($null -ne $dispatch) "the bottom run button dispatches by tab"
Check ($dispatch.Extent.Text -match '"addlinks"[\s\S]*WsApplyPart = "insert"[\s\S]*Start-WorksetApply') "«Вставить связи» = apply with inserts only"
Check ($dispatch.Extent.Text -match '"worksets"[\s\S]*WsApplyPart = "sets"[\s\S]*Start-WorksetApply') "«Записать наборы» = apply with worksets only"
Check (-not (Test-UiWorksetsOn)) "no worksets switch on tab 1 any more"
$applyFn = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq "Start-WorksetApply" }, $true) | Select-Object -First 1
Check ($null -eq $applyFn) "Start-WorksetApply lives in workset_ui.ps1, not in the old run"
$uiAst = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path $ToolSrc "workset_ui.ps1"), [ref]$null, [ref]$null)
$applyFn = $uiAst.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq "Start-WorksetApply" }, $true) | Select-Object -First 1
Check ($null -ne $applyFn) "Start-WorksetApply is defined"
Check ($applyFn.Extent.Text -notmatch "Save-JobFiles") "the workset run does not write job.json or mapping.csv"
Check ($applyFn.Extent.Text -match "Show-OpsPlanDialog") "the workset run asks with the same dry-run dialog"
Check ($applyFn.Extent.Text -match 'Get-WorksetRenameCreatePlan \$worksets \$edits \$needed \$pairs') "the plan includes the link rows"
$uiText = [System.IO.File]::ReadAllText((Join-Path $ToolSrc "workset_ui.ps1"))
Check ($uiText.Contains("Открыть отчёт чтения…")) "the read tab can open an existing csv"
Check ($uiText.Contains("Resolve-WorksetReportFile")) "read and apply use the pointer resolver"
Check ($uiText -notmatch '\(Read-Utf8BomLines \$pointer\)\[0\]') "a one-line pointer is not indexed as a string"
$editor = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq "Show-CellEditor" }, $true) | Select-Object -First 1
Check ($null -ne $editor) "Show-CellEditor is still in presets.ps1"

$tmp = Join-Path ([System.IO.Path]::GetTempPath()) ("wsui_" + [Guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $tmp -Force | Out-Null
try {
    $job = [pscustomobject]@{
        saveas_needed = $false; relink = $true; unc_move = $false; delete_old = $false; copy = $false
        source_year = "2024"; target_year = ""
    }
    $row = [pscustomobject]@{ old_path = "RSN://revit-server-2024.example.local/proj/a.rvt"; needs_buffer = $false }
    $passes = @(New-OpsPassQueue -Job $job -ToolDir $tmp -CodeDir $tmp -Rows @($row))
    $names = @($passes | ForEach-Object { $_.Name })
    Check ($names -contains "relink") "checkbox off keeps the relink pass (queue unchanged)"
    Check ($names -notcontains "workset-read") "the old queue does not grow a workset pass"
}
finally { Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue }

$readPass = New-WorksetReadPass "C:\tool" "C:\tool\code" "2024"
Eq $readPass.Name "workset-read" "read pass name"
Eq $readPass.Worksets "close_all" "fast read keeps worksets closed"
$fullPass = New-WorksetReadPass "C:\tool" "C:\tool\code" "2024" "open_all"
Eq $fullPass.Worksets "open_all" "full read opens every workset"
Check (-not $readPass.Detach) "read pass does not detach"
Check ($readPass.Task.EndsWith("read_worksets.py")) "read pass task is read_worksets.py"
Check ($readPass.List.EndsWith("rvt_list_worksets.txt")) "read pass list is the workset list"
$hidden = Get-WorksetReadStatusNote "fast" 10 0
Check ($hidden.Contains("не видны") -and $hidden.Contains("не считает")) "fast read with no links says they may be hidden"
$seen = Get-WorksetReadStatusNote "fast" 10 3
Check ($seen.Contains("не считает") -and $seen -notmatch "не видны") "fast read that saw links still skips element counts"
Eq (Get-WorksetReadStatusNote "full" 10 0) "" "full read does not warn about closed worksets"
$uiBody = [System.IO.File]::ReadAllText((Join-Path $ToolSrc "workset_ui.ps1"))
Check ($uiBody.Contains("Полное чтение")) "the links tab has a full-read button"
Eq (Get-WorksetYearCaption ([pscustomobject]@{ Ok = $true; Year = "2022"; Message = "" })) "Revit 2022 (по моделям)" "the read label shows the detected year"
Eq (Get-WorksetYearCaption ([pscustomobject]@{ Ok = $false; Year = ""; Message = "В списке модели разных лет Revit. Один запуск — один год." })) "В списке модели разных лет Revit. Один запуск — один год." "mixed years stay a stop message"
$readFn = $uiAst.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq "Start-WorksetRead" }, $true) | Select-Object -First 1
Check ($null -ne $readFn) "Start-WorksetRead is defined"
Check ($readFn.Extent.Text -notmatch "WsYear") "the read pass does not take a year from a combo"
Check ($readFn.Extent.Text -match 'Resolve-WorksetReadYear \$paths \$script:servers ""') "the read pass detects the year from the models"
$installFn = $uiAst.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq "Install-WorksetUi" }, $true) | Select-Object -First 1
$presetsText = [System.IO.File]::ReadAllText((Join-Path $ToolSrc "presets.ps1"))
Check ($presetsText.Contains('$script:Ui.WsYearLabel = $linksPane.Year') -and $presetsText.Contains('$script:Ui.WsYearLabel2 = $setsPane.Year')) "tabs 4 and 5 show the year next to the model counter"
Check ($presetsText -match 'New-OpsBar @\(\$p\.All, \$p\.None, \$p\.Del\) @\(\$p\.Year, \$p\.Count\)') "the year is its own cell in the row under the list"
$gridText = [System.IO.File]::ReadAllText((Join-Path $ToolSrc "ui_grid.ps1"))
Check ($gridText -match 'function New-OpsInfoLabel[\s\S]*?AutoSize = \$true[\s\S]*?AutoEllipsis = \$false') "the year label sizes to its text and is never cut"
Check ($installFn.Extent.Text -notmatch "WsYearLabel") "no year label squeezed into the read-button row"
Check ($installFn.Extent.Text -notmatch "из servers.cfg") "the year combo is gone from the links tab"
Check ($installFn.Extent.Text -match "Полное чтение") "full read stays on the links tab"
Check ($installFn.Extent.Text -match "Открыть отчёт чтения") "open-report stays on the links tab"
$applyBody = $applyFn.Extent.Text
Check ($applyBody -match 'WsApplyPart -eq "insert"') "insert tab does not send renames"
Check ($applyBody -match 'WsApplyPart -eq "sets"') "workset tab does not send inserts"

# ---------------------------------------------------------------- roles, paint, rename preview
$script:WsRoles = @{}
Eq (Get-WorksetRoleLabel "RSN://h/a.rvt") "целевая" "default role is target"
Set-WorksetRole @("RSN://h/a.rvt") "new"
Eq (Get-WorksetRoleLabel "rsn://H/a.rvt") "новая" "role is stored by path key"
Set-WorksetRole @("RSN://h/a.rvt") "target"
Eq (Get-WorksetRoleCode "RSN://h/a.rvt") "target" "role can be set back"

$paint = Get-WorksetLinkPaint "no_pair" "" $false
Eq $paint.Row "missing" "a row without a pair is highlighted"
$paint = Get-WorksetLinkPaint "insert" "" $false
Eq $paint.Cell "empty" "an insert without a workset name is an empty cell"
$paint = Get-WorksetLinkPaint "insert" "00_Связь_ВК" $true
Eq $paint.Cell "manual" "a typed workset name is marked manual"
$paint = Get-WorksetLinkPaint "exists" "00_Связь_ВК" $false
Check ($paint.Dim) "an existing link is dimmed"
Check (-not $paint.Cell) "an existing link does not need a new name"
$paint = Get-WorksetLinkPaint "insert" "00_Связь_ВК" $false
Eq $paint.Row "" "a normal insert is not a missing row"
Eq $paint.Cell "" "a proposed name is not an empty cell"

$script:WsRenameEdit = @{}
$ren = [pscustomobject]@{ Name = "00_Связь_ОВ"; NewName = ""; Host = "" }
$n = Update-WorksetRenamePreview @($ren) "ов" "ВК"
Eq $n 1 "replace visits the row"
Eq $ren.Name "00_Связь_ОВ" "Было stays"
Eq $ren.NewName "00_Связь_ОВ" "replace is case-sensitive"
$null = Update-WorksetRenamePreview @($ren) "ОВ" "ВК"
Eq $ren.NewName "00_Связь_ВК" "ОВ is replaced"
Eq $ren.Name "00_Связь_ОВ" "Было still stays"

# ---------------------------------------------------------------- inventory and pairs
$csvRows = @(
    [pscustomobject]@{ row_kind = "workset"; host = "RSN://h/PROJ_IOS_AR_ORG_K01_R24.rvt"; workset_id = "9"; workset_name = "30_ИОС_Элементы" },
    [pscustomobject]@{ row_kind = "workset"; host = "RSN://h/PROJ_IOS_AR_ORG_K01_R24.rvt"; workset_id = "2"; workset_name = "00_Связь_ОВ" },
    [pscustomobject]@{ row_kind = "link"; host = "RSN://h/PROJ_IOS_AR_ORG_K01_R24.rvt"; link_name = "PROJ_IOS_OV_ORG_K01_R24"; link_path = "RSN://h/PROJ_IOS_OV_ORG_K01_R24.rvt"; instance_workset = "00_Связь_ОВ"; type_workset = "00_Связь_ОВ"; is_nested = "0" },
    [pscustomobject]@{ row_kind = "link"; host = "RSN://h/PROJ_IOS_AR_ORG_K01_R24.rvt"; link_name = "PROJ_IOS_SS_ORG_K01_R24"; link_path = "RSN://h/PROJ_IOS_SS_ORG_K01_R24.rvt"; instance_workset = "30_ИОС_Элементы"; type_workset = "30_ИОС_Элементы"; is_nested = "0" },
    [pscustomobject]@{ row_kind = "not_workshared"; host = "RSN://h/bare.rvt" }
)
$inv = ConvertTo-WorksetInventoryFromRows $csvRows
Eq $inv.WorksetCount 2 "both user worksets are kept, including the default"
Eq $inv.LinkCount 2 "both links are kept"
Check ($inv.NotWorkshared.ContainsKey((ConvertTo-WorksetPathKey "RSN://h/bare.rvt"))) "a non-workshared model is marked"
$names = @($inv.Worksets | ForEach-Object { $_.Name })
Check ($names -contains "30_ИОС_Элементы") "30_ИОС_Элементы is a normal user workset in the table"

$script:WsCodes = @(Get-DefaultSectionCodes)
$script:WsManual = @{}
$roles = @("fixed", "fixed", "section", "fixed", "building", "fixed")
$built = New-WorksetUiPairs @(
    "RSN://h/PROJ_IOS_VK_ORG_K01_R24.rvt"
) @(
    "RSN://h/PROJ_IOS_AR_ORG_K01_R24.rvt"
) "_" $roles $inv
Eq $built.Count 1 "one pair inside K01"
Eq @($built.Rows)[0].WorksetName "00_Связь_ВК" "the mask comes from 00_Связь_ОВ, not from 30_ИОС_Элементы"
Eq @($built.Rows)[0].State "insert" "the new VK link is not already there"

# ---------------------------------------------------------------- tab 4 on the real names: 1 target + 10 new
$siteDir = "RSN://revit-server-2022.example.local/proj SampleSite/Стадия П/"
$siteT = $siteDir + "04_ЭОМ/SITE_ORG_EOM_K1_R22.rvt"
$siteN = New-Object System.Collections.Generic.List[string]
foreach ($b in @("K1", "K2", "K3", "K4")) { [void]$siteN.Add(($siteDir + "06_HC/SITE_ORG_HC_{0}_R22.rvt" -f $b)) }
foreach ($b in @("K1", "K2", "K3", "K4", "K5", "PRK")) { [void]$siteN.Add(($siteDir + "07_AVT/SITE_ORG_AVT_{0}_R22.rvt" -f $b)) }
$sitePrk = $siteDir + "07_AVT/SITE_ORG_AVT_PRK_R22.rvt"
$siteInvRows = @(
    [pscustomobject]@{ row_kind = "workset"; host = $siteT; workset_id = "1"; workset_name = "01_Связь_RVT_AR_K1" },
    [pscustomobject]@{ row_kind = "workset"; host = $siteT; workset_id = "2"; workset_name = "01_Связь_RVT_AR_PRK" },
    [pscustomobject]@{ row_kind = "link"; host = $siteT; link_name = "SITE_ORG_AR_K1_R22"; link_path = ($siteDir + "01_AR/SITE_ORG_AR_K1_R22.rvt"); instance_workset = "01_Связь_RVT_AR_K1"; type_workset = "01_Связь_RVT_AR_K1"; is_nested = "0" },
    [pscustomobject]@{ row_kind = "link"; host = $siteT; link_name = "SITE_ORG_AR_PRK_R22"; link_path = ($siteDir + "01_AR/SITE_ORG_AR_PRK_R22.rvt"); instance_workset = "01_Связь_RVT_AR_PRK"; type_workset = "01_Связь_RVT_AR_PRK"; is_nested = "0" },
    [pscustomobject]@{ row_kind = "link"; host = $siteT; link_name = "SITE_ORG_TASK_PRK_R22"; link_path = ($siteDir + "00_TASK/SITE_ORG_TASK_PRK_R22.rvt"); instance_workset = "01_Связь_RVT_TASK_PRK"; type_workset = "01_Связь_RVT_TASK_PRK"; is_nested = "0" }
)
$script:WsInventory = ConvertTo-WorksetInventoryFromRows $siteInvRows
$script:WsCodes = @(Get-DefaultSectionCodes)
$script:WsManual = @{}
$script:WsManualPairs = New-Object System.Collections.Generic.List[object]
$script:WsRoles = @{}
Set-WorksetRole $siteN.ToArray() "new"
$fakeItems = New-Object System.Collections.Generic.List[object]
[void]$fakeItems.Add([pscustomobject]@{ Checked = $true; Tag = $siteT })
foreach ($p in $siteN) { [void]$fakeItems.Add([pscustomobject]@{ Checked = $true; Tag = $p }) }
$script:Ui = @{ PlainLists = @{ addlinks = [pscustomobject]@{ Items = $fakeItems.ToArray() } } }
$script:WsActiveKey = "addlinks"
$script:WsSeparator = "_"
$script:WsSavedRoles = @()
$script:WsUserRoles = @()
$script:WsMaskRoles = @()
Build-WorksetPairsFromUi
$rowsNow = @($script:WsPairResult.Rows)
$insNow = @($rowsNow | Where-Object { $_.State -eq "insert" })
Eq $insNow.Count 2 "real names: two inserts without pressing anything"
Eq @($rowsNow | Where-Object { $_.State -eq "no_pair" }).Count 8 "real names: 8 new models without a pair"
Eq $script:WsMaskResolved.Source "codes" "no saved mask: detected from the section codes"
Eq (@($insNow | ForEach-Object { $_.WorksetName }) -join ",") "01_Связь_RVT_HC_K1,01_Связь_RVT_AVT_K1" "worksets follow the host's 01_Связь_RVT_<раздел>_<корпус>"
Eq @($script:WsSavedRoles).Count 0 "the detection does not pretend to be a saved mask"
Check ($script:WsPairNote.Contains("вставить 2")) ("status after the recompute: " + $script:WsPairNote)
Eq (Get-WorksetStateLabel "unparsed") "имя не разобрано" "unparsed rows say «имя не разобрано»"

# no ticked new model: zero rows and one plain status line
$script:WsRoles = @{}
Build-WorksetPairsFromUi
Eq @($script:WsPairResult.Rows).Count 0 "all «целевая»: the «Связи» table is empty"
Eq $script:WsPairNote "Отметьте новые модели в колонке «Роль»." "status line asks for new models"
Set-WorksetRole $siteN.ToArray() "new"
# an extra target of a building without new models is not a red row
$siteK9 = $siteDir + "04_ЭОМ/SITE_ORG_EOM_K9_R22.rvt"
$script:Ui = @{ PlainLists = @{ addlinks = [pscustomobject]@{ Items = @($fakeItems.ToArray()) + @([pscustomobject]@{ Checked = $true; Tag = $siteK9 }) } } }
Build-WorksetPairsFromUi
Eq @(@($script:WsPairResult.Rows) | Where-Object { $_.TargetPath -eq $siteK9 }).Count 0 "an unpaired target is not listed"
Eq @(@($script:WsPairResult.Rows) | Where-Object { $_.State -eq "no_pair" }).Count 8 "«нет пары» only for the new models"
$script:Ui = @{ PlainLists = @{ addlinks = [pscustomobject]@{ Items = $fakeItems.ToArray() } } }

# a saved mask survives every recompute
$script:WsSavedRoles = @("fixed", "fixed", "section", "building", "fixed")
Build-WorksetPairsFromUi
Build-WorksetPairsFromUi
Eq ($script:WsSavedRoles -join ",") "fixed,fixed,section,building,fixed" "the saved mask is not wiped by a recompute"
Eq $script:WsMaskResolved.Source "saved" "the saved mask is the one in use"

# an unfinished edit on «Маска имени»: rows stay in the table as «имя не разобрано»
$script:WsUserRoles = @("fixed", "fixed", "section", "section", "fixed")
Build-WorksetPairsFromUi
$rowsBad = @($script:WsPairResult.Rows)
Check ($rowsBad.Count -gt 0) "the table is not empty when the mask does not fit"
Eq @($rowsBad | Where-Object { $_.State -eq "unparsed" }).Count $rowsBad.Count "every row says «имя не разобрано»"
Check ($script:WsPairNote.Contains("Маска имени")) "the status points to «Маска имени»"
Eq ($script:WsMaskRoles -join ",") "fixed,fixed,section,section,fixed" "the unfinished edit stays on screen"
$script:WsUserRoles = @()

# hand-made pair: AVT_PRK into EOM_K1, kept through recomputes, removable
Check (Add-WorksetManualPair $siteT $sitePrk) "a hand pair is added"
Check (-not (Add-WorksetManualPair $siteT $sitePrk)) "the same hand pair is not added twice"
Check (-not (Add-WorksetManualPair $siteT $siteT)) "a model is not linked into itself"
Build-WorksetPairsFromUi
$manRow = @(@($script:WsPairResult.Rows) | Where-Object { $_.Source -eq "manual" })
Eq $manRow.Count 1 "the hand row is in the table"
Eq $manRow[0].State "insert" "the hand row is an insert"
Eq $manRow[0].WorksetName "01_Связь_RVT_AVT_PRK" "the hand row gets a workset by the same pattern"
Eq @(@($script:WsPairResult.Rows) | Where-Object { $_.State -eq "no_pair" }).Count 7 "AVT_PRK is no longer without a pair"
$script:WsManual[[string]$manRow[0].Key] = "01_Связь_RVT_АВТ_Паркинг"
Build-WorksetPairsFromUi
$manRow = @(@($script:WsPairResult.Rows) | Where-Object { $_.Source -eq "manual" })
Eq $manRow.Count 1 "the hand row survives a recompute"
Eq $manRow[0].WorksetName "01_Связь_RVT_АВТ_Паркинг" "the typed workset of a hand row is kept"
$planIn = Get-WorksetRenameCreatePlan @($script:WsInventory.Worksets) @() @(Get-WorksetNeededNames @($script:WsPairResult.Rows)) @($script:WsPairResult.Rows)
Eq @(@($planIn.Rows) | Where-Object { $_.Action -eq "insert" -and $_.Arg1 -eq $sitePrk }).Count 1 "the hand row goes into the apply plan like an automatic one"
Eq @(@($planIn.Rows) | Where-Object { $_.Action -eq "insert" }).Count 3 "plan: HC_K1, AVT_K1 and the hand AVT_PRK"
Eq @(@($planIn.Rows) | Where-Object { $_.Host -like "*AVT_K2*" -or $_.Host -like "*HC_K2*" }).Count 0 "models without a pair are not opened by the apply"
$res = Remove-WorksetManualRows @([string]$manRow[0].Key, [string]$insNow[0].Key)
Eq $res.Removed 1 "«Удалить строку» removes the hand row"
Eq $res.Auto 1 "an automatic row is not removed"
Build-WorksetPairsFromUi
Eq @(@($script:WsPairResult.Rows) | Where-Object { $_.Source -eq "manual" }).Count 0 "the removed hand row stays removed"
Eq @(@($script:WsPairResult.Rows) | Where-Object { $_.State -eq "no_pair" }).Count 8 "AVT_PRK is back without a pair"

# the existing-link check applies to hand rows too
$siteInvRows2 = @($siteInvRows) + @([pscustomobject]@{ row_kind = "link"; host = $siteT; link_name = "SITE_ORG_AVT_PRK_R22"; link_path = $sitePrk; instance_workset = "01_Связь_RVT_AVT_PRK"; type_workset = "01_Связь_RVT_AVT_PRK"; is_nested = "0" })
$script:WsInventory = ConvertTo-WorksetInventoryFromRows $siteInvRows2
[void](Add-WorksetManualPair $siteT $sitePrk)
Build-WorksetPairsFromUi
Eq @(@($script:WsPairResult.Rows) | Where-Object { $_.Source -eq "manual" })[0].State "exists" "a hand row of a link already in the model is «есть»"
# a hand pair whose model left the list is not shown
$script:Ui = @{ PlainLists = @{ addlinks = [pscustomobject]@{ Items = @($fakeItems.ToArray() | Where-Object { $_.Tag -ne $sitePrk }) } } }
Build-WorksetPairsFromUi
Eq @(@($script:WsPairResult.Rows) | Where-Object { $_.Source -eq "manual" }).Count 0 "a hand pair is hidden while its model is not in the list"
$script:WsManualPairs = New-Object System.Collections.Generic.List[object]
$script:WsSavedRoles = @()
$script:WsRoles = @{}

# role drop-downs
Eq ((Get-WorksetRoleClickPaths @("a", "b", "c") "b") -join ",") "a,b,c" "a click inside the selection applies to all selected rows"
Eq ((Get-WorksetRoleClickPaths @("a", "b") "z") -join ",") "z" "a click outside the selection applies to that row"
Eq ((Get-WorksetRoleClickPaths @() "z") -join ",") "z" "no selection: that row"
Eq ((Get-WorksetModelRoleChoices) -join ",") "новая,целевая" "model role choices"
Eq ((Get-WorksetMaskRoleChoices) -join ",") "раздел,корпус,как есть,не важен" "mask role choices"
foreach ($w in (Get-WorksetMaskRoleChoices)) { Eq (Get-WorksetRoleWord (ConvertFrom-WorksetRoleWord $w)) $w ("role word round trip: " + $w) }

# ---------------------------------------------------------------- year and paths
$servers = @(
    [pscustomobject]@{ Host = "revit-server-2024.example.local"; Label = "a"; Year = "2024" },
    [pscustomobject]@{ Host = "revit-server-2022.example.local"; Label = "b"; Year = "2022" }
)
$y = Resolve-WorksetReadYear @("RSN://revit-server-2024.example.local/proj/a.rvt") $servers ""
Check $y.Ok "year from servers.cfg"
Eq $y.Year "2024" "host revit-server-2024.example.local is 2024"
$mix = Resolve-WorksetReadYear @("RSN://revit-server-2024.example.local/a.rvt", "RSN://revit-server-2022.example.local/b.rvt") $servers ""
Check (-not $mix.Ok) "mixed years stop the read"
Eq $mix.Year "" "no year is chosen for a mix"
$picked = Resolve-WorksetReadYear @("RSN://revit-server-2024.example.local/a.rvt", "RSN://revit-server-2022.example.local/b.rvt") $servers "2024"
Check $picked.Ok "an explicit year overrides the mix"
Eq $picked.Year "2024" "the chosen year is kept"
$gate = Test-WorksetReadPaths @("RSN://h/a.rvt", "\\srv\share\a.rvt")
Check $gate.Ok "RSN and UNC are accepted"
$gate = Test-WorksetReadPaths @("D:\Models\a.rvt")
Check (-not $gate.Ok) "a drive path is rejected"
$gate = Test-WorksetReadPaths @("BIM 360://proj/a.rvt")
Check (-not $gate.Ok) "BIM 360 is rejected"

# ---------------------------------------------------------------- tabs 4 and 5 on the common grid
$installText = $installFn.Extent.Text
foreach ($need in @("Прочитать наборы и связи", "Полное чтение", "Открыть отчёт чтения…", "Добавить связь вручную…", "Удалить строку", "Запомнить маску",
        "Коды разделов", "Маска имени", "Модель:", "Заменить в имени:", "Удалить выбранные", "вместе с элементами",
        "Рабочий набор сейчас", "Станет", "Элементов", "Действие", "Связи в моделях")) {
    Check ($installText.Contains($need)) ("tabs 4/5 have «{0}»" -f $need)
}
Check ($installText -notmatch "CbWorksets" -and $installText -notmatch "Рабочие наборы и связи") "the old worksets checkbox is gone"
foreach ($gone in @("Предложить пары", "Обновить предложения", "Предложить роли")) {
    Check (-not $uiText.Contains($gone)) ("the button «{0}» is gone: pairs and roles come by themselves" -f $gone)
}
Check ($installText -notmatch 'maskList\.Add_MouseDoubleClick') "mask roles are not cycled by double-click any more"
Check ($installText -match 'maskList\.Add_MouseClick[\s\S]*?Show-WorksetCellCombo') "mask roles: a drop-down in the cell"
Check ($installText -match 'addList\.Add_MouseDown[\s\S]*?WsRoleSelBefore') "model list remembers the selection before the click"
Check ($installText -match 'addList\.Add_MouseClick[\s\S]*?Get-WorksetRoleClickPaths[\s\S]*?Show-WorksetCellCombo') "model roles: a drop-down in the «Роль» cell for the selected rows"
Check ($installText -match 'WsRecalcTimer') "a debounce timer recomputes the pairs"
$nextFn = $uiAst.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq "Get-NextWorksetRole" }, $true) | Select-Object -First 1
Check ($null -eq $nextFn) "removed: Get-NextWorksetRole (double-click cycling)"
$countFn = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq "Update-PlainCount" }, $true) | Select-Object -First 1
Check ($countFn.Extent.Text -match "Request-WorksetRecompute") "a tick on tabs 4/5 re-proposes the pairs"
$importFn = $uiAst.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq "Import-WorksetReadFile" }, $true) | Select-Object -First 1
Check ($importFn.Extent.Text -match "Update-WorksetViews" -and $importFn.Extent.Text -match "WsPairNote") "a read / opened report proposes the pairs and says the result"
$buildFn = $uiAst.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq "Build-WorksetPairsFromUi" }, $true) | Select-Object -First 1
Check ($buildFn.Extent.Text -notmatch 'WsSavedRoles\s*=') "the recompute never writes the saved mask"
Check ($buildFn.Extent.Text -notmatch "Rows = @\(\)") "the recompute never returns a silently empty table"
Check ($installText -notmatch "Набор сейчас»|Location = New-Object System.Drawing.Point") "no absolute coordinates in the worksets panes"
Check ($installText.Contains('$linkList.Columns.Add("Рабочий набор", 200)')) "the «Связи» column is «Рабочий набор»"
Check (-not $installText.Contains('Columns.Add("Набор"')) "no column called just «Набор»"
Check ($installText -match 'linkList\.Add_MouseClick[\s\S]*?"link-workset"') "one click on «Рабочий набор» opens the edit box"
Check ($installText -match 'linkList\.Add_KeyDown[\s\S]*?F2[\s\S]*?"link-workset"') "F2 opens the edit box"
$formFn = $uiAst.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq "New-WorksetInsertForm" }, $true) | Select-Object -First 1
Check ($null -ne $formFn) "tab 4 has its own confirm form"
$formText = $formFn.Extent.Text
foreach ($need in @('"Вставить"', '"Отмена"', '"Закрыть"', '"Модель"', '"Что вставить"', '"Рабочий набор"', '"Действие"', '"Пропущено"', '"Причина"', "Get-WorksetInsertActionText", "Set-OpsColumnsToContent", "ToolTipText")) {
    Check ($formText.Contains($need)) ("confirm form has {0}" -f $need)
}
Check ($formText -match 'AcceptButton = \$btnYes') "«Вставить» is the AcceptButton"
Check ($formText -match 'CancelButton = \$btnNo') "«Отмена» is the CancelButton"
Check ($formText -match 'New-OpsBar @\(\) \$right' -and $formText -match 'New-OpsColumnGrid') "buttons sit in a bar row of the common grid"
Check ($formText -notmatch 'Location = New-Object|Dock = "Bottom"|\.SelectAll\(') "no absolute positions, no docked strip, nothing pre-selected"
Check ($formText -match 'MinimumSize' -and $formText -match 'WorkingArea') "the form has a minimum size and fits the screen"
Check ($formText -match 'ShowItemToolTips') "full paths are a tool tip"
$applyText = $applyFn.Extent.Text
Check ($applyText -match 'Show-WorksetInsertDialog \$insertView "Пробный прогон" \$true') "the dry run of tab 4 uses the same form with «Закрыть» only"
Check ($applyText -match 'Show-WorksetInsertDialog \$insertView \$title \$false') "«Вставить связи» asks with the new form"
Check ($applyText -match 'a model whose rows are all skips is not opened') "skipped-only models are not opened"
$planFn = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq "New-OpsPlanForm" }, $true) | Select-Object -First 1
Check ($null -ne $planFn) "the generic plan dialog has a builder"
Check ($planFn.Extent.Text -notmatch 'Location = New-Object System.Drawing.Point' -and $planFn.Extent.Text -match 'New-OpsBar') "plan dialog buttons are in a grid bar, not at absolute points (they slid off-screen)"
Check ($planFn.Extent.Text -match 'ActiveControl = \$btnNo') "plan dialog text is not pre-selected"
$snapText = [System.IO.File]::ReadAllText((Join-Path $ToolSrc "snap_ui.ps1"))
Check ($formText -notmatch 'рабочий набор новый') "the «Действие» value is short"
Check ($formText -match 'Register-OpsFillColumn \$lv 3 @\(\)' -and $formText -match 'Register-OpsFillColumn \$sk 2 @\(\)') "the last column takes the rest; the others fit their text"
Check ($snapText.Contains("function Test-SnapListView") -and $snapText.Contains("MeasureText") -and $snapText.Contains("Test-SnapTab4Lists")) "snap_ui measures the ListView cells"
Check ($snapText -match 'Test-SnapListView \$parts\.List' -and $snapText -match 'Test-SnapListView \$parts\.Skipped') "snap_ui checks the cells of the confirm forms"
$presetText = [System.IO.File]::ReadAllText((Join-Path $ToolSrc "presets.ps1"))
Check ($presetText -match 'function Set-UiStatus\(\[string\]\$Text, \[string\]\$Tip = ""\)' -and $presetText.Contains("StatusTip    = `$statusTip")) "the full status text is the tool tip of the status line"
Check ($snapText.Contains("04_svyazi_bez_novyh_min") -and $snapText.Contains("06_vstavka") -and $snapText.Contains("07_probnyi_progon") -and $snapText.Contains("1024x550")) "snap_ui shoots the empty tab 4 and the confirm forms"
$busyCall = [regex]::Match($installText, 'Add-OpsBusy @\(([^)]*)\)').Groups[1].Value
Check ($busyCall -and $busyCall -notmatch "btnOpenRead") "«Открыть отчёт чтения…» is never disabled"
Check ($uiText -match 'WsDeleteElements = \$rbWith') "«вместе с элементами» is the radio, «перенести» is the default"
foreach ($gone in @("Sync-WorksetLayout", "Update-WorksetTabBounds", "Sync-WorksetExclusive", "Mount-WorksetTopTabs")) {
    $f = $uiAst.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $gone }, $true) | Select-Object -First 1
    Check ($null -eq $f) ("removed: {0}" -f $gone)
}

# ---------------------------------------------------------------- tabs 4/5: year like tab 3
$srv = @([pscustomobject]@{ Host = "revit-server-2024.example.local"; Label = "a"; Year = "2024" })
$unc22 = @('\\server\share\Sample\АР\PROJ_K1_AP_R22.rvt', '\\server\share\Sample\КР\PROJ_K1_KP_R22.rvt')
$y = Resolve-WorksetReadYear $unc22 $srv ""
Check ($y.Ok -and $y.Year -eq "2022") "network files _R22: year 2022 from the name"
Eq (Get-WorksetYearCaption $y) "Revit 2022 (по моделям)" "network files _R22: caption"
$up = Get-UpgradeYearInfo $unc22 $srv
Eq $up.Year $y.Year "tabs 4/5 and tab 3 find the same year"
$y = Resolve-WorksetReadYear @('\\server\P\A_R22.rvt', 'RSN://revit-server-2024.example.local/P/B.rvt') $srv ""
Check (-not $y.Ok) "_R22 and a 2024 server: mixed"
Eq (Get-WorksetYearCaption $y) "Разные годы Revit в списке" "mixed: short caption"
Check ((Get-WorksetYearText $y) -match "2022, 2024") "mixed: the message box lists the years"
$y = Resolve-WorksetReadYear @('\\server\P\A_R22.rvt', '\\server\P\B.rvt') $srv ""
Check (-not $y.Ok) "a network file without _Rxx: the read cannot start Revit"
Eq (Get-WorksetYearCaption $y) "Год Revit не определён" "no year: the short caption"
Check (-not @($y.Rows)[0].Reason -and @($y.Rows)[1].Reason -match "_Rxx") "no year: the reason is on that row only"
Check ((Get-WorksetYearText $y) -match "B: ") "no year: the message box names the model"
$y = Resolve-WorksetReadYear @('RSN://unknown.example.local/P/C_R22.rvt') $srv ""
Check (-not $y.Ok -and @($y.Rows)[0].Reason -match "unknown.example.local") "unknown Revit Server: reason names the host"
Check ((Get-WorksetYearCaption $y).Length -le 30) "the caption is short enough for its cell"

function New-FakeRow([string]$Path, [bool]$Checked = $true) {
    $subs = @(
        [pscustomobject]@{ Text = "x"; Tag = $null; ForeColor = $null },
        [pscustomobject]@{ Text = $Path; Tag = $null; ForeColor = $null },
        [pscustomobject]@{ Text = ""; Tag = $null; ForeColor = $null }
    )
    return [pscustomobject]@{ Checked = $Checked; Tag = $Path; SubItems = $subs }
}
$script:servers = $srv
$l1 = [pscustomobject]@{ Text = ""; ForeColor = $null }
$l2 = [pscustomobject]@{ Text = ""; ForeColor = $null }
$rowsA = @((New-FakeRow $unc22[0]), (New-FakeRow $unc22[1]))
$rowsB = @((New-FakeRow '\\server\P\A_R22.rvt'), (New-FakeRow '\\server\P\NoYear.rvt'), (New-FakeRow '\\server\P\Off.rvt' $false))
$script:Ui = @{ WsYearLabel = $l1; WsYearLabel2 = $l2; PlainLists = @{ addlinks = [pscustomobject]@{ Items = $rowsA }; worksets = [pscustomobject]@{ Items = $rowsB } } }
Update-WorksetYearLabel
Eq $l1.Text "Revit 2022 (по моделям)" "tab 4 with PROJ_..._R22: «Revit 2022 (по моделям)»"
Eq $l2.Text "Год Revit не определён" "tab 5 has its own models and its own year"
Eq $l2.ForeColor ([System.Drawing.Color]::Firebrick) "a failed year is red"
Check ($rowsB[1].SubItems[2].Text -match "_Rxx" -and $rowsB[1].SubItems[2].Tag -eq "reason") "the reason is in «Состояние» of that row"
Eq $rowsB[0].SubItems[2].Text "" "a good row has no reason"
Eq $rowsB[2].SubItems[2].Text "" "an unticked row is not part of the run"
$rowsB[1].Checked = $false
Update-WorksetYearLabel
Eq $l2.Text "Revit 2022 (по моделям)" "untick the bad row: the year is found"
Eq $rowsB[1].SubItems[2].Text "" "...and its reason is cleared"

# opening a read report fills the model list of the active tab with its hosts
$tmp2 = Join-Path ([System.IO.Path]::GetTempPath()) ("wsui2_" + [Guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $tmp2 -Force | Out-Null
try {
    $csv = Join-Path $tmp2 "workset_read_x.csv"
    $lines = @(
        "row_kind;host;workset_id;workset_name;link_name;link_path;instance_workset;type_workset;is_nested;loaded",
        "workset;RSN://revit-server-2024.example.local/p/A_K01.rvt;1;00_Связь_ОВ;;;;;;",
        "workset;RSN://revit-server-2024.example.local/p/B_K01.rvt;1;00_Связь_ОВ;;;;;;",
        "workset;RSN://revit-server-2024.example.local/p/A_K01.rvt;2;30_ИОС_Элементы;;;;;;"
    )
    [System.IO.File]::WriteAllLines($csv, $lines, (New-Object System.Text.UTF8Encoding $true))
    $script:AddedTo = ""
    $script:AddedPaths = @()
    function Add-PlainTabModels([string]$Key, $Paths) { $script:AddedTo = $Key; $script:AddedPaths = @($Paths); return @($Paths).Count }
    function Set-UiStatus([string]$Text) { $script:LastStatus = $Text }
    function Update-WorksetViews { }
    $script:WsActiveKey = "worksets"
    Import-WorksetReadFile $csv ""
    Eq $script:AddedTo "worksets" "report hosts go to the active tab"
    Eq @($script:AddedPaths).Count 2 "each host once"
    Eq $script:WsInventory.WorksetCount 3 "inventory is read"
    Check ($script:LastStatus -match "добавлено в список: 2") "the status line says how many models came from the report"
    $hostsYear = Resolve-WorksetReadYear @($script:AddedPaths) $script:servers ""
    Check ($hostsYear.Ok -and $hostsYear.Year -eq "2024") "report hosts get the year the same way (servers.cfg)"
    $script:Ui = @{ WsYearLabel = $l1; PlainLists = @{ addlinks = [pscustomobject]@{ Items = @($script:AddedPaths | ForEach-Object { New-FakeRow $_ }) } } }
    Update-WorksetYearLabel
    Eq $l1.Text "Revit 2024 (по моделям)" "after «Открыть отчёт чтения…» the label shows the year of the report models"
}
finally { Remove-Item -LiteralPath $tmp2 -Recurse -Force -ErrorAction SilentlyContinue }

# the active tab decides whose ticked models are used
$script:Ui = @{ PlainLists = @{
        addlinks = [pscustomobject]@{ Items = @([pscustomobject]@{ Checked = $true; Tag = "RSN://h/a.rvt" }) }
        worksets = [pscustomobject]@{ Items = @([pscustomobject]@{ Checked = $true; Tag = "RSN://h/b.rvt" }, [pscustomobject]@{ Checked = $false; Tag = "RSN://h/c.rvt" }) }
    }
}
$script:WsActiveKey = "addlinks"
Eq ((Get-CheckedWorksetPaths) -join ",") "RSN://h/a.rvt" "tab 4 uses its own models"
$script:WsActiveKey = "worksets"
Eq ((Get-CheckedWorksetPaths) -join ",") "RSN://h/b.rvt" "tab 5 uses its own ticked models"

Write-Host ""
Write-Host ("Checks passed: {0}, failed: {1}" -f $script:Pass, $script:Fail)
if ($script:Fail -gt 0) { exit 1 }
exit 0
