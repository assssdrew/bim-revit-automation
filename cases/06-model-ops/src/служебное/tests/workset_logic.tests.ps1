# Name: workset_logic.tests.ps1
# Version: 3.1.0
# What it does: See the case README.
# Notes: Paths and project names in this portfolio copy are placeholders.

# Tests for workset_logic.ps1 (pure logic, no WinForms, no Revit).
# Windows PowerShell 5.1 or PowerShell 7:
#   powershell -NoProfile -ExecutionPolicy Bypass -File tests\workset_logic.tests.ps1

$ToolSrc = Split-Path -Parent $PSScriptRoot
. (Join-Path $ToolSrc "workset_logic.ps1")

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
function Rows($Result) {
    if ($null -eq $Result -or $null -eq $Result.Rows) { return @() }
    return @($Result.Rows)
}

$Roles = @("fixed", "fixed", "section", "fixed", "building", "fixed")
$Codes = @(Get-DefaultSectionCodes)

# ---------------------------------------------------------------- seed
Eq (Get-WorksetDefaultIgnoreName) "30_ИОС_Элементы" "default workset name is the active IOS workset"
$wantCodes = @(
    @("AR", "АР"), @("KR", "КР"), @("OV", "ОВ"), @("VK", "ВК"), @("EOM", "ЭОМ"), @("SS", "СС"),
    @("PT", "ПТ"), @("PD", "ПД"), @("HS", "ХС"), @("OT", "ОТ"), @("IOS", "ИОС"),
    @("ITP", "ИТП"), @("AVT", "АВТ"), @("HC", "ХЦ")
)
Eq $Codes.Count $wantCodes.Count "seed has one row per section code"
foreach ($w in $wantCodes) {
    $hit = $false
    foreach ($c in $Codes) { if ($c.Latin -ceq $w[0] -and $c.Cyrillic -ceq $w[1]) { $hit = $true } }
    Check $hit ("seed pair {0}={1}" -f $w[0], $w[1])
}
$vk = Find-SectionCode $Codes "vk"
Eq $vk.Script "latin" "VK matches the latin column"
Eq $vk.Cyrillic "ВК" "VK row carries ВК"
$ov = Find-SectionCode $Codes "ов"
Eq $ov.Script "cyrillic" "ОВ matches the cyrillic column"
Eq $ov.Latin "OV" "ОВ row carries OV"
Check ($null -eq (Find-SectionCode $Codes "TASK")) "TASK is not a section code"

# ---------------------------------------------------------------- paths and year
Eq (Test-WorksetStoragePath "RSN://revit-server-2024.example.local/proj/a.rvt").Kind "rsn" "RSN is accepted"
Check (Test-WorksetStoragePath "rsn://h/a.rvt").Ok "RSN scheme is case-insensitive"
Eq (Test-WorksetStoragePath "\\srv\share\a.rvt").Kind "unc" "UNC is accepted"
Eq (Test-WorksetStoragePath "BIM 360://proj/a.rvt").Kind "cloud" "BIM 360 is rejected"
Eq (Test-WorksetStoragePath "ACC://proj/a.rvt").Kind "cloud" "ACC is rejected"
Eq (Test-WorksetStoragePath "Autodesk Docs://proj/a.rvt").Kind "cloud" "Autodesk Docs is rejected"
Eq (Test-WorksetStoragePath "D:\Models\a.rvt").Kind "local" "a drive path is rejected"
Eq (Test-WorksetStoragePath "  ").Kind "empty" "blank path is empty"
$oneYear = Test-WorksetSingleYear @("2024", " 2024 ", "")
Check $oneYear.Ok "one year is accepted"
Eq $oneYear.Year "2024" "the year is 2024"
Eq (Test-WorksetSingleYear @("2024", "2025")).Code "mixed_year" "two years stop the run"
Eq (Test-WorksetSingleYear @("", " ")).Code "no_year" "no year stops the run"

# ---------------------------------------------------------------- mask suggestion
$news = @(
    "PROJ_IOS_VK_ORG_K01_R24",
    "PROJ_IOS_VK_ORG_K02_R24",
    "PROJ_IOS_VK_ORG_P01_R24"
)
$targets = @(
    "PROJ_IOS_AR_ORG_K01_R24",
    "PROJ_IOS_OV_ORG_K01_R24",
    "PROJ_IOS_AR_ORG_K02_R24",
    "PROJ_IOS_KR_ORG_P01_R24",
    "PROJ_IOS_AR_ORG_K03_R24"
)
$sug = Get-WorksetNameMaskSuggestion $news $targets "_"
Check $sug.Ok "two buildings and one new section: suggestion applies"
Eq ($sug.Roles -join ",") "fixed,fixed,section,fixed,building,fixed" "roles are project, IOS, section, ORG, building, year"
Eq (Get-WorksetNameMaskSuggestion @("PROJ_IOS_VK_ORG_K01_R24") $targets).Reason "ambiguous" "a single new file does not reveal the building"
Eq (Get-WorksetNameMaskSuggestion @() $targets).Reason "empty" "no new names"
$mix = @($news) + @("PROJ_IOS_VK_ORG_K01_EXTRA_R24")
Eq (Get-WorksetNameMaskSuggestion $mix $targets).Reason "length" "a different token count is not a mask"
$igNew = @("A_VK_K01_Q1_W1", "A_VK_K02_Q2_W2")
$igTgt = @("A_AR_K01_T1_U1", "A_OV_K02_T2_U2")
$ig = Get-WorksetNameMaskSuggestion $igNew $igTgt "_"
Eq $ig.Reason "many_ignore" "two leftover positions are not applied"
Check (-not $ig.Apply) "many_ignore does not apply"
Eq $ig.RoleCount 5 "the guess is still returned for the user to edit"

# ---------------------------------------------------------------- pairs
$nk01 = "RSN://srv/new/PROJ_IOS_VK_ORG_K01_R24.rvt"
$nk02 = "RSN://srv/new/PROJ_IOS_VK_ORG_K02_R24.rvt"
$np01 = "RSN://srv/new/PROJ_IOS_VK_ORG_P01_R24.rvt"
$nbad = "RSN://srv/new/PROJ_IOS_VK_ORG_K01_EXTRA_R24.rvt"
$tar = "RSN://srv/old/PROJ_IOS_AR_ORG_K01_R24.rvt"
$tov = "RSN://srv/old/proj_ios_OV_org_K01_r24.rvt"
$tk02 = "RSN://srv/old/PROJ_IOS_AR_ORG_K02_R24.rvt"
$tp = "RSN://srv/old/PROJ_IOS_KR_ORG_P01_R24.rvt"
$tk03 = "RSN://srv/old/PROJ_IOS_AR_ORG_K03_R24.rvt"
$pairs = Get-WorksetLinkPairs @($nk01, $nk02, $np01, $nbad) @($tar, $tov, $tk02, $tp, $tk03) "_" $Roles
Eq $pairs.Count 6 "4 inserts + unparsed + K03 without a partner"
$inserts = @(Rows $pairs | Where-Object { $_.State -eq "insert" })
Eq $inserts.Count 4 "four real pairs"
$k01 = @($inserts | Where-Object { $_.LinkPath -eq $nk01 })
Eq $k01.Count 2 "VK K01 goes into both K01 models"
Eq @($k01 | Where-Object { $_.TargetPath -eq $tk02 }).Count 0 "VK K01 does not go into K02"
Eq @($inserts | Where-Object { $_.LinkPath -eq $np01 -and $_.TargetPath -eq $tp }).Count 1 "P01 pairs only with P01"
Eq @($inserts | Where-Object { $_.LinkPath -eq $np01 -and $_.TargetPath -eq $tar }).Count 0 "P01 does not pair with K01"
Eq @($inserts | Where-Object { $_.TargetPath -eq $tov }).Count 1 "fixed tokens match without case"
$self = Get-WorksetLinkPairs @($tar, $nk01) @($tar) "_" $Roles
Eq @((Rows $self) | Where-Object { $_.LinkPath -eq $tar -and $_.TargetPath -eq $tar }).Count 0 "a model is not paired with itself"
Eq @((Rows $self) | Where-Object { $_.State -eq "insert" -and $_.LinkPath -eq $nk01 }).Count 1 "the other new model still pairs"
$unp = @(Rows $pairs | Where-Object { $_.LinkPath -eq $nbad })
Eq $unp.Count 1 "extra token is one row"
Eq $unp[0].State "unparsed" "extra token is unparsed"
$k03 = @(Rows $pairs | Where-Object { $_.TargetPath -eq $tk03 })
Eq $k03.Count 1 "K03 stays in the table"
Eq $k03[0].State "no_pair" "K03 has no new model of that building"
Eq $k03[0].LinkPath "" "the unpaired target has no link"

# ---------------------------------------------------------------- inventory: full path, other path, nested
$mini = Get-WorksetLinkPairs @($nk01) @($tar) "_" $Roles
Eq (Rows $mini)[0].State "insert" "before inventory the pair is an insert"
$links = @(
    [pscustomobject]@{ Host = $tar; Path = "rsn://srv/new/PROJ_IOS_VK_ORG_K01_R24.rvt"; Name = ""; Nested = $false },
    [pscustomobject]@{ Host = $tov; Path = "RSN://srv/elsewhere/PROJ_IOS_VK_ORG_K01_R24.rvt"; Name = ""; Nested = $false },
    [pscustomobject]@{ Host = $tar; Path = $nk01; Name = ""; Nested = $true }
)
# dedicated rows so the nested copy on $tar does not hide the top-level match
$top = Get-WorksetLinkPairs @($nk01) @($tar) "_" $Roles
Update-WorksetPairInventory (Rows $top) @($links[0])
Eq (Rows $top)[0].State "exists" "the same file with a different slash case is already linked"
$other = Get-WorksetLinkPairs @($nk01) @($tov) "_" $Roles
Update-WorksetPairInventory (Rows $other) @($links[1])
Eq (Rows $other)[0].State "other_path" "the same file name under another path is not inserted"
$nest = Get-WorksetLinkPairs @($nk01) @($tar) "_" $Roles
Update-WorksetPairInventory (Rows $nest) @($links[2])
Eq (Rows $nest)[0].State "insert" "a nested link does not count as already inserted"

# ---------------------------------------------------------------- workset mask
function Link([string]$Name, [string]$Workset, [bool]$Nested) {
    return [pscustomobject]@{ Name = $Name; Workset = $Workset; Nested = $Nested }
}
$good = @(
    (Link "PROJ_IOS_OV_ORG_K01_R24" "00_Связь_ОВ" $false),
    (Link "PROJ_IOS_EOM_ORG_K01_R24" "00_Связь_ЭОМ" $false),
    (Link "PROJ_IOS_SS_ORG_K01_R24" "30_ИОС_Элементы" $false),
    (Link "PROJ_IOS_VK_ORG_K01_R24" "99_ВК" $true)
)
$mask = Get-WorksetLinkMask $good $Codes "_" $Roles $null
Check $mask.Ok "agreeing link worksets produce a mask"
Eq $mask.Mask "00_Связь_{section}" "the section token is the only placeholder"
Eq $mask.Script "cyrillic" "the model writes sections in Cyrillic"
Eq $mask.Votes 2 "the default workset and the nested link do not vote"
Eq (Get-ProposedWorksetName $mask.Mask $mask.Script "VK" $Codes) "00_Связь_ВК" "VK is written as ВК"
Eq (Get-ProposedWorksetName $mask.Mask $mask.Script "vk" $Codes) "00_Связь_ВК" "the spelling comes from the table, not from the letter case"
Eq (Get-ProposedWorksetName $mask.Mask $mask.Script "ZZ" $Codes) "" "an unknown section leaves the name empty"
$onlyDefault = @(
    (Link "PROJ_IOS_IOS_ORG_K01_R24" "30_ИОС_Элементы" $false),
    (Link "PROJ_IOS_OV_ORG_K01_R24" "30_иос_элементы" $false)
)
Eq (Get-WorksetLinkMask $onlyDefault $Codes "_" $Roles $null).Reason "no_votes" "30_ИОС_Элементы never becomes a link mask"
$ovik = @((Link "PROJ_IOS_OV_ORG_K01_R24" "00_Связь_ОВиК" $false))
Eq (Get-WorksetLinkMask $ovik $Codes "_" $Roles $null).Reason "no_votes" "ОВ inside ОВиК is not the section token"
$split = @(
    (Link "PROJ_IOS_OV_ORG_K01_R24" "00_Связь_ОВ" $false),
    (Link "PROJ_IOS_VK_ORG_K01_R24" "01_Связь_ВК" $false)
)
Eq (Get-WorksetLinkMask $split $Codes "_" $Roles $null).Reason "disagree" "two prefixes do not invent a mask"
$latin = @(
    (Link "PROJ_IOS_OV_ORG_K01_R24" "00_Link_OV" $false),
    (Link "PROJ_IOS_EOM_ORG_K01_R24" "00_Link_EOM" $false)
)
$lat = Get-WorksetLinkMask $latin $Codes "_" $Roles $null
Eq $lat.Script "latin" "a latin model keeps the latin script"
Eq (Get-ProposedWorksetName $lat.Mask $lat.Script "VK" $Codes) "00_Link_VK" "VK stays VK"

# ---------------------------------------------------------------- launch gate
function PairState([string]$State, [string]$Workset) {
    return [pscustomobject]@{ State = $State; WorksetName = $Workset }
}
$blocked = Test-WorksetLaunch @(
    (PairState "insert" ""),
    (PairState "insert" "00_Связь_ВК"),
    (PairState "no_pair" ""),
    (PairState "exists" ""),
    (PairState "other_path" ""),
    (PairState "unparsed" "")
)
Check $blocked.Block "an empty workset on an insert blocks the launch"
Eq $blocked.Empty 1 "the empty insert is the only empty cell"
Eq $blocked.Invalid 0 "the other rows are not invalid names"
$badName = Test-WorksetLaunch @((PairState "insert" "00_Связь:ВК"))
Check $badName.Block "an illegal workset name blocks the launch"
Eq $badName.Invalid 1 "one illegal name"
Check (-not (Test-WorksetLaunch @((PairState "no_pair" ""), (PairState "other_path" ""), (PairState "exists" ""))).Block) "skips do not block"
Check (Test-WorksetNameText "00_Связь_ВК").Ok "a normal link workset name is legal"
Eq (Test-WorksetNameText ("x" * 256)).Code "long" "256 characters is too long"
Eq (Test-WorksetNameText "a{b").Code "chars" "a brace is illegal"

# ---------------------------------------------------------------- rename table, find/replace, conflicts
$ws = New-Object System.Collections.Generic.List[object]
for ($i = 1; $i -le 12; $i++) {
    $modelPath = "RSN://srv/m{0:d2}.rvt" -f $i
    [void]$ws.Add([pscustomobject]@{ Host = $modelPath; Id = ("ov-{0}" -f $i); Name = "00_Связь_ОВ" })
    [void]$ws.Add([pscustomobject]@{ Host = $modelPath; Id = ("def-{0}" -f $i); Name = "30_ИОС_Элементы" })
    if ($i -le 2) { [void]$ws.Add([pscustomobject]@{ Host = $modelPath; Id = ("w1-{0}" -f $i); Name = "Workset1" }) }
}
$all = Group-WorksetRename $ws.ToArray() ""
$names = @((Rows $all) | ForEach-Object { $_.Name })
Check ($names -contains "Workset1") "Workset1 is listed"
Check ($names -contains "30_ИОС_Элементы") "the default workset is listed for rename"
Check ($names -contains "00_Связь_ОВ") "the link workset is listed"
Eq @(Rows $all | Where-Object { $_.Name -eq "00_Связь_ОВ" })[0].HostCount 12 "the link workset is grouped across 12 models"
$one = Group-WorksetRename $ws.ToArray() "RSN://srv/m01.rvt"
Eq $one.Count 3 "filter by one model shows each of its worksets"
Eq @((Rows $one) | Where-Object { $_.Name -eq "Workset1" }).Count 1 "the filtered model includes Workset1"
Eq (Rows $one)[0].HostCount 1 "a filtered row is that one model"
Eq @((Rows $one) | Where-Object { $_.Host -eq "RSN://srv/m02.rvt" }).Count 0 "other models stay out of the filter"
$was = "00_Связь_ОВ"
$became = ConvertTo-WorksetReplacedName $was "ОВ" "ВК"
Eq $became "00_Связь_ВК" "find/replace substitutes the fragment"
Eq $was "00_Связь_ОВ" "Было is left untouched"
Eq (ConvertTo-WorksetReplacedName $was "ов" "ВК") $was "find/replace is case-sensitive"
Eq (ConvertTo-WorksetReplacedName $was "ов" "ВК" $true) "00_Связь_ВК" "IgnoreCase is explicit"
Eq (ConvertTo-WorksetReplacedName "a.b" "." "X") "aXb" "the search text is literal, not a regex"
$ed = @([pscustomobject]@{ OldName = "00_Связь_ОВ"; NewName = "Workset1" })
$cf = Get-WorksetRenameConflicts $ws.ToArray() $ed
Eq $cf.Count 1 "one rename edit"
Eq (Rows $cf)[0].HostCount 12 "the workset exists in 12 models"
Eq (Rows $cf)[0].ConflictCount 2 "two models already have the new name"
Eq (Rows $cf)[0].OkCount 10 "the other ten can be renamed"
$caseOnly = Get-WorksetRenameConflicts $ws.ToArray() @([pscustomobject]@{ OldName = "00_Связь_ОВ"; NewName = "00_связь_ов" })
Eq $caseOnly.Count 0 "a case-only change is not a second workset"

# ---------------------------------------------------------------- apply plan and dry-run
$maskRow = [pscustomobject]@{ Host = $tar; Ok = $mask.Ok; Mask = $mask.Mask; Script = $mask.Script }
$planPairs = Get-WorksetLinkPairs @($nk01, "RSN://srv/new/PROJ_IOS_SS_ORG_K01_R24.rvt") @($tar) "_" $Roles
Add-WorksetProposals (Rows $planPairs) @($maskRow) $Codes
Eq (Rows $planPairs)[0].WorksetName "00_Связь_ВК" "the proposal is filled in"
Eq (Rows $planPairs)[1].WorksetName "00_Связь_СС" "SS follows the same mask"
# both rows of this host ask for the same name: the workset is created once
(Rows $planPairs)[1].WorksetName = "00_Связь_ВК"
$inv = @(
    [pscustomobject]@{ Host = $tar; Id = "1"; Name = "00_Связь_ОВ" },
    [pscustomobject]@{ Host = $tar; Id = "2"; Name = "30_ИОС_Элементы" },
    [pscustomobject]@{ Host = $tar; Id = "3"; Name = "Workset1" }
)
$act = Get-WorksetApplyActions (Rows $planPairs) $inv @()
$insAct = @(Rows $act | Where-Object { $_.Action -eq "insert" })
Eq $insAct.Count 2 "two inserts"
Check $insAct[0].CreateWorkset "the first insert creates the missing workset"
Check (-not $insAct[1].CreateWorkset) "the second insert reuses it"
# existing name is not created again
(Rows $planPairs)[0].WorksetName = "00_Связь_ОВ"
(Rows $planPairs)[1].WorksetName = "00_Связь_ОВ"
$act2 = Get-WorksetApplyActions (Rows $planPairs) $inv @()
Check (-not (Rows $act2 | Where-Object { $_.Action -eq "insert" })[0].CreateWorkset) "an existing workset is reused"
# other path is a skip, not an insert, and does not block
$skipPairs = Get-WorksetLinkPairs @($nk01) @($tov) "_" $Roles
Update-WorksetPairInventory (Rows $skipPairs) @($links[1])
$skipAct = Get-WorksetApplyActions (Rows $skipPairs) @() @()
Eq (Rows $skipAct)[0].Action "skip" "other path is a skip"
Eq (Rows $skipAct)[0].Code "SKIP_OTHER_PATH" "the report code names the other path"
Check (-not (Test-WorksetLaunch (Rows $skipPairs)).Block) "other path does not block"
$renAct = Get-WorksetApplyActions @() $ws.ToArray() (Rows $cf)
$renOk = @(Rows $renAct | Where-Object { $_.Action -eq "rename" })
$renBad = @(Rows $renAct | Where-Object { $_.Code -eq "SKIP_CONFLICT" })
Eq $renOk.Count 10 "ten renames are planned"
Eq $renBad.Count 2 "two conflicts are skipped"
Eq @($renOk | Where-Object { $_.Key -eq "ov-3" }).Count 1 "the rename keeps the workset id"
$bag = New-Object System.Collections.Generic.List[object]
foreach ($a in (Rows $renAct)) { [void]$bag.Add($a) }
foreach ($a in (Rows $skipAct)) { [void]$bag.Add($a) }
foreach ($a in (Rows $act)) { [void]$bag.Add($a) }
$dry = Get-WorksetDryRunText $bag.ToArray()
Check ($dry.Contains("Переименовать:")) "dry-run lists renames"
Check ($dry.Contains("00_Связь_ОВ → Workset1")) "dry-run shows было → станет"
Check ($dry.Contains("создать набор")) "dry-run shows the workset that will be created"
Check ($dry.Contains("SKIP_OTHER_PATH")) "dry-run shows the other-path skip"
Check ($dry.Contains("SKIP_CONFLICT")) "dry-run shows the rename conflict"

# ---------------------------------------------------------------- step 4 plan: rename and create, no insert
$h1 = "RSN://srv/m01.rvt"
$h2 = "RSN://srv/m02.rvt"
$sets = @(
    [pscustomobject]@{ Host = $h1; Id = "ov-1"; Name = "00_Связь_ОВ" },
    [pscustomobject]@{ Host = $h1; Id = "w1"; Name = "Workset1" },
    [pscustomobject]@{ Host = $h2; Id = "ov-2"; Name = "00_Связь_ОВ" }
)
$edits = @(
    [pscustomobject]@{ Host = $h1; Id = "ov-1"; OldName = "00_Связь_ОВ"; NewName = "Workset1" },
    [pscustomobject]@{ Host = $h2; Id = "ov-2"; OldName = "00_Связь_ОВ"; NewName = "00_Связь_ВК" }
)
$need = @(
    [pscustomobject]@{ Host = $h2; Name = "00_Связь_СС" },
    [pscustomobject]@{ Host = $h2; Name = "00_Связь_ВК" },
    [pscustomobject]@{ Host = $h1; Name = "00_Связь_СС" }
)
$step4 = Get-WorksetRenameCreatePlan $sets $edits $need
$s4 = @(Rows $step4)
Eq @($s4 | Where-Object { $_.Action -eq "insert" }).Count 0 "step 4 does not emit insert rows"
Eq @($s4 | Where-Object { $_.Action -eq "rename" -and $_.Key -eq "ov-2" }).Count 1 "the free rename keeps the workset id"
Eq @($s4 | Where-Object { $_.Action -eq "rename" -and $_.Key -eq "ov-2" })[0].Arg2 "00_Связь_ВК" "rename target is the new name"
Eq @($s4 | Where-Object { $_.Code -eq "SKIP_CONFLICT" -and $_.Key -eq "ov-1" }).Count 1 "a taken name is a conflict on that model only"
Eq @($s4 | Where-Object { $_.Action -eq "create" -and $_.Host -eq $h2 -and $_.Arg1 -eq "00_Связь_СС" }).Count 1 "a missing name is created"
Eq @($s4 | Where-Object { $_.Action -eq "create" -and $_.Arg1 -eq "00_Связь_ВК" }).Count 0 "a name produced by rename is not created again"
Eq @($s4 | Where-Object { $_.Action -eq "create" -and $_.Host -eq $h1 }).Count 1 "the other model still gets the missing name"
$order = @($s4 | ForEach-Object { $_.Action })
$renAt = [array]::IndexOf($order, "rename")
$makeAt = [array]::IndexOf($order, "create")
Check ($renAt -ge 0 -and $makeAt -gt $renAt) "renames come before creates"
$txt = Get-WorksetStep4Text $s4
Check ($txt.Contains("$h2 — 00_Связь_ОВ — 00_Связь_ВК")) "dry-run lists model, old name, new name"
Check ($txt.Contains("$h2 — создать набор 00_Связь_СС")) "dry-run lists the workset to create"
Check ($txt.Contains("SKIP_CONFLICT")) "dry-run lists the conflict"
Check ($txt.Contains("Вставка связей — следующий шаг.")) "dry-run says link insert is later"
$emptyIns = @([pscustomobject]@{ State = "insert"; WorksetName = "" })
$blk = Test-WorksetStep4Block $emptyIns @()
Check (-not $blk.Block) "an insert without a workset is skipped and listed, it does not stop the others"
Eq $blk.Empty 1 "the empty cell is counted"
$badEdit = @([pscustomobject]@{ OldName = "00_Связь_ОВ"; NewName = "a:b" })
$blk2 = Test-WorksetStep4Block @() $badEdit
Check $blk2.Block "an illegal rename blocks before the dialog"
Eq $blk2.BadRename 1 "one bad rename"
$caseEdit = @([pscustomobject]@{ Host = $h2; Id = "ov-2"; OldName = "00_Связь_ОВ"; NewName = "00_связь_ов" })
$casePlan = Get-WorksetRenameCreatePlan @($sets[2]) $caseEdit @()
Eq $casePlan.ApplyCount 0 "a case-only rename is not sent"
$inserts = @(
    [pscustomobject]@{ State = "insert"; TargetPath = $h2; LinkPath = "RSN://srv/new/PROJ_IOS_VK_ORG_K02_R24.rvt"; WorksetName = "00_Связь_ВК"; Key = "row-1"; Section = "VK" },
    [pscustomobject]@{ State = "exists"; TargetPath = $h2; LinkPath = "RSN://srv/old/PROJ_IOS_OV_ORG_K02_R24.rvt"; WorksetName = "00_Связь_ОВ"; Key = "row-2"; Section = "OV" },
    [pscustomobject]@{ State = "other_path"; TargetPath = $h1; LinkPath = "\\srv\other\PROJ_IOS_SS_ORG_K01_R24.rvt"; WorksetName = "00_Связь_СС"; Key = "row-3"; Section = "SS" }
)
$withLinks = Get-WorksetRenameCreatePlan $sets @() @(
    [pscustomobject]@{ Host = $h2; Name = "00_Связь_ВК" }
) $inserts
$linkRows = @(Rows $withLinks)
Eq @($linkRows | Where-Object { $_.Action -eq "insert" }).Count 1 "an insert row is planned for a new link"
Eq @($linkRows | Where-Object { $_.Action -eq "insert" })[0].Arg2 "00_Связь_ВК" "the insert names the target workset"
Eq @($linkRows | Where-Object { $_.Code -eq "SKIP_EXISTS" }).Count 1 "the same path is SKIP_EXISTS"
Eq @($linkRows | Where-Object { $_.Code -eq "SKIP_OTHER_PATH" }).Count 1 "the same file on another path is SKIP_OTHER_PATH"
$linkOrder = @($linkRows | ForEach-Object { $_.Action })
$createAt = [array]::IndexOf($linkOrder, "create")
$insertAt = [array]::IndexOf($linkOrder, "insert")
Check ($createAt -ge 0 -and $insertAt -gt $createAt) "create comes before insert"
$linkText = Get-WorksetStep4Text $linkRows
Check ($linkText.Contains("$h2 — вставить связь RSN://srv/new/PROJ_IOS_VK_ORG_K02_R24.rvt — в набор 00_Связь_ВК")) "dry-run lists the link and its workset"
Check (-not $linkText.Contains("следующий шаг")) "a plan with inserts does not postpone them"
$noPair = @([pscustomobject]@{ State = "no_pair"; WorksetName = "00_Связь_ВК"; TargetPath = ""; LinkPath = "" })
$blk3 = Test-WorksetStep4Block $noPair @()
Check (-not $blk3.Block) "a row without a pair is skipped, it does not stop the other inserts"
Eq $blk3.NoPair 1 "the unpaired row is counted"
$blk4 = Test-WorksetStep4Block @($noPair[0], [pscustomobject]@{ State = "unparsed"; WorksetName = "" }, [pscustomobject]@{ State = "insert"; WorksetName = "01_Связь_RVT_HC_K1" }) @()
Check (-not $blk4.Block) "no_pair and unparsed next to a good insert: the launch goes on"

# ---------------------------------------------------------------- mask from the section codes (real names, tab 4)
$siteTarget = "RSN://revit-server-2022.example.local/proj SampleSite/Стадия П/04_ЭОМ/SITE_ORG_EOM_K1_R22.rvt"
$siteNew = New-Object System.Collections.Generic.List[string]
foreach ($b in @("K1", "K2", "K3", "K4")) { [void]$siteNew.Add(("RSN://revit-server-2022.example.local/proj SampleSite/Стадия П/06_HC/SITE_ORG_HC_{0}_R22.rvt" -f $b)) }
foreach ($b in @("K1", "K2", "K3", "K4", "K5", "PRK")) { [void]$siteNew.Add(("RSN://revit-server-2022.example.local/proj SampleSite/Стадия П/07_AVT/SITE_ORG_AVT_{0}_R22.rvt" -f $b)) }
$siteNames = New-Object System.Collections.Generic.List[string]
[void]$siteNames.Add((Get-WorksetBaseName $siteTarget))
foreach ($p in $siteNew) { [void]$siteNames.Add((Get-WorksetBaseName $p)) }
$old = Get-WorksetNameMaskSuggestion @($siteNew | ForEach-Object { Get-WorksetBaseName $_ }) @((Get-WorksetBaseName $siteTarget)) "_"
Check (-not $old.Ok) "the old guess fails on 1 target + 10 new of two sections (the bug)"
$byCodes = Get-WorksetNameMaskByCodes $siteNames.ToArray() $Codes "_"
Check $byCodes.Ok "section codes find the mask on the same names"
Eq ($byCodes.Roles -join ",") "fixed,fixed,section,building,fixed" "SITE, ORG fixed; EOM/HC/AVT section; K1..PRK building; R22 fixed"
$sitePairs = Get-WorksetLinkPairs $siteNew.ToArray() @($siteTarget) "_" $byCodes.Roles
$siteIns = @(Rows $sitePairs | Where-Object { $_.State -eq "insert" })
Eq $siteIns.Count 2 "EOM_K1 gets two links"
Eq (@($siteIns | ForEach-Object { $_.LinkName }) -join ",") "SITE_ORG_HC_K1_R22,SITE_ORG_AVT_K1_R22" "HC_K1 and AVT_K1"
Eq @(Rows $sitePairs | Where-Object { $_.State -eq "no_pair" }).Count 8 "the other 8 new models have no pair (other buildings)"
Eq @(Rows $sitePairs | Where-Object { $_.State -eq "unparsed" }).Count 0 "nothing is unparsed"
$one = Get-WorksetNameMaskByCodes @("SITE_ORG_EOM_K1_R22", "SITE_ORG_HC_K1_R22") $Codes "_"
Eq ($one.Roles -join ",") "fixed,fixed,section,building,fixed" "one building only: the token next to the section is the building, not R22"
$cyr = Get-WorksetNameMaskByCodes @("PROJ_К1_АР_R22", "PROJ_К2_АР_R22", "PROJ_К1_КР_R22", "PROJ_К1_ОВ_R22") $Codes "_"
Eq ($cyr.Roles -join ",") "fixed,building,section,fixed" "Cyrillic codes and the building before the section"
$cyr1 = Get-WorksetNameMaskByCodes @("PROJ_К1_АР_R22", "PROJ_К1_ОВ_R22") $Codes "_"
Eq ($cyr1.Roles -join ",") "fixed,building,section,fixed" "R22 after the section is skipped, К1 before it is the building"
Eq (Get-WorksetNameMaskByCodes @("A_B_C", "A_B_C_D") $Codes "_").Reason "length" "different token counts"
Eq (Get-WorksetNameMaskByCodes @("X_Y_K1", "X_Z_K2") $Codes "_").Reason "no_codes" "no section code in the names"
Eq (Get-WorksetNameMaskByCodes @("P_AR_K1_V1", "P_OV_K2_V2") $Codes "_").Reason "ambiguous" "two varying positions besides the section"
Check (-not (Test-WorksetYearToken "K1") -and (Test-WorksetYearToken "R22") -and (Test-WorksetYearToken "r2024")) "R22 / R2024 are Revit years"

# ---------------------------------------------------------------- which mask is used: user > saved > codes > guess
$savedRoles = @("fixed", "fixed", "section", "building", "fixed")
$tgtNames = @((Get-WorksetBaseName $siteTarget))
$newNames = @($siteNew | ForEach-Object { Get-WorksetBaseName $_ })
$r1 = Resolve-WorksetNameMask @() $savedRoles $newNames $tgtNames $Codes "_"
Eq $r1.Source "saved" "a saved mask is used as it is"
$weird = @("fixed", "fixed", "building", "section", "fixed")
$r2 = Resolve-WorksetNameMask @() $weird $newNames $tgtNames $Codes "_"
Eq $r2.Source "saved" "a saved mask is never replaced by the detection"
Eq ($r2.Roles -join ",") ($weird -join ",") "the saved roles come back unchanged"
$r3 = Resolve-WorksetNameMask @() @("fixed", "fixed", "fixed", "fixed", "fixed") $newNames $tgtNames $Codes "_"
Eq $r3.Source "codes" "a saved file without section/building is not a mask: detection runs"
$r4 = Resolve-WorksetNameMask @("fixed", "fixed", "section", "section", "fixed") $savedRoles $newNames $tgtNames $Codes "_"
Check (-not $r4.Ok) "an unfinished user edit is not replaced by the saved mask"
Eq $r4.Reason "roles" "the reason is the role set"
$r5 = Resolve-WorksetNameMask @("fixed", "fixed", "section", "building", "ignore") @() $newNames $tgtNames $Codes "_"
Eq $r5.Source "user" "a valid user edit wins"
$r6 = Resolve-WorksetNameMask @() @() @("X_Y_K1") @("X_Z_K2") $Codes "_"
Check (-not $r6.Ok) "no codes and no guess: no mask"
Eq $r6.Source "none" "source none"

# ---------------------------------------------------------------- status line is never silent
# rows with workset names (the empty-workset message replaces the counts)
$siteNamed = @(Rows $sitePairs | ForEach-Object { $c = $_.PSObject.Copy(); $c | Add-Member -Force -NotePropertyName WorksetName -NotePropertyValue "x"; $c })
$sumOk = Get-WorksetPairSummary $siteNamed $r1 10 1
Check ($sumOk.Contains("вставить 2") -and $sumOk.Contains("нет пары 8") -and $sumOk.Contains("name_mask.cfg")) ("summary of a good run: " + $sumOk)
$unparsedRows = @(Rows (Get-WorksetLinkPairs @("RSN://h/X_Y_K1.rvt") @("RSN://h/X_Z_K2.rvt") "_" @()))
Eq @($unparsedRows | Where-Object { $_.State -eq "unparsed" }).Count 2 "no roles: both rows are unparsed, the table is not empty"
$sumBad = Get-WorksetPairSummary $unparsedRows $r6 1 1
Check ($sumBad.Contains("Имя не разобрано") -and $sumBad.Contains("Маска имени")) ("summary without a mask points to «Маска имени»: " + $sumBad)
Check ((Get-WorksetPairSummary @() $r4 1 1).Contains("ровно один")) "summary of an unfinished edit names the rule"
Eq (Get-WorksetPairSummary @() $r1 0 1) "Отметьте новые модели в колонке «Роль»." "summary without new models says so"

# ---------------------------------------------------------------- link workset mask keeps the building
$hostLinks = @(
    (Link "SITE_ORG_AR_K1_R22" "01_Связь_RVT_AR_K1" $false),
    (Link "SITE_ORG_AR_PRK_R22" "01_Связь_RVT_AR_PRK" $false),
    (Link "SITE_ORG_EOM_PRK_R22" "01_Связь_RVT_EOM_PRK" $false),
    (Link "SITE_ORG_TASK_PRK_R22" "01_Связь_RVT_TASK_PRK" $false)
)
$hm = Get-WorksetLinkMask $hostLinks $Codes "_" $savedRoles $null
Check $hm.Ok "links of K1 and of PRK agree once the building is a placeholder"
Eq $hm.Mask "01_Связь_RVT_{section}_{building}" "section and building placeholders"
Eq $hm.Script "latin" "latin section codes"
Eq (Get-ProposedWorksetName $hm.Mask $hm.Script "HC" $Codes "K1") "01_Связь_RVT_HC_K1" "HC_K1 goes to 01_Связь_RVT_HC_K1"
Eq (Get-ProposedWorksetName $hm.Mask $hm.Script "AVT" $Codes "PRK") "01_Связь_RVT_AVT_PRK" "AVT_PRK goes to 01_Связь_RVT_AVT_PRK"
Eq (Get-ProposedWorksetName $hm.Mask $hm.Script "AVT" $Codes "") "" "no building: no name instead of a broken one"
$maj = @(
    (Link "PROJ_IOS_OV_ORG_K01_R24" "00_Связь_ОВ" $false),
    (Link "PROJ_IOS_EOM_ORG_K01_R24" "00_Связь_ЭОМ" $false),
    (Link "PROJ_IOS_VK_ORG_K01_R24" "01_Связь_ВК" $false)
)
$mj = Get-WorksetLinkMask $maj $Codes "_" $Roles $null
Check ($mj.Ok -and $mj.Reason -eq "majority") "two of three links agree: their mask wins"
Eq $mj.Mask "00_Связь_{section}" "the majority mask"

# ---------------------------------------------------------------- what the «Связи» table shows
$tK9 = "RSN://revit-server-2022.example.local/proj SampleSite/Стадия П/04_ЭОМ/SITE_ORG_EOM_K9_R22.rvt"
$withK9 = Get-WorksetLinkPairs $siteNew.ToArray() @($siteTarget, $tK9) "_" $byCodes.Roles
Check (@(Rows $withK9 | Where-Object { $_.TargetPath -eq $tK9 -and -not $_.LinkPath }).Count -eq 1) "the pair builder still knows the unpaired target"
$vis = Select-WorksetVisiblePairs (Rows $withK9) 10
Eq @($vis.Rows | Where-Object { $_.TargetPath -and -not $_.LinkPath }).Count 0 "unpaired targets are not listed"
Eq @($vis.Rows | Where-Object { $_.State -eq "no_pair" }).Count 8 "«нет пары» rows are the new models only"
Eq @($vis.Rows | Where-Object { $_.State -eq "insert" }).Count 2 "the inserts stay"
$allTargets = Get-WorksetLinkPairs @() @($siteTarget, $tK9) "_" $byCodes.Roles
Check (@(Rows $allTargets).Count -gt 0) "without new models the builder returns one row per target"
$none = Select-WorksetVisiblePairs (Rows $allTargets) 0
Eq $none.Count 0 "no new model: the table is empty"
$manOnly = @([pscustomobject]@{ TargetPath = $siteTarget; LinkPath = $tK9; State = "insert"; Source = "manual" })
Eq (Select-WorksetVisiblePairs $manOnly 0).Count 1 "a hand-made row stays even without new models"

# ---------------------------------------------------------------- empty workset: skipped with a reason
$emptyRow = [pscustomobject]@{ State = "insert"; TargetPath = $siteTarget; LinkPath = $siteNew[0]; WorksetName = ""; Key = "e1"; Section = "HC" }
$okRow = [pscustomobject]@{ State = "insert"; TargetPath = $siteTarget; LinkPath = $siteNew[4]; WorksetName = "01_Связь_RVT_AVT_K1"; Key = "e2"; Section = "AVT" }
$ePlan = Get-WorksetRenameCreatePlan @() @() @([pscustomobject]@{ Host = $siteTarget; Name = "01_Связь_RVT_AVT_K1" }) @($emptyRow, $okRow)
Eq @(@($ePlan.Rows) | Where-Object { $_.Action -eq "insert" }).Count 1 "only the row with a workset is inserted"
Eq @(@($ePlan.Rows) | Where-Object { $_.Code -eq "SKIP_EMPTY_WORKSET" }).Count 1 "the row without a workset is a skip, not an insert with an empty name"
$eView = Get-WorksetInsertView @($ePlan.Rows) @()
Eq @($eView.Skipped).Count 1 "the confirm form lists the skipped row"
Eq @($eView.Skipped)[0].Reason "нет рабочего набора" "the skipped row says why, shortly"
Check (@($eView.Skipped)[0].Tip.Contains("введите имя в колонке «Рабочий набор»")) "the full reason is the tool tip"
Eq @($eView.Rows)[0].Action "вставить + новый набор" "short action of an insert with a new workset"
Eq (Get-WorksetInsertActionText $false) "вставить" "short action of a plain insert"
foreach ($code in @("SKIP_EXISTS", "SKIP_OTHER_PATH", "SKIP_EMPTY_WORKSET", "SKIP_BAD_NAME", "SKIP_CONFLICT", "no_pair", "unparsed")) {
    Check ((Get-WorksetSkipReason $code).Length -le 34) ("short reason fits a narrow column: " + $code)
    Check ((Get-WorksetSkipReasonLong $code).Length -gt (Get-WorksetSkipReason $code).Length) ("long reason for the tool tip: " + $code)
}
Eq @($eView.Skipped)[0].Link "SITE_ORG_HC_K1_R22" "short name, not the RSN path"
Eq @($eView.Rows)[0].Model "SITE_ORG_EOM_K1_R22" "the form shows short model names"
Check (@($eView.Rows)[0].IsNew) "a workset created by this run is marked new"
Eq $eView.Summary "В 1 модель будет вставлена 1 связь, будет создан 1 рабочий набор." "summary of the form"
$eNote = Get-WorksetPairSummary @($emptyRow, $okRow) $r1 10 1
Eq $eNote "Нет рабочего набора у 1 строки: прочитайте наборы и связи или введите имя." "status names the empty worksets, without the counts"
Eq (Get-WorksetEmptyWorksetText 3) "Нет рабочего набора у 3 строк: прочитайте наборы и связи или введите имя." "3 rows"

# ---------------------------------------------------------------- status line fits one line at the minimum width
$many = New-Object System.Collections.Generic.List[object]
for ($i = 0; $i -lt 120; $i++) {
    [void]$many.Add([pscustomobject]@{ State = "insert"; WorksetName = "x"; Source = "manual" })
    [void]$many.Add([pscustomobject]@{ State = "exists"; WorksetName = "x" })
    [void]$many.Add([pscustomobject]@{ State = "no_pair"; WorksetName = "" })
}
$rUser = [pscustomobject]@{ Ok = $true; Source = "user"; Roles = @("section", "building"); Reason = "" }
$texts = @(
    (Get-WorksetPairSummary $many.ToArray() $rUser 10 1),
    (Get-WorksetPairSummary @($emptyRow, $okRow) $r1 10 1),
    (Get-WorksetPairSummary $unparsedRows $r6 1 1),
    (Get-WorksetPairSummary @() $r4 1 1),
    (Get-WorksetPairSummary @() $r1 0 1),
    (Get-WorksetPairSummary @() $r1 1 0),
    (Get-WorksetEmptyWorksetText 1234),
    (Join-WorksetStatus "Добавлено строк вручную: 12." (Get-WorksetPairSummary $many.ToArray() $rUser 10 1)),
    (Join-WorksetStatus "Маска записана в name_mask.cfg." (Get-WorksetPairSummary $many.ToArray() $rUser 10 1))
)
foreach ($t in $texts) { Check ($t.Length -le $script:WsStatusMax) ("status fits one line ({0} chars): {1}" -f $t.Length, $t) }
Eq (Join-WorksetStatus "Маска записана в name_mask.cfg." "Пары: вставить 2, уже есть 0, нет пары 8. Маска имени: из name_mask.cfg.") "Маска записана в name_mask.cfg. Пары: вставить 2, уже есть 0, нет пары 8. Маска имени: из name_mask.cfg." "a short note keeps the prefix"
Eq (Join-WorksetStatus "Добавлено строк вручную: 1." "") "Добавлено строк вручную: 1." "no note: the prefix alone"
$npView = Get-WorksetInsertView @() @([pscustomobject]@{ State = "no_pair"; TargetPath = ""; LinkPath = $siteNew[1] })
Eq @($npView.Skipped)[0].Reason "нет целевой модели этого корпуса" "a new model without a pair is listed in «Пропущено»"
Eq $npView.Summary "Вставлять нечего: все строки пропущены." "nothing to insert"

# ---------------------------------------------------------------- plural forms
Eq (Get-WorksetInsertSummary 2 1 2) "В 1 модель будет вставлено 2 связи, будет создано 2 рабочих набора." "2 links, 2 worksets"
Eq (Get-WorksetInsertSummary 1 1 0) "В 1 модель будет вставлена 1 связь." "1 link, no worksets"
Eq (Get-WorksetInsertSummary 5 2 5) "В 2 модели будет вставлено 5 связей, будет создано 5 рабочих наборов." "5 links"
Eq (Get-WorksetInsertSummary 21 5 11) "В 5 моделей будет вставлена 21 связь, будет создано 11 рабочих наборов." "21 / 11"
Eq (Get-WorksetInsertSummary 12 22 1) "В 22 модели будет вставлено 12 связей, будет создан 1 рабочий набор." "12 / 22"
Eq (Get-WorksetInsertSummary 0 0 0) "Вставлять нечего." "nothing at all"
Eq (Get-RuPlural 111 "a" "b" "c") "c" "111 is many"
Eq (Get-RuPlural 104 "a" "b" "c") "b" "104 is few"

# ---------------------------------------------------------------- hand-made pairs
$prk = @($siteNew | Where-Object { $_ -like "*AVT_PRK*" })[0]
$manualIn = @([pscustomobject]@{ TargetPath = $siteTarget; LinkPath = $prk })
$merged = @(Merge-WorksetManualPairs (Rows $sitePairs) $manualIn "_" $savedRoles $Codes)
$man = @($merged | Where-Object { $_.Source -eq "manual" })
Eq $man.Count 1 "the hand pair is one row"
Eq $man[0].State "insert" "the hand pair is an insert"
Eq $man[0].Section "AVT" "section from the mask"
Eq $man[0].Building "PRK" "building from the mask (outside the same-building rule)"
Eq @($merged | Where-Object { $_.LinkPath -eq $prk -and $_.State -eq "no_pair" }).Count 0 "the no_pair row of AVT_PRK is gone"
Eq @($merged | Where-Object { $_.State -eq "no_pair" }).Count 7 "the other 7 stay without a pair"
Eq @($merged | Where-Object { $_.Source -eq "auto" }).Count ($merged.Count - 1) "every other row is marked auto"
$hc1 = @($siteNew | Where-Object { $_ -like "*HC_K1*" })[0]
$dup = @(Merge-WorksetManualPairs (Rows $sitePairs) @([pscustomobject]@{ TargetPath = $siteTarget; LinkPath = $hc1 }) "_" $savedRoles $Codes)
Eq @($dup | Where-Object { $_.LinkPath -eq $hc1 }).Count 1 "a hand pair equal to an automatic one is not doubled"
$selfM = @(Merge-WorksetManualPairs @() @([pscustomobject]@{ TargetPath = $siteTarget; LinkPath = $siteTarget }) "_" $savedRoles $Codes)
Eq $selfM.Count 0 "a model is not linked into itself by hand"
$noMask = @(Merge-WorksetManualPairs @() $manualIn "_" @() $Codes)
Eq $noMask[0].Section "AVT" "without a mask the section is the first code in the name"

# ---------------------------------------------------------------- csv roundtrip, seed file is not overwritten
$tmp = Join-Path ([System.IO.Path]::GetTempPath()) ("ws_" + [Guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $tmp -Force | Out-Null
try {
    $codesPath = Join-Path $tmp "section_codes.csv"
    $made = Initialize-SectionCodeFile $codesPath
    Eq $made.Count $wantCodes.Count "a missing code file is seeded"
    $bytes = [System.IO.File]::ReadAllBytes($codesPath)
    Check ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) "section_codes.csv is UTF-8 with BOM"
    Check (([System.IO.File]::ReadAllText($codesPath, (New-Object System.Text.UTF8Encoding $true))).Contains("ВК")) "ВК survives the file"
    Write-SectionCodeFile $codesPath @([pscustomobject]@{ Latin = "ZZ"; Cyrillic = "ЯЯ" })
    $again = Initialize-SectionCodeFile $codesPath
    Eq $again.Count 1 "an edited code file is not replaced by the seed"
    Eq (Rows $again)[0].Cyrillic "ЯЯ" "the user's row is kept"
    $maskPath = Join-Path $tmp "name_mask.cfg"
    Write-NameMaskFile $maskPath "_" $Roles
    $readMask = Read-NameMaskFile $maskPath
    Eq $readMask.Separator "_" "mask separator roundtrip"
    Eq ($readMask.Roles -join ",") ($Roles -join ",") "mask roles roundtrip"
    $planPath = Join-Path $tmp "workset_plan.csv"
    $srcActs = Rows $act
    $srcActs[0].Arg1 = "RSN://srv/папка/link.rvt"
    Write-WorksetPlanFile $planPath $srcActs
    $planBytes = [System.IO.File]::ReadAllBytes($planPath)
    Check ($planBytes[0] -eq 0xEF -and $planBytes[1] -eq 0xBB -and $planBytes[2] -eq 0xBF) "workset_plan.csv is UTF-8 with BOM"
    $back = Read-WorksetPlanFile $planPath
    Eq $back.Count $act.Count "plan row count roundtrip"
    Eq (Rows $back)[0].Arg1 "RSN://srv/папка/link.rvt" "Cyrillic path roundtrip"
    Eq (Rows $back)[0].Arg2 "00_Связь_ВК" "Cyrillic workset name roundtrip"
    Check (Rows $back)[0].CreateWorkset "create flag roundtrip"
    Check (-not (Rows $back)[1].CreateWorkset) "the second create flag stays off"
    $quoted = @([pscustomobject]@{ Action = "skip"; Host = "h"; Key = "k"; Arg1 = "a;b"; Arg2 = 'say "hi"'; Arg3 = ""; Code = "SKIP_OTHER_PATH"; CreateWorkset = $false })
    Write-WorksetPlanFile $planPath $quoted
    $q = Rows (Read-WorksetPlanFile $planPath)
    Eq $q[0].Arg1 "a;b" "a semicolon inside a field survives"
    Eq $q[0].Arg2 'say "hi"' "quotes inside a field survive"
}
finally {
    Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue
}

# ---------------------------------------------------------------- delete workset lines
$delSets = @(
    [pscustomobject]@{ Host = "RSN://h/a.rvt"; Id = "10"; Name = "00_Связь_ОВ" },
    [pscustomobject]@{ Host = "RSN://h/a.rvt"; Id = "11"; Name = "30_ИОС_Элементы" }
)
$delRows = @(Get-WorksetDeleteActions $delSets @("00_Связь_ОВ", "30_ИОС_Элементы") "move")
Eq $delRows.Count 1 "the landing workset is not deleted"
Eq $delRows[0].Arg2 "move" "the default delete moves elements"
Eq (Get-WorksetDeleteDryLine "RSN://h/a.rvt" "00_Связь_ОВ" "elements" -1) "ВМЕСТЕ С ЭЛЕМЕНТАМИ: RSN://h/a.rvt — удалить набор 00_Связь_ОВ и не посчитано элементов" "delete-with-elements names the unknown count"
$delEl = @(Get-WorksetDeleteActions $delSets @("00_Связь_ОВ") "elements")
Eq $delEl[0].Arg2 "elements" "the elements option is explicit"

# ---------------------------------------------------------------- one-line report pointer
$rep = Join-Path ([System.IO.Path]::GetTempPath()) ("wsrep_" + [Guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $rep -Force | Out-Null
try {
    $older = Join-Path $rep "workset_read_2026-11-10_1000.csv"
    $newer = Join-Path $rep "workset_read_2026-11-10_1006.csv"
    Write-Utf8BomLines $older @("row_kind;host")
    Write-Utf8BomLines $newer @("row_kind;host")
    [System.IO.File]::SetLastWriteTimeUtc($older, [datetime]::UtcNow.AddMinutes(-5))
    [System.IO.File]::SetLastWriteTimeUtc($newer, [datetime]::UtcNow)
    $pointer = Join-Path $rep "_active_workset_read.txt"
    Write-Utf8BomLines $pointer @($newer)
    $got = Get-WorksetPointerTarget $pointer
    Eq $got $newer "a one-line pointer is the whole path"
    Check ($got.Length -gt 1) "the pointer line is not a single character"
    $hit = Resolve-WorksetReportFile $pointer "workset_read"
    Check ($hit.Ok -and -not $hit.Fallback) "an existing pointer file is used as-is"
    Eq $hit.Path $newer "resolve returns the pointed csv"
    Write-Utf8BomLines $pointer @((Join-Path $rep "missing.csv"))
    $fall = Resolve-WorksetReportFile $pointer "workset_read"
    Check ($fall.Ok -and $fall.Fallback) "a broken pointer falls back to the newest csv"
    Eq $fall.Path $newer "the newest workset_read csv wins"
    $none = Resolve-WorksetReportFile (Join-Path $rep "absent.txt") "workset_apply"
    Check (-not $none.Ok) "a missing apply report stays unresolved"
}
finally { Remove-Item -LiteralPath $rep -Recurse -Force -ErrorAction SilentlyContinue }

Write-Host ""
Write-Host ("Checks passed: {0}, failed: {1}" -f $script:Pass, $script:Fail)
if ($script:Fail -gt 0) { exit 1 }
exit 0
