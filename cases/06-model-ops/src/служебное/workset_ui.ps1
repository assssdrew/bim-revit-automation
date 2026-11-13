# Name: workset_ui.ps1
# Version: 3.1.0
# What it does: See the case README.
# Notes: Paths and project names in this portfolio copy are placeholders.

# UTF-8 WITH BOM required (PowerShell 5.1).
# Window for "Рабочие наборы и связи": tabs, roles, read and apply.
# Dot-sourcing defines functions only. Install-WorksetUi builds the controls.

# The functions use System.Drawing / WinForms types (colors, sizes). presets.ps1 loads them
# first; loading here too keeps the file usable on its own (tests, Windows PowerShell 5.1
# does not load System.Drawing by itself as PowerShell 7 does). WinForms is missing off
# Windows (the logic tests run there too): a failed load must not stop the dot-sourcing.
foreach ($asm in @("System.Drawing", "System.Windows.Forms")) {
    try { Add-Type -AssemblyName $asm -ErrorAction Stop } catch { }
}

$script:WsRoleNew = "new"
$script:WsRoleTarget = "target"
$script:WsRoles = @{}
$script:WsManual = @{}
# pairs added with «Добавить связь вручную…»: { TargetPath; LinkPath }, kept for the session
$script:WsManualPairs = New-Object System.Collections.Generic.List[object]
# name mask: saved = name_mask.cfg (never wiped), user = edited on «Маска имени», not saved yet
$script:WsSavedRoles = @()
$script:WsUserRoles = @()
$script:WsMaskResolved = $null
$script:WsPairNote = ""
$script:WsRenameEdit = @{}
$script:WsApplyById = @{}
$script:WsApplyByPath = @{}
$script:WsReadMode = "fast"
$script:WsDeleteNames = New-Object System.Collections.Generic.List[string]
$script:WsApplyPart = ""

function Get-WorksetApplyResultLabel([string]$Code) {
    switch ($Code) {
        "OK" { return "записано" }
        "SKIP_CONFLICT" { return "конфликт" }
        "SKIP_BORROW" { return "занято" }
        "SKIP_NOT_WORKSHARED" { return "не совместная" }
        "SKIP_BAD_NAME" { return "плохое имя" }
        "ERROR_CREATE" { return "ошибка создания" }
        "ERROR_SYNC" { return "ошибка синхронизации" }
        "SKIP_EXISTS" { return "уже есть" }
        "SKIP_OTHER_PATH" { return "другой путь" }
        "SKIP_COORDS" { return "нет общих координат" }
        "SKIP_NOT_FOUND" { return "файл не найден" }
        "ERROR_LINK" { return "ошибка связи" }
        default { return [string]$Code }
    }
}

function Get-WorksetLinkApplyNote([string]$HostPath, [string]$LinkPath) {
    if (-not $script:WsApplyByPath) { return "" }
    $k = (ConvertTo-WorksetPathKey $HostPath) + "|" + (ConvertTo-WorksetPathKey $LinkPath)
    if (-not $script:WsApplyByPath.ContainsKey($k)) { return "" }
    return Get-WorksetApplyResultLabel ([string]$script:WsApplyByPath[$k])
}

function Get-WorksetApplyNote([string]$Name, [string]$HostPath, [string]$HostList) {
    if (-not $script:WsApplyById -or -not $script:WsInventory -or -not $script:WsInventory.Worksets) { return "" }
    $hosts = @{}
    if ($HostPath) { $hosts[(ConvertTo-WorksetPathKey $HostPath)] = $true }
    foreach ($h in ([string]$HostList).Split("|")) {
        if ($h) { $hosts[(ConvertTo-WorksetPathKey $h)] = $true }
    }
    $found = ""
    foreach ($w in @($script:WsInventory.Worksets)) {
        if (-not $hosts.ContainsKey((ConvertTo-WorksetPathKey ([string]$w.Host)))) { continue }
        if (-not (Test-WorksetSameText ([string]$w.Name) $Name)) { continue }
        $id = [string]$w.Id
        if (-not $script:WsApplyById.ContainsKey($id)) { continue }
        $label = Get-WorksetApplyResultLabel ([string]$script:WsApplyById[$id])
        if ($label -like "ошибка*") { return $label }
        $found = $label
    }
    return $found
}

# The worksets have their own tabs now («Связи: добавить», «Рабочие наборы»); tab 1 never runs them.
function Test-UiWorksetsOn {
    return $false
}

function Get-WorksetRoleCode([string]$Path) {
    if (-not $script:WsRoles) { return $script:WsRoleTarget }
    $key = ConvertTo-WorksetPathKey $Path
    if ($key -and $script:WsRoles.ContainsKey($key) -and $script:WsRoles[$key] -eq $script:WsRoleNew) {
        return $script:WsRoleNew
    }
    return $script:WsRoleTarget
}

function Get-WorksetRoleLabel([string]$Path) {
    if ((Get-WorksetRoleCode $Path) -eq $script:WsRoleNew) { return "новая" }
    return "целевая"
}

function Set-WorksetRole {
    param([string[]]$Paths, [string]$Role)
    if (-not $script:WsRoles) { $script:WsRoles = @{} }
    $code = $script:WsRoleTarget
    if ($Role -eq $script:WsRoleNew -or $Role -eq "новая") { $code = $script:WsRoleNew }
    foreach ($p in $Paths) {
        if ([string]::IsNullOrWhiteSpace($p)) { continue }
        $script:WsRoles[(ConvertTo-WorksetPathKey $p)] = $code
    }
}

function Get-WorksetStateLabel([string]$State) {
    switch ($State) {
        "insert" { return "вставить" }
        "exists" { return "есть" }
        "other_path" { return "другой путь" }
        "no_pair" { return "нет пары" }
        "unparsed" { return "имя не разобрано" }
        default { return [string]$State }
    }
}

# Paint tokens, not colors: the view maps them. Empty workset blocks only an insert.
function Get-WorksetLinkPaint {
    param([string]$State, [string]$WorksetName, [bool]$Manual)
    $row = ""
    if ($State -eq "no_pair" -or $State -eq "unparsed") { $row = "missing" }
    $cell = ""
    if ($Manual) { $cell = "manual" }
    elseif ($State -eq "insert" -and [string]::IsNullOrWhiteSpace($WorksetName)) { $cell = "empty" }
    $dim = ($State -eq "exists" -or $State -eq "other_path")
    return [pscustomobject]@{ Row = $row; Cell = $cell; Dim = $dim }
}

function Get-WorksetRoleWord([string]$Role) {
    switch ((ConvertTo-WorksetRole $Role)) {
        "section" { return "раздел" }
        "building" { return "корпус" }
        "ignore" { return "не важен" }
        default { return "как есть" }
    }
}

# Choices of the «Роль» drop-down on «Маска имени», in this order.
function Get-WorksetMaskRoleChoices {
    return @("раздел", "корпус", "как есть", "не важен")
}

function ConvertFrom-WorksetRoleWord([string]$Word) {
    switch (([string]$Word).Trim()) {
        "раздел" { return "section" }
        "корпус" { return "building" }
        "не важен" { return "ignore" }
        default { return "fixed" }
    }
}

# Choices of the «Роль» drop-down of the model list on «Связи: добавить».
function Get-WorksetModelRoleChoices {
    return @("новая", "целевая")
}

# Rows a click in the «Роль» cell applies to: the whole selection (taken before the click
# changed it) when the clicked row is part of it, otherwise only the clicked row.
function Get-WorksetRoleClickPaths {
    param([string[]]$SelectedBefore, [string]$Clicked)
    $list = New-Object System.Collections.Generic.List[string]
    $inSel = $false
    if ($null -ne $SelectedBefore) {
        foreach ($p in $SelectedBefore) { if ((ConvertTo-WorksetPathKey $p) -eq (ConvertTo-WorksetPathKey $Clicked)) { $inSel = $true } }
    }
    if ($inSel) { foreach ($p in $SelectedBefore) { if ($p) { [void]$list.Add([string]$p) } } }
    elseif ($Clicked) { [void]$list.Add($Clicked) }
    return $list.ToArray()
}

function New-WorksetReadPass {
    param([string]$ToolDir, [string]$CodeDir, [string]$Year, [string]$Worksets = "close_all")
    if ($Worksets -ne "open_all") { $Worksets = "close_all" }
    return [pscustomobject]@{
        Kind     = "rbp"
        Name     = "workset-read"
        Task     = [System.IO.Path]::Combine($CodeDir, "read_worksets.py")
        List     = [System.IO.Path]::Combine($ToolDir, "rvt_list_worksets.txt")
        Year     = $Year
        Detach   = $false
        Worksets = $Worksets
    }
}

function Get-WorksetReadStatusNote([string]$Mode, [int]$WorksetCount, [int]$LinkCount) {
    if ($Mode -eq "full") { return "" }
    if ($WorksetCount -gt 0 -and $LinkCount -eq 0) {
        return "При закрытых наборах связи могут быть не видны. Число элементов этот проход не считает."
    }
    return "Быстрое чтение не считает элементы в наборах. Связи видны, только если их набор открыт."
}

function Resolve-WorksetReadYear {
    param([string[]]$Paths, $Servers, [string]$Chosen)
    if (-not [string]::IsNullOrWhiteSpace($Chosen) -and $Chosen -ne "из servers.cfg") {
        return [pscustomobject]@{ Ok = $true; Year = $Chosen.Trim(); Message = ""; Detail = ""; Rows = @() }
    }
    # same detection as «Смена версии Revit» (ui_logic.ps1)
    return (Get-ReadYearInfo $Paths $Servers)
}

function Get-WorksetYearCaption($Resolved) {
    return (Get-ReadYearCaption $Resolved)
}

function Get-WorksetYearText($Resolved) {
    # message box text: the long reason when there is one
    if ($Resolved -and $Resolved.Detail) { return [string]$Resolved.Detail }
    if ($Resolved -and $Resolved.Message) { return [string]$Resolved.Message }
    return "Год Revit не определён"
}

function Get-WorksetTabPaths([string]$Key, [bool]$CheckedOnly = $true) {
    $list = New-Object System.Collections.Generic.List[string]
    if (-not $script:Ui -or -not $script:Ui.ContainsKey("PlainLists") -or -not $script:Ui.PlainLists) { return $list.ToArray() }
    if (-not $script:Ui.PlainLists.ContainsKey($Key)) { return $list.ToArray() }
    foreach ($it in @($script:Ui.PlainLists[$Key].Items)) {
        if ($null -eq $it) { continue }
        if ($CheckedOnly -and -not $it.Checked) { continue }
        [void]$list.Add([string]$it.Tag)
    }
    return $list.ToArray()
}

# year label of each tab (its own models) + the per-row reason in «Состояние»
function Update-WorksetYearLabel {
    if (-not $script:Ui) { return }
    $servers = $null
    if (Get-Variable -Name servers -Scope Script -ErrorAction SilentlyContinue) { $servers = $script:servers }
    foreach ($pair in @(@("WsYearLabel", "addlinks"), @("WsYearLabel2", "worksets"))) {
        $lbl = $null
        if ($script:Ui.ContainsKey($pair[0])) { $lbl = $script:Ui[$pair[0]] }
        if (-not $lbl) { continue }
        $paths = @(Get-WorksetTabPaths $pair[1])
        $info = $null
        if ($paths.Count -gt 0) { $info = Resolve-WorksetReadYear $paths $servers "" }
        $text = "Год Revit не определён"
        $red = $false
        if ($info) { $text = Get-WorksetYearCaption $info; $red = (-not $info.Ok) }
        if ($lbl.Text -ne $text) { $lbl.Text = $text }
        if ($red) { $lbl.ForeColor = [System.Drawing.Color]::Firebrick } else { $lbl.ForeColor = [System.Drawing.SystemColors]::ControlText }
        Set-WorksetYearReasons $pair[1] $info
    }
}

function Set-WorksetYearReasons([string]$Key, $Info) {
    if (-not $script:Ui.ContainsKey("PlainLists") -or -not $script:Ui.PlainLists -or -not $script:Ui.PlainLists.ContainsKey($Key)) { return }
    $byPath = @{}
    if ($Info) { foreach ($r in @($Info.Rows)) { if ($r.Reason) { $byPath[[string]$r.Path] = [string]$r.Reason } } }
    foreach ($it in @($script:Ui.PlainLists[$Key].Items)) {
        if ($null -eq $it -or $it.SubItems.Count -lt 3) { continue }
        $cell = $it.SubItems[2]
        $p = [string]$it.Tag
        if ($byPath.ContainsKey($p)) {
            if ($cell.Text -ne $byPath[$p]) { $cell.Text = $byPath[$p] }
            $cell.Tag = "reason"
            $cell.ForeColor = [System.Drawing.Color]::Firebrick
        }
        elseif ($cell.Tag -eq "reason") {
            $cell.Text = ""
            $cell.Tag = $null
            $cell.ForeColor = [System.Drawing.SystemColors]::WindowText
        }
    }
}

function Test-WorksetReadPaths {
    param([string[]]$Paths)
    $bad = New-Object System.Collections.Generic.List[string]
    foreach ($p in $Paths) {
        if ([string]::IsNullOrWhiteSpace($p)) { continue }
        $t = Test-WorksetStoragePath $p
        if (-not $t.Ok) { [void]$bad.Add($p) }
    }
    return [pscustomobject]@{ Ok = ($bad.Count -eq 0); Count = $bad.Count; Paths = $bad.ToArray() }
}

function ConvertTo-WorksetInventoryFromRows {
    param($Rows)
    $worksets = New-Object System.Collections.Generic.List[object]
    $links = New-Object System.Collections.Generic.List[object]
    $bare = @{}
    if ($null -ne $Rows) {
        foreach ($r in $Rows) {
            $kind = ""
            if ($r.PSObject.Properties["row_kind"]) { $kind = [string]$r.row_kind }
            $hostPath = ""
            if ($r.PSObject.Properties["host"]) { $hostPath = [string]$r.host }
            if ($kind -eq "not_workshared") {
                $bare[(ConvertTo-WorksetPathKey $hostPath)] = $true
                continue
            }
            if ($kind -eq "workset") {
                $name = ""
                $id = ""
                if ($r.PSObject.Properties["workset_name"]) { $name = [string]$r.workset_name }
                if ($r.PSObject.Properties["workset_id"]) { $id = [string]$r.workset_id }
                [void]$worksets.Add([pscustomobject]@{ Host = $hostPath; Id = $id; Name = $name })
                continue
            }
            if ($kind -eq "link") {
                $nested = $false
                if ($r.PSObject.Properties["is_nested"] -and [string]$r.is_nested -eq "1") { $nested = $true }
                $inst = ""
                $typeWs = ""
                $path = ""
                $name = ""
                if ($r.PSObject.Properties["instance_workset"]) { $inst = [string]$r.instance_workset }
                if ($r.PSObject.Properties["type_workset"]) { $typeWs = [string]$r.type_workset }
                if ($r.PSObject.Properties["link_path"]) { $path = [string]$r.link_path }
                if ($r.PSObject.Properties["link_name"]) { $name = [string]$r.link_name }
                [void]$links.Add([pscustomobject]@{
                        Host             = $hostPath
                        Path             = $path
                        Name             = $name
                        Nested           = $nested
                        InstanceWorkset  = $inst
                        TypeWorkset      = $typeWs
                    })
            }
        }
    }
    return [pscustomobject]@{
        Worksets      = $worksets.ToArray()
        Links         = $links.ToArray()
        NotWorkshared = $bare
        WorksetCount  = $worksets.Count
        LinkCount     = $links.Count
    }
}

function Get-WorksetPairKey($Pair) {
    $t = ConvertTo-WorksetPathKey ([string]$Pair.TargetPath)
    $l = ConvertTo-WorksetPathKey ([string]$Pair.LinkPath)
    if (-not $l) { $l = ConvertTo-WorksetPathKey ([string]$Pair.LinkName) }
    if (-not $t) { $t = ConvertTo-WorksetPathKey ([string]$Pair.TargetName) }
    return ($t + "|" + $l + "|" + [string]$Pair.State)
}

function New-WorksetUiPairs {
    param($NewPaths, $TargetPaths, [string]$Separator, [string[]]$Roles, $Inventory, $ManualPairs = $null)
    $NewPaths = @($NewPaths)
    $TargetPaths = @($TargetPaths)
    $built = Get-WorksetLinkPairs $NewPaths $TargetPaths $Separator $Roles
    $auto = @()
    if ($null -ne $built -and $null -ne $built.Rows) { $auto = @($built.Rows) }
    # hand-made pairs join the automatic rows here, so they get the same workset proposal
    # and the same existing-link check (exists / other_path) and go into the same plan
    $rows = @(Merge-WorksetManualPairs $auto $ManualPairs $Separator $Roles (Get-WorksetCodeTable))
    $linkItems = @()
    if ($null -ne $Inventory -and $null -ne $Inventory.Links) { $linkItems = @($Inventory.Links) }
    $masks = New-Object System.Collections.Generic.List[object]
    $seenHost = @{}
    foreach ($p in $rows) {
        if ([string]::IsNullOrWhiteSpace([string]$p.TargetPath)) { continue }
        $hk = ConvertTo-WorksetPathKey ([string]$p.TargetPath)
        if ($seenHost.ContainsKey($hk)) { continue }
        $seenHost[$hk] = $true
        $forMask = New-Object System.Collections.Generic.List[object]
        foreach ($e in $linkItems) {
            if ((ConvertTo-WorksetPathKey ([string]$e.Host)) -ne $hk) { continue }
            $ws = [string]$e.InstanceWorkset
            if (-not $ws) { $ws = [string]$e.TypeWorkset }
            [void]$forMask.Add([pscustomobject]@{
                    Name     = $(if ($e.Path) { [string]$e.Path } else { [string]$e.Name })
                    Workset  = $ws
                    Nested   = [bool]$e.Nested
                })
        }
        $mask = Get-WorksetLinkMask $forMask.ToArray() (Get-WorksetCodeTable) $Separator $Roles $null
        [void]$masks.Add([pscustomobject]@{
                Host   = [string]$p.TargetPath
                Ok     = [bool]$mask.Ok
                Mask   = [string]$mask.Mask
                Script = [string]$mask.Script
            })
    }
    if ($rows.Count -gt 0) {
        Add-WorksetProposals $rows $masks.ToArray() (Get-WorksetCodeTable)
        Update-WorksetPairInventory $rows $linkItems
    }
    foreach ($p in $rows) {
        $p.Key = Get-WorksetPairKey $p
        $p | Add-Member -NotePropertyName Manual -NotePropertyValue $false -Force
        if ($script:WsManual -and $script:WsManual.ContainsKey([string]$p.Key)) {
            $p.WorksetName = [string]$script:WsManual[[string]$p.Key]
            $p.Manual = $true
        }
    }
    return [pscustomobject]@{ Ok = [bool]$built.Ok; Count = $rows.Count; Rows = $rows }
}

function Get-WorksetCodeTable {
    if ($script:WsCodes -and $script:WsCodes.Count -gt 0) { return @($script:WsCodes) }
    return @(Get-DefaultSectionCodes)
}

function Get-WorksetRenameEditKey([string]$HostPath, [string]$Name) {
    $h = "*"
    if (-not [string]::IsNullOrWhiteSpace($HostPath)) { $h = ConvertTo-WorksetPathKey $HostPath }
    return ($h + "|" + ([string]$Name).ToLowerInvariant())
}

# Fills "Станет" from the current name. Case-sensitive. Does not change "Было".
function Update-WorksetRenamePreview {
    param($Rows, [string]$Find, [string]$Replace)
    if ($null -eq $Rows) { return 0 }
    $n = 0
    foreach ($r in $Rows) {
        $next = ConvertTo-WorksetReplacedName ([string]$r.Name) $Find $Replace
        $r.NewName = $next
        if ($script:WsRenameEdit) {
            $script:WsRenameEdit[(Get-WorksetRenameEditKey ([string]$r.Host) ([string]$r.Name))] = $next
        }
        $n++
    }
    return $n
}

# Models of the active worksets tab (each of the two tabs has its own model list).
function Get-WorksetModelList {
    if (-not $script:Ui) { return $null }
    $key = "worksets"
    if ($script:WsActiveKey) { $key = [string]$script:WsActiveKey }
    if ($script:Ui.ContainsKey("PlainLists") -and $script:Ui.PlainLists -and $script:Ui.PlainLists.ContainsKey($key)) {
        return $script:Ui.PlainLists[$key]
    }
    return $script:Ui.List
}

function Get-CheckedWorksetPaths {
    $list = New-Object System.Collections.Generic.List[string]
    $lv = Get-WorksetModelList
    if (-not $lv) { return $list.ToArray() }
    foreach ($it in $lv.Items) {
        if ($it.Checked) { [void]$list.Add([string]$it.Tag) }
    }
    return $list.ToArray()
}

function Set-WorksetActiveTab([string]$Key) {
    if ($Key -ne "addlinks" -and $Key -ne "worksets") { return }
    if ($script:WsActiveKey -eq $Key) { return }
    $script:WsActiveKey = $Key
    if (-not $script:OpsBuildOnly) { try { Update-WorksetViews } catch { } }
}

function Split-WorksetCheckedByRole {
    $news = New-Object System.Collections.Generic.List[string]
    $targets = New-Object System.Collections.Generic.List[string]
    foreach ($p in (Get-CheckedWorksetPaths)) {
        if ((Get-WorksetRoleCode $p) -eq $script:WsRoleNew) { [void]$news.Add($p) }
        else { [void]$targets.Add($p) }
    }
    # Lists stay lists: a one-element array on a pscustomobject unwraps to a string,
    # and a string then iterates as characters.
    return [pscustomobject]@{ New = $news; Target = $targets; NewCount = $news.Count; TargetCount = $targets.Count }
}

# Role column (новая / целевая) of the «Связи: добавить» model list.
function Update-WorksetRoleColumn {
    if (-not $script:Ui -or -not $script:Ui.ContainsKey("PlainLists") -or -not $script:Ui.PlainLists) { return }
    if (-not $script:Ui.PlainLists.ContainsKey("addlinks")) { return }
    foreach ($it in @($script:Ui.PlainLists["addlinks"].Items)) {
        if ($it.SubItems.Count -lt 4) { continue }
        $t = Get-WorksetRoleLabel ([string]$it.Tag)
        if ($it.SubItems[3].Text -ne $t) { $it.SubItems[3].Text = $t }
    }
}

function Save-WorksetCodeFile {
    if ($script:OpsBuildOnly) { return }
    if (-not $script:WsCodes) { return }
    $path = Join-Path $ToolDir "section_codes.csv"
    Write-SectionCodeFile $path @($script:WsCodes)
}

function Save-WorksetMaskFile {
    if (-not $script:WsMaskRoles) { return }
    $sep = "_"
    if ($script:Ui -and $script:Ui.ContainsKey("WsSep") -and $script:Ui.WsSep.Text) { $sep = $script:Ui.WsSep.Text }
    if ($script:WsSeparator) { $sep = $script:WsSeparator }
    # from now on this mask is the saved one; a recompute never replaces it
    $script:WsSavedRoles = @($script:WsMaskRoles)
    $script:WsUserRoles = @()
    if ($script:OpsBuildOnly) { return }
    Write-NameMaskFile (Join-Path $ToolDir "name_mask.cfg") $sep @($script:WsMaskRoles)
}

function Load-WorksetUserFiles {
    $codePath = Join-Path $ToolDir "section_codes.csv"
    $script:WsCodes = @()
    if ($script:OpsBuildOnly -and -not (Test-Path -LiteralPath $codePath)) {
        # snap_ui.ps1: never write into the tool folder
        $script:WsCodes = @(Get-DefaultSectionCodes)
    }
    else {
        $loaded = Initialize-SectionCodeFile $codePath
        if ($null -ne $loaded -and $null -ne $loaded.Rows) { $script:WsCodes = @($loaded.Rows) }
    }
    $maskPath = Join-Path $ToolDir "name_mask.cfg"
    $script:WsSeparator = "_"
    $script:WsMaskRoles = @()
    $script:WsSavedRoles = @()
    $script:WsUserRoles = @()
    if (Test-Path -LiteralPath $maskPath) {
        $mask = Read-NameMaskFile $maskPath
        if ($mask.Separator) { $script:WsSeparator = [string]$mask.Separator }
        # a file without exactly one section and one building is not a mask: auto-detect instead
        if ($mask.RoleCount -gt 0 -and (Test-WorksetRoleSet @($mask.Roles))) {
            $script:WsSavedRoles = @($mask.Roles)
            $script:WsMaskRoles = @($mask.Roles)
        }
    }
}

function Import-WorksetReadFile([string]$CsvPath, [string]$Note, [bool]$AfterRun = $false) {
    if (-not $CsvPath -or -not (Test-Path -LiteralPath $CsvPath)) {
        Set-UiStatus "Файл отчёта чтения не найден."
        return
    }
    $csv = Read-WorksetCsv $CsvPath
    $script:WsInventory = ConvertTo-WorksetInventoryFromRows $csv.Rows
    # the models of the report go into the model list of this tab (the year is then detected from them)
    $hosts = New-Object System.Collections.Generic.List[string]
    $seen = @{}
    foreach ($r in @($csv.Rows)) {
        if (-not $r.PSObject.Properties["host"]) { continue }
        $h = [string]$r.host
        $k = ConvertTo-WorksetPathKey $h
        if (-not $k -or $seen.ContainsKey($k)) { continue }
        $seen[$k] = $true
        [void]$hosts.Add($h)
    }
    $added = 0
    $key = "worksets"
    if ($script:WsActiveKey) { $key = [string]$script:WsActiveKey }
    if ($hosts.Count -gt 0 -and (Get-Command Add-PlainTabModels -ErrorAction SilentlyContinue)) {
        $added = Add-PlainTabModels $key $hosts.ToArray()
    }
    $bare = 0
    if ($script:WsInventory.NotWorkshared) { $bare = $script:WsInventory.NotWorkshared.Count }
    $where = $CsvPath
    if ($Note) { $where = ("{0} ({1})" -f $CsvPath, $Note) }
    # the status line keeps one short line; the report path and the read notes go to its tool tip
    $text = "Прочитано наборов: {0}, связей: {1}, несовместных моделей: {2}." -f $script:WsInventory.WorksetCount, $script:WsInventory.LinkCount, $bare
    if ($added -gt 0) { $text = ("Моделей из отчёта добавлено в список: {0}. " -f $added) + $text }
    $tip = $text + [Environment]::NewLine + $where
    if ($AfterRun) {
        $modeNote = Get-WorksetReadStatusNote $script:WsReadMode $script:WsInventory.WorksetCount $script:WsInventory.LinkCount
        if ($modeNote) { $tip = $tip + [Environment]::NewLine + $modeNote }
    }
    # pairs are proposed right away (no extra button); the timer started by the new rows is not needed
    Update-WorksetViews
    Stop-WorksetRecompute
    if ($script:WsActiveKey -eq "addlinks" -and $script:WsPairNote) {
        $tip = $script:WsPairNote + [Environment]::NewLine + $tip
        $text = Join-WorksetStatus $text $script:WsPairNote
    }
    Set-UiStatus $text $tip
}

# Pointer / newest report: the tool's отчёты folder first, then the old rbp\reports.
function Find-WorksetReportAnywhere([string]$PointerName, [string]$Prefix) {
    foreach ($dir in @(Get-OpsReportsReadDirs $ToolDir)) {
        $found = Resolve-WorksetReportFile (Join-Path $dir $PointerName) $Prefix
        if ($found.Ok) { return $found }
    }
    return [pscustomobject]@{ Ok = $false; Path = ""; Fallback = $false }
}

function Get-WorksetReportDialogDir {
    foreach ($dir in @(Get-OpsReportsReadDirs $ToolDir)) {
        if (Test-Path -LiteralPath $dir) { return $dir }
    }
    return ""
}

function Import-WorksetReadReport {
    $found = Find-WorksetReportAnywhere "_active_workset_read.txt" "workset_read"
    if (-not $found.Ok) {
        Set-UiStatus "Файл отчёта чтения не найден."
        return
    }
    $note = ""
    if ($found.Fallback) { $note = "последний отчёт чтения" }
    Import-WorksetReadFile $found.Path $note $true
}

function Show-WorksetReadReportDialog {
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Title = "Открыть отчёт чтения"
    $dlg.Filter = "Отчёт чтения (*.csv)|*.csv|Все файлы (*.*)|*.*"
    $dir = Get-WorksetReportDialogDir
    if ($dir) { $dlg.InitialDirectory = $dir }
    if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
    Import-WorksetReadFile $dlg.FileName ""
}

# Hand-made pairs that still make sense: both models are in the tab's model list.
function Get-WorksetManualPairsInList {
    $list = New-Object System.Collections.Generic.List[object]
    if (-not $script:WsManualPairs -or $script:WsManualPairs.Count -eq 0) { return $list.ToArray() }
    $key = "addlinks"
    if ($script:WsActiveKey) { $key = [string]$script:WsActiveKey }
    if ($key -ne "addlinks") { return $list.ToArray() }
    $inList = $null
    if ($script:Ui -and $script:Ui.ContainsKey("PlainLists") -and $script:Ui.PlainLists -and $script:Ui.PlainLists.ContainsKey("addlinks")) {
        $inList = @{}
        foreach ($p in @(Get-WorksetTabPaths "addlinks" $false)) { $inList[(ConvertTo-WorksetPathKey $p)] = $true }
    }
    foreach ($m in $script:WsManualPairs) {
        if ($null -ne $inList) {
            if (-not $inList.ContainsKey((ConvertTo-WorksetPathKey ([string]$m.TargetPath)))) { continue }
            if (-not $inList.ContainsKey((ConvertTo-WorksetPathKey ([string]$m.LinkPath)))) { continue }
        }
        [void]$list.Add($m)
    }
    return $list.ToArray()
}

function Add-WorksetManualPair([string]$TargetPath, [string]$LinkPath) {
    if ([string]::IsNullOrWhiteSpace($TargetPath) -or [string]::IsNullOrWhiteSpace($LinkPath)) { return $false }
    $tk = ConvertTo-WorksetPathKey $TargetPath
    $lk = ConvertTo-WorksetPathKey $LinkPath
    if ($tk -eq $lk) { return $false }
    if (-not $script:WsManualPairs) { $script:WsManualPairs = New-Object System.Collections.Generic.List[object] }
    foreach ($m in $script:WsManualPairs) {
        if ((ConvertTo-WorksetPathKey ([string]$m.TargetPath)) -eq $tk -and (ConvertTo-WorksetPathKey ([string]$m.LinkPath)) -eq $lk) { return $false }
    }
    [void]$script:WsManualPairs.Add([pscustomobject]@{ TargetPath = $TargetPath.Trim(); LinkPath = $LinkPath.Trim() })
    return $true
}

# Removes the hand-made rows with these keys. Automatic rows are not removed (count returned apart).
function Remove-WorksetManualRows([string[]]$Keys) {
    $removed = 0
    $autoHit = 0
    if (-not $script:WsPairResult -or $null -eq $script:WsPairResult.Rows) { return [pscustomobject]@{ Removed = 0; Auto = 0 } }
    $want = @{}
    foreach ($k in $Keys) { if ($k) { $want[[string]$k] = $true } }
    foreach ($r in @($script:WsPairResult.Rows)) {
        if (-not $want.ContainsKey([string]$r.Key)) { continue }
        $src = ""
        if ($r.PSObject.Properties["Source"]) { $src = [string]$r.Source }
        if ($src -ne "manual") { $autoHit++; continue }
        $tk = ConvertTo-WorksetPathKey ([string]$r.TargetPath)
        $lk = ConvertTo-WorksetPathKey ([string]$r.LinkPath)
        for ($i = $script:WsManualPairs.Count - 1; $i -ge 0; $i--) {
            $m = $script:WsManualPairs[$i]
            if ((ConvertTo-WorksetPathKey ([string]$m.TargetPath)) -eq $tk -and (ConvertTo-WorksetPathKey ([string]$m.LinkPath)) -eq $lk) {
                $script:WsManualPairs.RemoveAt($i)
                $removed++
            }
        }
        if ($script:WsManual) { $script:WsManual.Remove([string]$r.Key) }
    }
    return [pscustomobject]@{ Removed = $removed; Auto = $autoHit }
}

# Pairs of the ticked models. The mask: roles edited on «Маска имени» > name_mask.cfg >
# section codes > the older guess. A saved mask is never replaced here. When nothing fits,
# the rows still show up as «имя не разобрано» and $script:WsPairNote says why.
function Build-WorksetPairsFromUi {
    $sep = "_"
    if ($script:WsSeparator) { $sep = $script:WsSeparator }
    $split = Split-WorksetCheckedByRole
    $newNames = New-Object System.Collections.Generic.List[string]
    $targetNames = New-Object System.Collections.Generic.List[string]
    foreach ($p in $split.New) { [void]$newNames.Add((Get-WorksetBaseName ([string]$p))) }
    foreach ($p in $split.Target) { [void]$targetNames.Add((Get-WorksetBaseName ([string]$p))) }
    $user = @()
    if ($script:WsUserRoles) { $user = @($script:WsUserRoles) }
    $saved = @()
    if ($script:WsSavedRoles) { $saved = @($script:WsSavedRoles) }
    $res = Resolve-WorksetNameMask $user $saved $newNames.ToArray() $targetNames.ToArray() (Get-WorksetCodeTable) $sep
    $script:WsMaskResolved = $res
    $roles = @()
    if ($res.Ok) { $roles = @($res.Roles) }
    # what «Маска имени» shows: the roles in use, or the user's unfinished edit
    if ($res.Ok) { $script:WsMaskRoles = @($res.Roles) }
    elseif ($user.Count -gt 0) { $script:WsMaskRoles = @($user) }
    else { $script:WsMaskRoles = @() }
    $manual = @(Get-WorksetManualPairsInList)
    $all = New-WorksetUiPairs $split.New $split.Target $sep $roles $script:WsInventory $manual
    # no ticked new model: an empty table, not a red «нет пары» row per target;
    # with new models «нет пары» is shown for new models only
    $shown = Select-WorksetVisiblePairs @($all.Rows) $split.NewCount
    $script:WsPairResult = [pscustomobject]@{ Ok = [bool]$all.Ok; Count = $shown.Count; Rows = $shown.Rows }
    $script:WsPairNote = Get-WorksetPairSummary @($shown.Rows) $res $split.NewCount $split.TargetCount
}

# Recompute after a tick / untick or a role change, once the clicks stop (300 ms).
function Request-WorksetRecompute {
    if (-not $script:Ui -or -not $script:Ui.ContainsKey("WsRecalcTimer") -or -not $script:Ui.WsRecalcTimer) { return }
    $script:Ui.WsRecalcTimer.Stop()
    $script:Ui.WsRecalcTimer.Start()
}

function Stop-WorksetRecompute {
    if (-not $script:Ui -or -not $script:Ui.ContainsKey("WsRecalcTimer") -or -not $script:Ui.WsRecalcTimer) { return }
    $script:Ui.WsRecalcTimer.Stop()
}

function Invoke-WorksetRecompute {
    Stop-WorksetRecompute
    if ($script:Run -and $script:Run.Active) { return }
    Update-WorksetViews
    if ($script:WsActiveKey -eq "addlinks" -and $script:WsPairNote) { Set-UiStatus $script:WsPairNote }
}

function Update-WorksetLinkPaint($Item, $Row) {
    $paint = Get-WorksetLinkPaint ([string]$Row.State) ([string]$Row.WorksetName) ([bool]$Row.Manual)
    $missing = [System.Drawing.Color]::MistyRose
    $empty = [System.Drawing.Color]::FromArgb(255, 230, 150)
    $manual = [System.Drawing.Color]::FromArgb(255, 245, 190)
    $normal = [System.Drawing.SystemColors]::Window
    $dim = [System.Drawing.Color]::Gray
    $text = [System.Drawing.SystemColors]::WindowText
    foreach ($sub in $Item.SubItems) {
        $sub.BackColor = $(if ($paint.Row -eq "missing") { $missing } else { $normal })
        $sub.ForeColor = $(if ($paint.Dim) { $dim } else { $text })
    }
    if ($Item.SubItems.Count -gt 4) {
        if ($paint.Cell -eq "empty") { $Item.SubItems[4].BackColor = $empty }
        elseif ($paint.Cell -eq "manual") { $Item.SubItems[4].BackColor = $manual }
    }
}

function Update-WorksetLinkView {
    if (-not $script:Ui -or -not $script:Ui.ContainsKey("LinkList") -or -not $script:Ui.LinkList) { return }
    $lv = $script:Ui.LinkList
    $lv.BeginUpdate()
    try {
        $lv.Items.Clear()
        $rows = @()
        if ($script:WsPairResult -and $null -ne $script:WsPairResult.Rows) { $rows = @($script:WsPairResult.Rows) }
        foreach ($r in $rows) {
            $target = [string]$r.TargetName
            if (-not $target) { $target = "(нет пары)" }
            $it = New-Object System.Windows.Forms.ListViewItem $target
            $it.UseItemStyleForSubItems = $false
            $link = [string]$r.LinkName
            if (-not $link) { $link = [string]$r.LinkPath }
            [void]$it.SubItems.Add($link)
            [void]$it.SubItems.Add([string]$r.Section)
            [void]$it.SubItems.Add([string]$r.Building)
            [void]$it.SubItems.Add([string]$r.WorksetName)
            $stateText = Get-WorksetStateLabel ([string]$r.State)
            $applied = Get-WorksetLinkApplyNote ([string]$r.TargetPath) ([string]$r.LinkPath)
            if ($applied) { $stateText = $applied }
            if ($r.PSObject.Properties["Source"] -and [string]$r.Source -eq "manual") { $stateText = $stateText + " (вручную)" }
            [void]$it.SubItems.Add($stateText)
            $it.Tag = [string]$r.Key
            Update-WorksetLinkPaint $it $r
            [void]$lv.Items.Add($it)
        }
    }
    finally { $lv.EndUpdate() }
}

function Update-WorksetRenameView {
    if (-not $script:Ui -or -not $script:Ui.ContainsKey("RenameList")) { return }
    $lv = $script:Ui.RenameList
    $hostFilter = ""
    if ($script:Ui.ContainsKey("WsFilter") -and $script:Ui.WsFilter.SelectedIndex -gt 0) {
        $i = $script:Ui.WsFilter.SelectedIndex
        if ($script:WsFilterPaths -and $i -le $script:WsFilterPaths.Count) { $hostFilter = [string]$script:WsFilterPaths[$i - 1] }
    }
    $src = New-Object System.Collections.Generic.List[object]
    if ($script:WsInventory -and $script:WsInventory.Worksets) {
        $checked = @{}
        foreach ($p in (Get-CheckedWorksetPaths)) { $checked[(ConvertTo-WorksetPathKey $p)] = $p }
        foreach ($w in @($script:WsInventory.Worksets)) {
            $hk = ConvertTo-WorksetPathKey ([string]$w.Host)
            if (-not $checked.ContainsKey($hk)) { continue }
            [void]$src.Add($w)
        }
    }
    $grouped = Group-WorksetRename $src.ToArray() $hostFilter
    $rows = @()
    if ($null -ne $grouped.Rows) { $rows = @($grouped.Rows) }
    foreach ($r in $rows) {
        $key = Get-WorksetRenameEditKey ([string]$r.Host) ([string]$r.Name)
        $newName = [string]$r.Name
        if ($script:WsRenameEdit.ContainsKey($key)) { $newName = [string]$script:WsRenameEdit[$key] }
        $r | Add-Member -NotePropertyName NewName -NotePropertyValue $newName -Force
    }
    $edits = New-Object System.Collections.Generic.List[object]
    foreach ($r in $rows) {
        if (-not (Test-WorksetSameText ([string]$r.Name) ([string]$r.NewName))) {
            [void]$edits.Add([pscustomobject]@{ OldName = [string]$r.Name; NewName = [string]$r.NewName })
        }
    }
    $conf = Get-WorksetRenameConflicts $src.ToArray() $edits.ToArray()
    $byOld = @{}
    foreach ($c in @($conf.Rows)) { $byOld[([string]$c.OldName).ToLowerInvariant()] = $c }
    $delete = @{}
    foreach ($x in @($script:WsDeleteNames)) { $delete[([string]$x).ToLowerInvariant()] = $true }
    $withElements = ($script:Ui.ContainsKey("WsDeleteElements") -and $script:Ui.WsDeleteElements -and $script:Ui.WsDeleteElements.Checked)
    $script:WsRenameRows = $rows
    $red = [System.Drawing.Color]::Firebrick
    $lv.BeginUpdate()
    try {
        $lv.Items.Clear()
        foreach ($r in $rows) {
            # 0 Модель, 1 Рабочий набор сейчас, 2 Станет, 3 Элементов, 4 Действие
            $model = ""
            if ($r.Host) { $model = Get-WorksetBaseName ([string]$r.Host) }
            elseif ([int]$r.HostCount -eq 1) { $model = Get-WorksetBaseName ([string]$r.HostList) }
            else { $model = ("моделей: {0}" -f $r.HostCount) }
            $it = New-Object System.Windows.Forms.ListViewItem $model
            $it.UseItemStyleForSubItems = $false
            $it.Name = [string]$r.Name
            [void]$it.SubItems.Add([string]$r.Name)
            [void]$it.SubItems.Add([string]$r.NewName)
            [void]$it.SubItems.Add("—")
            $action = ""
            $renamed = -not (Test-WorksetSameText ([string]$r.Name) ([string]$r.NewName))
            $ck = ([string]$r.Name).ToLowerInvariant()
            if ($delete.ContainsKey($ck)) {
                if ($withElements) { $action = "удалить ВМЕСТЕ С ЭЛЕМЕНТАМИ" }
                else { $action = ("удалить, элементы → {0}" -f $script:WorksetDefaultIgnoreName) }
            }
            elseif ($renamed) { $action = "переименовать" }
            if ($byOld.ContainsKey($ck) -and [int]$byOld[$ck].ConflictCount -gt 0) {
                $action = ("конфликт имени: в {0} из {1}" -f $byOld[$ck].ConflictCount, $byOld[$ck].HostCount)
                $it.SubItems[2].ForeColor = $red
            }
            $applied = Get-WorksetApplyNote ([string]$r.Name) ([string]$r.Host) ([string]$r.HostList)
            if ($applied) {
                $action = $applied
                if ($applied -eq "записано") { $it.SubItems[2].ForeColor = [System.Drawing.Color]::DarkGreen }
                else { $it.SubItems[2].ForeColor = $red }
            }
            [void]$it.SubItems.Add($action)
            if ($action.StartsWith("удалить ВМЕСТЕ") -or $action.StartsWith("конфликт")) { $it.SubItems[4].ForeColor = $red }
            if ($renamed) { $it.SubItems[2].BackColor = [System.Drawing.Color]::FromArgb(255, 245, 190) }
            $it.Tag = Get-WorksetRenameEditKey ([string]$r.Host) ([string]$r.Name)
            [void]$lv.Items.Add($it)
        }
    }
    finally { $lv.EndUpdate() }
}

function Update-WorksetCodeView {
    if (-not $script:Ui -or -not $script:Ui.ContainsKey("CodeList")) { return }
    $lv = $script:Ui.CodeList
    $lv.BeginUpdate()
    try {
        $lv.Items.Clear()
        foreach ($c in @($script:WsCodes)) {
            $it = New-Object System.Windows.Forms.ListViewItem ([string]$c.Latin)
            [void]$it.SubItems.Add([string]$c.Cyrillic)
            [void]$lv.Items.Add($it)
        }
    }
    finally { $lv.EndUpdate() }
}

function Update-WorksetMaskView {
    if (-not $script:Ui -or -not $script:Ui.ContainsKey("MaskList")) { return }
    if ($script:Ui.ContainsKey("WsSep") -and $script:WsSeparator) { $script:Ui.WsSep.Text = $script:WsSeparator }
    $example = @()
    $split = Split-WorksetCheckedByRole
    $sample = ""
    if ($split.NewCount -gt 0) { $sample = [string]$split.New[0] }
    elseif ($split.TargetCount -gt 0) { $sample = [string]$split.Target[0] }
    if ($sample) { $example = @(Split-WorksetModelTokens (Get-WorksetBaseName $sample) $script:WsSeparator) }
    $lv = $script:Ui.MaskList
    $lv.BeginUpdate()
    try {
        $lv.Items.Clear()
        $n = $script:WsMaskRoles.Count
        if ($example.Count -gt $n) { $n = $example.Count }
        if ($n -eq 0) { $n = $example.Count }
        $words = New-Object System.Collections.Generic.List[string]
        for ($i = 0; $i -lt $n; $i++) {
            $role = "fixed"
            if ($i -lt @($script:WsMaskRoles).Count) { $role = [string]@($script:WsMaskRoles)[$i] }
            $tok = ""
            if ($i -lt $example.Count) { $tok = [string]$example[$i] }
            $it = New-Object System.Windows.Forms.ListViewItem ([string]($i + 1))
            [void]$it.SubItems.Add($tok)
            [void]$it.SubItems.Add((Get-WorksetRoleWord $role))
            $it.Tag = (ConvertTo-WorksetRole $role)
            if (-not $it.Tag) { $it.Tag = "fixed" }
            [void]$lv.Items.Add($it)
            [void]$words.Add($(if ($tok) { $tok } else { "?" }))
        }
        if ($script:Ui.ContainsKey("WsMaskPreview")) {
            $src = "none"
            if ($script:WsMaskResolved) { $src = [string]$script:WsMaskResolved.Source }
            if ($script:WsMaskResolved -and -not $script:WsMaskResolved.Ok) { $src = "none" }
            $script:Ui.WsMaskPreview.Text = ("{0}   ·   маска: {1}" -f ($words.ToArray() -join " _ "), (Get-WorksetMaskSourceWord $src))
        }
    }
    finally { $lv.EndUpdate() }
}

# «Модель: [все отмеченные ▾]» - names in the list, full paths kept aside.
function Update-WorksetFilter {
    if (-not $script:Ui -or -not $script:Ui.ContainsKey("WsFilter")) { return }
    $cmb = $script:Ui.WsFilter
    $prevPath = ""
    if ($cmb.SelectedIndex -gt 0 -and $script:WsFilterPaths -and $cmb.SelectedIndex -le $script:WsFilterPaths.Count) {
        $prevPath = [string]$script:WsFilterPaths[$cmb.SelectedIndex - 1]
    }
    $paths = New-Object System.Collections.Generic.List[string]
    $cmb.BeginUpdate()
    try {
        $cmb.Items.Clear()
        [void]$cmb.Items.Add("все отмеченные")
        foreach ($p in (Get-CheckedWorksetPaths)) {
            [void]$paths.Add([string]$p)
            [void]$cmb.Items.Add((Get-WorksetBaseName ([string]$p)))
        }
    }
    finally { $cmb.EndUpdate() }
    $script:WsFilterPaths = $paths
    $idx = 0
    if ($prevPath) {
        for ($i = 0; $i -lt $paths.Count; $i++) { if ($paths[$i] -eq $prevPath) { $idx = $i + 1; break } }
    }
    $script:WsFilterLock = $true
    try { $cmb.SelectedIndex = $idx } finally { $script:WsFilterLock = $false }
}

function Update-WorksetViews {
    Build-WorksetPairsFromUi
    Update-WorksetLinkView
    Update-WorksetFilter
    Update-WorksetRenameView
    Update-WorksetCodeView
    Update-WorksetMaskView
    Update-WorksetYearLabel
    Update-WorksetExistingLinks
}

function Read-WorksetCodesFromView {
    if (-not $script:Ui.ContainsKey("CodeList")) { return }
    $list = New-Object System.Collections.Generic.List[object]
    foreach ($it in $script:Ui.CodeList.Items) {
        $lat = [string]$it.Text.Trim()
        $cyr = ""
        if ($it.SubItems.Count -gt 1) { $cyr = [string]$it.SubItems[1].Text.Trim() }
        if (-not $lat -or -not $cyr) { continue }
        [void]$list.Add([pscustomobject]@{ Latin = $lat; Cyrillic = $cyr })
    }
    $script:WsCodes = $list.ToArray()
    Save-WorksetCodeFile
    # the codes decide the mask when none is saved
    Request-WorksetRecompute
}

function Read-WorksetRolesFromView {
    if (-not $script:Ui.ContainsKey("MaskList")) { return }
    $list = New-Object System.Collections.Generic.List[string]
    foreach ($it in $script:Ui.MaskList.Items) {
        $role = [string]$it.Tag
        if (-not $role) { $role = "fixed" }
        [void]$list.Add($role)
    }
    $script:WsMaskRoles = $list.ToArray()
    # an edit on «Маска имени» is used right away; «Запомнить маску» writes it to name_mask.cfg
    $script:WsUserRoles = $list.ToArray()
    if ($script:Ui.ContainsKey("WsSep") -and $script:Ui.WsSep.Text) { $script:WsSeparator = $script:Ui.WsSep.Text }
}

function Close-WorksetEditor([bool]$Commit) {
    $st = $script:WsEdit
    $script:WsEdit = $null
    $ed = $null
    if ($script:Ui.ContainsKey("WsEditor")) { $ed = $script:Ui.WsEditor }
    if ($ed) { $ed.Visible = $false }
    if (-not $Commit -or -not $st) { return }
    $text = ""
    if ($ed) { $text = $ed.Text.Trim() }
    if ($st.Kind -eq "link-workset") {
        if (-not $script:WsPairResult -or $null -eq $script:WsPairResult.Rows) { return }
        foreach ($r in @($script:WsPairResult.Rows)) {
            if ([string]$r.Key -ne [string]$st.Key) { continue }
            if ($text -eq "") {
                $script:WsManual.Remove([string]$r.Key)
                $r.Manual = $false
                $r.WorksetName = ""
            }
            else {
                $chk = Test-WorksetNameText $text
                if (-not $chk.Ok) { Set-UiStatus $chk.Message; return }
                $script:WsManual[[string]$r.Key] = $text
                $r.WorksetName = $text
                $r.Manual = $true
            }
        }
        Update-WorksetLinkView
        # the status line recounts the rows still without a workset
        $res = $script:WsMaskResolved
        $split = Split-WorksetCheckedByRole
        $script:WsPairNote = Get-WorksetPairSummary @($script:WsPairResult.Rows) $res $split.NewCount $split.TargetCount
        if ($script:WsPairNote) { Set-UiStatus $script:WsPairNote }
        return
    }
    if ($st.Kind -eq "link-path") {
        if (-not $script:WsPairResult -or $null -eq $script:WsPairResult.Rows) { return }
        foreach ($r in @($script:WsPairResult.Rows)) {
            if ([string]$r.Key -ne [string]$st.Key) { continue }
            $r.LinkPath = $text
            $r.LinkName = Get-WorksetBaseName $text
        }
        if ($script:WsInventory) { Update-WorksetPairInventory @($script:WsPairResult.Rows) @($script:WsInventory.Links) }
        Update-WorksetLinkView
        return
    }
    if ($st.Kind -eq "rename") {
        $script:WsRenameEdit[[string]$st.Key] = $text
        Update-WorksetRenameView
        return
    }
    if ($st.Kind -eq "code") {
        if ($st.Item -and $st.Item.SubItems.Count -gt $st.Col) {
            $st.Item.SubItems[$st.Col].Text = $text
            if ($st.Col -eq 0) { $st.Item.Text = $text }
        }
        Read-WorksetCodesFromView
    }
}

# The workset of a row can be typed when that row is (or would be) an insert.
function Test-WorksetRowEditable([string]$Key) {
    if (-not $script:WsPairResult -or $null -eq $script:WsPairResult.Rows) { return $false }
    foreach ($r in @($script:WsPairResult.Rows)) {
        if ([string]$r.Key -ne $Key) { continue }
        return ([string]$r.State -eq "insert" -and [bool]$r.TargetPath -and [bool]$r.LinkPath)
    }
    return $false
}

function Show-WorksetGridEditor($List, $Item, [int]$Col, [string]$Kind) {
    if ($script:Run.Active) { return }
    if (-not $script:Ui.ContainsKey("WsEditor") -or -not $script:Ui.WsEditor) {
        $ed = New-Object System.Windows.Forms.TextBox
        $ed.BorderStyle = "FixedSingle"
        $ed.Visible = $false
        $ed.Add_KeyDown({
                if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
                    $_.Handled = $true
                    $_.SuppressKeyPress = $true
                    Close-WorksetEditor $true
                }
                elseif ($_.KeyCode -eq [System.Windows.Forms.Keys]::Escape) {
                    $_.Handled = $true
                    $_.SuppressKeyPress = $true
                    Close-WorksetEditor $false
                }
            })
        $ed.Add_Leave({ Close-WorksetEditor $true })
        $script:Ui.WsEditor = $ed
    }
    $ed = $script:Ui.WsEditor
    if ($ed.Parent -ne $List) {
        if ($ed.Parent) { $ed.Parent.Controls.Remove($ed) }
        $List.Controls.Add($ed)
    }
    $b = $Item.SubItems[$Col].Bounds
    $ed.SetBounds($b.X, $b.Y, [Math]::Max($b.Width, 80), $b.Height)
    $ed.Font = $List.Font
    $ed.Text = [string]$Item.SubItems[$Col].Text
    $script:WsEdit = @{ Kind = $Kind; Key = [string]$Item.Tag; Col = $Col; Item = $Item; Opened = [DateTime]::UtcNow }
    $ed.Visible = $true
    $ed.BringToFront()
    $ed.Focus()
    $ed.SelectAll()
}

function Get-CheckedInventoryWorksets {
    $checked = @{}
    foreach ($p in (Get-CheckedWorksetPaths)) { $checked[(ConvertTo-WorksetPathKey $p)] = $true }
    $list = New-Object System.Collections.Generic.List[object]
    if ($script:WsInventory -and $script:WsInventory.Worksets) {
        foreach ($w in @($script:WsInventory.Worksets)) {
            if ($checked.ContainsKey((ConvertTo-WorksetPathKey ([string]$w.Host)))) { [void]$list.Add($w) }
        }
    }
    return $list.ToArray()
}

function Get-WorksetApplyEdits {
    param($Worksets)
    $list = New-Object System.Collections.Generic.List[object]
    if (-not $script:WsRenameEdit -or $null -eq $Worksets) { return $list.ToArray() }
    foreach ($key in @($script:WsRenameEdit.Keys)) {
        $newName = [string]$script:WsRenameEdit[$key]
        $bar = ([string]$key).IndexOf("|")
        if ($bar -lt 1) { continue }
        $hostKey = ([string]$key).Substring(0, $bar)
        $nameKey = ([string]$key).Substring($bar + 1)
        foreach ($w in @($Worksets)) {
            if (([string]$w.Name).ToLowerInvariant() -ne $nameKey) { continue }
            $hk = ConvertTo-WorksetPathKey ([string]$w.Host)
            if ($hostKey -ne "*" -and $hostKey -ne $hk) { continue }
            [void]$list.Add([pscustomobject]@{
                    Host    = [string]$w.Host
                    Id      = [string]$w.Id
                    OldName = [string]$w.Name
                    NewName = $newName
                })
        }
    }
    return $list.ToArray()
}

function Get-WorksetNeededNames {
    param($Pairs)
    $list = New-Object System.Collections.Generic.List[object]
    $seen = @{}
    if ($null -eq $Pairs) { return $list.ToArray() }
    foreach ($p in @($Pairs)) {
        if ([string]$p.State -ne "insert") { continue }
        $name = [string]$p.WorksetName
        $hostPath = [string]$p.TargetPath
        if (-not $name.Trim() -or -not $hostPath) { continue }
        $k = (ConvertTo-WorksetPathKey $hostPath) + "|" + $name.Trim().ToLowerInvariant()
        if ($seen.ContainsKey($k)) { continue }
        $seen[$k] = $true
        [void]$list.Add([pscustomobject]@{ Host = $hostPath; Name = $name.Trim() })
    }
    return $list.ToArray()
}

function Import-WorksetApplyReport {
    $found = Find-WorksetReportAnywhere "_active_workset_apply.txt" "workset_apply"
    if (-not $found.Ok) {
        Set-UiStatus "Файл отчёта записи наборов не найден."
        return
    }
    $csvPath = $found.Path
    $csv = Read-WorksetCsv $csvPath
    $script:WsApplyById = @{}
    $script:WsApplyByPath = @{}
    $ok = 0
    $err = 0
    $skip = 0
    foreach ($r in @($csv.Rows)) {
        $action = ""
        $result = ""
        $code = ""
        $key = ""
        if ($r.PSObject.Properties["action"]) { $action = [string]$r.action }
        if ($r.PSObject.Properties["result"]) { $result = [string]$r.result }
        if ($r.PSObject.Properties["code"]) { $code = [string]$r.code }
        if ($r.PSObject.Properties["key"]) { $key = [string]$r.key }
        if ($result -eq "OK") { $ok++ }
        elseif ($result -eq "ERROR") { $err++ }
        else { $skip++ }
        if ($action -eq "rename" -and $key) {
            $script:WsApplyById[$key] = $(if ($result -eq "OK") { "OK" } else { $code })
        }
        if ($action -eq "insert") {
            $hostKey = ""
            $linkKey = ""
            if ($r.PSObject.Properties["host"]) { $hostKey = ConvertTo-WorksetPathKey ([string]$r.host) }
            if ($r.PSObject.Properties["target"]) { $linkKey = ConvertTo-WorksetPathKey ([string]$r.target) }
            if (-not $script:WsApplyByPath) { $script:WsApplyByPath = @{} }
            $script:WsApplyByPath[($hostKey + "|" + $linkKey)] = $(if ($result -eq "OK") { "OK" } else { $code })
        }
    }
    Update-WorksetRenameView
    Update-WorksetLinkView
    $where = $csvPath
    if ($found.Fallback) { $where = ("{0} (последний отчёт записи)" -f $csvPath) }
    $text = "Записано {0}, пропусков {1}, ошибок {2}." -f $ok, $skip, $err
    Set-UiStatus $text ($text + [Environment]::NewLine + $where)
}

function Start-WorksetApply {
    if ($script:Run.Active) { return }
    if (-not $script:WsInventory -or -not $script:WsInventory.Worksets) {
        [void][System.Windows.Forms.MessageBox]::Show($script:Ui.Form, "Сначала прочитайте наборы и связи.", "Запуск")
        return
    }
    Build-WorksetPairsFromUi
    $worksets = @(Get-CheckedInventoryWorksets)
    $edits = @(Get-WorksetApplyEdits $worksets)
    $pairs = @()
    if ($script:WsPairResult -and $null -ne $script:WsPairResult.Rows) { $pairs = @($script:WsPairResult.Rows) }
    # each tab checks only what it sends: links on tab 4, worksets on tab 5
    if ($script:WsApplyPart -eq "insert") { $edits = @() }
    elseif ($script:WsApplyPart -eq "sets") { $pairs = @() }
    $block = Test-WorksetStep4Block $pairs $edits
    if ($block.Block) {
        [void][System.Windows.Forms.MessageBox]::Show($script:Ui.Form, $block.Message, "Запуск")
        Set-UiStatus $block.Message
        return
    }
    $needed = @(Get-WorksetNeededNames $pairs)
    $plan = Get-WorksetRenameCreatePlan $worksets $edits $needed $pairs
    $rowsNow = New-Object System.Collections.Generic.List[object]
    foreach ($a in @($plan.Rows)) { [void]$rowsNow.Add($a) }
    if ($script:WsApplyPart -eq "sets" -and $script:WsDeleteNames -and $script:WsDeleteNames.Count -gt 0) {
        $delMode = "move"
        if ($script:Ui.ContainsKey("WsDeleteElements") -and $script:Ui.WsDeleteElements -and $script:Ui.WsDeleteElements.Checked) {
            $delMode = "elements"
        }
        foreach ($a in @(Get-WorksetDeleteActions $worksets $script:WsDeleteNames.ToArray() $delMode)) {
            [void]$rowsNow.Add($a)
        }
    }
    $plan = [pscustomobject]@{ Count = $rowsNow.Count; ApplyCount = 0; Rows = $rowsNow.ToArray() }
    foreach ($a in $rowsNow) {
        if ($a.Action -eq "rename" -or $a.Action -eq "create" -or $a.Action -eq "insert" -or $a.Action -eq "delete") {
            $plan.ApplyCount = [int]$plan.ApplyCount + 1
        }
    }
    $text = Get-WorksetStep4Text $plan.Rows
    $extraLines = New-Object System.Collections.Generic.List[string]
    foreach ($a in $rowsNow) {
        if ($a.Action -ne "delete") { continue }
        [void]$extraLines.Add((Get-WorksetDeleteDryLine ([string]$a.Host) ([string]$a.Arg1) ([string]$a.Arg2) -1))
    }
    if ($extraLines.Count -gt 0) { $text = $text + [Environment]::NewLine + ($extraLines.ToArray() -join [Environment]::NewLine) }
    # rows without a pair are not inserted and do not stop the others: listed so nothing is hidden
    $loose = New-Object System.Collections.Generic.List[string]
    foreach ($p in $pairs) {
        $st = [string]$p.State
        if ($st -ne "no_pair" -and $st -ne "unparsed") { continue }
        $nm = [string]$p.LinkName
        if (-not $nm) { $nm = [string]$p.TargetName }
        [void]$loose.Add(("  {0} — {1}" -f $nm, (Get-WorksetStateLabel $st)))
    }
    if ($loose.Count -gt 0) {
        $text = $text + [Environment]::NewLine + "Без пары (не вставляются):" + [Environment]::NewLine + ($loose.ToArray() -join [Environment]::NewLine)
    }
    $title = "Рабочие наборы"
    if ($script:WsApplyPart -eq "insert") { $title = "Вставка связей" }
    $insertView = $null
    if ($script:WsApplyPart -eq "insert") { $insertView = Get-WorksetInsertView @($plan.Rows) $pairs }
    if ($script:DryRunOnly -and $insertView) {
        [void](Show-WorksetInsertDialog $insertView "Пробный прогон" $true)
        Set-UiStatus "Пробный прогон: ничего не запускалось."
        return
    }
    if ($script:DryRunOnly) {
        [void](Show-OpsPlanDialog -Title ("Пробный прогон: " + $title) -Text ("Пробный прогон: в модели ничего не записывается." + [Environment]::NewLine + [Environment]::NewLine + $text) -CanApply $false -YesText "" -InfoOnly $true)
        Set-UiStatus "Пробный прогон: ничего не запускалось."
        return
    }
    $can = ([int]$plan.ApplyCount -gt 0)
    if ($can) {
        $paths = New-Object System.Collections.Generic.List[string]
        $seen = @{}
        foreach ($a in @($plan.Rows)) { if ($a.Host) { [void]$paths.Add([string]$a.Host) } }
        $year = Resolve-WorksetReadYear $paths.ToArray() $script:servers ""
        if (-not $year.Ok) {
            [void][System.Windows.Forms.MessageBox]::Show($script:Ui.Form, (Get-WorksetYearText $year), "Запуск")
            return
        }
        $exe = Find-BatchRvt
        if (-not $exe) {
            [void][System.Windows.Forms.MessageBox]::Show($script:Ui.Form, "На этом ПК не найден Revit Batch Processor (BatchRvt.exe).", "Запуск")
            return
        }
    }
    if ($insertView) {
        $apply = Show-WorksetInsertDialog $insertView $title $false
        if (-not $apply) {
            Set-UiStatus "Вставка связей отменена"
            return
        }
    }
    else {
        $apply = Show-OpsPlanDialog -Title $title -Text $text -CanApply $can -YesText "Да, записать"
        if (-not $apply) {
            Set-UiStatus "Запись наборов отменена"
            return
        }
    }
    if (-not $can) { return }
    $planPath = Join-Path $ToolDir "workset_plan.csv"
    $jobPath = Join-Path $ToolDir "workset_job.json"
    Write-WorksetPlanFile $planPath @($plan.Rows)
    Write-WorksetJobFile $jobPath $year.Year "workset_plan.csv" (Initialize-OpsReportsDir $ToolDir)
    $hosts = New-Object System.Collections.Generic.List[string]
    $seenHost = @{}
    foreach ($a in @($plan.Rows)) {
        # a model whose rows are all skips is not opened
        if ($a.Action -ne "rename" -and $a.Action -ne "create" -and $a.Action -ne "insert" -and $a.Action -ne "delete") { continue }
        $h = [string]$a.Host
        $k = ConvertTo-WorksetPathKey $h
        if (-not $k -or $seenHost.ContainsKey($k)) { continue }
        $seenHost[$k] = $true
        [void]$hosts.Add($h)
    }
    Write-Utf8BomLines (Join-Path $ToolDir "rvt_list_worksets.txt") $hosts.ToArray()
    $pass = [pscustomobject]@{
        Kind     = "rbp"
        Name     = "workset-apply"
        Task     = [System.IO.Path]::Combine($CodeDir, "apply_workset_links.py")
        List     = [System.IO.Path]::Combine($ToolDir, "rvt_list_worksets.txt")
        Year     = $year.Year
        Detach   = $false
        Worksets = "open_all"
    }
    Reset-OpsReportSession
    $script:Run.Queue = @($pass)
    $script:Run.Index = 0
    $script:Run.ExitCodes = @()
    $script:Run.Exe = $exe
    $script:Run.ExeDir = Split-Path -Parent $exe
    $script:Run.ReadImport = $false
    $script:Run.ApplyImport = $true
    $script:WsApplyById = @{}
    $script:WsApplyByPath = @{}
    $script:Run.List = Get-WorksetModelList
    Set-UiBusy $true
    $script:Ui.Timer.Start()
    Start-NextPass
}

function Start-WorksetRead {
    param([string]$Worksets = "close_all")
    if ($script:Run.Active) { return }
    $paths = @(Get-CheckedWorksetPaths)
    if ($paths.Count -eq 0) {
        [void][System.Windows.Forms.MessageBox]::Show($script:Ui.Form, "Отметьте модели галочками.", "Чтение")
        return
    }
    $gate = Test-WorksetReadPaths $paths
    if (-not $gate.Ok) {
        [void][System.Windows.Forms.MessageBox]::Show($script:Ui.Form, "Нужны пути RSN или сетевые папки (\\сервер\...). BIM 360 и диск не поддерживаются.", "Чтение")
        return
    }
    $year = Resolve-WorksetReadYear $paths $script:servers ""
    if (-not $year.Ok) {
        [void][System.Windows.Forms.MessageBox]::Show($script:Ui.Form, (Get-WorksetYearText $year), "Чтение")
        return
    }
    if ($Worksets -ne "open_all") { $Worksets = "close_all" }
    $script:WsReadMode = $(if ($Worksets -eq "open_all") { "full" } else { "fast" })
    $listPath = Join-Path $ToolDir "rvt_list_worksets.txt"
    Write-Utf8BomLines $listPath $paths
    $pass = New-WorksetReadPass $ToolDir $CodeDir $year.Year $Worksets
    $exe = Find-BatchRvt
    if (-not $exe) {
        [void][System.Windows.Forms.MessageBox]::Show($script:Ui.Form, "На этом ПК не найден Revit Batch Processor (BatchRvt.exe).", "Чтение")
        return
    }
    Reset-OpsReportSession
    $script:Run.Queue = @($pass)
    $script:Run.Index = 0
    $script:Run.ExitCodes = @()
    $script:Run.Exe = $exe
    $script:Run.ExeDir = Split-Path -Parent $exe
    $script:Run.ReadImport = $true
    $script:Run.List = Get-WorksetModelList
    Set-UiBusy $true
    $script:Ui.Timer.Start()
    Start-NextPass
}

# Drop-down over one ListView cell («Роль» of the model list, «Роль» of «Маска имени»).
# The choice is applied on SelectedIndexChanged while the drop-down is armed; closing it only
# hides it, so the order of the close / select notifications does not matter.
function Show-WorksetCellCombo($List, $Item, [int]$Col, [string[]]$Choices, [string]$Current, [string]$Kind, $Paths) {
    if ($script:Run.Active) { return }
    if (-not $script:Ui.ContainsKey("WsCombo") -or -not $script:Ui.WsCombo) {
        $cb = New-Object System.Windows.Forms.ComboBox
        $cb.DropDownStyle = "DropDownList"
        $cb.Visible = $false
        $cb.Add_SelectedIndexChanged({
                if (-not $script:WsComboArmed -or -not $script:WsComboState) { return }
                $word = [string]$script:Ui.WsCombo.SelectedItem
                if ($word) { Set-WorksetComboChoice $script:WsComboState $word }
            })
        $cb.Add_DropDownClosed({ $script:Ui.WsCombo.Visible = $false })
        $cb.Add_Leave({ $script:Ui.WsCombo.Visible = $false })
        $cb.Add_KeyDown({
                if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Escape -or $_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
                    $_.Handled = $true
                    $script:Ui.WsCombo.Visible = $false
                }
            })
        $script:Ui.WsCombo = $cb
    }
    $cb = $script:Ui.WsCombo
    $script:WsComboArmed = $false
    if ($cb.Parent -ne $List) {
        if ($cb.Parent) { $cb.Parent.Controls.Remove($cb) }
        $List.Controls.Add($cb)
    }
    $cb.Items.Clear()
    foreach ($c in $Choices) { [void]$cb.Items.Add($c) }
    $idx = [Array]::IndexOf([string[]]$Choices, $Current)
    $cb.SelectedIndex = $idx
    $b = $Item.SubItems[$Col].Bounds
    $cb.Font = $List.Font
    $cb.SetBounds($b.X, $b.Y, [Math]::Max($b.Width, 110), $b.Height)
    $script:WsComboState = @{ Kind = $Kind; Item = $Item; Col = $Col; Paths = @($Paths) }
    $script:WsComboArmed = $true
    $cb.Visible = $true
    $cb.BringToFront()
    $cb.Focus()
    $cb.DroppedDown = $true
}

function Set-WorksetComboChoice($State, [string]$Word) {
    if ($State.Kind -eq "model-role") {
        $code = $script:WsRoleTarget
        if ($Word -eq "новая") { $code = $script:WsRoleNew }
        Set-WorksetRole @($State.Paths) $code
        Update-WorksetRoleColumn
        Request-WorksetRecompute
        return
    }
    if ($State.Kind -eq "mask-role") {
        $it = $State.Item
        if (-not $it) { return }
        $role = ConvertFrom-WorksetRoleWord $Word
        $it.Tag = $role
        if ($it.SubItems.Count -gt 2) { $it.SubItems[2].Text = Get-WorksetRoleWord $role }
        Read-WorksetRolesFromView
        Request-WorksetRecompute
    }
}

# «Добавить связь вручную…»: which models get which links, outside the same-building rule.
# Left: where to insert (pre-ticked: the rows selected in the model list). Right: what to insert.
function Show-WorksetManualPairDialog {
    if ($script:Run.Active) { return }
    $paths = @(Get-WorksetTabPaths "addlinks" $false)
    if ($paths.Count -eq 0) { Set-UiStatus "Список моделей пуст."; return }
    $selected = @{}
    foreach ($it in @($script:Ui.PlainLists["addlinks"].SelectedItems)) { $selected[(ConvertTo-WorksetPathKey ([string]$it.Tag))] = $true }

    $f = New-Object System.Windows.Forms.Form
    $f.Text = "Добавить связь вручную"
    $f.StartPosition = "CenterParent"
    $f.Font = $script:Ui.Form.Font
    $f.MinimizeBox = $false
    $f.ShowInTaskbar = $false
    $wa = [System.Windows.Forms.Screen]::FromControl($script:Ui.Form).WorkingArea
    $f.Size = New-Object System.Drawing.Size([Math]::Min(1000, $wa.Width - 40), [Math]::Min(560, $wa.Height - 40))
    $f.MinimumSize = New-Object System.Drawing.Size(640, 360)

    $lists = New-Object System.Collections.Generic.List[object]
    foreach ($head in @("Куда вставить", "Что вставить")) {
        $lv = New-OpsListView
        $lv.CheckBoxes = $true
        [void]$lv.Columns.Add($head, 300)
        [void]$lv.Columns.Add("Роль", 90)
        Register-OpsFillColumn $lv 0 @()
        foreach ($p in $paths) {
            $it = New-Object System.Windows.Forms.ListViewItem (Get-WorksetBaseName $p)
            [void]$it.SubItems.Add((Get-WorksetRoleLabel $p))
            $it.Tag = [string]$p
            [void]$lv.Items.Add($it)
        }
        [void]$lists.Add($lv)
    }
    foreach ($it in $lists[0].Items) { if ($selected.ContainsKey((ConvertTo-WorksetPathKey ([string]$it.Tag)))) { $it.Checked = $true } }

    $two = New-Object System.Windows.Forms.TableLayoutPanel
    $two.Dock = "Fill"
    $two.Margin = New-Object System.Windows.Forms.Padding(0)
    $two.ColumnCount = 2
    $two.RowCount = 1
    [void]$two.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle ([System.Windows.Forms.SizeType]::Percent, 50)))
    [void]$two.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle ([System.Windows.Forms.SizeType]::Percent, 50)))
    [void]$two.RowStyles.Add((New-Object System.Windows.Forms.RowStyle ([System.Windows.Forms.SizeType]::Percent, 100)))
    $lists[0].Margin = New-Object System.Windows.Forms.Padding(0, 0, 6, 0)
    $lists[1].Margin = New-Object System.Windows.Forms.Padding(6, 0, 0, 0)
    $two.Controls.Add($lists[0], 0, 0)
    $two.Controls.Add($lists[1], 1, 0)

    $btnOk = New-OpsButton "Добавить"
    $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $btnCancel = New-OpsButton "Отмена"
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $grid = New-OpsColumnGrid @(0, $script:OpsRowH)
    $grid.Controls.Add($two, 0, 0)
    $grid.Controls.Add((New-OpsBar @() @($btnOk, $btnCancel)), 0, 1)
    $f.Controls.Add($grid)
    $f.AcceptButton = $btnOk
    $f.CancelButton = $btnCancel
    $res = $f.ShowDialog($script:Ui.Form)
    $targets = New-Object System.Collections.Generic.List[string]
    $links = New-Object System.Collections.Generic.List[string]
    foreach ($it in $lists[0].Items) { if ($it.Checked) { [void]$targets.Add([string]$it.Tag) } }
    foreach ($it in $lists[1].Items) { if ($it.Checked) { [void]$links.Add([string]$it.Tag) } }
    $f.Dispose()
    if ($res -ne [System.Windows.Forms.DialogResult]::OK) { return }
    if ($targets.Count -eq 0 -or $links.Count -eq 0) { Set-UiStatus "Отметьте, куда вставить и что вставить."; return }
    $n = 0
    foreach ($t in $targets) { foreach ($l in $links) { if (Add-WorksetManualPair $t $l) { $n++ } } }
    Update-WorksetViews
    Set-UiStatus (Join-WorksetStatus ("Добавлено строк вручную: {0}." -f $n) $script:WsPairNote)
}

# Confirm form of «Вставить связи» (and of its «Пробный прогон»), on the common grid:
#   row 0  one plain sentence (Get-WorksetInsertSummary);
#   row 1  ListView «Модель | Что вставить | Рабочий набор | Действие», short names, full paths in the tool tip;
#   row 2  ListView «Пропущено | Модель | Причина», only when something is skipped;
#   row 3  bar with «Вставить» (AcceptButton) and «Отмена» (CancelButton), or only «Закрыть».
# Sized to the screen working area (works on a 1024x550 remote desktop), resizable, MinimumSize.
# Returns @{ Form; Yes; No; List; Skipped }; Show-WorksetInsertDialog shows it.
function New-WorksetInsertForm($View, [string]$Title, [bool]$DryRun) {
    $f = New-Object System.Windows.Forms.Form
    $f.Text = $Title
    $f.StartPosition = "CenterParent"
    $f.ShowInTaskbar = $false
    $f.MinimizeBox = $false
    $f.FormBorderStyle = "Sizable"
    if ($script:Ui -and $script:Ui.Form) { $f.Font = $script:Ui.Form.Font } else { $f.Font = $script:OpsFont }
    $f.MinimumSize = New-Object System.Drawing.Size(600, 320)
    $wa = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea
    if ($script:Ui -and $script:Ui.Form) { $wa = [System.Windows.Forms.Screen]::FromControl($script:Ui.Form).WorkingArea }
    $w = [Math]::Max(600, [Math]::Min(960, $wa.Width - 40))
    $h = [Math]::Max(320, [Math]::Min(600, $wa.Height - 40))
    $f.Size = New-Object System.Drawing.Size($w, $h)

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = [string]$View.Summary
    $lbl.AutoSize = $false
    $lbl.AutoEllipsis = $false
    $lbl.Dock = "Fill"
    $lbl.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $lbl.Margin = New-Object System.Windows.Forms.Padding(0)

    $lv = New-OpsListView
    $lv.ShowItemToolTips = $true
    $lv.MultiSelect = $false
    [void]$lv.Columns.Add("Модель", 200)
    [void]$lv.Columns.Add("Что вставить", 220)
    [void]$lv.Columns.Add("Рабочий набор", 220)
    [void]$lv.Columns.Add("Действие", 110)
    foreach ($r in @($View.Rows)) {
        $it = New-Object System.Windows.Forms.ListViewItem ([string]$r.Model)
        [void]$it.SubItems.Add([string]$r.Link)
        [void]$it.SubItems.Add([string]$r.Workset)
        # short: «вставить» / «вставить + новый набор»; the paths are the tool tip
        $act = [string]$r.Action
        if (-not $act) { $act = Get-WorksetInsertActionText ([bool]$r.IsNew) }
        [void]$it.SubItems.Add($act)
        $tip = "{0}{1}← {2}" -f $r.ModelPath, [Environment]::NewLine, $r.LinkPath
        if ($r.IsNew) { $tip = $tip + [Environment]::NewLine + ("Рабочий набор «{0}» будет создан." -f $r.Workset) }
        $it.ToolTipText = $tip
        [void]$lv.Items.Add($it)
    }
    # every column as wide as its longest text; the last one takes the rest of the width
    $needMain = Set-OpsColumnsToContent $lv $f.Font
    Register-OpsFillColumn $lv 3 @()

    $skipCount = @($View.Skipped).Count
    $sk = $null
    $needSkip = 0
    if ($skipCount -gt 0) {
        $sk = New-OpsListView
        $sk.MultiSelect = $false
        $sk.ShowItemToolTips = $true
        [void]$sk.Columns.Add("Пропущено", 220)
        [void]$sk.Columns.Add("Модель", 180)
        [void]$sk.Columns.Add("Причина", 300)
        foreach ($s in @($View.Skipped)) {
            $it = New-Object System.Windows.Forms.ListViewItem ([string]$s.Link)
            [void]$it.SubItems.Add([string]$s.Model)
            [void]$it.SubItems.Add([string]$s.Reason)
            $tip = [string]$s.Reason
            if ($s.PSObject.Properties["Tip"] -and $s.Tip) { $tip = [string]$s.Tip }
            $it.ToolTipText = $tip
            [void]$sk.Items.Add($it)
        }
        $needSkip = Set-OpsColumnsToContent $sk $f.Font
        Register-OpsFillColumn $sk 2 @()
    }
    # the minimum width shows every column whole (as far as the screen allows)
    $need = [Math]::Max($needMain, $needSkip)
    $need = [Math]::Max($need, [System.Windows.Forms.TextRenderer]::MeasureText([string]$View.Summary, $f.Font).Width)
    $need = $need + [System.Windows.Forms.SystemInformation]::VerticalScrollBarWidth + 2 * $script:OpsGap + 24
    $minW = [Math]::Max(600, [Math]::Min($need, $wa.Width - 40))
    $f.MinimumSize = New-Object System.Drawing.Size($minW, 320)
    if ($f.Width -lt $minW) { $f.Width = $minW }

    $btnYes = New-OpsButton "Вставить"
    $btnYes.DialogResult = [System.Windows.Forms.DialogResult]::Yes
    $btnYes.Enabled = ([int]$View.InsertCount -gt 0)
    $btnNo = New-OpsButton "Отмена"
    $btnNo.DialogResult = [System.Windows.Forms.DialogResult]::No
    $right = @($btnYes, $btnNo)
    if ($DryRun) {
        $btnNo.Text = "Закрыть"
        $btnYes.Visible = $false
        $right = @($btnNo)
    }
    $heights = @(($script:OpsFont.Height * 2 + 8), 0, $script:OpsRowH)
    if ($sk) { $heights = @(($script:OpsFont.Height * 2 + 8), 0, 0, $script:OpsRowH) }
    $grid = New-OpsColumnGrid $heights
    if ($sk) {
        # the inserts get the bigger part, the skipped list stays short
        $grid.RowStyles[1].SizeType = [System.Windows.Forms.SizeType]::Percent
        $grid.RowStyles[1].Height = 62
        $grid.RowStyles[2].SizeType = [System.Windows.Forms.SizeType]::Percent
        $grid.RowStyles[2].Height = 38
        $sk.Margin = New-Object System.Windows.Forms.Padding(0, 8, 0, 0)
    }
    $grid.Controls.Add($lbl, 0, 0)
    $grid.Controls.Add($lv, 0, 1)
    $row = 2
    if ($sk) { $grid.Controls.Add($sk, 0, 2); $row = 3 }
    $bar = New-OpsBar @() $right
    $bar.Margin = New-Object System.Windows.Forms.Padding(0, 8, 0, 0)
    $grid.Controls.Add($bar, 0, $row)
    $f.Controls.Add($grid)
    if ($DryRun) { $f.AcceptButton = $btnNo }
    elseif ($btnYes.Enabled) { $f.AcceptButton = $btnYes }
    $f.CancelButton = $btnNo
    # nothing pre-selected: the focus sits on the button, not on a list or a text
    if ($DryRun -or -not $btnYes.Enabled) { $f.ActiveControl = $btnNo } else { $f.ActiveControl = $btnYes }
    return @{ Form = $f; Yes = $btnYes; No = $btnNo; List = $lv; Skipped = $sk; Bar = $bar }
}

function Show-WorksetInsertDialog($View, [string]$Title, [bool]$DryRun) {
    $parts = New-WorksetInsertForm $View $Title $DryRun
    $f = $parts.Form
    try { $res = $f.ShowDialog($script:Ui.Form) }
    finally { $f.Dispose() }
    if ($DryRun) { return $false }
    return ($res -eq [System.Windows.Forms.DialogResult]::Yes)
}

# Builds the lower panes of «Связи: добавить» ($LinksHost) and «Рабочие наборы» ($SetsHost)
# on the common grid (ui_grid.ps1). The model lists above them belong to presets.ps1.
function Install-WorksetUi($LinksHost, $SetsHost) {
    $script:WsActiveKey = "addlinks"

    # ---- «Связи: добавить»: read row, then sub-tabs Связи / Коды разделов / Маска имени.
    # Pairs are proposed by themselves (after a read, a report, a role or a tick): no button.
    $btnRead = New-OpsButton "Прочитать наборы и связи"
    $btnReadFull = New-OpsButton "Полное чтение"
    $btnOpenRead = New-OpsButton "Открыть отчёт чтения…"
    $script:Ui.BtnWsRead = $btnRead
    $script:Ui.BtnWsReadFull = $btnReadFull
    $script:Ui.BtnWsOpenRead = $btnOpenRead
    $readBar = New-OpsBar @($btnRead, $btnReadFull, $btnOpenRead) @()
    $btnManual = New-OpsButton "Добавить связь вручную…"
    $btnDropRow = New-OpsButton "Удалить строку"
    $script:Ui.BtnWsManual = $btnManual
    $script:Ui.BtnWsDropRow = $btnDropRow

    $linkList = New-OpsListView
    [void]$linkList.Columns.Add("Целевая модель", 220)
    [void]$linkList.Columns.Add("Связь", 260)
    [void]$linkList.Columns.Add("Раздел", 80)
    [void]$linkList.Columns.Add("Корпус", 80)
    [void]$linkList.Columns.Add("Рабочий набор", 200)
    [void]$linkList.Columns.Add("Состояние", 140)
    Register-OpsFillColumn $linkList 5 @(0, 1, 4)
    $script:Ui.LinkList = $linkList

    $codeList = New-OpsListView
    [void]$codeList.Columns.Add("Код (латиница)", 200)
    [void]$codeList.Columns.Add("Код (кириллица)", 200)
    Register-OpsFillColumn $codeList 1 @(0)
    $script:Ui.CodeList = $codeList
    $btnCodeAdd = New-OpsButton "Добавить"
    $btnCodeDel = New-OpsButton "Удалить"
    $codeGrid = New-OpsColumnGrid @($script:OpsRowH, 0) 8
    $codeGrid.Controls.Add((New-OpsBar @($btnCodeAdd, $btnCodeDel) @()), 0, 0)
    $codeGrid.Controls.Add($codeList, 0, 1)

    $txtSep = New-Object System.Windows.Forms.TextBox
    $txtSep.Width = 48
    $txtSep.Text = "_"
    $txtSep.Anchor = "Left"
    $script:Ui.WsSep = $txtSep
    $btnSaveMask = New-OpsButton "Запомнить маску"
    $lblSep = New-OpsLabel "Разделитель"
    $txtSep.Margin = New-Object System.Windows.Forms.Padding(0, 0, 16, 0)
    $maskList = New-OpsListView
    [void]$maskList.Columns.Add("№", 50)
    [void]$maskList.Columns.Add("Часть имени", 260)
    [void]$maskList.Columns.Add("Роль", 160)
    Register-OpsFillColumn $maskList 2 @(1)
    $script:Ui.MaskList = $maskList
    $lblPreview = New-Object System.Windows.Forms.Label
    $lblPreview.AutoSize = $false
    $lblPreview.AutoEllipsis = $true
    $lblPreview.Dock = "Fill"
    $lblPreview.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $script:Ui.WsMaskPreview = $lblPreview
    $maskGrid = New-OpsColumnGrid @($script:OpsRowH, 0, 30) 8
    $maskGrid.Controls.Add((New-OpsBar @($lblSep, $txtSep, $btnSaveMask) @()), 0, 0)
    $maskGrid.Controls.Add($maskList, 0, 1)
    $maskGrid.Controls.Add($lblPreview, 0, 2)

    $linkTabs = New-Object System.Windows.Forms.TabControl
    $linkTabs.Dock = "Fill"
    $linkTabs.Margin = New-Object System.Windows.Forms.Padding(0)
    $tabLinks = New-Object System.Windows.Forms.TabPage
    $tabLinks.Text = "Связи"
    $tabLinks.UseVisualStyleBackColor = $true
    $linkGrid = New-OpsColumnGrid @($script:OpsRowH, 0) 8
    $linkGrid.Controls.Add((New-OpsBar @($btnManual, $btnDropRow) @()), 0, 0)
    $linkGrid.Controls.Add($linkList, 0, 1)
    $tabLinks.Controls.Add($linkGrid)
    $tabCodes = New-Object System.Windows.Forms.TabPage
    $tabCodes.Text = "Коды разделов"
    $tabCodes.UseVisualStyleBackColor = $true
    $tabCodes.Controls.Add($codeGrid)
    $tabMask = New-Object System.Windows.Forms.TabPage
    $tabMask.Text = "Маска имени"
    $tabMask.UseVisualStyleBackColor = $true
    $tabMask.Controls.Add($maskGrid)
    $linkTabs.TabPages.AddRange(@($tabLinks, $tabCodes, $tabMask))
    $script:Ui.Tabs = $linkTabs
    $script:Ui.TabLinks = $tabLinks
    $script:Ui.TabCodes = $tabCodes
    $script:Ui.TabMask = $tabMask

    $linksGrid = New-OpsColumnGrid @($script:OpsRowH, 0) 0
    $linksGrid.Controls.Add($readBar, 0, 0)
    $linksGrid.Controls.Add($linkTabs, 0, 1)
    $LinksHost.Controls.Add($linksGrid)

    # ---- «Рабочие наборы»: read row, filter row, sub-tabs Наборы / Связи в моделях, delete row
    $btnRead2 = New-OpsButton "Прочитать наборы и связи"
    $btnReadFull2 = New-OpsButton "Полное чтение"
    $btnOpenRead2 = New-OpsButton "Открыть отчёт чтения…"
    $readBar2 = New-OpsBar @($btnRead2, $btnReadFull2, $btnOpenRead2) @()

    $lblFilter = New-OpsLabel "Модель:"
    $cmbFilter = New-Object System.Windows.Forms.ComboBox
    $cmbFilter.DropDownStyle = "DropDownList"
    $cmbFilter.Width = 280
    $cmbFilter.Anchor = "Left"
    $cmbFilter.Margin = New-Object System.Windows.Forms.Padding(0, 0, 24, 0)
    [void]$cmbFilter.Items.Add("все отмеченные")
    $cmbFilter.SelectedIndex = 0
    $script:Ui.WsFilter = $cmbFilter
    $lblRepl = New-OpsLabel "Заменить в имени:"
    $txtFind = New-Object System.Windows.Forms.TextBox
    $txtFind.Width = 150
    $txtFind.Anchor = "Left"
    $txtFind.Margin = New-Object System.Windows.Forms.Padding(0)
    $arrow = New-OpsLabel "→"
    $arrow.Margin = New-Object System.Windows.Forms.Padding(6, 0, 6, 0)
    $txtRepl = New-Object System.Windows.Forms.TextBox
    $txtRepl.Width = 150
    $txtRepl.Anchor = "Left"
    $txtRepl.Margin = New-Object System.Windows.Forms.Padding(0, 0, 8, 0)
    Set-OpsCue $txtFind "было"
    Set-OpsCue $txtRepl "стало"
    $script:Ui.WsFind = $txtFind
    $script:Ui.WsRepl = $txtRepl
    $btnReplace = New-OpsButton "Заменить"
    $filterBar = New-OpsBar @($lblFilter, $cmbFilter, $lblRepl, $txtFind, $arrow, $txtRepl, $btnReplace) @()

    $renameList = New-OpsListView
    [void]$renameList.Columns.Add("Модель", 200)
    [void]$renameList.Columns.Add("Рабочий набор сейчас", 240)
    [void]$renameList.Columns.Add("Станет", 240)
    [void]$renameList.Columns.Add("Элементов", 90)
    [void]$renameList.Columns.Add("Действие", 260)
    Register-OpsFillColumn $renameList 4 @(0, 1, 2)
    $script:Ui.RenameList = $renameList

    $exist = New-OpsListView
    [void]$exist.Columns.Add("Модель", 240)
    [void]$exist.Columns.Add("Связь", 260)
    [void]$exist.Columns.Add("Рабочий набор экземпляра", 200)
    [void]$exist.Columns.Add("Рабочий набор типа", 200)
    Register-OpsFillColumn $exist 3 @(0, 1, 2)
    $script:Ui.ExistingLinks = $exist

    $setTabs = New-Object System.Windows.Forms.TabControl
    $setTabs.Dock = "Fill"
    $setTabs.Margin = New-Object System.Windows.Forms.Padding(0)
    $tabSets = New-Object System.Windows.Forms.TabPage
    $tabSets.Text = "Наборы"
    $tabSets.Padding = New-Object System.Windows.Forms.Padding(8)
    $tabSets.UseVisualStyleBackColor = $true
    $tabSets.Controls.Add($renameList)
    $tabExist = New-Object System.Windows.Forms.TabPage
    $tabExist.Text = "Связи в моделях (только просмотр)"
    $tabExist.Padding = New-Object System.Windows.Forms.Padding(8)
    $tabExist.UseVisualStyleBackColor = $true
    $tabExist.Controls.Add($exist)
    $setTabs.TabPages.AddRange(@($tabSets, $tabExist))
    $script:Ui.SetTabs = $setTabs
    $script:Ui.TabSets = $tabSets

    $btnMark = New-OpsButton "Удалить выбранные"
    $rbMove = New-OpsRadio ("перенести элементы в {0}" -f $script:WorksetDefaultIgnoreName) $true
    $rbWith = New-OpsRadio "вместе с элементами"
    $rbWith.ForeColor = [System.Drawing.Color]::Firebrick
    $btnMark.Margin = New-Object System.Windows.Forms.Padding(0, 0, 16, 0)
    $rbMove.Margin = New-Object System.Windows.Forms.Padding(0, 0, 16, 0)
    $rbMove.Anchor = "Left"
    $rbWith.Anchor = "Left"
    $script:Ui.WsDeleteElements = $rbWith
    $script:Ui.WsDeleteMove = $rbMove
    $deleteBar = New-OpsBar @($btnMark, $rbMove, $rbWith) @()

    $setsGrid = New-OpsColumnGrid @($script:OpsRowH, $script:OpsRowH, 0, $script:OpsRowH) 0
    $setsGrid.Controls.Add($readBar2, 0, 0)
    $setsGrid.Controls.Add($filterBar, 0, 1)
    $setsGrid.Controls.Add($setTabs, 0, 2)
    $setsGrid.Controls.Add($deleteBar, 0, 3)
    $SetsHost.Controls.Add($setsGrid)

    Add-OpsBusy @($btnRead, $btnReadFull, $btnManual, $btnDropRow, $btnCodeAdd, $btnCodeDel, $btnSaveMask,
        $btnRead2, $btnReadFull2, $btnReplace, $btnMark, $rbMove, $rbWith, $cmbFilter)
    # «Открыть отчёт чтения…» is never switched off: it only reads a file

    Load-WorksetUserFiles
    Update-WorksetCodeView
    Update-WorksetMaskView

    # one recompute after the clicks stop (ticks, roles, mask edits)
    $recalc = New-Object System.Windows.Forms.Timer
    $recalc.Interval = 300
    $recalc.Add_Tick({ Invoke-WorksetRecompute })
    $script:Ui.WsRecalcTimer = $recalc

    foreach ($b in @($btnRead, $btnRead2)) { $b.Add_Click({ Start-WorksetRead -Worksets "close_all" }) }
    foreach ($b in @($btnReadFull, $btnReadFull2)) { $b.Add_Click({ Start-WorksetRead -Worksets "open_all" }) }
    foreach ($b in @($btnOpenRead, $btnOpenRead2)) { $b.Add_Click({ Show-WorksetReadReportDialog }) }
    $btnManual.Add_Click({ Show-WorksetManualPairDialog })
    $btnDropRow.Add_Click({
            $keys = New-Object System.Collections.Generic.List[string]
            foreach ($it in @($script:Ui.LinkList.SelectedItems)) { [void]$keys.Add([string]$it.Tag) }
            if ($keys.Count -eq 0) { Set-UiStatus "Выделите строки, добавленные вручную."; return }
            $res = Remove-WorksetManualRows $keys.ToArray()
            Update-WorksetViews
            $text = "Удалено строк: {0}." -f $res.Removed
            if ($res.Auto -gt 0) { $text = $text + (" Автоматических строк не удалено: {0} — снимите галочку у модели или смените её роль." -f $res.Auto) }
            Set-UiStatus $text
        })
    $cmbFilter.Add_SelectedIndexChanged({ if (-not $script:WsFilterLock) { Update-WorksetRenameView } })
    $btnReplace.Add_Click({
            $null = Update-WorksetRenamePreview $script:WsRenameRows $script:Ui.WsFind.Text $script:Ui.WsRepl.Text
            Update-WorksetRenameView
            Set-UiStatus "«Станет» заполнено. В модели ничего не записано."
        })
    $btnMark.Add_Click({
            $n = 0
            foreach ($it in @($script:Ui.RenameList.SelectedItems)) {
                $nm = [string]$it.Name
                if (-not $nm) { continue }
                $idx = -1
                for ($i = 0; $i -lt $script:WsDeleteNames.Count; $i++) { if ($script:WsDeleteNames[$i] -eq $nm) { $idx = $i; break } }
                if ($idx -ge 0) { $script:WsDeleteNames.RemoveAt($idx) } else { [void]$script:WsDeleteNames.Add($nm) }
                $n++
            }
            if ($n -eq 0) { Set-UiStatus "Выделите наборы в таблице."; return }
            Update-WorksetRenameView
            Set-UiStatus ("К удалению наборов: {0}. Повторное нажатие снимает пометку." -f $script:WsDeleteNames.Count)
        })
    $rbWith.Add_CheckedChanged({ Update-WorksetRenameView })
    $btnCodeAdd.Add_Click({
            $it = New-Object System.Windows.Forms.ListViewItem ""
            [void]$it.SubItems.Add("")
            [void]$script:Ui.CodeList.Items.Add($it)
        })
    $btnCodeDel.Add_Click({
            foreach ($it in @($script:Ui.CodeList.SelectedItems)) { $script:Ui.CodeList.Items.Remove($it) }
            Read-WorksetCodesFromView
        })
    $btnSaveMask.Add_Click({
            Read-WorksetRolesFromView
            if (-not (Test-WorksetRoleSet @($script:WsMaskRoles))) {
                [void][System.Windows.Forms.MessageBox]::Show($script:Ui.Form, "Нужны ровно один «раздел» и один «корпус».", "Маска")
                return
            }
            Save-WorksetMaskFile
            Update-WorksetViews
            Set-UiStatus (Join-WorksetStatus "Маска записана в name_mask.cfg." $script:WsPairNote)
        })
    $txtSep.Add_TextChanged({
            if ($script:Ui.WsSep.Text -and $script:Ui.WsSep.Text -ne $script:WsSeparator) {
                $script:WsSeparator = $script:Ui.WsSep.Text
                Request-WorksetRecompute
            }
        })
    $linkList.Add_MouseDoubleClick({
            $h = $script:Ui.LinkList.HitTest($_.Location)
            if (-not $h.Item -or -not $h.SubItem) { return }
            $col = $h.Item.SubItems.IndexOf($h.SubItem)
            if ($col -eq 1) { Show-WorksetGridEditor $script:Ui.LinkList $h.Item 1 "link-path" }
            elseif ($col -eq 4) { Show-WorksetGridEditor $script:Ui.LinkList $h.Item 4 "link-workset" }
        })
    # «Рабочий набор»: a single click on the cell (or F2 on the row) opens the edit box too,
    # so an empty yellow cell is filled in by hand without hunting for a double-click
    $linkList.Add_MouseClick({
            if ($_.Button -ne [System.Windows.Forms.MouseButtons]::Left) { return }
            $mods = [System.Windows.Forms.Control]::ModifierKeys
            if (($mods -band [System.Windows.Forms.Keys]::Control) -or ($mods -band [System.Windows.Forms.Keys]::Shift)) { return }
            $h = $script:Ui.LinkList.HitTest($_.Location)
            if (-not $h.Item -or -not $h.SubItem) { return }
            if ($h.Item.SubItems.IndexOf($h.SubItem) -ne 4) { return }
            if (-not (Test-WorksetRowEditable ([string]$h.Item.Tag))) { return }
            Show-WorksetGridEditor $script:Ui.LinkList $h.Item 4 "link-workset"
        })
    $linkList.Add_KeyDown({
            if ($_.KeyCode -ne [System.Windows.Forms.Keys]::F2) { return }
            if ($script:Ui.LinkList.SelectedItems.Count -lt 1) { return }
            $it = $script:Ui.LinkList.SelectedItems[0]
            if (-not (Test-WorksetRowEditable ([string]$it.Tag))) { return }
            $_.Handled = $true
            Show-WorksetGridEditor $script:Ui.LinkList $it 4 "link-workset"
        })
    $renameList.Add_MouseDoubleClick({
            $h = $script:Ui.RenameList.HitTest($_.Location)
            if (-not $h.Item -or -not $h.SubItem) { return }
            $col = $h.Item.SubItems.IndexOf($h.SubItem)
            if ($col -eq 2) { Show-WorksetGridEditor $script:Ui.RenameList $h.Item 2 "rename" }
        })
    $codeList.Add_MouseDoubleClick({
            $h = $script:Ui.CodeList.HitTest($_.Location)
            if (-not $h.Item -or -not $h.SubItem) { return }
            $col = $h.Item.SubItems.IndexOf($h.SubItem)
            if ($col -eq 0 -or $col -eq 1) { Show-WorksetGridEditor $script:Ui.CodeList $h.Item $col "code" }
        })
    # «Роль» of a name part: a drop-down in the cell (раздел / корпус / как есть / не важен)
    $maskList.Add_MouseClick({
            if ($_.Button -ne [System.Windows.Forms.MouseButtons]::Left) { return }
            $h = $script:Ui.MaskList.HitTest($_.Location)
            if (-not $h.Item -or -not $h.SubItem) { return }
            if ($h.Item.SubItems.IndexOf($h.SubItem) -ne 2) { return }
            Show-WorksetCellCombo $script:Ui.MaskList $h.Item 2 (Get-WorksetMaskRoleChoices) ([string]$h.Item.SubItems[2].Text) "mask-role" @()
        })
    $linkTabs.Add_SelectedIndexChanged({
            if ($script:Ui.Tabs.SelectedTab -eq $script:Ui.TabCodes) { return }
            if ($script:Ui.CodeList) { Read-WorksetCodesFromView }
        })

    # roles of the models (новая / целевая): context menu of the «Связи: добавить» model list
    $addList = $null
    if ($script:Ui.PlainLists -and $script:Ui.PlainLists.ContainsKey("addlinks")) { $addList = $script:Ui.PlainLists["addlinks"] }
    if ($addList) {
        $menu = New-Object System.Windows.Forms.ContextMenuStrip
        $miNew = $menu.Items.Add("Пометить как новые")
        $miTarget = $menu.Items.Add("Пометить как целевые")
        $script:Ui.MiRoleNew = $miNew
        $script:Ui.MiRoleTarget = $miTarget
        $miNew.Add_Click({
                $paths = New-Object System.Collections.Generic.List[string]
                foreach ($it in $script:Ui.PlainLists["addlinks"].SelectedItems) { [void]$paths.Add([string]$it.Tag) }
                Set-WorksetRole $paths.ToArray() $script:WsRoleNew
                Update-WorksetRoleColumn
                Invoke-WorksetRecompute
            })
        $miTarget.Add_Click({
                $paths = New-Object System.Collections.Generic.List[string]
                foreach ($it in $script:Ui.PlainLists["addlinks"].SelectedItems) { [void]$paths.Add([string]$it.Tag) }
                Set-WorksetRole $paths.ToArray() $script:WsRoleTarget
                Update-WorksetRoleColumn
                Invoke-WorksetRecompute
            })
        $menu.Add_Opening({
                $has = ($script:Ui.PlainLists["addlinks"].SelectedItems.Count -gt 0) -and -not $script:Run.Active
                $script:Ui.MiRoleNew.Enabled = $has
                $script:Ui.MiRoleTarget.Enabled = $has
            }) | Out-Null
        $addList.ContextMenuStrip = $menu
        # «Роль» cell: a drop-down (новая / целевая) for every selected row. MouseDown runs before
        # the ListView changes the selection, so the selection is remembered there and put back.
        $addList.Add_MouseDown({
                param($s, $e)
                $sel = New-Object System.Collections.Generic.List[string]
                foreach ($it in $s.SelectedItems) { [void]$sel.Add([string]$it.Tag) }
                $script:WsRoleSelBefore = $sel.ToArray()
            })
        $addList.Add_MouseClick({
                param($s, $e)
                if ($e.Button -ne [System.Windows.Forms.MouseButtons]::Left) { return }
                $mods = [System.Windows.Forms.Control]::ModifierKeys
                if (($mods -band [System.Windows.Forms.Keys]::Control) -or ($mods -band [System.Windows.Forms.Keys]::Shift)) { return }
                $h = $s.HitTest($e.Location)
                if (-not $h.Item -or -not $h.SubItem) { return }
                if ($h.Item.SubItems.IndexOf($h.SubItem) -ne 3) { return }
                $paths = @(Get-WorksetRoleClickPaths @($script:WsRoleSelBefore) ([string]$h.Item.Tag))
                $keep = @{}
                foreach ($p in $paths) { $keep[[string]$p] = $true }
                foreach ($it in $s.Items) {
                    $want = $keep.ContainsKey([string]$it.Tag)
                    if ($it.Selected -ne $want) { $it.Selected = $want }
                }
                Show-WorksetCellCombo $s $h.Item 3 (Get-WorksetModelRoleChoices) ([string]$h.Item.SubItems[3].Text) "model-role" $paths
            })
    }
    Update-WorksetYearLabel
    Update-WorksetExistingLinks
}

function Update-WorksetExistingLinks {
    if (-not $script:Ui -or -not $script:Ui.ContainsKey("ExistingLinks") -or -not $script:Ui.ExistingLinks) { return }
    $lv = $script:Ui.ExistingLinks
    $lv.BeginUpdate()
    try {
        $lv.Items.Clear()
        if (-not $script:WsInventory -or -not $script:WsInventory.Links) { return }
        foreach ($e in @($script:WsInventory.Links)) {
            $it = New-Object System.Windows.Forms.ListViewItem ([string]$e.Host)
            [void]$it.SubItems.Add([string]$e.Name)
            [void]$it.SubItems.Add([string]$e.InstanceWorkset)
            [void]$it.SubItems.Add([string]$e.TypeWorkset)
            [void]$lv.Items.Add($it)
        }
    }
    finally { $lv.EndUpdate() }
}

