# Name: workset_logic.ps1
# Version: 3.1.0
# What it does: See the case README.
# Notes: Paths and project names in this portfolio copy are placeholders.

# UTF-8 WITH BOM required (PowerShell 5.1).
# Pure logic for the future "worksets and links" mode of batch_model_ops.
# No WinForms and no Revit. The window and the RBP task scripts call these functions later.
# Lists of rows are List[object] and leave the function via .ToArray(), never via @($list):
# Windows PowerShell 5.1 throws "Argument types do not match" on that wrap.

$script:WorksetSectionPlaceholder = "{section}"
$script:WorksetBuildingPlaceholder = "{building}"
$script:WorksetDefaultIgnoreName = "30_ИОС_Элементы"
$script:WorksetPlanHeader = @("action", "host", "key", "arg1", "arg2", "arg3", "code", "create_workset")

function Get-WorksetDefaultIgnoreName {
    return $script:WorksetDefaultIgnoreName
}

function Test-WorksetSameText([string]$A, [string]$B) {
    return [string]::Equals([string]$A, [string]$B, [StringComparison]::OrdinalIgnoreCase)
}

function Get-WorksetBaseName([string]$PathOrName) {
    if ([string]::IsNullOrWhiteSpace($PathOrName)) { return "" }
    $name = ($PathOrName.Trim() -replace "/", "\").Split("\")[-1].Trim()
    if ($name.Length -ge 4 -and $name.Substring($name.Length - 4).ToLowerInvariant() -eq ".rvt") {
        $name = $name.Substring(0, $name.Length - 4)
    }
    return $name
}

function ConvertTo-WorksetPathKey([string]$Path) {
    if ([string]::IsNullOrWhiteSpace($Path)) { return "" }
    return $Path.Trim().Replace("\", "/").TrimEnd("/").ToLowerInvariant()
}

# RSN and UNC only. BIM 360 / ACC / a drive letter are rejected.
function Test-WorksetStoragePath([string]$Path) {
    if ([string]::IsNullOrWhiteSpace($Path)) {
        return [pscustomobject]@{ Ok = $false; Kind = "empty" }
    }
    $t = $Path.Trim()
    if ($t.StartsWith("RSN://", [StringComparison]::OrdinalIgnoreCase)) {
        return [pscustomobject]@{ Ok = $true; Kind = "rsn" }
    }
    if ($t.StartsWith("\\")) {
        return [pscustomobject]@{ Ok = $true; Kind = "unc" }
    }
    $up = $t.ToUpperInvariant()
    if ($up.StartsWith("BIM 360://") -or $up.StartsWith("BIM360://") -or $up.StartsWith("ACC://") -or $up.StartsWith("AUTODESK DOCS://")) {
        return [pscustomobject]@{ Ok = $false; Kind = "cloud" }
    }
    return [pscustomobject]@{ Ok = $false; Kind = "local" }
}

# One run is one Revit year. An empty list and a mix of years both refuse the run.
function Test-WorksetSingleYear {
    param([string[]]$Years)
    $seen = New-Object System.Collections.Generic.List[string]
    $have = @{}
    foreach ($y in $Years) {
        if ([string]::IsNullOrWhiteSpace($y)) { continue }
        $t = $y.Trim()
        if (-not $have.ContainsKey($t)) {
            $have[$t] = $true
            [void]$seen.Add($t)
        }
    }
    $list = $seen.ToArray()
    if ($seen.Count -eq 0) {
        return [pscustomobject]@{ Ok = $false; Code = "no_year"; Year = ""; YearCount = 0; Years = $list }
    }
    if ($seen.Count -eq 1) {
        return [pscustomobject]@{ Ok = $true; Code = ""; Year = $seen[0]; YearCount = 1; Years = $list }
    }
    return [pscustomobject]@{ Ok = $false; Code = "mixed_year"; Year = ""; YearCount = $seen.Count; Years = $list }
}

# Seed for section_codes.csv. Both spellings are stored; the user edits the file afterwards.
# Pairs follow the section catalog already used by the model names in this repo
# (PROJ_IOS_*_ORG_Kxx: OV VK EOM SS PT PD HS OT, plus ITP AVT HC) and the
# architecture/structure codes named for this tool (AR KR), plus IOS/ИОС.
function Get-DefaultSectionCodes {
    $pairs = @(
        @("AR", "АР"), @("KR", "КР"), @("OV", "ОВ"), @("VK", "ВК"),
        @("EOM", "ЭОМ"), @("SS", "СС"), @("PT", "ПТ"), @("PD", "ПД"),
        @("HS", "ХС"), @("OT", "ОТ"), @("IOS", "ИОС"), @("ITP", "ИТП"),
        @("AVT", "АВТ"), @("HC", "ХЦ")
    )
    $list = New-Object System.Collections.Generic.List[object]
    foreach ($p in $pairs) {
        [void]$list.Add([pscustomobject]@{ Latin = [string]$p[0]; Cyrillic = [string]$p[1] })
    }
    return $list.ToArray()
}

function Find-SectionCode {
    param($Codes, [string]$Token)
    if ([string]::IsNullOrWhiteSpace($Token) -or $null -eq $Codes) { return $null }
    foreach ($c in $Codes) {
        if (Test-WorksetSameText $c.Latin $Token) {
            return [pscustomobject]@{ Latin = [string]$c.Latin; Cyrillic = [string]$c.Cyrillic; Script = "latin" }
        }
        if (Test-WorksetSameText $c.Cyrillic $Token) {
            return [pscustomobject]@{ Latin = [string]$c.Latin; Cyrillic = [string]$c.Cyrillic; Script = "cyrillic" }
        }
    }
    return $null
}

function Split-WorksetModelTokens([string]$Name, [string]$Separator) {
    $list = New-Object System.Collections.Generic.List[string]
    if ([string]::IsNullOrEmpty($Separator)) { return $list.ToArray() }
    $base = Get-WorksetBaseName $Name
    $parts = $base.Split([string[]]@($Separator), [System.StringSplitOptions]::None)
    foreach ($p in $parts) { [void]$list.Add([string]$p) }
    return $list.ToArray()
}

function Test-WorksetRoleSet {
    param([string[]]$Roles)
    if ($null -eq $Roles -or $Roles.Count -eq 0) { return $false }
    $section = 0
    $building = 0
    foreach ($r in $Roles) {
        $role = ([string]$r).Trim().ToLowerInvariant()
        if ($role -eq "fixed") { continue }
        if ($role -eq "section") { $section++; continue }
        if ($role -eq "building") { $building++; continue }
        if ($role -eq "ignore") { continue }
        return $false
    }
    return ($section -eq 1 -and $building -eq 1)
}

function ConvertTo-WorksetRole([string]$Role) {
    if ([string]::IsNullOrWhiteSpace($Role)) { return "" }
    return $Role.Trim().ToLowerInvariant()
}

function Get-WorksetRoleToken {
    param([string[]]$Tokens, [string[]]$Roles, [string]$Role)
    $want = ConvertTo-WorksetRole $Role
    for ($i = 0; $i -lt $Roles.Count; $i++) {
        if ((ConvertTo-WorksetRole $Roles[$i]) -eq $want) { return [string]$Tokens[$i] }
    }
    return ""
}

function Test-WorksetFixedEqual {
    param([string[]]$Left, [string[]]$Right, [string[]]$Roles)
    for ($i = 0; $i -lt $Roles.Count; $i++) {
        if ((ConvertTo-WorksetRole $Roles[$i]) -ne "fixed") { continue }
        if (-not (Test-WorksetSameText $Left[$i] $Right[$i])) { return $false }
    }
    return $true
}

# Guess roles from the new-section files and the existing files.
# A position is the section when every new name carries one token and the targets
# carry at least one other token there. A position is the building when the new
# names carry several tokens and each of them occurs among the targets.
# More than one "ignore" position, or two candidates for the same role, is not applied.
function Get-WorksetNameMaskSuggestion {
    param(
        [string[]]$NewNames,
        [string[]]$TargetNames,
        [string]$Separator = "_"
    )
    $none = @()
    $fail = {
        param($Reason)
        return [pscustomobject]@{
            Ok        = $false
            Apply     = $false
            Reason    = $Reason
            Separator = $Separator
            Roles     = $none
            RoleCount = 0
        }
    }
    $news = New-Object System.Collections.Generic.List[string]
    $targets = New-Object System.Collections.Generic.List[string]
    foreach ($n in $NewNames) { if (-not [string]::IsNullOrWhiteSpace($n)) { [void]$news.Add($n.Trim()) } }
    foreach ($n in $TargetNames) { if (-not [string]::IsNullOrWhiteSpace($n)) { [void]$targets.Add($n.Trim()) } }
    if ($news.Count -eq 0 -or $targets.Count -eq 0) { return (& $fail "empty") }
    if ([string]::IsNullOrEmpty($Separator)) { return (& $fail "separator") }

    $count = -1
    $allNames = New-Object System.Collections.Generic.List[string]
    foreach ($n in $news) { [void]$allNames.Add($n) }
    foreach ($n in $targets) { [void]$allNames.Add($n) }
    foreach ($n in $allNames) {
        $toks = @(Split-WorksetModelTokens $n $Separator)
        if ($count -lt 0) { $count = $toks.Count }
        elseif ($toks.Count -ne $count -or $count -eq 0) { return (& $fail "length") }
    }

    $roles = New-Object System.Collections.Generic.List[string]
    $sectionAt = New-Object System.Collections.Generic.List[int]
    $buildingAt = New-Object System.Collections.Generic.List[int]
    $ignoreAt = New-Object System.Collections.Generic.List[int]
    for ($i = 0; $i -lt $count; $i++) {
        $nu = @{}
        $tu = @{}
        foreach ($n in $news) {
            $toks = @(Split-WorksetModelTokens $n $Separator)
            $nu[$toks[$i].ToLowerInvariant()] = $true
        }
        foreach ($n in $targets) {
            $toks = @(Split-WorksetModelTokens $n $Separator)
            $tu[$toks[$i].ToLowerInvariant()] = $true
        }
        $role = "ignore"
        if ($nu.Count -eq 1 -and $tu.Count -eq 1) {
            $nv = @($nu.Keys)[0]
            $tv = @($tu.Keys)[0]
            if ($nv -eq $tv) { $role = "fixed" }
        }
        if ($role -eq "ignore" -and $nu.Count -eq 1) {
            $nv = @($nu.Keys)[0]
            $other = $false
            foreach ($tv in $tu.Keys) { if ($tv -ne $nv) { $other = $true } }
            if ($other) { $role = "section" }
        }
        if ($role -eq "ignore" -and $nu.Count -gt 1) {
            $allIn = $true
            foreach ($nv in $nu.Keys) { if (-not $tu.ContainsKey($nv)) { $allIn = $false } }
            if ($allIn) { $role = "building" }
        }
        [void]$roles.Add($role)
        if ($role -eq "section") { [void]$sectionAt.Add($i) }
        elseif ($role -eq "building") { [void]$buildingAt.Add($i) }
        elseif ($role -eq "ignore") { [void]$ignoreAt.Add($i) }
    }

    $roleArr = $roles.ToArray()
    $reason = "ok"
    $apply = $true
    if ($sectionAt.Count -ne 1 -or $buildingAt.Count -ne 1) {
        $apply = $false
        $reason = "ambiguous"
    }
    elseif ($ignoreAt.Count -gt 1) {
        $apply = $false
        $reason = "many_ignore"
    }
    return [pscustomobject]@{
        Ok        = $apply
        Apply     = $apply
        Reason    = $reason
        Separator = $Separator
        Roles     = $roleArr
        RoleCount = $roles.Count
    }
}

# A Revit year token such as R22 / R2024 is never the building.
function Test-WorksetYearToken([string]$Token) {
    if ([string]::IsNullOrWhiteSpace($Token)) { return $false }
    return ($Token.Trim() -match '^[RrРр]\d{2,4}$')
}

# Roles from the section-code table. Works with any mix of new and target models,
# one target and many new ones included.
#   section  = the position whose tokens are section codes (EOM, HC, AVT, АР...) in most names;
#   building = the other position whose tokens differ between the names (K1..K5, PRK);
#              when every name has the same building, the constant token next to the section
#              (not a Revit year like R22) is taken;
#   fixed    = every other position (the same text in all names).
# Reason: ok, empty, separator, length, no_codes, ambiguous.
function Get-WorksetNameMaskByCodes {
    param(
        [string[]]$Names,
        $Codes,
        [string]$Separator = "_"
    )
    $none = @()
    $fail = {
        param($Reason)
        return [pscustomobject]@{ Ok = $false; Apply = $false; Reason = $Reason; Separator = $Separator; Roles = $none; RoleCount = 0 }
    }
    $all = New-Object System.Collections.Generic.List[string]
    foreach ($n in $Names) { if (-not [string]::IsNullOrWhiteSpace($n)) { [void]$all.Add($n.Trim()) } }
    if ($all.Count -eq 0) { return (& $fail "empty") }
    if ([string]::IsNullOrEmpty($Separator)) { return (& $fail "separator") }
    $tokens = New-Object System.Collections.Generic.List[object]
    $count = -1
    foreach ($n in $all) {
        $toks = @(Split-WorksetModelTokens $n $Separator)
        if ($count -lt 0) { $count = $toks.Count }
        elseif ($toks.Count -ne $count) { return (& $fail "length") }
        [void]$tokens.Add($toks)
    }
    if ($count -lt 2) { return (& $fail "length") }
    $codeHits = New-Object System.Collections.Generic.List[int]
    $distinct = New-Object System.Collections.Generic.List[int]
    for ($i = 0; $i -lt $count; $i++) {
        $hits = 0
        $seen = @{}
        foreach ($toks in $tokens) {
            $t = [string]$toks[$i]
            $seen[$t.ToLowerInvariant()] = $true
            if ($null -ne (Find-SectionCode $Codes $t)) { $hits++ }
        }
        [void]$codeHits.Add($hits)
        [void]$distinct.Add($seen.Count)
    }
    $best = -1
    $bestHits = 0
    $tie = $false
    for ($i = 0; $i -lt $count; $i++) {
        if ($codeHits[$i] -gt $bestHits) { $best = $i; $bestHits = $codeHits[$i]; $tie = $false }
        elseif ($codeHits[$i] -eq $bestHits -and $bestHits -gt 0) { $tie = $true }
    }
    if ($best -lt 0 -or ($bestHits * 2) -lt $all.Count) { return (& $fail "no_codes") }
    if ($tie) { return (& $fail "ambiguous") }
    $varying = New-Object System.Collections.Generic.List[int]
    for ($i = 0; $i -lt $count; $i++) {
        if ($i -eq $best) { continue }
        if ($distinct[$i] -gt 1) { [void]$varying.Add($i) }
    }
    $building = -1
    if ($varying.Count -eq 1) { $building = $varying[0] }
    elseif ($varying.Count -gt 1) { return (& $fail "ambiguous") }
    else {
        # every name has the same building: the constant token next to the section
        $order = New-Object System.Collections.Generic.List[int]
        for ($d = 1; $d -lt $count; $d++) {
            if (($best + $d) -lt $count) { [void]$order.Add($best + $d) }
            if (($best - $d) -ge 0) { [void]$order.Add($best - $d) }
        }
        foreach ($i in $order) {
            $t = [string]$tokens[0][$i]
            if (Test-WorksetYearToken $t) { continue }
            if ($null -ne (Find-SectionCode $Codes $t)) { continue }
            $building = $i
            break
        }
        if ($building -lt 0) { return (& $fail "ambiguous") }
    }
    $roles = New-Object System.Collections.Generic.List[string]
    for ($i = 0; $i -lt $count; $i++) {
        if ($i -eq $best) { [void]$roles.Add("section") }
        elseif ($i -eq $building) { [void]$roles.Add("building") }
        else { [void]$roles.Add("fixed") }
    }
    return [pscustomobject]@{ Ok = $true; Apply = $true; Reason = "ok"; Separator = $Separator; Roles = $roles.ToArray(); RoleCount = $roles.Count }
}

# Which mask the pairs use, in this order:
#   user  - roles changed on «Маска имени» in this session (not saved yet);
#   saved - name_mask.cfg (never wiped by a recompute);
#   codes - detected from the section-code table (Get-WorksetNameMaskByCodes);
#   guess - the older guess from new vs target names (Get-WorksetNameMaskSuggestion);
#   none  - nothing fits: every row is «имя не разобрано».
function Resolve-WorksetNameMask {
    param(
        [string[]]$UserRoles,
        [string[]]$SavedRoles,
        [string[]]$NewNames,
        [string[]]$TargetNames,
        $Codes,
        [string]$Separator = "_"
    )
    if ([string]::IsNullOrEmpty($Separator)) { $Separator = "_" }
    if ($null -ne $UserRoles -and $UserRoles.Count -gt 0) {
        # an unfinished edit is not replaced by a guess: the user sees what is missing
        if (Test-WorksetRoleSet $UserRoles) {
            return [pscustomobject]@{ Ok = $true; Source = "user"; Reason = "ok"; Separator = $Separator; Roles = [string[]]$UserRoles }
        }
        return [pscustomobject]@{ Ok = $false; Source = "user"; Reason = "roles"; Separator = $Separator; Roles = [string[]]$UserRoles }
    }
    if ($null -ne $SavedRoles -and $SavedRoles.Count -gt 0 -and (Test-WorksetRoleSet $SavedRoles)) {
        return [pscustomobject]@{ Ok = $true; Source = "saved"; Reason = "ok"; Separator = $Separator; Roles = [string[]]$SavedRoles }
    }
    $names = New-Object System.Collections.Generic.List[string]
    foreach ($n in $NewNames) { if (-not [string]::IsNullOrWhiteSpace($n)) { [void]$names.Add($n) } }
    foreach ($n in $TargetNames) { if (-not [string]::IsNullOrWhiteSpace($n)) { [void]$names.Add($n) } }
    if ($names.Count -eq 0) {
        return [pscustomobject]@{ Ok = $false; Source = "none"; Reason = "empty"; Separator = $Separator; Roles = [string[]]@() }
    }
    $byCodes = Get-WorksetNameMaskByCodes $names.ToArray() $Codes $Separator
    if ($byCodes.Ok) {
        return [pscustomobject]@{ Ok = $true; Source = "codes"; Reason = "ok"; Separator = $Separator; Roles = [string[]]$byCodes.Roles }
    }
    $sug = Get-WorksetNameMaskSuggestion $NewNames $TargetNames $Separator
    if ($sug.Ok) {
        return [pscustomobject]@{ Ok = $true; Source = "guess"; Reason = "ok"; Separator = $Separator; Roles = [string[]]$sug.Roles }
    }
    return [pscustomobject]@{ Ok = $false; Source = "none"; Reason = [string]$byCodes.Reason; Separator = $Separator; Roles = [string[]]@() }
}

function Get-WorksetMaskSourceWord([string]$Source) {
    switch ($Source) {
        "user" { return "изменена вручную, не записана" }
        "saved" { return "из name_mask.cfg" }
        "codes" { return "по кодам разделов" }
        "guess" { return "по именам моделей" }
        default { return "не определена" }
    }
}

# One status line after the pairs are recomputed. Never silent: an empty or unparsed
# table always says why and where to fix it.
function Get-WorksetPairSummary {
    param($Rows, $Resolved, [int]$NewCount, [int]$TargetCount)
    $ins = 0; $ex = 0; $np = 0; $unp = 0; $man = 0
    $roleCount = 0
    if ($null -ne $Resolved -and $null -ne $Resolved.Roles) { $roleCount = @($Resolved.Roles).Count }
    if ($null -ne $Rows) {
        foreach ($r in $Rows) {
            $st = [string]$r.State
            if ($st -eq "insert") { $ins++ }
            elseif ($st -eq "exists" -or $st -eq "other_path") { $ex++ }
            elseif ($st -eq "no_pair") { $np++ }
            elseif ($st -eq "unparsed") { $unp++ }
            if ($r.PSObject.Properties["Source"] -and [string]$r.Source -eq "manual") { $man++ }
        }
    }
    if ($NewCount -eq 0 -and $man -eq 0) {
        return "Отметьте новые модели в колонке «Роль»."
    }
    if ($TargetCount -eq 0 -and $man -eq 0) {
        return "Нет отмеченных целевых моделей: смените роль в колонке «Роль» списка моделей."
    }
    # every text fits one status line at the minimum window width (about 110 characters);
    # a warning replaces the pair counts instead of being glued in front of them
    if ($null -ne $Resolved -and -not $Resolved.Ok -and [string]$Resolved.Reason -eq "roles") {
        return "В маске имени нужен ровно один «раздел» и один «корпус» (вкладка «Маска имени»)."
    }
    if ($null -eq $Resolved -or -not $Resolved.Ok) {
        return "Имя не разобрано: задайте раздел и корпус на вкладке «Маска имени»."
    }
    if ($unp -gt 0) {
        return ("Имя не разобрано у {0} {1}: проверьте вкладку «Маска имени»." -f $unp, (Get-RuPlural $unp "строки" "строк" "строк"))
    }
    $emptyWs = 0
    if ($null -ne $Rows) {
        foreach ($r in $Rows) {
            if ([string]$r.State -eq "insert" -and [string]::IsNullOrWhiteSpace([string]$r.WorksetName)) { $emptyWs++ }
        }
    }
    if ($emptyWs -gt 0) { return (Get-WorksetEmptyWorksetText $emptyWs) }
    $text = "Пары: вставить {0}, уже есть {1}, нет пары {2}" -f $ins, $ex, $np
    if ($man -gt 0) { $text = $text + (", вручную {0}" -f $man) }
    return $text + (". Маска имени: {0}." -f (Get-WorksetMaskSourceWord ([string]$Resolved.Source)))
}

# Longest status line that fits one line at the minimum window width.
$script:WsStatusMax = 110

# "Добавлено вручную: 1." + the pair note, or the note alone when both do not fit one line.
function Join-WorksetStatus([string]$Prefix, [string]$Note) {
    if (-not $Prefix) { return $Note }
    if (-not $Note) { return $Prefix }
    $both = $Prefix.TrimEnd() + " " + $Note
    if ($both.Length -le $script:WsStatusMax) { return $both }
    return $Note
}

# Russian plural: 1 / 21 связь, 2-4 / 22 связи, 5-20 / 11-14 связей.
function Get-RuPlural([int]$N, [string]$One, [string]$Few, [string]$Many) {
    $n = [Math]::Abs($N)
    $d = $n % 10
    $h = $n % 100
    if ($d -eq 1 -and $h -ne 11) { return $One }
    if ($d -ge 2 -and $d -le 4 -and ($h -lt 12 -or $h -gt 14)) { return $Few }
    return $Many
}

function Get-WorksetEmptyWorksetText([int]$Count) {
    return ("Нет рабочего набора у {0} {1}: прочитайте наборы и связи или введите имя." -f $Count, (Get-RuPlural $Count "строки" "строк" "строк"))
}

# Rows the «Связи» table shows on tab 4:
#   no ticked new model -> only hand-made rows (usually none: the table is empty);
#   otherwise every row except a target that got no link («нет пары» belongs to new models only).
function Select-WorksetVisiblePairs {
    param($Rows, [int]$NewCount)
    $out = New-Object System.Collections.Generic.List[object]
    $hidden = 0
    if ($null -ne $Rows) {
        foreach ($r in $Rows) {
            $manual = ($r.PSObject.Properties["Source"] -and [string]$r.Source -eq "manual")
            if ($NewCount -le 0 -and -not $manual) { $hidden++; continue }
            if ($r.TargetPath -and -not $r.LinkPath) { $hidden++; continue }
            [void]$out.Add($r)
        }
    }
    return [pscustomobject]@{ Rows = $out.ToArray(); Count = $out.Count; Hidden = $hidden }
}

# Short reason of a skipped row in the confirm form (the column stays narrow) ...
function Get-WorksetSkipReason([string]$Code) {
    switch ($Code) {
        "SKIP_EXISTS" { return "связь уже есть" }
        "SKIP_OTHER_PATH" { return "стоит по другому пути" }
        "SKIP_EMPTY_WORKSET" { return "нет рабочего набора" }
        "SKIP_BAD_NAME" { return "недопустимое имя набора" }
        "SKIP_CONFLICT" { return "имя набора уже занято" }
        "no_pair" { return "нет целевой модели этого корпуса" }
        "unparsed" { return "имя не разобрано" }
        default { return [string]$Code }
    }
}

# ... and the full one, shown as the row tool tip.
function Get-WorksetSkipReasonLong([string]$Code) {
    switch ($Code) {
        "SKIP_EXISTS" { return "Эта связь уже вставлена в модель." }
        "SKIP_OTHER_PATH" { return "В модели уже стоит этот файл, но по другому пути." }
        "SKIP_EMPTY_WORKSET" { return "Не определён рабочий набор: прочитайте наборы и связи целевых моделей или введите имя в колонке «Рабочий набор»." }
        "SKIP_BAD_NAME" { return "Имя рабочего набора содержит недопустимые символы." }
        "SKIP_CONFLICT" { return "Рабочий набор с таким именем уже есть в модели." }
        "no_pair" { return "Для этой новой модели нет отмеченной целевой модели того же корпуса." }
        "unparsed" { return "Раздел и корпус не найдены в имени: проверьте вкладку «Маска имени»." }
        default { return [string]$Code }
    }
}

# «Действие» of an insert row in the confirm form.
function Get-WorksetInsertActionText([bool]$IsNew) {
    if ($IsNew) { return "вставить + новый набор" }
    return "вставить"
}

# One sentence on top of the confirm form, with Russian plural forms.
function Get-WorksetInsertSummary([int]$Inserts, [int]$Models, [int]$Creates, [int]$Skipped = 0) {
    if ($Inserts -le 0) {
        if ($Skipped -gt 0) { return "Вставлять нечего: все строки пропущены." }
        return "Вставлять нечего."
    }
    $verb = Get-RuPlural $Inserts "вставлена" "вставлено" "вставлено"
    $text = "В {0} {1} будет {2} {3} {4}" -f $Models, (Get-RuPlural $Models "модель" "модели" "моделей"), $verb, $Inserts, (Get-RuPlural $Inserts "связь" "связи" "связей")
    if ($Creates -gt 0) {
        $made = Get-RuPlural $Creates "создан" "создано" "создано"
        $text = $text + (", будет {0} {1} {2}" -f $made, $Creates, (Get-RuPlural $Creates "рабочий набор" "рабочих набора" "рабочих наборов"))
    }
    return $text + "."
}

# What the tab-4 confirm form shows, from the plan actions and the pair rows.
# Rows: Model / Link (short names), ModelPath / LinkPath (tool tip), Workset, IsNew.
# Skipped: Model, Link, Reason. Pairs give the new models without a pair.
function Get-WorksetInsertView {
    param($Actions, $Pairs)
    $creates = @{}
    if ($null -ne $Actions) {
        foreach ($a in $Actions) {
            if ([string]$a.Action -ne "create") { continue }
            $creates[(ConvertTo-WorksetPathKey ([string]$a.Host)) + "|" + ([string]$a.Arg1).Trim().ToLowerInvariant()] = $true
        }
    }
    $rows = New-Object System.Collections.Generic.List[object]
    $skipped = New-Object System.Collections.Generic.List[object]
    $hosts = @{}
    $newSets = @{}
    if ($null -ne $Actions) {
        foreach ($a in $Actions) {
            $act = [string]$a.Action
            if ($act -eq "insert") {
                $hk = ConvertTo-WorksetPathKey ([string]$a.Host)
                $ws = ([string]$a.Arg2).Trim()
                $ck = $hk + "|" + $ws.ToLowerInvariant()
                $isNew = ($creates.ContainsKey($ck) -or [bool]$a.CreateWorkset)
                if ($isNew) { $newSets[$ck] = $true }
                $hosts[$hk] = $true
                [void]$rows.Add([pscustomobject]@{
                        Model     = (Get-WorksetBaseName ([string]$a.Host))
                        ModelPath = [string]$a.Host
                        Link      = (Get-WorksetBaseName ([string]$a.Arg1))
                        LinkPath  = [string]$a.Arg1
                        Workset   = $ws
                        IsNew     = $isNew
                        Action    = (Get-WorksetInsertActionText $isNew)
                    })
            }
            elseif ($act -eq "skip") {
                [void]$skipped.Add([pscustomobject]@{
                        Model  = (Get-WorksetBaseName ([string]$a.Host))
                        Link   = (Get-WorksetBaseName ([string]$a.Arg1))
                        Reason = (Get-WorksetSkipReason ([string]$a.Code))
                        Tip    = (Get-WorksetSkipReasonLong ([string]$a.Code))
                    })
            }
        }
    }
    if ($null -ne $Pairs) {
        foreach ($p in $Pairs) {
            $st = [string]$p.State
            if ($st -ne "no_pair" -and $st -ne "unparsed") { continue }
            if (-not $p.LinkPath) { continue }
            [void]$skipped.Add([pscustomobject]@{
                    Model  = "—"
                    Link   = (Get-WorksetBaseName ([string]$p.LinkPath))
                    Reason = (Get-WorksetSkipReason $st)
                    Tip    = (Get-WorksetSkipReasonLong $st)
                })
        }
    }
    $summary = Get-WorksetInsertSummary $rows.Count $hosts.Count $newSets.Count $skipped.Count
    return [pscustomobject]@{
        Summary     = $summary
        Rows        = $rows.ToArray()
        Skipped     = $skipped.ToArray()
        InsertCount = $rows.Count
        ModelCount  = $hosts.Count
        CreateCount = $newSets.Count
    }
}

function ConvertTo-WorksetModel($Item) {
    if ($null -eq $Item) { return $null }
    if ($Item -is [string]) {
        $path = ([string]$Item).Trim()
        return [pscustomobject]@{ Path = $path; Name = (Get-WorksetBaseName $path) }
    }
    $path = ""
    if ($Item.PSObject.Properties["Path"] -and $Item.Path) { $path = ([string]$Item.Path).Trim() }
    $name = ""
    if ($Item.PSObject.Properties["Name"] -and $Item.Name) { $name = ([string]$Item.Name).Trim() }
    if (-not $name) { $name = Get-WorksetBaseName $path }
    if (-not $path) { $path = $name }
    return [pscustomobject]@{ Path = $path; Name = $name }
}

function New-WorksetPairRow {
    param(
        [string]$Key,
        [string]$TargetPath,
        [string]$TargetName,
        [string]$LinkPath,
        [string]$LinkName,
        [string]$Section,
        [string]$Building,
        [string]$State
    )
    return [pscustomobject]@{
        Key         = $Key
        TargetPath  = $TargetPath
        TargetName  = $TargetName
        LinkPath    = $LinkPath
        LinkName    = $LinkName
        Section     = $Section
        Building    = $Building
        State       = $State
        WorksetName = ""
    }
}

# Pair every new model with every target of the same building and a different section.
# A row without a partner stays in the list with state no_pair. A name that does not
# fit the mask is unparsed. A model is not paired with itself.
function Get-WorksetLinkPairs {
    param(
        $NewModels,
        $TargetModels,
        [string]$Separator,
        [string[]]$Roles
    )
    $news = New-Object System.Collections.Generic.List[object]
    $targets = New-Object System.Collections.Generic.List[object]
    foreach ($m in $NewModels) {
        $c = ConvertTo-WorksetModel $m
        if ($null -ne $c) { [void]$news.Add($c) }
    }
    foreach ($m in $TargetModels) {
        $c = ConvertTo-WorksetModel $m
        if ($null -ne $c) { [void]$targets.Add($c) }
    }
    $rows = New-Object System.Collections.Generic.List[object]
    $roleOk = Test-WorksetRoleSet $Roles
    $newHit = @{}
    $tgtHit = @{}
    if ($roleOk) {
        foreach ($n in $news) {
            $nt = @(Split-WorksetModelTokens $n.Name $Separator)
            $nOk = ($nt.Count -eq $Roles.Count)
            if (-not $nOk) { continue }
            $nKey = ConvertTo-WorksetPathKey $n.Path
            foreach ($t in $targets) {
                $tt = @(Split-WorksetModelTokens $t.Name $Separator)
                if ($tt.Count -ne $Roles.Count) { continue }
                $tKey = ConvertTo-WorksetPathKey $t.Path
                if ($nKey -and $nKey -eq $tKey) { continue }
                $ns = Get-WorksetRoleToken $nt $Roles "section"
                $ts = Get-WorksetRoleToken $tt $Roles "section"
                $nb = Get-WorksetRoleToken $nt $Roles "building"
                $tb = Get-WorksetRoleToken $tt $Roles "building"
                if (-not (Test-WorksetSameText $nb $tb)) { continue }
                if (Test-WorksetSameText $ns $ts) { continue }
                if (-not (Test-WorksetFixedEqual $nt $tt $Roles)) { continue }
                $newHit[$nKey] = $true
                $tgtHit[$tKey] = $true
                [void]$rows.Add((New-WorksetPairRow "" $t.Path $t.Name $n.Path $n.Name $ns $nb "insert"))
            }
        }
    }
    foreach ($n in $news) {
        $nKey = ConvertTo-WorksetPathKey $n.Path
        if ($newHit.ContainsKey($nKey)) { continue }
        $nt = @(Split-WorksetModelTokens $n.Name $Separator)
        $state = "no_pair"
        if (-not $roleOk -or $nt.Count -ne $Roles.Count) { $state = "unparsed" }
        [void]$rows.Add((New-WorksetPairRow "" "" "" $n.Path $n.Name "" "" $state))
    }
    foreach ($t in $targets) {
        $tKey = ConvertTo-WorksetPathKey $t.Path
        if ($tgtHit.ContainsKey($tKey)) { continue }
        $tt = @(Split-WorksetModelTokens $t.Name $Separator)
        $state = "no_pair"
        if (-not $roleOk -or $tt.Count -ne $Roles.Count) { $state = "unparsed" }
        [void]$rows.Add((New-WorksetPairRow "" $t.Path $t.Name "" "" "" "" $state))
    }
    $i = 1
    foreach ($r in $rows) { $r.Key = ("row-{0}" -f $i); $i++ }
    return [pscustomobject]@{ Ok = [bool]$roleOk; Count = $rows.Count; Rows = $rows.ToArray() }
}

# Full path wins. The same file name stored under a different path is not inserted.
# Nested links do not count as "already there": the host still needs its own instance.
function Get-WorksetInsertStatus {
    param(
        [string]$LinkPath,
        $ExistingLinks
    )
    $want = ConvertTo-WorksetPathKey $LinkPath
    $wantName = Get-WorksetBaseName $LinkPath
    $other = $false
    if ($null -eq $ExistingLinks) {
        return [pscustomobject]@{ State = "insert"; Code = "" }
    }
    foreach ($e in $ExistingLinks) {
        $nested = $false
        if ($e.PSObject.Properties["Nested"] -and $e.Nested) { $nested = $true }
        if ($nested) { continue }
        $path = ""
        if ($e.PSObject.Properties["Path"]) { $path = [string]$e.Path }
        $name = ""
        if ($e.PSObject.Properties["Name"] -and $e.Name) { $name = [string]$e.Name }
        if (-not $name) { $name = Get-WorksetBaseName $path }
        $key = ConvertTo-WorksetPathKey $path
        if ($want -and $key -and $want -eq $key) {
            return [pscustomobject]@{ State = "exists"; Code = "SKIP_EXISTS" }
        }
        if ($wantName -and (Test-WorksetSameText $wantName $name) -and $want -ne $key) { $other = $true }
    }
    if ($other) { return [pscustomobject]@{ State = "other_path"; Code = "SKIP_OTHER_PATH" } }
    return [pscustomobject]@{ State = "insert"; Code = "" }
}

function Update-WorksetPairInventory {
    param($Pairs, $ExistingLinks)
    if ($null -eq $Pairs) { return }
    foreach ($p in $Pairs) {
        if ($p.State -ne "insert") { continue }
        if ([string]::IsNullOrWhiteSpace($p.LinkPath) -or [string]::IsNullOrWhiteSpace($p.TargetPath)) { continue }
        $mine = New-Object System.Collections.Generic.List[object]
        $hostKey = ConvertTo-WorksetPathKey $p.TargetPath
        if ($null -ne $ExistingLinks) {
            foreach ($e in $ExistingLinks) {
                $h = ""
                if ($e.PSObject.Properties["Host"]) { $h = [string]$e.Host }
                if ((ConvertTo-WorksetPathKey $h) -ne $hostKey) { continue }
                [void]$mine.Add($e)
            }
        }
        $st = Get-WorksetInsertStatus $p.LinkPath $mine.ToArray()
        $p.State = $st.State
    }
}

function Test-WorksetIgnoredName {
    param([string]$Name, [string[]]$IgnoreNames)
    if ($null -eq $IgnoreNames) { $IgnoreNames = @($script:WorksetDefaultIgnoreName) }
    $n = ""
    if ($null -ne $Name) { $n = $Name.Trim() }
    foreach ($g in $IgnoreNames) {
        if (Test-WorksetSameText $n $g) { return $true }
    }
    return $false
}

# Split a workset name into tokens on "_" "-" and whitespace, keeping the separators.
function Get-WorksetNamePieces([string]$Name) {
    $tokens = New-Object System.Collections.Generic.List[string]
    $seps = New-Object System.Collections.Generic.List[string]
    if ($null -eq $Name) { $Name = "" }
    $rx = [regex]'[^_\-\s]+'
    $idx = 0
    foreach ($m in $rx.Matches($Name)) {
        [void]$seps.Add($Name.Substring($idx, $m.Index - $idx))
        [void]$tokens.Add($m.Value)
        $idx = $m.Index + $m.Length
    }
    $tail = ""
    if ($idx -lt $Name.Length) { $tail = $Name.Substring($idx) }
    # Keep the List[string]: a one-element array stored on a pscustomobject
    # would unwrap to a string and Tokens[0] would be the first character.
    return [pscustomobject]@{ Tokens = $tokens; Seps = $seps; Tail = $tail; TokenCount = $tokens.Count }
}

function Join-WorksetNamePieces($Pieces, [int]$ReplaceAt, [string]$Replacement) {
    $sb = New-Object System.Text.StringBuilder
    for ($i = 0; $i -lt $Pieces.TokenCount; $i++) {
        [void]$sb.Append([string]$Pieces.Seps[$i])
        if ($i -eq $ReplaceAt) { [void]$sb.Append($Replacement) }
        else { [void]$sb.Append([string]$Pieces.Tokens[$i]) }
    }
    [void]$sb.Append([string]$Pieces.Tail)
    return $sb.ToString()
}

# Same as Join-WorksetNamePieces with two placeholders: section and (optional) building.
function Join-WorksetNamePiecesMask($Pieces, [int]$SectionAt, [int]$BuildingAt) {
    $sb = New-Object System.Text.StringBuilder
    for ($i = 0; $i -lt $Pieces.TokenCount; $i++) {
        [void]$sb.Append([string]$Pieces.Seps[$i])
        if ($i -eq $SectionAt) { [void]$sb.Append($script:WorksetSectionPlaceholder) }
        elseif ($i -eq $BuildingAt) { [void]$sb.Append($script:WorksetBuildingPlaceholder) }
        else { [void]$sb.Append([string]$Pieces.Tokens[$i]) }
    }
    [void]$sb.Append([string]$Pieces.Tail)
    return $sb.ToString()
}

# The mask is learned from link worksets of ONE model.
# "30_ИОС_Элементы" is the active default workset (elements and links with no
# workset of their own land there). It is not a link workset and does not vote,
# even when its tokens look like a section code. Nested links do not vote.
# Every voting link must produce the same mask and the same script, otherwise the
# proposal is empty and the user types the name.
function Get-WorksetLinkMask {
    param(
        $Links,
        $Codes,
        [string]$Separator,
        [string[]]$Roles,
        [string[]]$IgnoreNames
    )
    if ($null -eq $IgnoreNames) { $IgnoreNames = @($script:WorksetDefaultIgnoreName) }
    $votes = New-Object System.Collections.Generic.List[object]
    if ($null -ne $Links) {
        foreach ($link in $Links) {
            $nested = $false
            if ($link.PSObject.Properties["Nested"] -and $link.Nested) { $nested = $true }
            if ($nested) { continue }
            $ws = ""
            if ($link.PSObject.Properties["Workset"]) { $ws = [string]$link.Workset }
            if (Test-WorksetIgnoredName $ws $IgnoreNames) { continue }
            $name = ""
            if ($link.PSObject.Properties["Name"]) { $name = [string]$link.Name }
            if (-not $name -and $link.PSObject.Properties["Path"]) { $name = Get-WorksetBaseName ([string]$link.Path) }
            $toks = @(Split-WorksetModelTokens $name $Separator)
            if ($toks.Count -ne $Roles.Count) { continue }
            $section = Get-WorksetRoleToken $toks $Roles "section"
            $row = Find-SectionCode $Codes $section
            if ($null -eq $row) { continue }
            $pieces = Get-WorksetNamePieces $ws
            $hitAt = -1
            $hitScript = ""
            $hits = 0
            for ($i = 0; $i -lt $pieces.TokenCount; $i++) {
                $token = [string]$pieces.Tokens[$i]
                $script = ""
                if (Test-WorksetSameText $token $row.Cyrillic) { $script = "cyrillic" }
                elseif (Test-WorksetSameText $token $row.Latin) { $script = "latin" }
                if (-not $script) { continue }
                $hits++
                $hitAt = $i
                $hitScript = $script
            }
            if ($hits -ne 1) { continue }
            # the building of the link inside the workset name (01_Связь_RVT_AR_K1) becomes
            # {building}, so links of other buildings (…_AR_PRK) vote for the same mask
            $buildingAt = -1
            $building = Get-WorksetRoleToken $toks $Roles "building"
            if ($building) {
                $bHits = 0
                for ($i = 0; $i -lt $pieces.TokenCount; $i++) {
                    if ($i -eq $hitAt) { continue }
                    if (Test-WorksetSameText ([string]$pieces.Tokens[$i]) $building) { $bHits++; $buildingAt = $i }
                }
                if ($bHits -ne 1) { $buildingAt = -1 }
            }
            $mask = Join-WorksetNamePiecesMask $pieces $hitAt $buildingAt
            [void]$votes.Add([pscustomobject]@{ Mask = $mask; Script = $hitScript })
        }
    }
    if ($votes.Count -eq 0) {
        return [pscustomobject]@{ Ok = $false; Mask = ""; Script = ""; Reason = "no_votes"; Votes = 0 }
    }
    $mask0 = [string]$votes[0].Mask
    $script0 = [string]$votes[0].Script
    $agree = $true
    foreach ($v in $votes) {
        if ($v.Mask -cne $mask0 -or $v.Script -cne $script0) { $agree = $false }
    }
    if ($agree) {
        return [pscustomobject]@{ Ok = $true; Mask = $mask0; Script = $script0; Reason = "ok"; Votes = $votes.Count }
    }
    # more than half of the links agree: that mask wins; a tie proposes nothing
    $tally = @{}
    $order = New-Object System.Collections.Generic.List[string]
    foreach ($v in $votes) {
        $k = [string]$v.Script + "|" + [string]$v.Mask
        if (-not $tally.ContainsKey($k)) { $tally[$k] = 0; [void]$order.Add($k) }
        $tally[$k] = [int]$tally[$k] + 1
    }
    foreach ($k in $order) {
        if (([int]$tally[$k] * 2) -gt $votes.Count) {
            $bar = $k.IndexOf("|")
            return [pscustomobject]@{ Ok = $true; Mask = $k.Substring($bar + 1); Script = $k.Substring(0, $bar); Reason = "majority"; Votes = $votes.Count }
        }
    }
    return [pscustomobject]@{ Ok = $false; Mask = ""; Script = ""; Reason = "disagree"; Votes = $votes.Count }
}

function Get-ProposedWorksetName {
    param(
        [string]$Mask,
        [string]$Script,
        [string]$SectionToken,
        $Codes,
        [string]$BuildingToken = ""
    )
    if ([string]::IsNullOrWhiteSpace($Mask) -or $Mask.IndexOf($script:WorksetSectionPlaceholder) -lt 0) { return "" }
    $row = Find-SectionCode $Codes $SectionToken
    if ($null -eq $row) { return "" }
    $spell = $row.Latin
    if ($Script -eq "cyrillic") { $spell = $row.Cyrillic }
    $at = $Mask.IndexOf($script:WorksetSectionPlaceholder)
    $out = $Mask.Substring(0, $at) + $spell + $Mask.Substring($at + $script:WorksetSectionPlaceholder.Length)
    $bAt = $out.IndexOf($script:WorksetBuildingPlaceholder)
    if ($bAt -ge 0) {
        if ([string]::IsNullOrWhiteSpace($BuildingToken)) { return "" }
        $out = $out.Substring(0, $bAt) + $BuildingToken.Trim() + $out.Substring($bAt + $script:WorksetBuildingPlaceholder.Length)
    }
    return $out
}

function Add-WorksetProposals {
    param($Pairs, $Masks, $Codes)
    if ($null -eq $Pairs) { return }
    foreach ($p in $Pairs) {
        $p.WorksetName = ""
        if ([string]::IsNullOrWhiteSpace($p.TargetPath) -or [string]::IsNullOrWhiteSpace($p.LinkPath)) { continue }
        if ($p.State -eq "unparsed" -or $p.State -eq "no_pair") { continue }
        $key = ConvertTo-WorksetPathKey $p.TargetPath
        $mask = $null
        if ($null -ne $Masks) {
            foreach ($m in $Masks) {
                $h = ""
                if ($m.PSObject.Properties["Host"]) { $h = [string]$m.Host }
                if ((ConvertTo-WorksetPathKey $h) -eq $key) { $mask = $m; break }
            }
        }
        if ($null -eq $mask -or -not $mask.Ok) { continue }
        $building = ""
        if ($p.PSObject.Properties["Building"]) { $building = [string]$p.Building }
        $p.WorksetName = Get-ProposedWorksetName $mask.Mask $mask.Script $p.Section $Codes $building
    }
}

# Section and building of one model name: by the mask roles when the name fits them,
# otherwise the first token found in the section-code table (building stays empty).
function Get-WorksetNameParts {
    param([string]$Name, [string]$Separator, [string[]]$Roles, $Codes)
    $toks = @(Split-WorksetModelTokens $Name $Separator)
    if ((Test-WorksetRoleSet $Roles) -and $toks.Count -eq $Roles.Count) {
        return [pscustomobject]@{
            Section  = (Get-WorksetRoleToken $toks $Roles "section")
            Building = (Get-WorksetRoleToken $toks $Roles "building")
            Parsed   = $true
        }
    }
    foreach ($t in $toks) {
        if ($null -ne (Find-SectionCode $Codes ([string]$t))) {
            return [pscustomobject]@{ Section = [string]$t; Building = ""; Parsed = $false }
        }
    }
    return [pscustomobject]@{ Section = ""; Building = ""; Parsed = $false }
}

# Pairs the user added by hand («Добавить связь вручную…»): Manual is a list of
# { TargetPath; LinkPath }. Each becomes an insert row with Source = manual, outside the
# same-building rule. An automatic row for the same target and link wins (no duplicate).
# One-sided rows (no_pair / unparsed) of the models now covered by a hand pair are dropped.
# Every row leaves with a Source: auto or manual.
function Merge-WorksetManualPairs {
    param($Rows, $Manual, [string]$Separator, [string[]]$Roles, $Codes)
    $out = New-Object System.Collections.Generic.List[object]
    $have = @{}
    if ($null -ne $Rows) {
        foreach ($r in $Rows) {
            if (-not $r.PSObject.Properties["Source"]) { $r | Add-Member -NotePropertyName Source -NotePropertyValue "auto" -Force }
            if ($r.TargetPath -and $r.LinkPath) {
                $have[(ConvertTo-WorksetPathKey $r.TargetPath) + "|" + (ConvertTo-WorksetPathKey $r.LinkPath)] = $true
            }
        }
    }
    $added = New-Object System.Collections.Generic.List[object]
    $coveredLink = @{}
    $coveredTarget = @{}
    if ($null -ne $Manual) {
        foreach ($m in $Manual) {
            if ($null -eq $m) { continue }
            $tp = ([string]$m.TargetPath).Trim()
            $lp = ([string]$m.LinkPath).Trim()
            if (-not $tp -or -not $lp) { continue }
            $tk = ConvertTo-WorksetPathKey $tp
            $lk = ConvertTo-WorksetPathKey $lp
            if ($tk -eq $lk) { continue }
            $coveredLink[$lk] = $true
            $coveredTarget[$tk] = $true
            $k = $tk + "|" + $lk
            if ($have.ContainsKey($k)) { continue }
            $have[$k] = $true
            $parts = Get-WorksetNameParts (Get-WorksetBaseName $lp) $Separator $Roles $Codes
            $row = New-WorksetPairRow "" $tp (Get-WorksetBaseName $tp) $lp (Get-WorksetBaseName $lp) $parts.Section $parts.Building "insert"
            $row | Add-Member -NotePropertyName Source -NotePropertyValue "manual" -Force
            [void]$added.Add($row)
        }
    }
    if ($null -ne $Rows) {
        foreach ($r in $Rows) {
            $st = [string]$r.State
            $oneSided = (-not $r.TargetPath) -or (-not $r.LinkPath)
            if ($oneSided -and ($st -eq "no_pair" -or $st -eq "unparsed")) {
                if ($r.LinkPath -and $coveredLink.ContainsKey((ConvertTo-WorksetPathKey $r.LinkPath))) { continue }
                if ($r.TargetPath -and $coveredTarget.ContainsKey((ConvertTo-WorksetPathKey $r.TargetPath))) { continue }
            }
            [void]$out.Add($r)
        }
    }
    foreach ($r in $added) { [void]$out.Add($r) }
    return $out.ToArray()
}

function Test-WorksetNameText([string]$Name) {
    if ([string]::IsNullOrWhiteSpace($Name)) {
        return [pscustomobject]@{ Ok = $false; Code = "empty"; Message = "Имя набора пустое" }
    }
    $n = $Name.Trim()
    if ($n.Length -gt 255) {
        return [pscustomobject]@{ Ok = $false; Code = "long"; Message = "Имя набора длиннее 255 знаков" }
    }
    $bad = @('\', ':', '{', '}', '[', ']', '|', ';', '<', '>', '?', '~')
    foreach ($ch in $n.ToCharArray()) {
        if ($bad -contains [string]$ch) {
            return [pscustomobject]@{ Ok = $false; Code = "chars"; Message = "В имени набора недопустимый символ" }
        }
    }
    return [pscustomobject]@{ Ok = $true; Code = ""; Message = "" }
}

# Empty or illegal workset on a row that would be inserted blocks the launch.
# no_pair, unparsed, exists and other_path do not.
function Test-WorksetLaunch {
    param($Pairs)
    $empty = 0
    $invalid = 0
    if ($null -ne $Pairs) {
        foreach ($p in $Pairs) {
            if ($p.State -ne "insert") { continue }
            $check = Test-WorksetNameText ([string]$p.WorksetName)
            if ($check.Code -eq "empty") { $empty++ }
            elseif (-not $check.Ok) { $invalid++ }
        }
    }
    return [pscustomobject]@{
        Block   = (($empty + $invalid) -gt 0)
        Empty   = $empty
        Invalid = $invalid
    }
}

# Literal find/replace. Case-sensitive unless IgnoreCase is set.
# The source string is not changed; the caller keeps it as "Было".
function ConvertTo-WorksetReplacedName {
    param(
        [string]$Name,
        [string]$Find,
        [string]$Replace,
        [bool]$IgnoreCase = $false
    )
    if ($null -eq $Name) { return "" }
    if ([string]::IsNullOrEmpty($Find)) { return $Name }
    if ($null -eq $Replace) { $Replace = "" }
    $cmp = [StringComparison]::Ordinal
    if ($IgnoreCase) { $cmp = [StringComparison]::OrdinalIgnoreCase }
    $sb = New-Object System.Text.StringBuilder
    $i = 0
    while ($i -le $Name.Length) {
        $hit = $Name.IndexOf($Find, $i, $cmp)
        if ($hit -lt 0) {
            [void]$sb.Append($Name.Substring($i))
            break
        }
        if ($hit -gt $i) { [void]$sb.Append($Name.Substring($i, $hit - $i)) }
        [void]$sb.Append($Replace)
        $i = $hit + $Find.Length
        if ($Find.Length -eq 0) { break }
    }
    return $sb.ToString()
}

function Group-WorksetRename {
    param($Worksets, [string]$HostFilter = "")
    $rows = New-Object System.Collections.Generic.List[object]
    $filterKey = ConvertTo-WorksetPathKey $HostFilter
    if ($filterKey) {
        if ($null -ne $Worksets) {
            foreach ($w in $Worksets) {
                if ((ConvertTo-WorksetPathKey ([string]$w.Host)) -ne $filterKey) { continue }
                [void]$rows.Add([pscustomobject]@{
                        Name      = [string]$w.Name
                        Id        = [string]$w.Id
                        Host      = [string]$w.Host
                        HostCount = 1
                        HostList  = [string]$w.Host
                    })
            }
        }
        return [pscustomobject]@{ Count = $rows.Count; Filtered = $true; Rows = $rows.ToArray() }
    }
    $order = New-Object System.Collections.Generic.List[string]
    $by = @{}
    if ($null -ne $Worksets) {
        foreach ($w in $Worksets) {
            $key = ([string]$w.Name).ToLowerInvariant()
            if (-not $by.ContainsKey($key)) {
                $by[$key] = New-Object System.Collections.Generic.List[object]
                [void]$order.Add($key)
            }
            [void]$by[$key].Add($w)
        }
    }
    foreach ($key in $order) {
        $grp = $by[$key]
        $hosts = New-Object System.Collections.Generic.List[string]
        foreach ($w in $grp) { [void]$hosts.Add([string]$w.Host) }
        [void]$rows.Add([pscustomobject]@{
                Name      = [string]$grp[0].Name
                Id        = ""
                Host      = ""
                HostCount = $grp.Count
                HostList  = ($hosts.ToArray() -join "|")
            })
    }
    return [pscustomobject]@{ Count = $rows.Count; Filtered = $false; Rows = $rows.ToArray() }
}

# A new name that already belongs to another workset of that model (and that
# workset is not itself renamed away in this batch) is a conflict for that model only.
function Get-WorksetRenameConflicts {
    param($Worksets, $Edits)
    $byHost = @{}
    $hostOrder = New-Object System.Collections.Generic.List[string]
    if ($null -ne $Worksets) {
        foreach ($w in $Worksets) {
            $hk = ConvertTo-WorksetPathKey ([string]$w.Host)
            if (-not $byHost.ContainsKey($hk)) {
                $byHost[$hk] = New-Object System.Collections.Generic.List[object]
                [void]$hostOrder.Add($hk)
            }
            [void]$byHost[$hk].Add($w)
        }
    }
    $out = New-Object System.Collections.Generic.List[object]
    if ($null -eq $Edits) {
        return [pscustomobject]@{ Count = 0; Rows = $out.ToArray() }
    }
    foreach ($ed in $Edits) {
        $old = [string]$ed.OldName
        $new = [string]$ed.NewName
        if ([string]::IsNullOrWhiteSpace($old) -or [string]::IsNullOrWhiteSpace($new)) { continue }
        if (Test-WorksetSameText $old $new) { continue }
        $okHosts = New-Object System.Collections.Generic.List[string]
        $badHosts = New-Object System.Collections.Generic.List[string]
        $seenHost = 0
        foreach ($hk in $hostOrder) {
            $grp = $byHost[$hk]
            $has = $false
            foreach ($w in $grp) { if (Test-WorksetSameText ([string]$w.Name) $old) { $has = $true } }
            if (-not $has) { continue }
            $seenHost++
            $final = @{}
            foreach ($w in $grp) {
                $name = [string]$w.Name
                foreach ($other in $Edits) {
                    if (Test-WorksetSameText $name ([string]$other.OldName) -and -not (Test-WorksetSameText ([string]$other.OldName) ([string]$other.NewName))) {
                        $name = [string]$other.NewName
                        break
                    }
                }
                $fk = $name.ToLowerInvariant()
                if (-not $final.ContainsKey($fk)) { $final[$fk] = 0 }
                $final[$fk] = [int]$final[$fk] + 1
            }
            $nk = $new.ToLowerInvariant()
            if ($final.ContainsKey($nk) -and [int]$final[$nk] -gt 1) { [void]$badHosts.Add([string]$grp[0].Host) }
            else { [void]$okHosts.Add([string]$grp[0].Host) }
        }
        if ($seenHost -eq 0) { continue }
        [void]$out.Add([pscustomobject]@{
                OldName       = $old
                NewName       = $new
                HostCount     = $seenHost
                ConflictCount = $badHosts.Count
                OkCount       = $okHosts.Count
                ConflictList  = ($badHosts.ToArray() -join "|")
                OkList        = ($okHosts.ToArray() -join "|")
            })
    }
    return [pscustomobject]@{ Count = $out.Count; Rows = $out.ToArray() }
}

function New-WorksetAction {
    param(
        [string]$Action,
        [string]$ModelHost,
        [string]$Key,
        [string]$Arg1,
        [string]$Arg2,
        [string]$Arg3,
        [string]$Code,
        [bool]$CreateWorkset
    )
    return [pscustomobject]@{
        Action         = $Action
        Host           = $ModelHost
        Key            = $Key
        Arg1           = $Arg1
        Arg2           = $Arg2
        Arg3           = $Arg3
        Code           = $Code
        CreateWorkset  = $CreateWorkset
    }
}

# Rename actions, then one insert per row. Creating a missing workset is part of
# that insert (later the RBP script does it in the same transaction as the
# placement and moves both the instance and the type onto the new workset).
# The default workset "30_ИОС_Элементы" is never proposed here.
function Get-WorksetApplyActions {
    param($Pairs, $Worksets, $RenameRows)
    $actions = New-Object System.Collections.Generic.List[object]
    $have = @{}
    if ($null -ne $Worksets) {
        foreach ($w in $Worksets) {
            $hk = ConvertTo-WorksetPathKey ([string]$w.Host)
            if (-not $have.ContainsKey($hk)) { $have[$hk] = @{} }
            $have[$hk][([string]$w.Name).ToLowerInvariant()] = $true
        }
    }
    if ($null -ne $RenameRows) {
        foreach ($rn in $RenameRows) {
            $ok = @{}
            foreach ($h in ([string]$rn.OkList).Split("|")) {
                if ($h) { $ok[(ConvertTo-WorksetPathKey $h)] = $h }
            }
            $bad = @{}
            foreach ($h in ([string]$rn.ConflictList).Split("|")) {
                if ($h) { $bad[(ConvertTo-WorksetPathKey $h)] = $h }
            }
            foreach ($hk in @($ok.Keys)) {
                $id = ""
                if ($null -ne $Worksets) {
                    foreach ($w in $Worksets) {
                        if ((ConvertTo-WorksetPathKey ([string]$w.Host)) -eq $hk -and (Test-WorksetSameText ([string]$w.Name) $rn.OldName)) {
                            $id = [string]$w.Id
                        }
                    }
                }
                [void]$actions.Add((New-WorksetAction "rename" $ok[$hk] $id $rn.OldName $rn.NewName "" "" $false))
                if (-not $have.ContainsKey($hk)) { $have[$hk] = @{} }
                [void]$have[$hk].Remove($rn.OldName.ToLowerInvariant())
                $have[$hk][$rn.NewName.ToLowerInvariant()] = $true
            }
            foreach ($hk in $bad.Keys) {
                [void]$actions.Add((New-WorksetAction "skip" $bad[$hk] "" $rn.OldName $rn.NewName "" "SKIP_CONFLICT" $false))
            }
        }
    }
    if ($null -ne $Pairs) {
        foreach ($p in $Pairs) {
            if ($p.State -eq "insert") {
                $check = Test-WorksetNameText ([string]$p.WorksetName)
                if (-not $check.Ok) {
                    $code = "SKIP_BAD_NAME"
                    if ($check.Code -eq "empty") { $code = "SKIP_EMPTY_WORKSET" }
                    [void]$actions.Add((New-WorksetAction "skip" $p.TargetPath $p.Key $p.LinkPath $p.WorksetName $p.Section $code $false))
                    continue
                }
                $hk = ConvertTo-WorksetPathKey $p.TargetPath
                if (-not $have.ContainsKey($hk)) { $have[$hk] = @{} }
                $nk = $p.WorksetName.Trim().ToLowerInvariant()
                $create = -not $have[$hk].ContainsKey($nk)
                if ($create) { $have[$hk][$nk] = $true }
                [void]$actions.Add((New-WorksetAction "insert" $p.TargetPath $p.Key $p.LinkPath $p.WorksetName.Trim() $p.Section "" $create))
                continue
            }
            $code = ""
            if ($p.State -eq "exists") { $code = "SKIP_EXISTS" }
            elseif ($p.State -eq "other_path") { $code = "SKIP_OTHER_PATH" }
            elseif ($p.State -eq "no_pair") { $code = "SKIP_NO_PAIR" }
            elseif ($p.State -eq "unparsed") { $code = "SKIP_UNPARSED" }
            else { $code = "SKIP" }
            $modelHost = $p.TargetPath
            if (-not $modelHost) { $modelHost = $p.LinkPath }
            [void]$actions.Add((New-WorksetAction "skip" $modelHost $p.Key $p.LinkPath $p.WorksetName $p.Section $code $false))
        }
    }
    return [pscustomobject]@{ Count = $actions.Count; Rows = $actions.ToArray() }
}

# Rename by id, then create missing names, then insert links (step 5).
# Edits are already per host: Host, Id, OldName, NewName.
# NeededNames: Host + Name to create when it is not already there.
# Inserts: pair rows. Only state "insert" becomes an insert action.
# "exists" and "other_path" are skips. Omit Inserts to plan renames and creates only.
function Get-WorksetRenameCreatePlan {
    param($Worksets, $Edits, $NeededNames, $Inserts)
    $byHost = @{}
    $hostOrder = New-Object System.Collections.Generic.List[string]
    if ($null -ne $Worksets) {
        foreach ($w in $Worksets) {
            $hk = ConvertTo-WorksetPathKey ([string]$w.Host)
            if (-not $byHost.ContainsKey($hk)) {
                $byHost[$hk] = New-Object System.Collections.Generic.List[object]
                [void]$hostOrder.Add($hk)
            }
            [void]$byHost[$hk].Add($w)
        }
    }
    $editsByHost = @{}
    if ($null -ne $Edits) {
        foreach ($e in $Edits) {
            if ([string]::IsNullOrWhiteSpace([string]$e.OldName)) { continue }
            if (Test-WorksetSameText ([string]$e.OldName) ([string]$e.NewName)) { continue }
            $hk = ConvertTo-WorksetPathKey ([string]$e.Host)
            if (-not $editsByHost.ContainsKey($hk)) { $editsByHost[$hk] = New-Object System.Collections.Generic.List[object] }
            [void]$editsByHost[$hk].Add($e)
            if (-not $byHost.ContainsKey($hk)) {
                $byHost[$hk] = New-Object System.Collections.Generic.List[object]
                [void]$hostOrder.Add($hk)
            }
        }
    }
    $actions = New-Object System.Collections.Generic.List[object]
    $have = @{}
    foreach ($hk in $hostOrder) {
        $grp = $byHost[$hk]
        $hostPath = ""
        if ($grp.Count -gt 0) { $hostPath = [string]$grp[0].Host }
        $final = @{}
        foreach ($w in $grp) {
            $final[[string]$w.Id] = [string]$w.Name
        }
        $hostEdits = @()
        if ($editsByHost.ContainsKey($hk)) { $hostEdits = @($editsByHost[$hk].ToArray()) }
        foreach ($e in $hostEdits) {
            $id = [string]$e.Id
            if ($final.ContainsKey($id)) { $final[$id] = [string]$e.NewName }
        }
        $counts = @{}
        foreach ($id in @($final.Keys)) {
            $nk = ([string]$final[$id]).ToLowerInvariant()
            if (-not $counts.ContainsKey($nk)) { $counts[$nk] = 0 }
            $counts[$nk] = [int]$counts[$nk] + 1
        }
        foreach ($e in $hostEdits) {
            $newName = [string]$e.NewName
            $chk = Test-WorksetNameText $newName
            if (-not $chk.Ok) {
                [void]$actions.Add((New-WorksetAction "skip" $hostPath ([string]$e.Id) ([string]$e.OldName) $newName "" "SKIP_BAD_NAME" $false))
                continue
            }
            $nk = $newName.ToLowerInvariant()
            if ($counts.ContainsKey($nk) -and [int]$counts[$nk] -gt 1) {
                [void]$actions.Add((New-WorksetAction "skip" $hostPath ([string]$e.Id) ([string]$e.OldName) $newName "" "SKIP_CONFLICT" $false))
                continue
            }
            if (-not $hostPath) { $hostPath = [string]$e.Host }
            [void]$actions.Add((New-WorksetAction "rename" $hostPath ([string]$e.Id) ([string]$e.OldName) $newName "" "" $false))
        }
        $names = @{}
        foreach ($id in @($final.Keys)) {
            $shown = [string]$final[$id]
            $skip = $false
            foreach ($e in $hostEdits) {
                if ([string]$e.Id -eq [string]$id) {
                    foreach ($a in $actions) {
                        if ($a.Action -eq "skip" -and [string]$a.Key -eq [string]$id -and (ConvertTo-WorksetPathKey ([string]$a.Host)) -eq $hk) {
                            $shown = [string]$e.OldName
                        }
                    }
                }
            }
            $names[$shown.ToLowerInvariant()] = $true
        }
        $have[$hk] = $names
        if (-not $hostPath -and $grp.Count -gt 0) { $hostPath = [string]$grp[0].Host }
        $have[$hk + "|path"] = $hostPath
    }
    if ($null -ne $NeededNames) {
        foreach ($n in $NeededNames) {
            $name = ""
            if ($n.PSObject.Properties["Name"]) { $name = [string]$n.Name }
            $hostPath = ""
            if ($n.PSObject.Properties["Host"]) { $hostPath = [string]$n.Host }
            $chk = Test-WorksetNameText $name
            if (-not $chk.Ok) { continue }
            $hk = ConvertTo-WorksetPathKey $hostPath
            if (-not $have.ContainsKey($hk)) { $have[$hk] = @{} }
            $nk = $name.Trim().ToLowerInvariant()
            if ($have[$hk].ContainsKey($nk)) { continue }
            $have[$hk][$nk] = $true
            [void]$actions.Add((New-WorksetAction "create" $hostPath "" $name.Trim() "" "" "" $true))
        }
    }
    if ($null -ne $Inserts) {
        foreach ($p in $Inserts) {
            $state = [string]$p.State
            $hostPath = [string]$p.TargetPath
            $link = [string]$p.LinkPath
            $wsName = ""
            if ($p.PSObject.Properties["WorksetName"]) { $wsName = [string]$p.WorksetName }
            if ($state -eq "exists") {
                [void]$actions.Add((New-WorksetAction "skip" $hostPath ([string]$p.Key) $link $wsName "" "SKIP_EXISTS" $false))
                continue
            }
            if ($state -eq "other_path") {
                [void]$actions.Add((New-WorksetAction "skip" $hostPath ([string]$p.Key) $link $wsName "" "SKIP_OTHER_PATH" $false))
                continue
            }
            if ($state -ne "insert") { continue }
            if (-not $link -or -not $hostPath) { continue }
            $wsCheck = Test-WorksetNameText $wsName
            if (-not $wsCheck.Ok) {
                $wsCode = "SKIP_BAD_NAME"
                if ($wsCheck.Code -eq "empty") { $wsCode = "SKIP_EMPTY_WORKSET" }
                [void]$actions.Add((New-WorksetAction "skip" $hostPath ([string]$p.Key) $link $wsName "" $wsCode $false))
                continue
            }
            [void]$actions.Add((New-WorksetAction "insert" $hostPath ([string]$p.Key) $link $wsName.Trim() ([string]$p.Section) "" $false))
        }
    }
    $renames = New-Object System.Collections.Generic.List[object]
    $creates = New-Object System.Collections.Generic.List[object]
    $inserts = New-Object System.Collections.Generic.List[object]
    $skips = New-Object System.Collections.Generic.List[object]
    foreach ($a in $actions) {
        if ($a.Action -eq "rename") { [void]$renames.Add($a) }
        elseif ($a.Action -eq "create") { [void]$creates.Add($a) }
        elseif ($a.Action -eq "insert") { [void]$inserts.Add($a) }
        else { [void]$skips.Add($a) }
    }
    $ordered = New-Object System.Collections.Generic.List[object]
    foreach ($a in $renames) { [void]$ordered.Add($a) }
    foreach ($a in $creates) { [void]$ordered.Add($a) }
    foreach ($a in $inserts) { [void]$ordered.Add($a) }
    foreach ($a in $skips) { [void]$ordered.Add($a) }
    $apply = 0
    foreach ($a in $ordered) {
        if ($a.Action -eq "rename" -or $a.Action -eq "create" -or $a.Action -eq "insert") { $apply++ }
    }
    return [pscustomobject]@{ Count = $ordered.Count; ApplyCount = $apply; Rows = $ordered.ToArray() }
}

function Test-WorksetStep4Block {
    param($Pairs, $Edits)
    $launch = Test-WorksetLaunch $Pairs
    $bad = 0
    if ($null -ne $Edits) {
        foreach ($e in $Edits) {
            if (Test-WorksetSameText ([string]$e.OldName) ([string]$e.NewName)) { continue }
            $chk = Test-WorksetNameText ([string]$e.NewName)
            if (-not $chk.Ok) { $bad++ }
        }
    }
    $noPair = 0
    if ($null -ne $Pairs) {
        foreach ($p in $Pairs) {
            $st = [string]$p.State
            if ($st -eq "no_pair" -or $st -eq "unparsed") { $noPair++ }
        }
    }
    $msgs = New-Object System.Collections.Generic.List[string]
    if ($bad -gt 0) { [void]$msgs.Add(("Недопустимое новое имя набора: {0}." -f $bad)) }
    return [pscustomobject]@{
        # rows without a pair or without a workset are skipped and listed in the confirm form,
        # they do not stop the inserts of the other rows; a bad rename on tab 5 still stops
        Block     = ($bad -gt 0)
        Empty     = $launch.Empty
        Invalid   = $launch.Invalid
        NoPair    = $noPair
        BadRename = $bad
        Message   = ($msgs.ToArray() -join " ")
    }
}

function Get-WorksetDeleteActions($Worksets, $Names, [string]$Mode) {
    $target = Get-WorksetDefaultIgnoreName
    $modeArg = $(if ($Mode -eq "elements") { "elements" } else { "move" })
    $actions = New-Object System.Collections.Generic.List[object]
    $wanted = @{}
    foreach ($name in @($Names)) {
        $n = [string]$name
        if ($n) { $wanted[$n.ToLowerInvariant()] = $n }
    }
    if ($null -eq $Worksets) { return $actions.ToArray() }
    foreach ($w in $Worksets) {
        $nm = [string]$w.Name
        if (-not $nm) { continue }
        if (-not $wanted.ContainsKey($nm.ToLowerInvariant())) { continue }
        if ($nm.ToLowerInvariant() -eq $target.ToLowerInvariant()) { continue }
        [void]$actions.Add((New-WorksetAction "delete" ([string]$w.Host) ([string]$w.Id) $nm $modeArg $target "" $false))
    }
    return $actions.ToArray()
}

function Get-WorksetDeleteDryLine([string]$ModelHost, [string]$Name, [string]$Mode, [int]$Count) {
    if ($Mode -eq "elements") {
        $shown = $(if ($Count -lt 0) { "не посчитано" } else { [string]$Count })
        return ("ВМЕСТЕ С ЭЛЕМЕНТАМИ: {0} — удалить набор {1} и {2} элементов" -f $ModelHost, $Name, $shown)
    }
    return ("{0} — перенести элементы набора {1} в 30_ИОС_Элементы" -f $ModelHost, $Name)
}

function Get-WorksetStep4Text {
    param($Actions)
    $ren = New-Object System.Collections.Generic.List[string]
    $make = New-Object System.Collections.Generic.List[string]
    $ins = New-Object System.Collections.Generic.List[string]
    $sk = New-Object System.Collections.Generic.List[string]
    if ($null -ne $Actions) {
        foreach ($a in $Actions) {
            if ($a.Action -eq "rename") {
                [void]$ren.Add(("{0} — {1} — {2}" -f $a.Host, $a.Arg1, $a.Arg2))
            }
            elseif ($a.Action -eq "create") {
                [void]$make.Add(("{0} — создать набор {1}" -f $a.Host, $a.Arg1))
            }
            elseif ($a.Action -eq "insert") {
                [void]$ins.Add(("{0} — вставить связь {1} — в набор {2}" -f $a.Host, $a.Arg1, $a.Arg2))
            }
            else {
                [void]$sk.Add(("{0} — {1} — {2}" -f $a.Host, $a.Arg1, $a.Code))
            }
        }
    }
    $lines = New-Object System.Collections.Generic.List[string]
    [void]$lines.Add("Переименовать:")
    if ($ren.Count -eq 0) { [void]$lines.Add("  (нет)") } else { foreach ($s in $ren) { [void]$lines.Add("  " + $s) } }
    [void]$lines.Add("Создать наборы:")
    if ($make.Count -eq 0) { [void]$lines.Add("  (нет)") } else { foreach ($s in $make) { [void]$lines.Add("  " + $s) } }
    [void]$lines.Add("Вставить:")
    if ($ins.Count -eq 0) { [void]$lines.Add("  (нет)") } else { foreach ($s in $ins) { [void]$lines.Add("  " + $s) } }
    [void]$lines.Add("Пропуск:")
    if ($sk.Count -eq 0) { [void]$lines.Add("  (нет)") } else { foreach ($s in $sk) { [void]$lines.Add("  " + $s) } }
    if ($ins.Count -eq 0) { [void]$lines.Add("Вставка связей — следующий шаг.") }
    return ($lines.ToArray() -join [Environment]::NewLine)
}

function Write-WorksetJobFile {
    param([string]$Path, [string]$Year, [string]$PlanName, [string]$ReportsDir = "")
    $obj = [pscustomobject]@{
        mode      = "apply"
        year      = $Year
        plan      = $PlanName
        saved_utc = [DateTime]::UtcNow.ToString("o")
    }
    # the RBP script writes its CSV / progress into this folder (<tool root>\отчёты)
    if ($ReportsDir) { $obj | Add-Member -NotePropertyName reports_dir -NotePropertyValue $ReportsDir }
    $json = $obj | ConvertTo-Json -Depth 4
    Write-Utf8BomLines $Path @($json)
}

function Get-WorksetDryRunText {
    param($Actions)
    $ren = New-Object System.Collections.Generic.List[string]
    $ins = New-Object System.Collections.Generic.List[string]
    $sk = New-Object System.Collections.Generic.List[string]
    if ($null -ne $Actions) {
        foreach ($a in $Actions) {
            if ($a.Action -eq "rename") {
                [void]$ren.Add(("{0} — {1} → {2}" -f $a.Host, $a.Arg1, $a.Arg2))
            }
            elseif ($a.Action -eq "insert") {
                $extra = ""
                if ($a.CreateWorkset) { $extra = ", создать набор" }
                [void]$ins.Add(("{0} ← {1}, набор {2}{3}" -f $a.Host, $a.Arg1, $a.Arg2, $extra))
            }
            else {
                [void]$sk.Add(("{0} — {1} — {2}" -f $a.Host, $a.Arg1, $a.Code))
            }
        }
    }
    $lines = New-Object System.Collections.Generic.List[string]
    [void]$lines.Add("Переименовать:")
    if ($ren.Count -eq 0) { [void]$lines.Add("  (нет)") } else { foreach ($s in $ren) { [void]$lines.Add("  " + $s) } }
    [void]$lines.Add("Вставить:")
    if ($ins.Count -eq 0) { [void]$lines.Add("  (нет)") } else { foreach ($s in $ins) { [void]$lines.Add("  " + $s) } }
    [void]$lines.Add("Пропуск:")
    if ($sk.Count -eq 0) { [void]$lines.Add("  (нет)") } else { foreach ($s in $sk) { [void]$lines.Add("  " + $s) } }
    return ($lines.ToArray() -join [Environment]::NewLine)
}

function ConvertTo-WorksetCsvField([string]$Val) {
    if ($null -eq $Val) { return "" }
    $s = [string]$Val
    $s = $s.Replace("`r", " ").Replace("`n", " ")
    if ($s.Contains(";") -or $s.Contains('"')) { return '"' + $s.Replace('"', '""') + '"' }
    return $s
}

function Split-WorksetCsvLine([string]$Line) {
    $list = New-Object System.Collections.Generic.List[string]
    if ($null -eq $Line) { return $list.ToArray() }
    $cur = New-Object System.Text.StringBuilder
    $q = $false
    $i = 0
    while ($i -lt $Line.Length) {
        $ch = $Line[$i]
        if ($q) {
            if ($ch -eq '"') {
                if (($i + 1) -lt $Line.Length -and $Line[$i + 1] -eq '"') { [void]$cur.Append('"'); $i += 2; continue }
                $q = $false
                $i++
                continue
            }
            [void]$cur.Append($ch)
            $i++
            continue
        }
        if ($ch -eq '"') { $q = $true; $i++; continue }
        if ($ch -eq ';') { [void]$list.Add($cur.ToString()); $cur = New-Object System.Text.StringBuilder; $i++; continue }
        [void]$cur.Append($ch)
        $i++
    }
    [void]$list.Add($cur.ToString())
    return $list.ToArray()
}

function Write-Utf8BomLines {
    param([string]$Path, [string[]]$Lines)
    $dir = Split-Path -Parent $Path
    if ($dir -and -not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
    $utf8 = New-Object System.Text.UTF8Encoding $true
    [System.IO.File]::WriteAllLines($Path, $Lines, $utf8)
}

function Read-Utf8BomLines([string]$Path) {
    $utf8 = New-Object System.Text.UTF8Encoding $true
    return [System.IO.File]::ReadAllLines($Path, $utf8)
}

# A one-line file comes back from Read-Utf8BomLines as a string.
# [0] on that string is the first character, not the line. Wrap first.
function Get-WorksetPointerTarget([string]$PointerPath) {
    if ([string]::IsNullOrWhiteSpace($PointerPath)) { return "" }
    if (-not (Test-Path -LiteralPath $PointerPath)) { return "" }
    $lines = @(Read-Utf8BomLines $PointerPath)
    if ($lines.Count -lt 1) { return "" }
    $text = [string]$lines[0]
    if ([string]::IsNullOrWhiteSpace($text)) { return "" }
    return $text.Trim().TrimStart([char]0xFEFF)
}

function Get-NewestWorksetReport([string]$Folder, [string]$NamePrefix) {
    if ([string]::IsNullOrWhiteSpace($Folder) -or [string]::IsNullOrWhiteSpace($NamePrefix)) { return "" }
    if (-not (Test-Path -LiteralPath $Folder)) { return "" }
    $best = $null
    foreach ($f in [System.IO.Directory]::GetFiles($Folder, ($NamePrefix + "_*.csv"))) {
        $info = New-Object System.IO.FileInfo $f
        if ($null -eq $best -or $info.LastWriteTimeUtc -gt $best.LastWriteTimeUtc) { $best = $info }
    }
    if ($null -eq $best) { return "" }
    return [string]$best.FullName
}

function Resolve-WorksetReportFile {
    param([string]$PointerPath, [string]$NamePrefix)
    $pointed = Get-WorksetPointerTarget $PointerPath
    if ($pointed -and (Test-Path -LiteralPath $pointed)) {
        return [pscustomobject]@{ Ok = $true; Path = $pointed; Fallback = $false }
    }
    $folder = ""
    if (-not [string]::IsNullOrWhiteSpace($PointerPath)) { $folder = Split-Path -Parent $PointerPath }
    $newest = Get-NewestWorksetReport $folder $NamePrefix
    if ($newest -and (Test-Path -LiteralPath $newest)) {
        return [pscustomobject]@{ Ok = $true; Path = $newest; Fallback = $true }
    }
    return [pscustomobject]@{ Ok = $false; Path = ""; Fallback = $false }
}

function Write-WorksetCsv {
    param([string]$Path, [string[]]$Header, $Rows)
    $lines = New-Object System.Collections.Generic.List[string]
    $head = New-Object System.Collections.Generic.List[string]
    foreach ($h in $Header) { [void]$head.Add((ConvertTo-WorksetCsvField $h)) }
    [void]$lines.Add(($head.ToArray() -join ";"))
    if ($null -ne $Rows) {
        foreach ($r in $Rows) {
            $cells = New-Object System.Collections.Generic.List[string]
            foreach ($h in $Header) {
                $val = ""
                if ($null -ne $r -and $r.PSObject.Properties[$h]) { $val = [string]$r.PSObject.Properties[$h].Value }
                [void]$cells.Add((ConvertTo-WorksetCsvField $val))
            }
            [void]$lines.Add(($cells.ToArray() -join ";"))
        }
    }
    Write-Utf8BomLines $Path $lines.ToArray()
}

function Read-WorksetCsv([string]$Path) {
    $lines = @(Read-Utf8BomLines $Path)
    $header = @()
    $rows = New-Object System.Collections.Generic.List[object]
    $first = $true
    foreach ($line in $lines) {
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        if ($line.TrimStart().StartsWith("#")) { continue }
        $cells = @(Split-WorksetCsvLine $line)
        if ($first) { $header = $cells; $first = $false; continue }
        $obj = [pscustomobject]@{}
        for ($i = 0; $i -lt $header.Count; $i++) {
            $val = ""
            if ($i -lt $cells.Count) { $val = [string]$cells[$i] }
            $obj | Add-Member -NotePropertyName $header[$i] -NotePropertyValue $val
        }
        [void]$rows.Add($obj)
    }
    return [pscustomobject]@{ Header = $header; Count = $rows.Count; Rows = $rows.ToArray() }
}

function Write-SectionCodeFile {
    param([string]$Path, $Codes)
    $rows = New-Object System.Collections.Generic.List[object]
    if ($null -ne $Codes) {
        foreach ($c in $Codes) {
            [void]$rows.Add([pscustomobject]@{ latin = [string]$c.Latin; cyrillic = [string]$c.Cyrillic })
        }
    }
    Write-WorksetCsv $Path @("latin", "cyrillic") $rows.ToArray()
}

function Read-SectionCodeFile([string]$Path) {
    $csv = Read-WorksetCsv $Path
    $list = New-Object System.Collections.Generic.List[object]
    foreach ($r in $csv.Rows) {
        $lat = ""
        $cyr = ""
        if ($r.PSObject.Properties["latin"]) { $lat = [string]$r.latin }
        if ($r.PSObject.Properties["cyrillic"]) { $cyr = [string]$r.cyrillic }
        if (-not $lat -or -not $cyr) { continue }
        [void]$list.Add([pscustomobject]@{ Latin = $lat; Cyrillic = $cyr })
    }
    return [pscustomobject]@{ Count = $list.Count; Rows = $list.ToArray() }
}

# Writes the seed only when the file is absent. An edited table is left as it is.
function Initialize-SectionCodeFile([string]$Path) {
    if (Test-Path -LiteralPath $Path) { return (Read-SectionCodeFile $Path) }
    $seed = @(Get-DefaultSectionCodes)
    Write-SectionCodeFile $Path $seed
    return (Read-SectionCodeFile $Path)
}

function Write-NameMaskFile {
    param([string]$Path, [string]$Separator, [string[]]$Roles)
    $roleText = ""
    if ($null -ne $Roles) { $roleText = ($Roles -join ",") }
    $lines = @(
        "# name mask: separator and a role per token (fixed, section, building, ignore)",
        ("separator={0}" -f $Separator),
        ("roles={0}" -f $roleText)
    )
    Write-Utf8BomLines $Path $lines
}

function Read-NameMaskFile([string]$Path) {
    $sep = "_"
    $roles = @()
    foreach ($line in @(Read-Utf8BomLines $Path)) {
        $t = $line.Trim()
        if (-not $t -or $t.StartsWith("#")) { continue }
        $eq = $t.IndexOf("=")
        if ($eq -lt 1) { continue }
        $k = $t.Substring(0, $eq).Trim().ToLowerInvariant()
        $v = $t.Substring($eq + 1).Trim()
        if ($k -eq "separator") { $sep = $v }
        elseif ($k -eq "roles" -and $v) { $roles = @($v.Split(",")) }
    }
    return [pscustomobject]@{ Separator = $sep; Roles = $roles; RoleCount = $roles.Count }
}

function Write-WorksetPlanFile {
    param([string]$Path, $Actions)
    $rows = New-Object System.Collections.Generic.List[object]
    if ($null -ne $Actions) {
        foreach ($a in $Actions) {
            $flag = "0"
            if ($a.CreateWorkset) { $flag = "1" }
            [void]$rows.Add([pscustomobject]@{
                    action         = [string]$a.Action
                    host           = [string]$a.Host
                    key            = [string]$a.Key
                    arg1           = [string]$a.Arg1
                    arg2           = [string]$a.Arg2
                    arg3           = [string]$a.Arg3
                    code           = [string]$a.Code
                    create_workset = $flag
                })
        }
    }
    Write-WorksetCsv $Path $script:WorksetPlanHeader $rows.ToArray()
}

function Read-WorksetPlanFile([string]$Path) {
    $csv = Read-WorksetCsv $Path
    $list = New-Object System.Collections.Generic.List[object]
    foreach ($r in $csv.Rows) {
        $create = $false
        if ($r.PSObject.Properties["create_workset"] -and [string]$r.create_workset -eq "1") { $create = $true }
        [void]$list.Add((New-WorksetAction ([string]$r.action) ([string]$r.host) ([string]$r.key) ([string]$r.arg1) ([string]$r.arg2) ([string]$r.arg3) ([string]$r.code) $create))
    }
    return [pscustomobject]@{ Count = $list.Count; Rows = $list.ToArray() }
}
