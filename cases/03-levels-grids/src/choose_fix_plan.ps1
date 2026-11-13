param(
    [string]$HostKind = "BF"
)

# WinForms Fix Plan UI. ASCII-only source; Russian via U().
if ([Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
    $hk = if (($HostKind + "").Trim().ToUpperInvariant() -eq "SECTIONS") { "SECTIONS" } else { "BF" }
    $arg = "-NoProfile -ExecutionPolicy Bypass -STA -File `"$PSCommandPath`" -HostKind $hk"
    Start-Process -FilePath "powershell.exe" -ArgumentList $arg -Wait -NoNewWindow
    exit $LASTEXITCODE
}

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$csvDir = Join-Path $env:USERPROFILE "Documents\doc\script\rbp\reports"
if (-not (Test-Path -LiteralPath $csvDir)) {
    New-Item -ItemType Directory -Path $csvDir -Force | Out-Null
}
$fixPlanPath = Join-Path $scriptDir "fix_plan.csv"
$rvtListPath = Join-Path $scriptDir "rvt_list.txt"
$activeDetails = Join-Path $csvDir "_active_levels_grids_audit_details.txt"
$activeReport = Join-Path $csvDir "_active_levels_grids_audit.txt"

function U {
    param([int[]]$Codes)
    return (-join ($Codes | ForEach-Object { [char]([int]$_) }))
}

$H_Model = U @(0x0418, 0x043C, 0x044F, 0x0020, 0x043C, 0x043E, 0x0434, 0x0435, 0x043B, 0x0438)
$H_Category = U @(0x041A, 0x0430, 0x0442, 0x0435, 0x0433, 0x043E, 0x0440, 0x0438, 0x044F)
$H_Problem = U @(0x041F, 0x0440, 0x043E, 0x0431, 0x043B, 0x0435, 0x043C, 0x0430)
$H_Element = U @(0x042D, 0x043B, 0x0435, 0x043C, 0x0435, 0x043D, 0x0442)
$H_ValHost = U @(0x0417, 0x043D, 0x0430, 0x0447, 0x0435, 0x043D, 0x0438, 0x0435, 0x0020, 0x0432, 0x0020, 0x043C, 0x043E, 0x0434, 0x0435, 0x043B, 0x0438)
$H_ValEx = U @(0x0417, 0x043D, 0x0430, 0x0447, 0x0435, 0x043D, 0x0438, 0x0435, 0x0020, 0x0432, 0x0020, 0x044D, 0x0442, 0x0430, 0x043B, 0x043E, 0x043D, 0x0435)
$H_Delta = U @(0x0420, 0x0430, 0x0437, 0x043D, 0x0438, 0x0446, 0x0430)
$H_Status = U @(0x0421, 0x0442, 0x0430, 0x0442, 0x0443, 0x0441)
$H_Code = U @(0x041A, 0x043E, 0x0434)
$H_Path = U @(0x041F, 0x0443, 0x0442, 0x044C, 0x0020, 0x043C, 0x043E, 0x0434, 0x0435, 0x043B, 0x0438)
$CatLevels = U @(0x0423, 0x0440, 0x043E, 0x0432, 0x043D, 0x0438)
$CatGrids = U @(0x041E, 0x0441, 0x0438)

$T_Modify = U @(0x0418, 0x0437, 0x043C, 0x0435, 0x043D, 0x0438, 0x0442, 0x044C)
$T_Rename = U @(0x041F, 0x0435, 0x0440, 0x0435, 0x0438, 0x043C, 0x0435, 0x043D, 0x043E, 0x0432, 0x0430, 0x0442, 0x044C)
$T_Accept = U @(0x041F, 0x0440, 0x0438, 0x043D, 0x044F, 0x0442, 0x044C, 0x0020, 0x0440, 0x0430, 0x0437, 0x043B, 0x0438, 0x0447, 0x0438, 0x0435)
$T_Reject = U @(0x041E, 0x0442, 0x043A, 0x043B, 0x043E, 0x043D, 0x0438, 0x0442, 0x044C)
$T_KeepPos = U @(0x0421, 0x043E, 0x0445, 0x0440, 0x0430, 0x043D, 0x0438, 0x0442, 0x044C, 0x0020, 0x043F, 0x043E, 0x043B, 0x043E, 0x0436, 0x0435, 0x043D, 0x0438, 0x0435)
$T_KeepName = U @(0x0421, 0x043E, 0x0445, 0x0440, 0x0430, 0x043D, 0x0438, 0x0442, 0x044C, 0x0020, 0x0438, 0x043C, 0x044F)
$T_Models = U @(0x041C, 0x043E, 0x0434, 0x0435, 0x043B, 0x0438, 0x0020, 0x0411, 0x0424)
$T_Issues = U @(0x041F, 0x0440, 0x043E, 0x0431, 0x043B, 0x0435, 0x043C, 0x044B, 0x0020, 0x0438, 0x0437, 0x0020, 0x043E, 0x0442, 0x0447, 0x0451, 0x0442, 0x0430)
$T_Title = U @(0x0418, 0x0441, 0x043F, 0x0440, 0x0430, 0x0432, 0x0438, 0x0442, 0x044C, 0x0020, 0x0411, 0x0424, 0x0020, 0x0438, 0x0437, 0x0020, 0x043E, 0x0442, 0x0447, 0x0451, 0x0442, 0x0430)
$T_Ok = "OK"
$T_Cancel = U @(0x041E, 0x0442, 0x043C, 0x0435, 0x043D, 0x0430)
$T_AllModels = U @(0x0412, 0x0441, 0x0435, 0x0020, 0x0411, 0x0424)
$T_NoModels = U @(0x041D, 0x0438, 0x0020, 0x0411, 0x0424)
$T_SyncAfter = U @(0x0421, 0x0438, 0x043D, 0x0445, 0x0440, 0x043E, 0x043D, 0x0438, 0x0437, 0x0438, 0x0440, 0x043E, 0x0432, 0x0430, 0x0442, 0x044C, 0x0020, 0x043F, 0x043E, 0x0441, 0x043B, 0x0435, 0x0020, 0x0438, 0x0437, 0x043C, 0x0435, 0x043D, 0x0435, 0x043D, 0x0438, 0x0439)
$T_Filter = U @(0x0424, 0x0438, 0x043B, 0x044C, 0x0442, 0x0440)
$T_ShowLevels = U @(0x0423, 0x0440, 0x043E, 0x0432, 0x043D, 0x0438)
$T_ShowGrids = U @(0x041E, 0x0441, 0x0438)
$T_OnlyRed = U @(0x0422, 0x043E, 0x043B, 0x044C, 0x043A, 0x043E, 0x0020, 0x0052, 0x0045, 0x0044)
$T_ToChange = U @(0x041A, 0x0020, 0x0438, 0x0437, 0x043C, 0x0435, 0x043D, 0x0435, 0x043D, 0x0438, 0x044E)
$T_PbpGroup = U @(0x0411, 0x0430, 0x0437, 0x043E, 0x0432, 0x0430, 0x044F, 0x0020, 0x0442, 0x043E, 0x0447, 0x043A, 0x0430)
$T_PbpCheck = U @(0x0421, 0x0434, 0x0432, 0x0438, 0x043D, 0x0443, 0x0442, 0x044C, 0x0020, 0x043F, 0x043E, 0x0020, 0x005A)
$T_PbpCheckX = U @(0x0421, 0x0434, 0x0432, 0x0438, 0x043D, 0x0443, 0x0442, 0x044C, 0x0020, 0x043F, 0x043E, 0x0020, 0x0058)
$T_PbpCheckY = U @(0x0421, 0x0434, 0x0432, 0x0438, 0x043D, 0x0443, 0x0442, 0x044C, 0x0020, 0x043F, 0x043E, 0x0020, 0x0059)
$T_PbpLabel = "Z, mm (minus = down):"
$T_PbpLabelX = "X, mm:"
$T_PbpLabelY = "Y, mm:"
$T_PbpFromReport = U @(0x041F, 0x043E, 0x0020, 0x043E, 0x0442, 0x0447, 0x0451, 0x0442, 0x0443)
$T_PbpFromGrids = U @(0x041F, 0x043E, 0x0020, 0x043E, 0x0441, 0x044F, 0x043C)
$T_PbpNone = U @(0x0412, 0x0020, 0x043E, 0x0442, 0x0447, 0x0451, 0x0442, 0x0435, 0x0020, 0x043D, 0x0435, 0x0442, 0x0020, 0x0441, 0x0442, 0x0440, 0x043E, 0x043A, 0x0438, 0x0020, 0x0431, 0x0430, 0x0437, 0x043E, 0x0432, 0x043E, 0x0439, 0x0020, 0x0442, 0x043E, 0x0447, 0x043A, 0x0438)
$CatCoords = U @(0x041A, 0x043E, 0x043E, 0x0440, 0x0434, 0x0438, 0x043D, 0x0430, 0x0442, 0x044B)
$T_PbpElem = U @(0x0411, 0x0430, 0x0437, 0x043E, 0x0432, 0x0430, 0x044F, 0x0020, 0x0442, 0x043E, 0x0447, 0x043A, 0x0430, 0x0020, 0x043F, 0x0440, 0x043E, 0x0435, 0x043A, 0x0442, 0x0430)
$T_PbpMoved = U @(0x0411, 0x0430, 0x0437, 0x043E, 0x0432, 0x0430, 0x044F, 0x0020, 0x0442, 0x043E, 0x0447, 0x043A, 0x0430, 0x0020, 0x0441, 0x043C, 0x0435, 0x0449, 0x0435, 0x043D, 0x0430)
$T_BadZ = "Enter a number for PBP (example: -300)."

if (($HostKind + "").Trim().ToUpperInvariant() -eq "SECTIONS") {
    $HostKind = "SECTIONS"
    $T_Models = U @(0x041C, 0x043E, 0x0434, 0x0435, 0x043B, 0x0438, 0x0020, 0x0440, 0x0430, 0x0437, 0x0434, 0x0435, 0x043B, 0x043E, 0x0432)
    $T_Title = U @(0x0418, 0x0441, 0x043F, 0x0440, 0x0430, 0x0432, 0x0438, 0x0442, 0x044C, 0x0020, 0x0440, 0x0430, 0x0437, 0x0434, 0x0435, 0x043B, 0x044B, 0x0020, 0x0438, 0x0437, 0x0020, 0x043E, 0x0442, 0x0447, 0x0451, 0x0442, 0x0430)
    $T_AllModels = U @(0x0412, 0x0441, 0x0435, 0x0020, 0x0440, 0x0430, 0x0437, 0x0434, 0x0435, 0x043B, 0x044B)
    $T_NoModels = U @(0x041D, 0x0438, 0x0020, 0x043E, 0x0434, 0x043D, 0x043E, 0x0433, 0x043E, 0x0020, 0x0440, 0x0430, 0x0437, 0x0434, 0x0435, 0x043B, 0x0430)
    $NeedAuditMsg = "Details CSV not found. Run menu 3 (Audit sections) first."
    $NeedModelsMsg = "No section models in this report. Run menu 3 first (not menu 1)."
    $NeedOneMsg = "Select at least one section model."
    $NeedPlanMsg = "No visible issues to include in plan. Adjust filters / section selection."
    $RvtListNote = "section models for Apply"
} else {
    $HostKind = "BF"
    $NeedAuditMsg = "Details CSV not found. Run menu 1 (Audit BF) first."
    $NeedModelsMsg = "No BF models in this report. Run menu 1 first (not menu 3)."
    $NeedOneMsg = "Select at least one BF model."
    $NeedPlanMsg = "No visible issues to include in plan. Adjust filters / BF selection."
    $RvtListNote = "BF models for Apply"
}

$ActModify = "modify"
$ActRename = "rename"
$ActAccept = "accept"
$ActReject = "reject"
$ActRu = @{
    modify = $T_Modify
    rename = $T_Rename
    accept = $T_Accept
    reject = $T_Reject
}

function ConvertFrom-SemicolonCsv([string]$Path) {
    $raw = [System.IO.File]::ReadAllText($Path, [System.Text.Encoding]::UTF8)
    $raw = $raw -replace "^\uFEFF", ""
    $field = New-Object System.Text.StringBuilder
    $fields = New-Object System.Collections.Generic.List[string]
    $inQuotes = $false
    $i = 0; $len = $raw.Length
    $script:headersLocal = $null
    $script:rowsLocal = @()
    function Flush-Row {
        if ($fields.Count -eq 0 -and $field.Length -eq 0) { return }
        $fields.Add($field.ToString()); [void]$field.Clear()
        $vals = $fields.ToArray(); $fields.Clear()
        if ($null -eq $script:headersLocal) { $script:headersLocal = $vals; return }
        $obj = [ordered]@{}
        for ($c = 0; $c -lt $script:headersLocal.Count; $c++) {
            $obj[$script:headersLocal[$c]] = if ($c -lt $vals.Count) { $vals[$c] } else { "" }
        }
        $script:rowsLocal += [pscustomobject]$obj
    }
    while ($i -lt $len) {
        $ch = $raw[$i]
        if ($inQuotes) {
            if ($ch -eq '"') {
                if (($i + 1) -lt $len -and $raw[$i + 1] -eq '"') { [void]$field.Append('"'); $i += 2; continue }
                $inQuotes = $false; $i++; continue
            }
            [void]$field.Append($ch); $i++; continue
        }
        if ($ch -eq '"') { $inQuotes = $true; $i++; continue }
        if ($ch -eq ';') { $fields.Add($field.ToString()); [void]$field.Clear(); $i++; continue }
        if ($ch -eq "`r") { $i++; continue }
        if ($ch -eq "`n") { Flush-Row; $i++; continue }
        [void]$field.Append($ch); $i++
    }
    if ($field.Length -gt 0 -or $fields.Count -gt 0) { Flush-Row }
    return @{ Headers = @($script:headersLocal); Rows = @($script:rowsLocal) }
}

function Get-Cell($row, [string]$name) {
    if ($row.PSObject.Properties.Name -contains $name) { return [string]$row.$name }
    return ""
}

function Resolve-DetailsPath {
    if (Test-Path -LiteralPath $activeDetails) {
        $dn = (Get-Content -LiteralPath $activeDetails -TotalCount 1 -ErrorAction SilentlyContinue)
        if ($dn) {
            $p = Join-Path $csvDir $dn.Trim()
            if (Test-Path -LiteralPath $p) { return $p }
        }
    }
    if (Test-Path -LiteralPath $activeReport) {
        $sn = (Get-Content -LiteralPath $activeReport -TotalCount 1 -ErrorAction SilentlyContinue)
        if ($sn) {
            $p = Join-Path $csvDir ($sn.Trim() -replace '\.csv$', '_details.csv')
            if (Test-Path -LiteralPath $p) { return $p }
        }
    }
    $latest = Get-ChildItem -LiteralPath $csvDir -Filter "levels_grids_audit_*_details.csv" -File -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if (-not $latest) {
        $latest = Get-ChildItem -LiteralPath $csvDir -Filter "lg_*_details.csv" -File -ErrorAction SilentlyContinue |
            Sort-Object LastWriteTime -Descending | Select-Object -First 1
    }
    if ($latest) { return $latest.FullName }
    return $null
}

function Get-AxisNumber([string]$Text, [string]$Axis) {
    if ([string]::IsNullOrWhiteSpace($Text)) { return $null }
    $t = $Text.Trim()
    $pat = '(?i)\b' + [regex]::Escape($Axis) + '\s*=\s*(-?[0-9]+(?:[.,][0-9]+)?)'
    if ($t -match $pat) {
        return [double](($Matches[1] -replace ',', '.'))
    }
    return $null
}

function Get-ZNumber([string]$Text) {
    $z = Get-AxisNumber $Text "Z"
    if ($null -ne $z) { return $z }
    if ([string]::IsNullOrWhiteSpace($Text)) { return $null }
    $plain = ($Text.Trim() -replace ',', '.').Trim()
    if ($plain -match '(?i)[XY]\s*=') { return $null }
    $n = 0.0
    if ([double]::TryParse($plain, [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$n)) {
        return $n
    }
    return $null
}

function Format-PbpNum($n) {
    if ($null -eq $n) { return "?" }
    return ([double]$n).ToString("0.###", [Globalization.CultureInfo]::InvariantCulture)
}

function Test-PbpShiftOn($n) {
    if ($null -eq $n) { return $false }
    return ([math]::Abs([double]$n) -ge 0.001)
}

function Format-PbpHintText($info) {
    if ($null -eq $info) { return $T_PbpNone }
    $src = $T_PbpFromReport
    $bits = New-Object System.Collections.Generic.List[string]
    foreach ($ax in @("X", "Y", "Z")) {
        $h = $info."Host$ax"
        $e = $info."Ex$ax"
        $s = $info."Shift$ax"
        if ($null -eq $h -and $null -eq $e -and $null -eq $s) { continue }
        [void]$bits.Add(("{0}: {1}={2} -> {3}  d={4}" -f $ax, $H_ValHost, (Format-PbpNum $h), (Format-PbpNum $e), (Format-PbpNum $s)))
    }
    if ($bits.Count -eq 0) { return $T_PbpNone }
    return ("{0}:  {1}" -f $src, ($bits -join "    "))
}

function Get-ProblemKind([string]$code, [string]$problemRu) {
    $c = ($code + "").Trim().ToLowerInvariant()
    if ($c -eq "name_mismatch" -or $c -eq "cr_name_changed") { return "name" }
    if ($c -in @(
        "elevation_delta", "position_delta", "angle_delta",
        "grid_relative_moved", "cr_grid_moved", "cr_level_moved"
    )) { return "position" }
    if ($c -in @(
        "no_monitoring", "missing_in_host", "leftover_exemplar", "monitor_broken",
        "duplicate_host", "duplicate_exemplar",
        "pbp_delta", "pbp_missing", "survey_delta", "survey_missing",
        "cr_watched_link_moved", "cr_link_needs_review", "cr_alert"
    )) {
        return "other"
    }
    $pName = U @(0x0438, 0x043C, 0x044F)
    $pMove = U @(0x043F, 0x0435, 0x0440, 0x0435, 0x043C, 0x0435, 0x0449)
    if ($problemRu -and $problemRu.ToLowerInvariant().Contains($pName)) { return "name" }
    if ($problemRu -and $problemRu.ToLowerInvariant().Contains($pMove)) { return "position" }
    return "other"
}

function Test-LooksLikeBf([string]$Name) {
    $b = ($Name + "").ToLowerInvariant()
    if ($b.Contains("_bf_")) { return $true }
    if ($b.StartsWith("bf_")) { return $true }
    if ($b.Contains("_bf.")) { return $true }
    if ($b.Contains("/bf_")) { return $true }
    if ($b.Contains("\bf_")) { return $true }
    return $false
}

function Test-KeepHostModel([string]$Name) {
    $isBf = Test-LooksLikeBf $Name
    if ($HostKind -eq "SECTIONS") { return -not $isBf }
    return $isBf
}

function Get-CategoryKind([string]$cat) {
    if ($cat -eq $CatLevels) { return "Level" }
    if ($cat -eq $CatGrids) { return "Grid" }
    $c = ($cat + "").ToLowerInvariant()
    if ($c.StartsWith("level")) { return "Level" }
    if ($c.StartsWith("grid")) { return "Grid" }
    if ($c.StartsWith((U @(0x043E, 0x0441)))) { return "Grid" }
    return "Other"
}

function Escape-CsvField([string]$Text) {
    if ($null -eq $Text) { $Text = "" }
    if ($Text.Contains(';') -or $Text.Contains('"') -or $Text.Contains("`n") -or $Text.Contains("`r")) {
        return '"' + ($Text.Replace('"', '""')) + '"'
    }
    return $Text
}

function Write-FixPlan([string]$Path, $Items, [bool]$SyncAfter) {
    $headersRu = @(
        $H_Model, $H_Path, $H_Category, $H_Element,
        (U @(0x0418, 0x043C, 0x044F, 0x0020, 0x0432, 0x0020, 0x044D, 0x0442, 0x0430, 0x043B, 0x043E, 0x043D, 0x0435)),
        $H_Code, $H_Problem,
        (U @(0x041F, 0x0440, 0x0430, 0x0432, 0x0438, 0x043B, 0x043E)),
        "action",
        (U @(0x0414, 0x0435, 0x0439, 0x0441, 0x0442, 0x0432, 0x0438, 0x0435)),
        $H_ValHost, $H_ValEx, $H_Delta, $H_Status,
        "sync_after_change"
    )
    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add(($headersRu | ForEach-Object { Escape-CsvField $_ }) -join ";")
    foreach ($it in $Items) {
        $syncFlag = "0"
        if ($SyncAfter) { $syncFlag = "1" }
        $vals = @(
            $it.model_name, $it.model_path, $it.category, $it.element, $it.name_exemplar,
            $it.problem_code, $it.problem, $it.rule, $it.action, $it.action_ru,
            $it.value_host, $it.value_exemplar, $it.delta, $it.status,
            $syncFlag
        )
        $lines.Add(($vals | ForEach-Object { Escape-CsvField ([string]$_) }) -join ";")
    }
    $enc = New-Object System.Text.UTF8Encoding($true)
    [System.IO.File]::WriteAllText($Path, (($lines -join "`r`n") + "`r`n"), $enc)
}

function Write-RvtList([string]$Path, [string[]]$ModelPaths) {
    $uniq = @()
    foreach ($p in $ModelPaths) {
        if ([string]::IsNullOrWhiteSpace($p)) { continue }
        if ($p -match '(?i)^RSN://x(/|$)') { continue }
        if ($uniq -notcontains $p) { $uniq += $p }
    }
    $body = @(
        "# Generated by choose_fix_plan.ps1 - $RvtListNote",
        "# $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    ) + $uniq
    $enc = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($Path, (($body -join "`r`n") + "`r`n"), $enc)
}

# ---- load data ----
$detailsPath = Resolve-DetailsPath
if (-not $detailsPath) {
    [System.Windows.Forms.MessageBox]::Show(
        $NeedAuditMsg,
        $T_Title, "OK", "Error") | Out-Null
    exit 1
}

$data = ConvertFrom-SemicolonCsv -Path $detailsPath
$issues = New-Object System.Collections.Generic.List[object]
$modelOrder = New-Object System.Collections.Generic.List[string]
$modelPaths = @{}
$pbpHint = @{}

foreach ($r in @($data.Rows)) {
    $code = Get-Cell $r $H_Code
    $problem = Get-Cell $r $H_Problem
    $cat = Get-Cell $r $H_Category
    $el = Get-Cell $r $H_Element
    $model = Get-Cell $r $H_Model
    if ([string]::IsNullOrWhiteSpace($model)) { continue }
    if ($problem -match '^\d+$' -and [string]::IsNullOrWhiteSpace($code) -and [string]::IsNullOrWhiteSpace($el)) { continue }

    if (-not $modelPaths.ContainsKey($model)) {
        [void]$modelOrder.Add($model)
        $modelPaths[$model] = (Get-Cell $r $H_Path)
    } elseif ([string]::IsNullOrWhiteSpace($modelPaths[$model])) {
        $modelPaths[$model] = (Get-Cell $r $H_Path)
    }

    $kind = Get-ProblemKind $code $problem
    $catKind = Get-CategoryKind $cat
    $codeL = ($code + "").Trim().ToLowerInvariant()
    if ($kind -eq "position" -and $catKind -eq "Other" -and (
            $codeL -in @("grid_relative_moved", "cr_grid_moved")
        )) {
        $catKind = "Grid"
    }
    if ($kind -eq "position" -and $catKind -eq "Other" -and $codeL -eq "cr_level_moved") {
        $catKind = "Level"
    }

    $elLow = ($el + "").ToLowerInvariant()
    $isPbp = (
        $codeL -eq "pbp_delta" -or $codeL -eq "pbp_missing" -or
        (
            ($cat -eq $CatCoords -or $cat -eq "Coords") -and
            $codeL -notmatch 'survey' -and
            ($elLow.Contains("base") -or $elLow.Contains((U @(0x0431, 0x0430, 0x0437, 0x043E, 0x0432))))
        )
    )
    if ($isPbp) {
        $vh = Get-Cell $r $H_ValHost
        $ve = Get-Cell $r $H_ValEx
        $vd = Get-Cell $r $H_Delta
        $hint = [pscustomobject]@{
            HostX = (Get-AxisNumber $vh "X"); ExX = (Get-AxisNumber $ve "X"); ShiftX = $null
            HostY = (Get-AxisNumber $vh "Y"); ExY = (Get-AxisNumber $ve "Y"); ShiftY = $null
            HostZ = (Get-ZNumber $vh);       ExZ = (Get-ZNumber $ve);       ShiftZ = $null
        }
        foreach ($ax in @("X", "Y", "Z")) {
            $h = $hint."Host$ax"
            $e = $hint."Ex$ax"
            if ($null -ne $h -and $null -ne $e) {
                $hint."Shift$ax" = [double]($e - $h)
            }
        }
        if ($null -eq $hint.ShiftZ) {
            $zd = Get-ZNumber $vd
            if ($null -ne $zd) { $hint.ShiftZ = -$zd }
        }
        $pbpHint[$model] = $hint
    }

    if ($kind -eq "other" -or $catKind -eq "Other") { continue }
    if ([string]::IsNullOrWhiteSpace($el)) { $el = Get-Cell $r $H_ValHost }
    $doLabel = if ($kind -eq "name") { $T_Rename } else { $T_Modify }
    $doAction = if ($kind -eq "name") { $ActRename } else { $ActModify }
    [void]$issues.Add([pscustomobject]@{
        model_name     = $model
        model_path     = (Get-Cell $r $H_Path)
        category       = $cat
        category_kind  = $catKind
        element        = $el
        name_exemplar  = (Get-Cell $r $H_ValEx)
        problem_code   = $code
        problem        = $problem
        kind           = $kind
        value_host     = (Get-Cell $r $H_ValHost)
        value_exemplar = (Get-Cell $r $H_ValEx)
        delta          = (Get-Cell $r $H_Delta)
        status         = (Get-Cell $r $H_Status)
        do_label       = $doLabel
        do_action      = $doAction
        action         = $doAction
    })
}

$keptModels = New-Object System.Collections.Generic.List[string]
foreach ($m in $modelOrder) {
    if (Test-KeepHostModel $m) { [void]$keptModels.Add($m) }
}
if ($keptModels.Count -eq 0) {
    [System.Windows.Forms.MessageBox]::Show(
        $NeedModelsMsg,
        $T_Title, "OK", "Error") | Out-Null
    exit 1
}
$modelOrder = $keptModels
$keepSet = @{}
foreach ($m in $modelOrder) { $keepSet[$m] = $true }
$keptIssues = New-Object System.Collections.Generic.List[object]
foreach ($it in $issues) {
    if ($keepSet.ContainsKey($it.model_name)) { [void]$keptIssues.Add($it) }
}
$issues = $keptIssues
$pbpKept = @{}
foreach ($k in @($pbpHint.Keys)) {
    if ($keepSet.ContainsKey($k)) { $pbpKept[$k] = $pbpHint[$k] }
}
$pbpHint = $pbpKept

# Sort: Levels first, then Grids; RED before others; then element name
$sorted = @($issues | Sort-Object `
    @{ Expression = { if ($_.category_kind -eq "Level") { 0 } else { 1 } } }, `
    @{ Expression = { if ((($_.status + "").ToUpperInvariant()) -eq "RED") { 0 } else { 1 } } }, `
    @{ Expression = { $_.element } })
$issues.Clear()
foreach ($s in $sorted) { [void]$issues.Add($s) }

# ---- exclusive checkbox helper ----
function Set-ExclusiveAction {
    param($CbDo, $CbAccept, $CbReject, $Which)
    $script:suppress = $true
    try {
        if ($Which -eq "do") {
            $CbDo.Checked = $true; $CbAccept.Checked = $false; $CbReject.Checked = $false
        } elseif ($Which -eq "accept") {
            $CbDo.Checked = $false; $CbAccept.Checked = $true; $CbReject.Checked = $false
        } elseif ($Which -eq "reject") {
            $CbDo.Checked = $false; $CbAccept.Checked = $false; $CbReject.Checked = $true
        }
    } finally { $script:suppress = $false }
}

function New-ActionTriple {
    param([string]$DoText, [string]$Default = "do")
    $p = New-Object System.Windows.Forms.FlowLayoutPanel
    $p.AutoSize = $true
    $p.WrapContents = $false
    $p.Margin = New-Object System.Windows.Forms.Padding(0)
    $cbDo = New-Object System.Windows.Forms.CheckBox
    $cbDo.Text = $DoText
    $cbDo.AutoSize = $true
    $cbAcc = New-Object System.Windows.Forms.CheckBox
    $cbAcc.Text = $T_Accept
    $cbAcc.AutoSize = $true
    $cbRej = New-Object System.Windows.Forms.CheckBox
    $cbRej.Text = $T_Reject
    $cbRej.AutoSize = $true
    $cbDo.Add_CheckedChanged({
        if ($script:suppress) { return }
        if ($cbDo.Checked) { Set-ExclusiveAction $cbDo $cbAcc $cbRej "do" }
        elseif (-not $cbAcc.Checked -and -not $cbRej.Checked) { Set-ExclusiveAction $cbDo $cbAcc $cbRej "reject" }
    }.GetNewClosure())
    $cbAcc.Add_CheckedChanged({
        if ($script:suppress) { return }
        if ($cbAcc.Checked) { Set-ExclusiveAction $cbDo $cbAcc $cbRej "accept" }
        elseif (-not $cbDo.Checked -and -not $cbRej.Checked) { Set-ExclusiveAction $cbDo $cbAcc $cbRej "reject" }
    }.GetNewClosure())
    $cbRej.Add_CheckedChanged({
        if ($script:suppress) { return }
        if ($cbRej.Checked) { Set-ExclusiveAction $cbDo $cbAcc $cbRej "reject" }
        elseif (-not $cbDo.Checked -and -not $cbAcc.Checked) { Set-ExclusiveAction $cbDo $cbAcc $cbRej "reject" }
    }.GetNewClosure())
    Set-ExclusiveAction $cbDo $cbAcc $cbRej $Default
    [void]$p.Controls.Add($cbDo)
    [void]$p.Controls.Add($cbAcc)
    [void]$p.Controls.Add($cbRej)
    return [pscustomobject]@{ Panel = $p; Do = $cbDo; Accept = $cbAcc; Reject = $cbRej }
}

function Get-TripleAction($triple, [string]$doAction) {
    if ($triple.Do.Checked) { return $doAction }
    if ($triple.Accept.Checked) { return $ActAccept }
    return $ActReject
}

function Set-TripleAction($triple, [string]$action, [string]$doAction) {
    if ($action -eq $doAction -or $action -eq $ActModify -or $action -eq $ActRename) {
        if ($action -eq $ActModify -or $action -eq $ActRename) {
            Set-ExclusiveAction $triple.Do $triple.Accept $triple.Reject "do"
        } else {
            Set-ExclusiveAction $triple.Do $triple.Accept $triple.Reject "do"
        }
    } elseif ($action -eq $ActAccept) {
        Set-ExclusiveAction $triple.Do $triple.Accept $triple.Reject "accept"
    } else {
        Set-ExclusiveAction $triple.Do $triple.Accept $triple.Reject "reject"
    }
}

# ---- build form ----
$form = New-Object System.Windows.Forms.Form
$form.Text = $T_Title
$form.StartPosition = "CenterScreen"
$form.Size = New-Object System.Drawing.Size(1200, 760)
$form.MinimumSize = New-Object System.Drawing.Size(980, 600)
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

$root = New-Object System.Windows.Forms.TableLayoutPanel
$root.Dock = "Fill"
$root.ColumnCount = 1
$root.RowCount = 4
[void]$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
[void]$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 48)))
[void]$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 52)))
[void]$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$form.Controls.Add($root)

# info
$lblInfo = New-Object System.Windows.Forms.Label
$lblInfo.AutoSize = $true
$lblInfo.Text = "Details: $detailsPath"
$lblInfo.Margin = New-Object System.Windows.Forms.Padding(8, 8, 8, 4)
[void]$root.Controls.Add($lblInfo, 0, 0)

# models group (scroll — do not AutoSize the group)
$grpModels = New-Object System.Windows.Forms.GroupBox
$grpModels.Text = "$T_Models  /  $T_PbpGroup"
$grpModels.Dock = "Fill"
$grpModels.Padding = New-Object System.Windows.Forms.Padding(8)
$grpModels.Margin = New-Object System.Windows.Forms.Padding(8, 4, 8, 4)
$modelsLayout = New-Object System.Windows.Forms.TableLayoutPanel
$modelsLayout.Dock = "Fill"
$modelsLayout.ColumnCount = 1
$modelsLayout.RowCount = 2
[void]$modelsLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
[void]$modelsLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$grpModels.Controls.Add($modelsLayout)

$scrollModels = New-Object System.Windows.Forms.Panel
$scrollModels.Dock = "Fill"
$scrollModels.AutoScroll = $true
$flModels = New-Object System.Windows.Forms.FlowLayoutPanel
$flModels.Location = New-Object System.Drawing.Point(0, 0)
$flModels.AutoSize = $true
$flModels.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
$flModels.FlowDirection = "TopDown"
$flModels.WrapContents = $false
$scrollModels.Controls.Add($flModels)
[void]$modelsLayout.Controls.Add($scrollModels, 0, 0)
$modelChecks = @{}
$modelPbp = @{}

function New-PbpAxis($textCb, $textLbl, $shiftVal) {
    $cb = New-Object System.Windows.Forms.CheckBox
    $cb.Text = $textCb
    $cb.AutoSize = $true
    $cb.Checked = (Test-PbpShiftOn $shiftVal)
    $cb.Margin = New-Object System.Windows.Forms.Padding(3, 8, 6, 3)
    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = $textLbl
    $lbl.AutoSize = $true
    $lbl.Margin = New-Object System.Windows.Forms.Padding(3, 10, 4, 3)
    $tb = New-Object System.Windows.Forms.TextBox
    $tb.Width = 70
    $tb.Margin = New-Object System.Windows.Forms.Padding(3, 6, 10, 3)
    if ($null -ne $shiftVal) { $tb.Text = (Format-PbpNum $shiftVal) }
    return @{ Cb = $cb; Lbl = $lbl; Tb = $tb }
}

foreach ($m in $modelOrder) {
    $rowFl = New-Object System.Windows.Forms.FlowLayoutPanel
    $rowFl.AutoSize = $true
    $rowFl.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
    $rowFl.FlowDirection = "LeftToRight"
    $rowFl.WrapContents = $true
    $rowFl.Margin = New-Object System.Windows.Forms.Padding(0, 2, 0, 2)

    $cb = New-Object System.Windows.Forms.CheckBox
    $cnt = @($issues | Where-Object { $_.model_name -eq $m }).Count
    $cb.Text = "$m  ($cnt)"
    $cb.AutoSize = $true
    $cb.Checked = $true
    $cb.Margin = New-Object System.Windows.Forms.Padding(3, 8, 16, 3)
    $modelChecks[$m] = $cb

    $info = $null
    if ($pbpHint.ContainsKey($m)) { $info = $pbpHint[$m] }

    $axX = New-PbpAxis $T_PbpCheckX $T_PbpLabelX $(if ($info) { $info.ShiftX } else { $null })
    $axY = New-PbpAxis $T_PbpCheckY $T_PbpLabelY $(if ($info) { $info.ShiftY } else { $null })
    $axZ = New-PbpAxis $T_PbpCheck $T_PbpLabel $(if ($info) { $info.ShiftZ } else { $null })

    $lblPbpHint = New-Object System.Windows.Forms.Label
    $lblPbpHint.AutoSize = $true
    $lblPbpHint.Margin = New-Object System.Windows.Forms.Padding(3, 10, 3, 3)
    $lblPbpHint.Text = (Format-PbpHintText $info)

    [void]$rowFl.Controls.Add($cb)
    foreach ($ax in @($axX, $axY, $axZ)) {
        [void]$rowFl.Controls.Add($ax.Cb)
        [void]$rowFl.Controls.Add($ax.Lbl)
        [void]$rowFl.Controls.Add($ax.Tb)
    }
    [void]$rowFl.Controls.Add($lblPbpHint)
    [void]$flModels.Controls.Add($rowFl)
    $modelPbp[$m] = @{
        Row  = $rowFl
        Hint = $lblPbpHint
        X    = $axX
        Y    = $axY
        Z    = $axZ
    }
}
$flModelBtns = New-Object System.Windows.Forms.FlowLayoutPanel
$flModelBtns.AutoSize = $true
$btnAllM = New-Object System.Windows.Forms.Button
$btnAllM.Text = $T_AllModels
$btnAllM.AutoSize = $true
$btnAllM.Add_Click({ foreach ($c in $modelChecks.Values) { $c.Checked = $true }; Update-IssueVisibility })
$btnNoM = New-Object System.Windows.Forms.Button
$btnNoM.Text = $T_NoModels
$btnNoM.AutoSize = $true
$btnNoM.Add_Click({ foreach ($c in $modelChecks.Values) { $c.Checked = $false }; Update-IssueVisibility })
[void]$flModelBtns.Controls.Add($btnAllM)
[void]$flModelBtns.Controls.Add($btnNoM)
[void]$modelsLayout.Controls.Add($flModelBtns, 0, 1)
[void]$root.Controls.Add($grpModels, 0, 1)

# issues group (scroll)
$grpIss = New-Object System.Windows.Forms.GroupBox
$grpIss.Text = $T_Issues
$grpIss.Dock = "Fill"
$grpIss.Padding = New-Object System.Windows.Forms.Padding(8)
$grpIss.Margin = New-Object System.Windows.Forms.Padding(8, 4, 8, 4)
$issLayout = New-Object System.Windows.Forms.TableLayoutPanel
$issLayout.Dock = "Fill"
$issLayout.ColumnCount = 1
$issLayout.RowCount = 2
[void]$issLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
[void]$issLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$grpIss.Controls.Add($issLayout)

$flFilter = New-Object System.Windows.Forms.FlowLayoutPanel
$flFilter.AutoSize = $true
$flFilter.Dock = "Top"
$flFilter.WrapContents = $true
$flFilter.Padding = New-Object System.Windows.Forms.Padding(0, 0, 0, 4)
$lblFilter = New-Object System.Windows.Forms.Label
$lblFilter.Text = ($T_Filter + ":")
$lblFilter.AutoSize = $true
$lblFilter.Margin = New-Object System.Windows.Forms.Padding(3, 8, 8, 3)
$cbShowLevels = New-Object System.Windows.Forms.CheckBox
$cbShowLevels.Text = $T_ShowLevels
$cbShowLevels.Checked = $true
$cbShowLevels.AutoSize = $true
$cbShowGrids = New-Object System.Windows.Forms.CheckBox
$cbShowGrids.Text = $T_ShowGrids
$cbShowGrids.Checked = $true
$cbShowGrids.AutoSize = $true
$cbOnlyRed = New-Object System.Windows.Forms.CheckBox
$cbOnlyRed.Text = $T_OnlyRed
$cbOnlyRed.Checked = $false
$cbOnlyRed.AutoSize = $true
$lblVisibleCount = New-Object System.Windows.Forms.Label
$lblVisibleCount.AutoSize = $true
$lblVisibleCount.Margin = New-Object System.Windows.Forms.Padding(16, 8, 3, 3)
[void]$flFilter.Controls.Add($lblFilter)
[void]$flFilter.Controls.Add($cbShowLevels)
[void]$flFilter.Controls.Add($cbShowGrids)
[void]$flFilter.Controls.Add($cbOnlyRed)
[void]$flFilter.Controls.Add($lblVisibleCount)
[void]$issLayout.Controls.Add($flFilter, 0, 0)

$scroll = New-Object System.Windows.Forms.Panel
$scroll.Dock = "Fill"
$scroll.AutoScroll = $true
$flIssues = New-Object System.Windows.Forms.FlowLayoutPanel
# Do not Dock=Top inside AutoScroll: docked children are ignored by AutoScrollMinSize
# and the last rows are clipped (and can drop out of fix_plan via Panel.Visible).
$flIssues.Location = New-Object System.Drawing.Point(0, 0)
$flIssues.AutoSize = $true
$flIssues.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
$flIssues.FlowDirection = "TopDown"
$flIssues.WrapContents = $false
$scroll.Controls.Add($flIssues)
[void]$issLayout.Controls.Add($scroll, 0, 1)
[void]$root.Controls.Add($grpIss, 0, 2)

$tip = New-Object System.Windows.Forms.ToolTip
$rowH = [Math]::Max(36, $form.Font.Height + 20)
$issueRows = New-Object System.Collections.Generic.List[object]
foreach ($it in $issues) {
    $rowPanel = New-Object System.Windows.Forms.Panel
    $rowPanel.Width = 1020
    $rowPanel.Height = $rowH
    $rowPanel.Margin = New-Object System.Windows.Forms.Padding(2, 3, 2, 3)

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.AutoEllipsis = $false
    $lbl.AutoSize = $true
    $lbl.MaximumSize = New-Object System.Drawing.Size(690, 0)
    $lbl.Location = New-Object System.Drawing.Point(0, 1)
    $lbl.TextAlign = [System.Drawing.ContentAlignment]::TopLeft
    $catRu = if ($it.category_kind -eq "Level") { $T_ShowLevels } else { $T_ShowGrids }
    $lbl.Text = ("{0} | {1} | {2} | {3}: {4} | {5}: {6} | {7}: {8}" -f `
        $it.model_name, $catRu, $it.problem, `
        $H_Element, $it.element, $H_ValHost, $it.value_host, $H_ValEx, $it.value_exemplar)
    $tip.SetToolTip($lbl, $lbl.Text)

    $triple = New-ActionTriple -DoText $it.do_label -Default "do"
    $triple.Panel.Location = New-Object System.Drawing.Point(700, [Math]::Max(0, [int](($rowH - 26) / 2)))

    $rowPanel.Controls.Add($lbl)
    $rowPanel.Controls.Add($triple.Panel)
    [void]$flIssues.Controls.Add($rowPanel)

    $rec = [pscustomobject]@{
        Issue    = $it
        Panel    = $rowPanel
        Lbl      = $lbl
        Triple   = $triple
        Included = $true
    }
    [void]$issueRows.Add($rec)
}

function Update-ModelsLayout {
    $w = $scrollModels.ClientSize.Width
    if ($w -lt 80) { return }
    $inner = [Math]::Max(240, $w - 22)
    $flModels.Width = $inner
    foreach ($m in @($modelPbp.Keys)) {
        $row = $modelPbp[$m]
        $row.Row.MaximumSize = New-Object System.Drawing.Size($inner, 0)
        $row.Row.Width = $inner
        $row.Hint.MaximumSize = New-Object System.Drawing.Size([Math]::Max(160, $inner - 24), 0)
    }
    $prefH = $flModels.PreferredSize.Height
    $scrollModels.AutoScrollMinSize = New-Object System.Drawing.Size(0, $prefH)
}

function Update-IssuesLayout {
    $w = $scroll.ClientSize.Width
    if ($w -lt 80) { return }
    $inner = [Math]::Max(240, $w - 22)
    $flIssues.Width = $inner
    $actionsW = 330
    foreach ($rec in $issueRows) {
        $rec.Panel.Width = [Math]::Max(220, $inner - 8)
        $lblW = [Math]::Max(120, $rec.Panel.Width - $actionsW)
        $rec.Lbl.MaximumSize = New-Object System.Drawing.Size($lblW, 0)
        $rec.Lbl.AutoSize = $true
        $hLbl = [Math]::Max($rowH - 2, $rec.Lbl.PreferredSize.Height)
        $hAct = [Math]::Max(26, $rec.Triple.Panel.PreferredSize.Height)
        $h = [Math]::Max($rowH, $hLbl + 6)
        $h = [Math]::Max($h, $hAct + 6)
        $rec.Panel.Height = $h
        $rec.Lbl.Location = New-Object System.Drawing.Point(0, 2)
        $yAct = [Math]::Max(0, [int](($h - $hAct) / 2))
        $rec.Triple.Panel.Location = New-Object System.Drawing.Point(($lblW + 6), $yAct)
    }
    $prefH = $flIssues.PreferredSize.Height
    $scroll.AutoScrollMinSize = New-Object System.Drawing.Size(0, $prefH)
}

function Test-IssueIncluded($it) {
    $m = $it.model_name
    $on = $modelChecks.ContainsKey($m) -and $modelChecks[$m].Checked
    if ($on -and $it.category_kind -eq "Level" -and -not $cbShowLevels.Checked) { $on = $false }
    if ($on -and $it.category_kind -eq "Grid" -and -not $cbShowGrids.Checked) { $on = $false }
    if ($on -and $cbOnlyRed.Checked) {
        $st = ($it.status + "").ToUpperInvariant()
        if ($st -ne "RED") { $on = $false }
    }
    return $on
}

function Update-SelectedCount {
    $nVis = 0
    $nDo = 0
    foreach ($rec in $issueRows) {
        if (-not $rec.Included) { continue }
        $nVis++
        if ($rec.Triple.Do.Checked) { $nDo++ }
    }
    $lblVisibleCount.Text = ("{0}: {1}    {2}: {3}" -f $T_Issues, $nVis, $T_ToChange, $nDo)
}

function Update-PbpRowState {
    foreach ($m in @($modelPbp.Keys)) {
        $on = $modelChecks.ContainsKey($m) -and $modelChecks[$m].Checked
        $row = $modelPbp[$m]
        $row.Hint.Enabled = $on
        foreach ($ax in @("X", "Y", "Z")) {
            $row[$ax].Cb.Enabled = $on
            $row[$ax].Tb.Enabled = $on
            $row[$ax].Lbl.Enabled = $on
        }
    }
}

function Update-IssueVisibility {
    foreach ($rec in $issueRows) {
        $on = Test-IssueIncluded $rec.Issue
        $rec.Included = $on
        $rec.Panel.Visible = $on
    }
    Update-SelectedCount
    Update-ModelsLayout
    Update-IssuesLayout
    Update-PbpRowState
}
foreach ($cb in $modelChecks.Values) {
    $cb.Add_CheckedChanged({ Update-IssueVisibility })
}
$cbShowLevels.Add_CheckedChanged({ Update-IssueVisibility })
$cbShowGrids.Add_CheckedChanged({ Update-IssueVisibility })
$cbOnlyRed.Add_CheckedChanged({ Update-IssueVisibility })
foreach ($rec in $issueRows) {
    $rec.Triple.Do.Add_CheckedChanged({ Update-SelectedCount })
    $rec.Triple.Accept.Add_CheckedChanged({ Update-SelectedCount })
    $rec.Triple.Reject.Add_CheckedChanged({ Update-SelectedCount })
}
$form.Add_Shown({ Update-IssueVisibility })
$form.Add_Resize({ Update-ModelsLayout; Update-IssuesLayout })

# bottom buttons
$flBottom = New-Object System.Windows.Forms.FlowLayoutPanel
$flBottom.Dock = "Fill"
$flBottom.FlowDirection = "RightToLeft"
$flBottom.AutoSize = $true
$flBottom.Padding = New-Object System.Windows.Forms.Padding(8)
$btnOk = New-Object System.Windows.Forms.Button
$btnOk.Text = $T_Ok
$btnOk.AutoSize = $true
$btnOk.DialogResult = [System.Windows.Forms.DialogResult]::None
$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Text = $T_Cancel
$btnCancel.AutoSize = $true
$btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
[void]$flBottom.Controls.Add($btnOk)
[void]$flBottom.Controls.Add($btnCancel)
[void]$root.Controls.Add($flBottom, 0, 3)
$form.CancelButton = $btnCancel

$cbSync = New-Object System.Windows.Forms.CheckBox
$cbSync.Text = $T_SyncAfter
$cbSync.Checked = $true
$cbSync.AutoSize = $true
$cbSync.Margin = New-Object System.Windows.Forms.Padding(6, 8, 18, 3)
[void]$flBottom.Controls.Add($cbSync)

$script:dialogOk = $false
$script:pbpByModel = @{}
$btnOk.Add_Click({
    $anyModel = $false
    foreach ($c in $modelChecks.Values) { if ($c.Checked) { $anyModel = $true; break } }
    if (-not $anyModel) {
        [System.Windows.Forms.MessageBox]::Show($NeedOneMsg, $T_Title, "OK", "Warning") | Out-Null
        return
    }
    $script:pbpByModel = @{}
    foreach ($m in @($modelChecks.Keys)) {
        if (-not $modelChecks[$m].Checked) { continue }
        if (-not $modelPbp.ContainsKey($m)) { continue }
        $row = $modelPbp[$m]
        $shift = @{}
        foreach ($ax in @("X", "Y", "Z")) {
            if (-not $row[$ax].Cb.Checked) { continue }
            $raw = ($row[$ax].Tb.Text + "").Trim() -replace ',', '.'
            $n = 0.0
            if (-not [double]::TryParse($raw, [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$n)) {
                [System.Windows.Forms.MessageBox]::Show(($T_BadZ + "  " + $ax + "  " + $m), $T_Title, "OK", "Warning") | Out-Null
                return
            }
            if ([math]::Abs($n) -ge 0.001) { $shift[$ax] = $n }
        }
        if ($shift.Count -gt 0) { $script:pbpByModel[$m] = $shift }
    }
    $script:dialogOk = $true
    $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Close()
})

Update-IssueVisibility

$result = $form.ShowDialog()
if (-not $script:dialogOk) {
    Write-Host "Cancelled."
    exit 1
}

# ---- build plan from UI (filter selection, not WinForms Visible) ----
$selectedModels = @($modelChecks.Keys | Where-Object { $modelChecks[$_].Checked })
$planItems = @()
foreach ($rec in $issueRows) {
    if (-not $rec.Included) { continue }
    $it = $rec.Issue
    if ($selectedModels -notcontains $it.model_name) { continue }
    $action = Get-TripleAction $rec.Triple $it.do_action
    $rule = if ($it.kind -eq "name") { $T_KeepName } else { $T_KeepPos }
    $exName = $it.name_exemplar
    if ($it.kind -eq "name" -and [string]::IsNullOrWhiteSpace($exName)) { $exName = $it.value_exemplar }
    $planItems += [pscustomobject]@{
        model_name     = $it.model_name
        model_path     = $it.model_path
        category       = $it.category_kind
        element        = $it.element
        name_exemplar  = $exName
        problem_code   = $it.problem_code
        problem        = $it.problem
        rule           = $rule
        action         = $action
        action_ru      = $ActRu[$action]
        value_host     = $it.value_host
        value_exemplar = $it.value_exemplar
        delta          = $it.delta
        status         = $it.status
    }
}

if ($null -ne $script:pbpByModel -and $script:pbpByModel.Count -gt 0) {
    foreach ($m in $selectedModels) {
        if (-not $script:pbpByModel.ContainsKey($m)) { continue }
        $sh = $script:pbpByModel[$m]
        $parts = New-Object System.Collections.Generic.List[string]
        foreach ($ax in @("X", "Y", "Z")) {
            if (-not $sh.ContainsKey($ax)) { continue }
            $txt = ([double]$sh[$ax]).ToString("0.###", [Globalization.CultureInfo]::InvariantCulture)
            [void]$parts.Add(("{0}={1}" -f $ax, $txt))
        }
        if ($parts.Count -eq 0) { continue }
        $planItems += [pscustomobject]@{
            model_name     = $m
            model_path     = $modelPaths[$m]
            category       = "Coords"
            element        = $T_PbpElem
            name_exemplar  = $T_PbpElem
            problem_code   = "pbp_delta"
            problem        = $T_PbpMoved
            rule           = "PBP XYZ"
            action         = $ActModify
            action_ru      = $ActRu[$ActModify]
            value_host     = ""
            value_exemplar = ""
            delta          = ($parts -join "; ")
            status         = "RED"
        }
    }
}

if ($planItems.Count -eq 0) {
    [System.Windows.Forms.MessageBox]::Show(
        $NeedPlanMsg,
        $T_Title, "OK", "Warning") | Out-Null
    exit 1
}

Write-FixPlan -Path $fixPlanPath -Items $planItems -SyncAfter $cbSync.Checked
$paths = @()
foreach ($m in $selectedModels) {
    $p = $modelPaths[$m]
    if ($p) { $paths += $p } else { $paths += $m }
}
Write-RvtList -Path $rvtListPath -ModelPaths $paths

$doM = @($planItems | Where-Object { $_.action -eq $ActModify }).Count
$doN = @($planItems | Where-Object { $_.action -eq $ActRename }).Count
$doS = @($planItems | Where-Object { $_.action -eq $ActAccept -or $_.action -eq $ActReject }).Count
$doP = @($planItems | Where-Object { $_.problem_code -eq "pbp_delta" -and $_.action -eq $ActModify }).Count
Write-Host "Saved: $fixPlanPath"
Write-Host "Saved: $rvtListPath"
Write-Host ("Plan: modify={0} rename={1} skip={2} pbp={3}" -f $doM, $doN, $doS, $doP)
Write-Host ""
Write-Host "NEXT in Revit Batch Processor:"
Write-Host "  Task Script = levels_grids_apply.py"
Write-Host "  File List   = rvt_list.txt"
Write-Host "  Worksets    = Open All"
exit 0
