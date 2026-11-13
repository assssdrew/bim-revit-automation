# Name: csv_to_colored_xlsx.ps1
# Version: 1.0
# What it does: Build colored Excel levels/grids report from latest audit CSV files.
# Inputs: Optional -CsvPath, -DetailsCsvPath, -OutputDir; xlsx_out_path.cfg.
# Outputs: levels_grids_*.xlsx in configured folder.
# How to run: color_latest_report.cmd or powershell -File csv_to_colored_xlsx.ps1
# Notes: OpenXML; Summary and Details sheets.
param(
    [Parameter(Mandatory = $false)][string]$CsvPath,
    [Parameter(Mandatory = $false)][string]$DetailsCsvPath,
    [Parameter(Mandatory = $false)][string]$OutputDir,
    [Parameter(Mandatory = $false)][string]$LogPath
)

# ASCII-only. Russian via Unicode code points.
$ErrorActionPreference = "Stop"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$assetsDir = Join-Path $scriptDir "assets"

function U {
    param([int[]]$Codes)
    return (-join ($Codes | ForEach-Object { [char]([int]$_) }))
}

$H_Status = U @(0x0421, 0x0442, 0x0430, 0x0442, 0x0443, 0x0441)
$H_Model = U @(0x0418, 0x043C, 0x044F, 0x0020, 0x043C, 0x043E, 0x0434, 0x0435, 0x043B, 0x0438)
$H_ModelFile = U @(0x0418, 0x043C, 0x044F, 0x0020, 0x0444, 0x0430, 0x0439, 0x043B, 0x0430)
$H_Codes = U @(0x041A, 0x043E, 0x0434, 0x044B)
$H_Problem = U @(0x041F, 0x0440, 0x043E, 0x0431, 0x043B, 0x0435, 0x043C, 0x0430)
$H_Code = U @(0x041A, 0x043E, 0x0434)
$H_Category = U @(0x041A, 0x0430, 0x0442, 0x0435, 0x0433, 0x043E, 0x0440, 0x0438, 0x044F)
$H_Total = U @(0x0418, 0x0442, 0x043E, 0x0433, 0x043E)
$H_Delta = U @(0x0420, 0x0430, 0x0437, 0x043D, 0x0438, 0x0446, 0x0430)
$H_Tol = U @(0x0414, 0x043E, 0x043F, 0x0443, 0x0441, 0x043A)
$H_IssuesTotal = U @(0x0420, 0x0430, 0x0441, 0x0445, 0x043E, 0x0436, 0x0434, 0x0435, 0x043D, 0x0438, 0x0439)
$H_Red = U @(0x041A, 0x0440, 0x0438, 0x0442, 0x0438, 0x0447, 0x0435, 0x0441, 0x043A, 0x0438, 0x0445)
$H_Yellow = U @(0x041F, 0x0440, 0x0435, 0x0434, 0x0443, 0x043F, 0x0440, 0x0435, 0x0436, 0x0434, 0x0435, 0x043D, 0x0438, 0x0439)
$CatLevel = U @(0x0423, 0x0440, 0x043E, 0x0432, 0x0435, 0x043D, 0x044C)          # Уровень
$CatLevels = U @(0x0423, 0x0440, 0x043E, 0x0432, 0x043D, 0x0438)                   # Уровни
$CatGrid = U @(0x041E, 0x0441, 0x044C)                                              # Ось
$CatGrids = U @(0x041E, 0x0441, 0x0438)                                             # Оси

$SheetSummary = U @(0x0421, 0x0432, 0x043E, 0x0434, 0x043A, 0x0430)
$SheetDetails = U @(0x0414, 0x0435, 0x0442, 0x0430, 0x043B, 0x0438)
$SheetDash = U @(0x0414, 0x0430, 0x0448, 0x0431, 0x043E, 0x0440, 0x0434)

# Revit Coordination Review messages (flat columns on Dashboard)
$D_NameChanged = U @(0x0418, 0x043C, 0x044F, 0x0020, 0x0438, 0x0437, 0x043C, 0x0435, 0x043D, 0x0435, 0x043D, 0x043E)
$D_LevelMoved = U @(0x0423, 0x0440, 0x043E, 0x0432, 0x0435, 0x043D, 0x044C, 0x0020, 0x043F, 0x0435, 0x0440, 0x0435, 0x043C, 0x0435, 0x0449, 0x0435, 0x043D)
$D_GridMoved = U @(0x0421, 0x0435, 0x0442, 0x043A, 0x0430, 0x0020, 0x043F, 0x0435, 0x0440, 0x0435, 0x043C, 0x0435, 0x0449, 0x0435, 0x043D, 0x0430)
$D_Deleted = U @(0x042D, 0x043B, 0x0435, 0x043C, 0x0435, 0x043D, 0x0442, 0x0020, 0x0443, 0x0434, 0x0430, 0x043B, 0x0435, 0x043D)
$D_PbpMoved = U @(0x0411, 0x0430, 0x0437, 0x043E, 0x0432, 0x0430, 0x044F, 0x0020, 0x0442, 0x043E, 0x0447, 0x043A, 0x0430, 0x0020, 0x0441, 0x043C, 0x0435, 0x0449, 0x0435, 0x043D, 0x0430)
$D_SurveyMoved = U @(0x0422, 0x043E, 0x0447, 0x043A, 0x0430, 0x0020, 0x0441, 0x044A, 0x0435, 0x043C, 0x043A, 0x0438, 0x0020, 0x0441, 0x043C, 0x0435, 0x0449, 0x0435, 0x043D, 0x0430)
$D_Other = U @(0x041F, 0x0440, 0x043E, 0x0447, 0x0435, 0x0435)

$H_ShortProblems = U @(0x041A, 0x043E, 0x0440, 0x043E, 0x0442, 0x043A, 0x043E, 0x0020, 0x043E, 0x0020, 0x043F, 0x0440, 0x043E, 0x0431, 0x043B, 0x0435, 0x043C, 0x0430, 0x0445)
$H_ExemplarLink = U @(0x0421, 0x0432, 0x044F, 0x0437, 0x044C, 0x002D, 0x044D, 0x0442, 0x0430, 0x043B, 0x043E, 0x043D) # Связь-эталон
$H_ModelPath = U @(0x041F, 0x0443, 0x0442, 0x044C, 0x0020, 0x043C, 0x043E, 0x0434, 0x0435, 0x043B, 0x0438) # Путь модели
$H_Timestamp = U @(0x0414, 0x0430, 0x0442, 0x0430, 0x002F, 0x0432, 0x0440, 0x0435, 0x043C, 0x044F) # Дата/время


function Get-NativePath([string]$Path) {
    if ([string]::IsNullOrWhiteSpace($Path)) { return $null }
    $p = $Path.Trim().Trim('"')
    $marker = "FileSystem::"
    $idx = $p.IndexOf($marker, [StringComparison]::OrdinalIgnoreCase)
    if ($idx -ge 0) { $p = $p.Substring($idx + $marker.Length) }
    return $p
}

function Write-LogLine([string]$Text, [string]$Target) {
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $dir = Split-Path -Parent $Target
    if ($dir -and !(Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    Add-Content -LiteralPath $Target -Value "$ts | $Text" -Encoding UTF8
}

function Write-Utf8NoBom([string]$Path, [string]$Content) {
    $enc = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($Path, $Content, $enc)
}

function Test-LegacyDesktopReports([string]$Path) {
    if ([string]::IsNullOrWhiteSpace($Path)) { return $false }
    $p = $Path.Replace('/', '\').ToLowerInvariant()
    return $p.EndsWith('\batch_levels_grids_reports')
}

function Get-XlsxOutputDir([string]$Preferred) {
    if (-not [string]::IsNullOrWhiteSpace($Preferred)) {
        $p = Get-NativePath $Preferred
        if ($p -and -not (Test-LegacyDesktopReports $p)) { return $p }
    }
    $cfg = Join-Path $scriptDir "xlsx_out_path.cfg"
    if (Test-Path -LiteralPath $cfg) {
        $line = Get-Content -LiteralPath $cfg -ErrorAction SilentlyContinue |
            Where-Object { $_ -and -not $_.StartsWith("#") } | Select-Object -First 1
        if ($line) {
            $p = Get-NativePath $line.Trim()
            if ($p -and -not (Test-LegacyDesktopReports $p)) { return $p }
        }
    }
    return $null
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

function Escape-Xml([string]$Text) {
    if ($null -eq $Text) { return "" }
    return ($Text -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;' -replace '"', '&quot;')
}

function Get-DashBucket([string]$Code, [string]$Category) {
    $c = ($Code + "").Trim().ToLowerInvariant()
    switch ($c) {
        "name_mismatch" { return $D_NameChanged }
        "elevation_delta" { return $D_LevelMoved }
        "position_delta" { return $D_GridMoved }
        "angle_delta" { return $D_GridMoved }
        "missing_in_host" { return $D_Other }
        "leftover_exemplar" { return $D_Other }
        "extra_in_host" { return $D_Deleted }
        "whitelist_missing" { return $D_Deleted }
        "pbp_delta" { return $D_PbpMoved }
        "pbp_missing" { return $D_PbpMoved }
        "survey_delta" { return $D_SurveyMoved }
        "survey_missing" { return $D_SurveyMoved }
        "cr_grid_moved" { return $D_GridMoved }
        "grid_relative_moved" { return $D_GridMoved }
        "cr_level_moved" { return $D_LevelMoved }
        "cr_name_changed" { return $D_NameChanged }
        "cr_element_deleted" { return $D_Deleted }
        "cr_watched_link_moved" { return $D_GridMoved }
        default { return $D_Other }
    }
}

function Build-Dashboard($Details) {
    $dashCols = @($D_NameChanged, $D_LevelMoved, $D_GridMoved, $D_PbpMoved, $D_SurveyMoved, $D_Deleted, $D_Other)
    $headers = @($H_Model) + $dashCols + @($H_Total)
    if (-not $Details -or -not $Details.Rows -or $Details.Rows.Count -eq 0) {
        return @{ Headers = $headers; Rows = @(); FirstRowByModel = @{} }
    }
    $modelKey = $H_Model
    $sample = $Details.Rows[0]
    if ($sample.PSObject.Properties.Name -notcontains $modelKey -and $sample.PSObject.Properties.Name -contains $H_ModelFile) {
        $modelKey = $H_ModelFile
    }
    $models = New-Object System.Collections.Generic.List[string]
    $counts = @{}; $firstRow = @{}; $excelRow = 2
    foreach ($row in $Details.Rows) {
        $m = if ($row.PSObject.Properties.Name -contains $modelKey) { [string]$row.$modelKey } else { "?" }
        if ([string]::IsNullOrWhiteSpace($m)) { $m = "?" }
        if (-not $models.Contains($m)) {
            [void]$models.Add($m); $firstRow[$m] = $excelRow
            foreach ($col in $dashCols) { $counts["$m|$col"] = 0 }
        }
        $code = ""; $cat = ""
        if ($row.PSObject.Properties.Name -contains $H_Code) { $code = [string]$row.$H_Code }
        if ($row.PSObject.Properties.Name -contains $H_Category) { $cat = [string]$row.$H_Category }
        # Also accept already-localized Revit message text in "Проблема"
        if ([string]::IsNullOrWhiteSpace($code) -and $row.PSObject.Properties.Name -contains $H_Problem) {
            $p = [string]$row.$H_Problem
            if ($p -eq $D_NameChanged) { $code = "name_mismatch" }
            elseif ($p -eq $D_LevelMoved) { $code = "elevation_delta" }
            elseif ($p -eq $D_GridMoved) { $code = "position_delta" }
            elseif ($p -eq $D_Deleted) { $code = "missing_in_host" }
            elseif ($p -eq $D_PbpMoved) { $code = "pbp_delta" }
            elseif ($p -eq $D_SurveyMoved) { $code = "survey_delta" }
        }
        $bucket = Get-DashBucket $code $cat
        $k = "$m|$bucket"
        if (-not $counts.ContainsKey($k)) { $counts[$k] = 0 }
        $counts[$k] = [int]$counts[$k] + 1
        $excelRow++
    }
    $rows = @()
    foreach ($m in $models) {
        $obj = [ordered]@{}; $obj[$H_Model] = $m; $total = 0
        foreach ($col in $dashCols) {
            $n = if ($counts.ContainsKey("$m|$col")) { [int]$counts["$m|$col"] } else { 0 }
            $obj[$col] = "$n"; $total += $n
        }
        $obj[$H_Total] = "$total"
        $rows += [pscustomobject]$obj
    }
    return @{ Headers = $headers; Rows = $rows; FirstRowByModel = $firstRow }
}


function New-SheetRowMeta {
    param($Row, [int]$OutlineLevel = 0, [bool]$Hidden = $false, [bool]$Collapsed = $false, [bool]$IsBlank = $false)
    return [pscustomobject]@{
        Row         = $Row
        OutlineLevel = $OutlineLevel
        Hidden      = $Hidden
        Collapsed   = $Collapsed
        IsBlank     = $IsBlank
    }
}

function Split-ShortProblems([string]$Text) {
    if ([string]::IsNullOrWhiteSpace($Text)) { return @() }
    $parts = New-Object System.Collections.ArrayList
    foreach ($p in ($Text -split ';')) {
        $t = $p.Trim()
        if ($t) { [void]$parts.Add($t) }
    }
    return @($parts.ToArray())
}

function Get-HeaderOnlyColWidth([string]$Header) {
    # Excel width ~ character count; size by header only (values may truncate).
    $n = if ($null -eq $Header) { 4 } else { $Header.Length }
    $w = [Math]::Max(3, $n + 1)
    if ($w -gt 45) { $w = 45 }
    return [Math]::Round([double]$w, 2)
}

function Expand-SummaryOutlineRows($Headers, $Rows) {
    <#
      Parent: full model row, "Коротко о проблемах" = count.
      Children: one problem per row; other columns empty. Collapsed by default.
      Blank separator between models (outside outline).
    #>
    $out = @()
    $dataRows = @()
    foreach ($r in @($Rows)) {
        $any = $false
        foreach ($h in $Headers) {
            if ($r.PSObject.Properties.Name -contains $h -and -not [string]::IsNullOrWhiteSpace([string]$r.$h)) {
                $any = $true; break
            }
        }
        if ($any) { $dataRows += ,$r }
    }
    for ($i = 0; $i -lt $dataRows.Count; $i++) {
        $src = $dataRows[$i]
        $problems = @()
        if ($src.PSObject.Properties.Name -contains $H_ShortProblems) {
            $problems = @(Split-ShortProblems ([string]$src.$H_ShortProblems))
        }
        $parent = [ordered]@{}
        foreach ($h in $Headers) {
            if ($src.PSObject.Properties.Name -contains $h) { $parent[$h] = [string]$src.$h } else { $parent[$h] = "" }
        }
        $parent[$H_ShortProblems] = "$($problems.Count)"
        $hasKids = $problems.Count -gt 0
        $out += ,(New-SheetRowMeta -Row ([pscustomobject]$parent) -OutlineLevel 0 -Hidden $false -Collapsed $hasKids)
        foreach ($p in $problems) {
            $child = [ordered]@{}
            foreach ($h in $Headers) { $child[$h] = "" }
            $child[$H_ShortProblems] = $p
            $out += ,(New-SheetRowMeta -Row ([pscustomobject]$child) -OutlineLevel 1 -Hidden $true -Collapsed $false)
        }
        if ($i -lt ($dataRows.Count - 1)) {
            $blank = [ordered]@{}
            foreach ($h in $Headers) { $blank[$h] = "" }
            $out += ,(New-SheetRowMeta -Row ([pscustomobject]$blank) -IsBlank $true)
        }
    }
    return @($out)
}

function Expand-DetailsOutlineRows($Headers, $Rows) {
    <#
      Group by model: parent = model + exemplar + path + date + worst status + problem count;
      children = detail rows with Status cleared. Collapsed by default.
    #>
    $metaList = @()
    $firstRowByModel = @{}
    $rowArr = @($Rows)
    if ($rowArr.Count -eq 0) {
        return @{ MetaRows = @(); FirstRowByModel = $firstRowByModel }
    }
    $modelKey = $H_Model
    $sample = $rowArr[0]
    if ($sample.PSObject.Properties.Name -notcontains $modelKey -and $sample.PSObject.Properties.Name -contains $H_ModelFile) {
        $modelKey = $H_ModelFile
    }

    $groupNames = @()
    $groupMap = @{}
    foreach ($row in $rowArr) {
        $m = "?"
        if ($row.PSObject.Properties.Name -contains $modelKey) { $m = [string]$row.$modelKey }
        if ([string]::IsNullOrWhiteSpace($m)) { $m = "?" }
        if (-not $groupMap.ContainsKey($m)) {
            $groupNames += $m
            $groupMap[$m] = New-Object System.Collections.ArrayList
        }
        [void]$groupMap[$m].Add($row)
    }

    $keepOnParent = @($modelKey, $H_ExemplarLink, $H_ModelPath, $H_Timestamp)
    $excelRow = 2
    foreach ($gName in $groupNames) {
        $childRows = @($groupMap[$gName].ToArray())
        $worst = ""
        $rank = @{ "GREEN" = 0; "OK" = 0; "YELLOW" = 1; "RED" = 2; "FAIL" = 2; "ERROR" = 2 }
        $bestRank = -1
        foreach ($it in $childRows) {
            if ($it.PSObject.Properties.Name -contains $H_Status) {
                $st = ([string]$it.$H_Status).ToUpperInvariant()
                if ($rank.ContainsKey($st) -and $rank[$st] -gt $bestRank) {
                    $bestRank = $rank[$st]
                    $worst = $st
                }
            }
        }
        if ($worst -eq "OK") { $worst = "GREEN" }
        if ($worst -eq "FAIL" -or $worst -eq "ERROR") { $worst = "RED" }

        $parent = [ordered]@{}
        foreach ($h in $Headers) { $parent[$h] = "" }
        $parent[$modelKey] = $gName
        if ($Headers -contains $H_Problem) { $parent[$H_Problem] = "$($childRows.Count)" }
        if ($Headers -contains $H_Status -and $worst) { $parent[$H_Status] = $worst }
        $first = $childRows[0]
        foreach ($h in $keepOnParent) {
            if ($h -eq $modelKey) { continue }
            if ($Headers -contains $h -and $first.PSObject.Properties.Name -contains $h) {
                $parent[$h] = [string]$first.$h
            }
        }

        $firstRowByModel[$gName] = $excelRow
        $hasKids = $childRows.Count -gt 0
        $metaList += ,(New-SheetRowMeta -Row ([pscustomobject]$parent) -OutlineLevel 0 -Hidden $false -Collapsed $hasKids)
        $excelRow++
        foreach ($it in $childRows) {
            $child = [ordered]@{}
            foreach ($h in $Headers) {
                if ($it.PSObject.Properties.Name -contains $h) { $child[$h] = [string]$it.$h } else { $child[$h] = "" }
            }
            if ($Headers -contains $H_Status) { $child[$H_Status] = "" }
            $metaList += ,(New-SheetRowMeta -Row ([pscustomobject]$child) -OutlineLevel 1 -Hidden $true -Collapsed $false)
            $excelRow++
        }
    }
    return @{ MetaRows = $metaList; FirstRowByModel = $firstRowByModel }
}

function New-OpenXmlXlsx {
    param($XlsxPath, $Summary, $Details, $Dashboard)

    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $tmp = Join-Path $env:TEMP ("lg_openxml_" + [guid]::NewGuid().ToString("N"))
    New-Item -ItemType Directory -Path $tmp -Force | Out-Null
    foreach ($p in @("_rels", "xl", "xl\_rels", "xl\worksheets", "xl\worksheets\_rels", "xl\media", "xl\drawings")) {
        New-Item -ItemType Directory -Path (Join-Path $tmp $p) -Force | Out-Null
    }

    $shared = New-Object System.Collections.Generic.List[string]
    $sharedIndex = @{}
    function Get-SharedIndex([string]$Text) {
        if ($null -eq $Text) { $Text = "" }
        if ($sharedIndex.ContainsKey($Text)) { return $sharedIndex[$Text] }
        $idx = $shared.Count; $shared.Add($Text); $sharedIndex[$Text] = $idx; return $idx
    }
    function ColName([int]$Index1Based) {
        $n = [int]$Index1Based; $name = ""
        while ($n -gt 0) { $n = $n - 1; $name = ([char](65 + ($n % 26))) + $name; $n = [int][math]::Floor([double]$n / 26.0) }
        return $name
    }

    # fills: 0 none,1 gray125,2 header blue,3 green,4 yellow,5 red,6 bandA,7 bandB
    $fillXml = New-Object System.Text.StringBuilder
    [void]$fillXml.Append('<fill><patternFill patternType="none"/></fill>')
    [void]$fillXml.Append('<fill><patternFill patternType="gray125"/></fill>')
    [void]$fillXml.Append('<fill><patternFill patternType="solid"><fgColor rgb="FFD9E1F2"/></patternFill></fill>')
    [void]$fillXml.Append('<fill><patternFill patternType="solid"><fgColor rgb="FFC6EFCE"/></patternFill></fill>')
    [void]$fillXml.Append('<fill><patternFill patternType="solid"><fgColor rgb="FFFFEB9C"/></patternFill></fill>')
    [void]$fillXml.Append('<fill><patternFill patternType="solid"><fgColor rgb="FFFFC7CE"/></patternFill></fill>')
    [void]$fillXml.Append('<fill><patternFill patternType="solid"><fgColor rgb="FFE2EFDA"/></patternFill></fill>')
    [void]$fillXml.Append('<fill><patternFill patternType="solid"><fgColor rgb="FFF2F2F2"/></patternFill></fill>')
    $fillCount = 8

    # cellXfs: 0 def,1 header,2 OK,3 YEL,4 RED,5 int,6 wrap,7 red-int,8 bandA,9 bandB,10 bandA-int,11 link
    $xfXml = New-Object System.Text.StringBuilder
    [void]$xfXml.Append('<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>')
    [void]$xfXml.Append('<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>')
    [void]$xfXml.Append('<xf numFmtId="0" fontId="0" fillId="3" borderId="0" xfId="0" applyFill="1"/>')
    [void]$xfXml.Append('<xf numFmtId="0" fontId="0" fillId="4" borderId="0" xfId="0" applyFill="1"/>')
    [void]$xfXml.Append('<xf numFmtId="0" fontId="0" fillId="5" borderId="0" xfId="0" applyFill="1"/>')
    [void]$xfXml.Append('<xf numFmtId="1" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>')
    [void]$xfXml.Append('<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment wrapText="1" vertical="top"/></xf>')
    [void]$xfXml.Append('<xf numFmtId="1" fontId="0" fillId="5" borderId="0" xfId="0" applyNumberFormat="1" applyFill="1"/>')
    [void]$xfXml.Append('<xf numFmtId="0" fontId="0" fillId="6" borderId="0" xfId="0" applyFill="1"/>')
    [void]$xfXml.Append('<xf numFmtId="0" fontId="0" fillId="7" borderId="0" xfId="0" applyFill="1"/>')
    [void]$xfXml.Append('<xf numFmtId="1" fontId="0" fillId="6" borderId="0" xfId="0" applyNumberFormat="1" applyFill="1"/>')
    [void]$xfXml.Append('<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>')
    $xfCount = 12

    $stylesXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="3">
    <font><sz val="11"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><name val="Calibri"/></font>
    <font><u/><sz val="11"/><color theme="10"/><name val="Calibri"/></font>
  </fonts>
  <fills count="$fillCount">$($fillXml.ToString())</fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="$xfCount">$($xfXml.ToString())</cellXfs>
</styleSheet>
"@

    function Build-SheetXml {
        param(
            $Headers,
            $MetaRows,
            $StatusColName,
            $IntColNames,
            [bool]$ColorPositiveInts,
            [bool]$BandByFirstCol,
            $HyperlinkMap,
            [bool]$ProblemCountCols,
            $ExtraRows,
            [bool]$EnableOutline
        )

        $sb = New-Object System.Text.StringBuilder
        [void]$sb.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
        [void]$sb.Append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')
        if ($EnableOutline) {
            [void]$sb.Append('<sheetPr><outlinePr summaryBelow="0" summaryRight="0"/></sheetPr>')
        }
        [void]$sb.Append('<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>')
        if ($EnableOutline) {
            [void]$sb.Append('<sheetFormatPr defaultRowHeight="15" outlineLevelRow="1"/>')
        }

        # Column widths = header text only
        if ($Headers -and $Headers.Count -gt 0) {
            [void]$sb.Append('<cols>')
            for ($c = 0; $c -lt $Headers.Count; $c++) {
                $w = Get-HeaderOnlyColWidth ([string]$Headers[$c])
                $ci = $c + 1
                [void]$sb.Append("<col min=`"$ci`" max=`"$ci`" width=`"$w`" customWidth=`"1`"/>")
            }
            [void]$sb.Append('</cols>')
        }

        [void]$sb.Append('<sheetData>')

        [void]$sb.Append('<row r="1" ht="22">')
        for ($c = 0; $c -lt $Headers.Count; $c++) {
            $ref = (ColName ($c + 1)) + "1"
            $idx = Get-SharedIndex ([string]$Headers[$c])
            [void]$sb.Append("<c r=`"$ref`" t=`"s`" s=`"1`"><v>$idx</v></c>")
        }
        [void]$sb.Append('</row>')

        $statusIdx = -1
        if ($StatusColName) {
            for ($i = 0; $i -lt $Headers.Count; $i++) { if ($Headers[$i] -eq $StatusColName) { $statusIdx = $i; break } }
        }
        $intSet = @{}; foreach ($n in $IntColNames) { if ($n) { $intSet[$n] = $true } }
        $probCols = @{}
        if ($ProblemCountCols) { for ($n = 1; $n -le 10; $n++) { $probCols["$n"] = $n } }

        # Normalize plain rows to meta if needed
        $normalized = @()
        foreach ($item in @($MetaRows)) {
            if ($null -eq $item) { continue }
            if ($item.PSObject.Properties.Name -contains "Row") { $normalized += $item }
            else { $normalized += (New-SheetRowMeta -Row $item) }
        }

        $band = 0; $prevKey = $null; $rNum = 2
        foreach ($meta in $normalized) {
            $row = $meta.Row
            $ol = 0; if ($meta.PSObject.Properties.Name -contains "OutlineLevel") { $ol = [int]$meta.OutlineLevel }
            $hid = $false; if ($meta.PSObject.Properties.Name -contains "Hidden") { $hid = [bool]$meta.Hidden }
            $colp = $false; if ($meta.PSObject.Properties.Name -contains "Collapsed") { $colp = [bool]$meta.Collapsed }
            $isBlank = $false; if ($meta.PSObject.Properties.Name -contains "IsBlank") { $isBlank = [bool]$meta.IsBlank }

            $bandStyle = 0
            if ($BandByFirstCol -and $Headers.Count -gt 0 -and $ol -eq 0 -and -not $isBlank) {
                $hk = $Headers[0]; $key = ""
                if ($row.PSObject.Properties.Name -contains $hk) { $key = [string]$row.$hk }
                if ($prevKey -ne $null -and $key -ne $prevKey -and $key -ne "") { $band = 1 - $band }
                if ($key -ne "") { $prevKey = $key }
                $bandStyle = if ($band -eq 0) { 8 } else { 9 }
            }

            $rowAttrs = "r=`"$rNum`""
            if ($ol -gt 0) { $rowAttrs += " outlineLevel=`"$ol`"" }
            if ($hid) { $rowAttrs += ' hidden="1"' }
            if ($colp) { $rowAttrs += ' collapsed="1"' }
            [void]$sb.Append("<row $rowAttrs>")

            for ($c = 0; $c -lt $Headers.Count; $c++) {
                $h = $Headers[$c]
                $val = ""; if ($row.PSObject.Properties.Name -contains $h) { $val = [string]$row.$h }
                $ref = (ColName ($c + 1)) + "$rNum"
                $style = $bandStyle
                # Status fill only on parent (outlineLevel 0) rows
                if ($c -eq $statusIdx -and $ol -eq 0 -and -not $isBlank) {
                    switch ($val.ToUpperInvariant()) {
                        "OK" { $style = 2 }; "GREEN" { $style = 2 }; "YELLOW" { $style = 3 }
                        "FAIL" { $style = 4 }; "RED" { $style = 4 }; "ERROR" { $style = 4 }
                    }
                }
                $asNum = $null
                $isProb = $probCols.ContainsKey($h)
                # Count in short_problems / problem on parent is integer
                $forceInt = ($ol -eq 0 -and (($h -eq $H_ShortProblems) -or ($h -eq $H_Problem)) -and $val -match '^\d+$')
                if (($intSet.ContainsKey($h) -or $isProb -or $forceInt) -and $val -ne "" -and [double]::TryParse(($val.Replace(',', '.')), [ref]$asNum)) {
                    $num = [math]::Round($asNum)
                    $nStyle = 5
                    if ($isProb -and $num -gt 0) { $nStyle = 7 }
                    elseif ($ColorPositiveInts -and $num -gt 0) { $nStyle = 7 }
                    elseif ($bandStyle -eq 8) { $nStyle = 10 }
                    [void]$sb.Append("<c r=`"$ref`" s=`"$nStyle`"><v>$num</v></c>")
                }
                else {
                    if ($HyperlinkMap -and $HyperlinkMap.ContainsKey($ref)) { $style = 11 }
                    elseif (($h -eq $H_ShortProblems -or $h -eq $H_Problem) -and $ol -gt 0) { if ($style -eq 0) { $style = 6 } }
                    $idx = Get-SharedIndex $val
                    if ($style -gt 0) { [void]$sb.Append("<c r=`"$ref`" t=`"s`" s=`"$style`"><v>$idx</v></c>") }
                    else { [void]$sb.Append("<c r=`"$ref`" t=`"s`"><v>$idx</v></c>") }
                }
            }
            [void]$sb.Append('</row>'); $rNum++
        }

        if ($ExtraRows) {
            foreach ($er in $ExtraRows) {
                [void]$sb.Append("<row r=`"$rNum`" ht=`"$($er.Height)`">")
                foreach ($cell in $er.Cells) {
                    $ref = (ColName ([int]$cell.Col)) + "$rNum"
                    if ($cell.ContainsKey("Num")) {
                        [void]$sb.Append("<c r=`"$ref`" s=`"$($cell.Style)`"><v>$($cell.Num)</v></c>")
                    }
                    else {
                        $idx = Get-SharedIndex ([string]$cell.Text)
                        [void]$sb.Append("<c r=`"$ref`" t=`"s`" s=`"$($cell.Style)`"><v>$idx</v></c>")
                    }
                }
                [void]$sb.Append('</row>'); $rNum++
            }
        }

        [void]$sb.Append('</sheetData>')
        $dataLast = [Math]::Max(1, $rNum - 1)
        if ($Headers.Count -gt 0) {
            $lastCol = ColName $Headers.Count
            [void]$sb.Append("<autoFilter ref=`"A1:${lastCol}${dataLast}`"/>")
        }
        if ($HyperlinkMap -and $HyperlinkMap.Count -gt 0) {
            [void]$sb.Append('<hyperlinks>')
            foreach ($key in $HyperlinkMap.Keys) {
                [void]$sb.Append("<hyperlink ref=`"$key`" location=`"$($HyperlinkMap[$key])`" display=`"`"/>")
            }
            [void]$sb.Append('</hyperlinks>')
        }
        [void]$sb.Append('</worksheet>')
        return @{ Xml = $sb.ToString(); LastRow = ($rNum - 1) }
    }

    # ---- Summary: count on parent + collapsed problem children ----
    if (-not $Summary) { $Summary = @{ Headers = @(); Rows = @() } }
    if (-not $Details) { $Details = @{ Headers = @($H_Model, $H_Status); Rows = @() } }
    if (-not $Dashboard) { $Dashboard = @{ Headers = @($H_Model); Rows = @(); FirstRowByModel = @{} } }

    $sumHeaders = @($Summary.Headers)
    $sumMeta = Expand-SummaryOutlineRows $sumHeaders @($Summary.Rows)

    $sumInt = @($H_IssuesTotal, $H_Red, $H_Yellow, $H_ShortProblems)
    foreach ($h in $sumHeaders) {
        if ($h -match '^[0-9]+$') { $sumInt += $h }
    }

    $sheet1Pack = Build-SheetXml -Headers $sumHeaders -MetaRows $sumMeta -StatusColName $H_Status `
        -IntColNames $sumInt -ColorPositiveInts $false -BandByFirstCol $false `
        -HyperlinkMap $null -ProblemCountCols $false -ExtraRows $null -EnableOutline $true

    # ---- Details: parent per model + collapsed detail children ----
    $detHeaders = @($Details.Headers); $detRowsRaw = @($Details.Rows)
    if ($detHeaders.Count -eq 0) { $detHeaders = @($H_Model, $H_Status) }
    $detExpand = Expand-DetailsOutlineRows $detHeaders $detRowsRaw
    $detMeta = @($detExpand.MetaRows)
    if ($Dashboard -and $detExpand.FirstRowByModel) {
        $Dashboard.FirstRowByModel = $detExpand.FirstRowByModel
    }

    $sheet2Pack = Build-SheetXml -Headers $detHeaders -MetaRows $detMeta -StatusColName $H_Status `
        -IntColNames @($H_Delta, $H_Tol, $H_Problem) -ColorPositiveInts $false -BandByFirstCol $false `
        -HyperlinkMap $null -ProblemCountCols $false -ExtraRows $null -EnableOutline $true

    # dashboard
    $dashHeaders = @($Dashboard.Headers); $dashRows = @($Dashboard.Rows)
    $linkMap = @{}; $dashR = 2
    foreach ($row in $dashRows) {
        $m = [string]$row.$H_Model
        if ($Dashboard.FirstRowByModel -and $Dashboard.FirstRowByModel.ContainsKey($m)) {
            $linkMap["A$dashR"] = "$SheetDetails!A$([int]$Dashboard.FirstRowByModel[$m])"
        }
        $dashR++
    }
    $dashInt = @(); foreach ($h in $dashHeaders) { if ($h -ne $H_Model) { $dashInt += $h } }
    $dashMeta = @($dashRows | ForEach-Object { New-SheetRowMeta -Row $_ })
    $sheet3Pack = Build-SheetXml -Headers $dashHeaders -MetaRows $dashMeta -StatusColName $null `
        -IntColNames $dashInt -ColorPositiveInts $true -BandByFirstCol $false `
        -HyperlinkMap $linkMap -ProblemCountCols $false -ExtraRows $null -EnableOutline $false

    $hasImg = $false

    $sst = New-Object System.Text.StringBuilder
    [void]$sst.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    [void]$sst.Append('<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="' + $shared.Count + '" uniqueCount="' + $shared.Count + '">')
    foreach ($s in $shared) { [void]$sst.Append('<si><t xml:space="preserve">' + (Escape-Xml $s) + '</t></si>') }
    [void]$sst.Append('</sst>')

    $ct = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>
"@
    Write-Utf8NoBom (Join-Path $tmp "[Content_Types].xml") $ct
    Write-Utf8NoBom (Join-Path $tmp "_rels\.rels") @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
"@
    Write-Utf8NoBom (Join-Path $tmp "xl\_rels\workbook.xml.rels") @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/>
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>
"@
    Write-Utf8NoBom (Join-Path $tmp "xl\workbook.xml") @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="$SheetSummary" sheetId="1" r:id="rId1"/>
    <sheet name="$SheetDetails" sheetId="2" r:id="rId2"/>
    <sheet name="$SheetDash" sheetId="3" r:id="rId3"/>
  </sheets>
</workbook>
"@
    Write-Utf8NoBom (Join-Path $tmp "xl\styles.xml") $stylesXml
    Write-Utf8NoBom (Join-Path $tmp "xl\sharedStrings.xml") $sst.ToString()
    Write-Utf8NoBom (Join-Path $tmp "xl\worksheets\sheet1.xml") $sheet1Pack.Xml
    Write-Utf8NoBom (Join-Path $tmp "xl\worksheets\sheet2.xml") $sheet2Pack.Xml
    Write-Utf8NoBom (Join-Path $tmp "xl\worksheets\sheet3.xml") $sheet3Pack.Xml

    if (Test-Path -LiteralPath $XlsxPath) { Remove-Item -LiteralPath $XlsxPath -Force }
    [System.IO.Compression.ZipFile]::CreateFromDirectory($tmp, $XlsxPath)
    Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue
}

# -------- main --------
if ([string]::IsNullOrWhiteSpace($CsvPath)) {
    $csvDir = Join-Path $env:USERPROFILE "Documents\doc\script\rbp\reports"
    if (!(Test-Path -LiteralPath $csvDir)) {
        New-Item -ItemType Directory -Path $csvDir -Force | Out-Null
    }
    $latestCsv = Get-ChildItem -LiteralPath $csvDir -Filter "levels_grids_audit_*.csv" -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -notmatch '_details\.csv$' } |
        Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if (-not $latestCsv) {
        $latestCsv = Get-ChildItem -LiteralPath $csvDir -Filter "lg_*.csv" -File -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -notmatch '_details\.csv$' } |
            Sort-Object LastWriteTime -Descending | Select-Object -First 1
    }
    if (-not $latestCsv) { throw "No levels_grids_audit_*.csv found in $csvDir" }
    $CsvPath = $latestCsv.FullName
}
$CsvPath = Get-NativePath $CsvPath
if (!(Test-Path -LiteralPath $CsvPath)) { throw "CSV not found: $CsvPath" }
if ([string]::IsNullOrWhiteSpace($DetailsCsvPath)) {
    $candidate = [System.IO.Path]::ChangeExtension($CsvPath, $null).TrimEnd('.') + "_details.csv"
    if (Test-Path -LiteralPath $candidate) { $DetailsCsvPath = $candidate }
} else { $DetailsCsvPath = Get-NativePath $DetailsCsvPath }

$outDir = Get-XlsxOutputDir $OutputDir
if ([string]::IsNullOrWhiteSpace($outDir)) { $outDir = Join-Path $env:USERPROFILE "Documents\doc\script\rbp\reports" }
if (!(Test-Path -LiteralPath $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }

$xlsxPath = Join-Path $outDir ([System.IO.Path]::GetFileNameWithoutExtension($CsvPath) + ".xlsx")
$logTarget = if ([string]::IsNullOrWhiteSpace($LogPath)) { Join-Path $outDir "levels_grids_audit_xlsx_export.log" } else { Get-NativePath $LogPath }

try {
    Write-LogLine "OpenXML export started. Csv=$CsvPath Xlsx=$xlsxPath" $logTarget
    $summary = ConvertFrom-SemicolonCsv -Path $CsvPath
    $details = $null
    if ($DetailsCsvPath -and (Test-Path -LiteralPath $DetailsCsvPath)) {
        $details = ConvertFrom-SemicolonCsv -Path $DetailsCsvPath
    }
    $dashboard = Build-Dashboard $details
    New-OpenXmlXlsx -XlsxPath $xlsxPath -Summary $summary -Details $details -Dashboard $dashboard
    Write-LogLine "OpenXML export success. Saved=$xlsxPath" $logTarget
    Write-Output "Created: $xlsxPath"
}
catch {
    Write-LogLine ("OpenXML export failed: " + $_.Exception.Message) $logTarget
    throw
}
