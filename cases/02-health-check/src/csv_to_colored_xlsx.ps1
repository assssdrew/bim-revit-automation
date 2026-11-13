# Name: csv_to_colored_xlsx.ps1
# Version: 1.0
# What it does: Build a colored Excel health report from the latest health_*.csv and *_links.csv.
# Inputs: Optional -CsvPath, -LinksCsvPath, -OutputDir; reads xlsx_out_path.cfg when omitted.
# Outputs: health_*.xlsx in the configured local folder; xlsx_export.log on errors.
# How to run: color_latest_report.cmd or: powershell -File csv_to_colored_xlsx.ps1
# Notes: OpenXML only (no Excel COM); status highlighting from thresholds.
param(
    [Parameter(Mandatory = $false)]
    [string]$CsvPath,
    [Parameter(Mandatory = $false)]
    [string]$LinksCsvPath,
    [Parameter(Mandatory = $false)]
    [string]$OutputDir,
    [Parameter(Mandatory = $false)]
    [string]$LogPath
)

# ASCII-only source. Russian UI strings are built from Unicode code points.
$ErrorActionPreference = "Stop"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

function U {
    param([int[]]$Codes)
    # Explicit [int] — PS may box code points as Decimal; [char]Decimal throws.
    return (-join ($Codes | ForEach-Object { [char]([int]$_) }))
}

# Russian headers / sheet names
$H_Status = U @(0x0421, 0x0442, 0x0430, 0x0442, 0x0443, 0x0441) # Status
$H_SizeMb = (U @(0x0420, 0x0430, 0x0437, 0x043C, 0x0435, 0x0440, 0x0020, 0x0444, 0x0430, 0x0439, 0x043B, 0x0430, 0x002C, 0x0020, 0x041C, 0x0411)) # Razmer faila, MB
$H_WarnCount = (U @(0x041A, 0x043E, 0x043B, 0x0438, 0x0447, 0x0435, 0x0441, 0x0442, 0x0432, 0x043E, 0x0020, 0x043F, 0x0440, 0x0435, 0x0434, 0x0443, 0x043F, 0x0440, 0x0435, 0x0436, 0x0434, 0x0435, 0x043D, 0x0438, 0x0439))
$H_RvtTotal = (U @(0x0421, 0x0432, 0x044F, 0x0437, 0x0438, 0x0020, 0x0052, 0x0056, 0x0054, 0x002C, 0x0020, 0x0432, 0x0441, 0x0435, 0x0433, 0x043E))
$H_RvtNf = (U @(0x0421, 0x0432, 0x044F, 0x0437, 0x0438, 0x0020, 0x0052, 0x0056, 0x0054, 0x002C, 0x0020, 0x006E, 0x006F, 0x0074, 0x0020, 0x0066, 0x006F, 0x0075, 0x006E, 0x0064))
$H_RvtUl = (U @(0x0421, 0x0432, 0x044F, 0x0437, 0x0438, 0x0020, 0x0052, 0x0056, 0x0054, 0x002C, 0x0020, 0x0075, 0x006E, 0x006C, 0x006F, 0x0061, 0x0064, 0x0065, 0x0064))
$H_CadTotal = (U @(0x0421, 0x0432, 0x044F, 0x0437, 0x0438, 0x0020, 0x0043, 0x0041, 0x0044, 0x002C, 0x0020, 0x0432, 0x0441, 0x0435, 0x0433, 0x043E))
$H_CadNf = (U @(0x0421, 0x0432, 0x044F, 0x0437, 0x0438, 0x0020, 0x0043, 0x0041, 0x0044, 0x002C, 0x0020, 0x006E, 0x006F, 0x0074, 0x0020, 0x0066, 0x006F, 0x0075, 0x006E, 0x0064))
$H_LinkedRvt = (U @(0x0421, 0x0432, 0x044F, 0x0437, 0x0430, 0x043D, 0x043D, 0x044B, 0x0435, 0x0020, 0x043C, 0x043E, 0x0434, 0x0435, 0x043B, 0x0438, 0x0020, 0x0052, 0x0056, 0x0054, 0x002C, 0x0020, 0x0448, 0x0442))
$H_WsTotal = (U @(0x0420, 0x0430, 0x0431, 0x043E, 0x0447, 0x0438, 0x0435, 0x0020, 0x043D, 0x0430, 0x0431, 0x043E, 0x0440, 0x044B, 0x002C, 0x0020, 0x0432, 0x0441, 0x0435, 0x0433, 0x043E))
$H_WsOwned = (U @(0x0420, 0x0430, 0x0431, 0x043E, 0x0447, 0x0438, 0x0435, 0x0020, 0x043D, 0x0430, 0x0431, 0x043E, 0x0440, 0x044B, 0x0020, 0x0437, 0x0430, 0x043D, 0x044F, 0x0442, 0x044B, 0x0020, 0x0434, 0x0440, 0x0443, 0x0433, 0x0438, 0x043C, 0x0438))
$H_Hub = (U @(0x041C, 0x043E, 0x0434, 0x0435, 0x043B, 0x044C, 0x002D, 0x0445, 0x0430, 0x0431))
$H_LinkName = (U @(0x0418, 0x043C, 0x044F, 0x0020, 0x0441, 0x0432, 0x044F, 0x0437, 0x0438))
$H_Path = (U @(0x041F, 0x0443, 0x0442, 0x044C))
$H_LinkStatus = (U @(0x0421, 0x0442, 0x0430, 0x0442, 0x0443, 0x0441, 0x0020, 0x0441, 0x0432, 0x044F, 0x0437, 0x0438))
$H_SavedPath = (U @(0x0421, 0x043E, 0x0445, 0x0440, 0x0430, 0x043D, 0x0435, 0x043D, 0x043D, 0x044B, 0x0439, 0x0020, 0x043F, 0x0443, 0x0442, 0x044C))
$H_PathType = (U @(0x0422, 0x0438, 0x043F, 0x0020, 0x043F, 0x0443, 0x0442, 0x044F))
$H_AbsPath = (U @(0x0410, 0x0431, 0x0441, 0x043E, 0x043B, 0x044E, 0x0442, 0x043D, 0x044B, 0x0439, 0x0020, 0x043F, 0x0443, 0x0442, 0x044C))
$H_RefType = (U @(0x0422, 0x0438, 0x043F, 0x0020, 0x0441, 0x0432, 0x044F, 0x0437, 0x0438))
$H_FileExists = (U @(0x0424, 0x0430, 0x0439, 0x043B, 0x0020, 0x0441, 0x0443, 0x0449, 0x0435, 0x0441, 0x0442, 0x0432, 0x0443, 0x0435, 0x0442))
$H_HostPath = (U @(0x041F, 0x0443, 0x0442, 0x044C, 0x0020, 0x0445, 0x043E, 0x0441, 0x0442, 0x0430))
$H_Timestamp = (U @(0x0414, 0x0430, 0x0442, 0x0430, 0x002F, 0x0432, 0x0440, 0x0435, 0x043C, 0x044F))
$H_LinkCount = (U @(0x0421, 0x0432, 0x044F, 0x0437, 0x0435, 0x0439))
$H_ModelFile = (U @(0x0418, 0x043C, 0x044F, 0x0020, 0x0444, 0x0430, 0x0439, 0x043B, 0x0430))
$H_Issues = (U @(0x041F, 0x0440, 0x043E, 0x0431, 0x043B, 0x0435, 0x043C, 0x044B))
$H_Version = (U @(0x0412, 0x0435, 0x0440, 0x0441, 0x0438, 0x044F, 0x0020, 0x0052, 0x0065, 0x0076, 0x0069, 0x0074))
$H_ScriptVer = (U @(0x0412, 0x0435, 0x0440, 0x0441, 0x0438, 0x044F, 0x0020, 0x0441, 0x043A, 0x0440, 0x0438, 0x043F, 0x0442, 0x0430))
$H_ModelPath = (U @(0x041F, 0x0443, 0x0442, 0x044C, 0x0020, 0x043A, 0x0020, 0x043C, 0x043E, 0x0434, 0x0435, 0x043B, 0x0438))
$SheetSummary = (U @(0x0421, 0x0432, 0x043E, 0x0434, 0x043A, 0x0430))
$SheetLinks = (U @(0x0421, 0x0432, 0x044F, 0x0437, 0x0438))
$SheetMetrics = (U @(0x041C, 0x0435, 0x0442, 0x0440, 0x0438, 0x043A, 0x0438))
$TitleOverview = (U @(0x041E, 0x0431, 0x0437, 0x043E, 0x0440))
$TitleContent = (U @(0x041A, 0x043E, 0x043D, 0x0442, 0x0435, 0x043D, 0x0442))
$TitleOther = (U @(0x041F, 0x0440, 0x043E, 0x0447, 0x0435, 0x0435))
$StLoaded = (U @(0x0417, 0x0430, 0x0433, 0x0440, 0x0443, 0x0436, 0x0435, 0x043D, 0x043E))
$StNotFound = (U @(0x041D, 0x0435, 0x0020, 0x043D, 0x0430, 0x0439, 0x0434, 0x0435, 0x043D, 0x043E))
$StUnloaded = (U @(0x0412, 0x044B, 0x0433, 0x0440, 0x0443, 0x0436, 0x0435, 0x043D, 0x043E))
$StYes = (U @(0x0434, 0x0430))
$StNo = (U @(0x043D, 0x0435, 0x0442))
$DefaultXlsxDir = Join-Path $env:USERPROFILE "Documents\doc\script\rbp\reports"

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
    if ($dir -and !(Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
    Add-Content -LiteralPath $Target -Value "$ts | $Text" -Encoding UTF8
}

function Write-Utf8NoBom([string]$Path, [string]$Content) {
    $enc = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($Path, $Content, $enc)
}

function Get-XlsxOutputDir([string]$Preferred) {
    if (-not [string]::IsNullOrWhiteSpace($Preferred)) {
        $p = Get-NativePath $Preferred
        if ($p) {
            if (!(Test-Path -LiteralPath $p)) { New-Item -ItemType Directory -Path $p -Force | Out-Null }
            return $p
        }
    }
    $fallback = $DefaultXlsxDir
    if (!(Test-Path -LiteralPath $fallback)) {
        New-Item -ItemType Directory -Path $fallback -Force | Out-Null
    }
    return $fallback
}

function Split-CsvLine([string]$Line) {
    $result = New-Object System.Collections.Generic.List[string]
    $sb = New-Object System.Text.StringBuilder
    $inQuotes = $false
    for ($i = 0; $i -lt $Line.Length; $i++) {
        $ch = $Line[$i]
        if ($ch -eq '"') {
            if ($inQuotes -and ($i + 1) -lt $Line.Length -and $Line[$i + 1] -eq '"') {
                [void]$sb.Append('"'); $i++
            }
            else { $inQuotes = -not $inQuotes }
        }
        elseif ($ch -eq ';' -and -not $inQuotes) {
            $result.Add($sb.ToString())
            [void]$sb.Clear()
        }
        else { [void]$sb.Append($ch) }
    }
    $result.Add($sb.ToString())
    return , $result.ToArray()
}

function ConvertFrom-SemicolonCsv([string]$Path) {
    $raw = [System.IO.File]::ReadAllText($Path, [System.Text.Encoding]::UTF8)
    $raw = $raw -replace "^\uFEFF", ""
    $lines = $raw -split "`r?`n"
    if (-not $lines -or $lines.Count -lt 1) { return @{ Headers = @(); Rows = @() } }

    $headers = Split-CsvLine $lines[0]
    $rows = @()
    for ($r = 1; $r -lt $lines.Count; $r++) {
        if ([string]::IsNullOrWhiteSpace($lines[$r])) { continue }
        $vals = Split-CsvLine $lines[$r]
        $obj = [ordered]@{}
        for ($c = 0; $c -lt $headers.Count; $c++) {
            $key = $headers[$c]
            $val = if ($c -lt $vals.Count) { $vals[$c] } else { "" }
            $obj[$key] = $val
        }
        $rows += [pscustomobject]$obj
    }
    return @{ Headers = $headers; Rows = $rows }
}

function Escape-Xml([string]$Text) {
    if ($null -eq $Text) { return "" }
    return ($Text -replace '&', '&amp;' -replace '<', '&lt;' -replace '>', '&gt;' -replace '"', '&quot;')
}

function Get-RowVal($row, [string]$Name) {
    if ($null -eq $row -or [string]::IsNullOrWhiteSpace($Name)) { return "" }
    if ($row.PSObject.Properties.Name -contains $Name) { return [string]$row.$Name }
    return ""
}

function New-SheetRowMeta {
    param($Row, [int]$OutlineLevel = 0, [bool]$Hidden = $false, [bool]$Collapsed = $false)
    return [pscustomobject]@{
        Row          = $Row
        OutlineLevel = $OutlineLevel
        Hidden       = $Hidden
        Collapsed    = $Collapsed
    }
}

function Get-ColWidth([string]$Header) {
    if ($Header -eq $H_SavedPath -or $Header -eq $H_AbsPath -or $Header -eq $H_HostPath -or $Header -eq $H_Path -or $Header -eq $H_ModelPath) { return 48 }
    if ($Header -eq $H_Hub -or $Header -eq $H_LinkName -or $Header -eq $H_ModelFile -or $Header -eq $H_Issues) { return 28 }
    if ($Header -eq $H_Timestamp) { return 18 }
    if ($Header -eq $H_LinkCount -or $Header -eq $H_Status) { return 10 }
    $n = if ($null -eq $Header) { 8 } else { $Header.Length }
    $w = [Math]::Max(10, $n + 2)
    if ($w -gt 22) { $w = 22 }
    return [Math]::Round([double]$w, 2)
}

function Pick-PresentHeaders($All, $Wanted) {
    $out = @()
    foreach ($w in @($Wanted)) {
        if ($All -contains $w) { $out += $w }
    }
    return @($out)
}

function Get-OverviewHeaders($All) {
    $wanted = @(
        $H_ModelFile, $H_Status, $H_Issues, $H_WarnCount,
        $H_RvtTotal, $H_RvtNf, $H_RvtUl, $H_LinkedRvt,
        $H_CadTotal, $H_SizeMb, $H_Version, $H_Timestamp
    )
    $picked = @(Pick-PresentHeaders $All $wanted)
    if ($picked.Count -gt 0) { return $picked }
    return @($All | Select-Object -First 12)
}

function Get-MetricSections($All, $Overview) {
    $used = @{}
    foreach ($h in @($Overview)) { if ($h) { $used[$h] = $true } }
    $used[$H_ScriptVer] = $true
    $rest = @($All | Where-Object { $_ -and -not $used.ContainsKey($_) })
    if ($rest.Count -eq 0) { return @() }
    $prefix = @(Pick-PresentHeaders $All @($H_ModelFile, $H_Status))
    $mid = [Math]::Max(1, [Math]::Ceiling($rest.Count / 2.0))
    $a = @($rest | Select-Object -First $mid)
    $b = @($rest | Select-Object -Skip $mid)
    $sections = @()
    $sections += ,@{ Title = $TitleContent; Headers = @($prefix + $a) }
    if ($b.Count -gt 0) {
        $sections += ,@{ Title = $TitleOther; Headers = @($prefix + $b) }
    }
    return @($sections)
}

function Test-RealLinkRow($row) {
    $name = Get-RowVal $row $H_LinkName
    $saved = Get-RowVal $row $H_SavedPath
    if ([string]::IsNullOrWhiteSpace($saved)) { $saved = Get-RowVal $row $H_Path }
    if (-not [string]::IsNullOrWhiteSpace($name)) { return $true }
    if ([string]::IsNullOrWhiteSpace($saved)) { return $false }
    if ($saved -eq "(path n/a)") { return $false }
    return $true
}

function Expand-LinksOutlineRows($Headers, $Rows) {
    $metaList = @()
    $rowArr = @($Rows)
    if ($rowArr.Count -eq 0) { return @() }

    $groupNames = @()
    $groupMap = @{}
    foreach ($row in $rowArr) {
        $m = Get-RowVal $row $H_Hub
        if ([string]::IsNullOrWhiteSpace($m)) { $m = "?" }
        if (-not $groupMap.ContainsKey($m)) {
            $groupNames += $m
            $groupMap[$m] = New-Object System.Collections.ArrayList
        }
        [void]$groupMap[$m].Add($row)
    }

    $keepOnParent = @($H_Hub, $H_LinkCount, $H_HostPath, $H_Timestamp)
    $clearOnChild = @($H_Hub, $H_LinkCount, $H_HostPath, $H_Timestamp, $H_AbsPath)
    foreach ($gName in $groupNames) {
        $all = @($groupMap[$gName].ToArray())
        $childRows = @($all | Where-Object { Test-RealLinkRow $_ })
        $first = $all[0]
        $parent = [ordered]@{}
        foreach ($h in $Headers) { $parent[$h] = "" }
        $parent[$H_Hub] = $gName
        if ($Headers -contains $H_LinkCount) { $parent[$H_LinkCount] = "$($childRows.Count)" }
        foreach ($h in $keepOnParent) {
            if ($h -eq $H_Hub -or $h -eq $H_LinkCount) { continue }
            if ($Headers -contains $h) {
                $val = Get-RowVal $first $h
                if ($val) { $parent[$h] = $val }
            }
        }
        $hasKids = $childRows.Count -gt 0
        $metaList += ,(New-SheetRowMeta -Row ([pscustomobject]$parent) -OutlineLevel 0 -Hidden $false -Collapsed $hasKids)
        foreach ($it in $childRows) {
            $child = [ordered]@{}
            foreach ($h in $Headers) {
                if ($clearOnChild -contains $h) { $child[$h] = ""; continue }
                $val = Get-RowVal $it $h
                if ([string]::IsNullOrWhiteSpace($val) -and $h -eq $H_SavedPath) {
                    $val = Get-RowVal $it $H_Path
                }
                $child[$h] = $val
            }
            $metaList += ,(New-SheetRowMeta -Row ([pscustomobject]$child) -OutlineLevel 1 -Hidden $true -Collapsed $false)
        }
    }
    return @($metaList)
}

function New-OpenXmlXlsx {
    param(
        [string]$XlsxPath,
        [object]$Summary,
        [object]$Links
    )

    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $tmp = Join-Path $env:TEMP ("health_openxml_" + [guid]::NewGuid().ToString("N"))
    New-Item -ItemType Directory -Path $tmp -Force | Out-Null
    New-Item -ItemType Directory -Path (Join-Path $tmp "_rels") -Force | Out-Null
    New-Item -ItemType Directory -Path (Join-Path $tmp "xl") -Force | Out-Null
    New-Item -ItemType Directory -Path (Join-Path $tmp "xl\_rels") -Force | Out-Null
    New-Item -ItemType Directory -Path (Join-Path $tmp "xl\worksheets") -Force | Out-Null

    $shared = New-Object System.Collections.Generic.List[string]
    $sharedIndex = @{}

    function Get-SharedIndex([string]$Text) {
        if ($null -eq $Text) { $Text = "" }
        if ($sharedIndex.ContainsKey($Text)) { return $sharedIndex[$Text] }
        $idx = $shared.Count
        $shared.Add($Text)
        $sharedIndex[$Text] = $idx
        return $idx
    }

    function ColName([int]$Index1Based) {
        # Must keep $n as [int]. [math]::Floor returns Double/Decimal;
        # [char](65 + Decimal) fails — broke XLSX once summary exceeded 26 columns (v1.4).
        $n = [int]$Index1Based
        $name = ""
        while ($n -gt 0) {
            $n = $n - 1
            $code = [int](65 + ($n % 26))
            $name = ([char]$code) + $name
            $n = [int][math]::Floor([double]$n / 26.0)
        }
        return $name
    }

    $stylesXml = @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2">
    <font><sz val="11"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><name val="Calibri"/></font>
  </fonts>
  <fills count="7">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFD9E1F2"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFC6EFCE"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFEB9C"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFC7CE"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFF2F2F2"/></patternFill></fill>
  </fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="8">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
    <xf numFmtId="0" fontId="0" fillId="3" borderId="0" xfId="0" applyFill="1"/>
    <xf numFmtId="0" fontId="0" fillId="4" borderId="0" xfId="0" applyFill="1"/>
    <xf numFmtId="0" fontId="0" fillId="5" borderId="0" xfId="0" applyFill="1"/>
    <xf numFmtId="1" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
    <xf numFmtId="0" fontId="1" fillId="6" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
    <xf numFmtId="1" fontId="1" fillId="6" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1"/>
  </cellXfs>
</styleSheet>
'@

    function Build-SheetXml {
        param(
            $Headers,
            $MetaRows,
            $StatusColName,
            $IntColNames,
            [bool]$EnableOutline,
            $FileExistsColName,
            [bool]$FreezeFirstCol,
            [bool]$EnableAutoFilter = $true
        )

        $sb = New-Object System.Text.StringBuilder
        [void]$sb.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
        [void]$sb.Append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">')
        if ($EnableOutline) {
            [void]$sb.Append('<sheetPr><outlinePr summaryBelow="0" summaryRight="0"/></sheetPr>')
        }
        if ($FreezeFirstCol) {
            [void]$sb.Append('<sheetViews><sheetView workbookViewId="0"><pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"/></sheetView></sheetViews>')
        }
        else {
            [void]$sb.Append('<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>')
        }
        if ($EnableOutline) {
            [void]$sb.Append('<sheetFormatPr defaultRowHeight="15" outlineLevelRow="1"/>')
        }
        if ($Headers -and $Headers.Count -gt 0) {
            [void]$sb.Append('<cols>')
            for ($c = 0; $c -lt $Headers.Count; $c++) {
                $w = Get-ColWidth ([string]$Headers[$c])
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
            for ($i = 0; $i -lt $Headers.Count; $i++) {
                if ($Headers[$i] -eq $StatusColName) { $statusIdx = $i; break }
            }
        }
        $existsIdx = -1
        if ($FileExistsColName) {
            for ($i = 0; $i -lt $Headers.Count; $i++) {
                if ($Headers[$i] -eq $FileExistsColName) { $existsIdx = $i; break }
            }
        }
        $intSet = @{}
        foreach ($n in $IntColNames) { if ($n) { $intSet[$n] = $true } }

        $normalized = @()
        foreach ($item in @($MetaRows)) {
            if ($null -eq $item) { continue }
            if ($item.PSObject.Properties.Name -contains "Row") { $normalized += $item }
            else { $normalized += (New-SheetRowMeta -Row $item) }
        }

        $rNum = 2
        foreach ($meta in $normalized) {
            $row = $meta.Row
            $ol = 0; if ($meta.PSObject.Properties.Name -contains "OutlineLevel") { $ol = [int]$meta.OutlineLevel }
            $hid = $false; if ($meta.PSObject.Properties.Name -contains "Hidden") { $hid = [bool]$meta.Hidden }
            $colp = $false; if ($meta.PSObject.Properties.Name -contains "Collapsed") { $colp = [bool]$meta.Collapsed }

            $rowAttrs = "r=`"$rNum`""
            if ($ol -gt 0) { $rowAttrs += " outlineLevel=`"$ol`"" }
            if ($hid) { $rowAttrs += ' hidden="1"' }
            if ($colp) { $rowAttrs += ' collapsed="1"' }
            [void]$sb.Append("<row $rowAttrs>")

            for ($c = 0; $c -lt $Headers.Count; $c++) {
                $h = $Headers[$c]
                $val = ""
                if ($row.PSObject.Properties.Name -contains $h) { $val = [string]$row.$h }
                $ref = (ColName ($c + 1)) + "$rNum"
                $style = 0
                if ($ol -eq 0 -and $EnableOutline) { $style = 6 }
                if ($c -eq $statusIdx) {
                    $up = $val.ToUpperInvariant()
                    if ($up -eq "GREEN" -or $up -eq "OK") { $style = 2 }
                    elseif ($up -eq "YELLOW" -or $val -eq $StUnloaded -or $up -match 'UNLOAD') { $style = 3 }
                    elseif ($up -eq "RED" -or $up -eq "ERROR" -or $up -eq "FAIL" -or $val -eq $StNotFound -or $up -match 'NOTFOUND|NOT FOUND') { $style = 4 }
                }
                if ($c -eq $existsIdx -and $ol -gt 0) {
                    if ($val -eq $StNo) { $style = 4 }
                }
                $asNum = $null
                if ($intSet.ContainsKey($h) -and $val -ne "" -and [double]::TryParse(($val.Replace(',', '.')), [ref]$asNum)) {
                    $num = [math]::Round($asNum)
                    $nStyle = 5
                    if ($ol -eq 0 -and $EnableOutline) { $nStyle = 7 }
                    [void]$sb.Append("<c r=`"$ref`" s=`"$nStyle`"><v>$num</v></c>")
                }
                else {
                    $idx = Get-SharedIndex $val
                    if ($style -gt 0) {
                        [void]$sb.Append("<c r=`"$ref`" t=`"s`" s=`"$style`"><v>$idx</v></c>")
                    }
                    else {
                        [void]$sb.Append("<c r=`"$ref`" t=`"s`"><v>$idx</v></c>")
                    }
                }
            }
            [void]$sb.Append('</row>')
            $rNum++
        }

        [void]$sb.Append('</sheetData>')
        $dataLast = [Math]::Max(1, $rNum - 1)
        if ($EnableAutoFilter -and $Headers.Count -gt 0) {
            $lastCol = ColName $Headers.Count
            [void]$sb.Append("<autoFilter ref=`"A1:${lastCol}${dataLast}`"/>")
        }
        [void]$sb.Append('</worksheet>')
        return $sb.ToString()
    }

    function Build-StackedSheetXml {
        param($Sections, $StatusColName, $IntColNames)

        $sb = New-Object System.Text.StringBuilder
        [void]$sb.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
        [void]$sb.Append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">')
        [void]$sb.Append('<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>')

        $maxCols = 1
        foreach ($sec in @($Sections)) {
            if ($sec.Headers -and $sec.Headers.Count -gt $maxCols) { $maxCols = $sec.Headers.Count }
        }
        [void]$sb.Append('<cols>')
        $firstSec = @($Sections)[0]
        $widthHeaders = if ($firstSec) { @($firstSec.Headers) } else { @() }
        for ($c = 0; $c -lt $maxCols; $c++) {
            $h = if ($c -lt $widthHeaders.Count) { [string]$widthHeaders[$c] } else { "" }
            $w = Get-ColWidth $h
            $ci = $c + 1
            [void]$sb.Append("<col min=`"$ci`" max=`"$ci`" width=`"$w`" customWidth=`"1`"/>")
        }
        [void]$sb.Append('</cols><sheetData>')

        $intSet = @{}
        foreach ($n in $IntColNames) { if ($n) { $intSet[$n] = $true } }
        $rNum = 1
        $secIndex = 0
        foreach ($sec in @($Sections)) {
            $headers = @($sec.Headers)
            $rows = @($sec.Rows)
            if ($secIndex -gt 0) {
                [void]$sb.Append("<row r=`"$rNum`"/>"); $rNum++
            }
            [void]$sb.Append("<row r=`"$rNum`" ht=`"20`">")
            $titleIdx = Get-SharedIndex ([string]$sec.Title)
            [void]$sb.Append("<c r=`"A$rNum`" t=`"s`" s=`"1`"><v>$titleIdx</v></c>")
            [void]$sb.Append("</row>"); $rNum++

            [void]$sb.Append("<row r=`"$rNum`" ht=`"22`">")
            for ($c = 0; $c -lt $headers.Count; $c++) {
                $ref = (ColName ($c + 1)) + "$rNum"
                $idx = Get-SharedIndex ([string]$headers[$c])
                [void]$sb.Append("<c r=`"$ref`" t=`"s`" s=`"1`"><v>$idx</v></c>")
            }
            [void]$sb.Append("</row>"); $rNum++

            $statusIdx = -1
            if ($StatusColName) {
                for ($i = 0; $i -lt $headers.Count; $i++) {
                    if ($headers[$i] -eq $StatusColName) { $statusIdx = $i; break }
                }
            }
            foreach ($row in $rows) {
                [void]$sb.Append("<row r=`"$rNum`">")
                for ($c = 0; $c -lt $headers.Count; $c++) {
                    $h = $headers[$c]
                    $val = Get-RowVal $row $h
                    $ref = (ColName ($c + 1)) + "$rNum"
                    $style = 0
                    if ($c -eq $statusIdx) {
                        $up = $val.ToUpperInvariant()
                        if ($up -eq "GREEN" -or $up -eq "OK") { $style = 2 }
                        elseif ($up -eq "YELLOW") { $style = 3 }
                        elseif ($up -eq "RED" -or $up -eq "ERROR" -or $up -eq "FAIL") { $style = 4 }
                    }
                    $asNum = $null
                    if ($intSet.ContainsKey($h) -and $val -ne "" -and [double]::TryParse(($val.Replace(',', '.')), [ref]$asNum)) {
                        $num = [math]::Round($asNum)
                        [void]$sb.Append("<c r=`"$ref`" s=`"5`"><v>$num</v></c>")
                    }
                    else {
                        $idx = Get-SharedIndex $val
                        if ($style -gt 0) { [void]$sb.Append("<c r=`"$ref`" t=`"s`" s=`"$style`"><v>$idx</v></c>") }
                        else { [void]$sb.Append("<c r=`"$ref`" t=`"s`"><v>$idx</v></c>") }
                    }
                }
                [void]$sb.Append("</row>"); $rNum++
            }
            $secIndex++
        }
        [void]$sb.Append('</sheetData></worksheet>')
        return $sb.ToString()
    }

    $allSumHeaders = @($Summary.Headers)
    $sumRows = @($Summary.Rows)
    $overviewHeaders = @(Get-OverviewHeaders $allSumHeaders)
    $sumMeta = @($sumRows | ForEach-Object { New-SheetRowMeta -Row $_ })

    $preferredLinkHeaders = @(
        $H_Hub, $H_LinkCount, $H_LinkName, $H_SavedPath, $H_PathType,
        $H_LinkStatus, $H_RefType, $H_FileExists, $H_HostPath, $H_Timestamp
    )
    $linkHeaders = $preferredLinkHeaders
    $linkRows = @()
    if ($Links -and $Links.Rows) {
        $linkRows = @($Links.Rows)
        $available = @($Links.Headers)
        if ($available -notcontains $H_SavedPath -and $available -contains $H_Path) {
            $linkHeaders = @(
                $H_Hub, $H_LinkCount, $H_LinkName, $H_Path, $H_LinkStatus
            )
        }
    }

    $intCols = @(
        $H_SizeMb, $H_WarnCount, $H_RvtTotal, $H_RvtNf, $H_RvtUl,
        $H_CadTotal, $H_CadNf, $H_LinkedRvt, $H_WsTotal, $H_WsOwned
    )

    $sheet1 = Build-SheetXml -Headers $overviewHeaders -MetaRows $sumMeta -StatusColName $H_Status `
        -IntColNames $intCols -EnableOutline $false -FileExistsColName $null -FreezeFirstCol $true
    $linkMeta = Expand-LinksOutlineRows $linkHeaders $linkRows
    $sheet2 = Build-SheetXml -Headers $linkHeaders -MetaRows $linkMeta -StatusColName $H_LinkStatus `
        -IntColNames @($H_LinkCount) -EnableOutline $true -FileExistsColName $H_FileExists -FreezeFirstCol $true

    $metricSecsRaw = @(Get-MetricSections $allSumHeaders $overviewHeaders)
    $metricSections = @()
    foreach ($sec in $metricSecsRaw) {
        $metricSections += ,@{ Title = $sec.Title; Headers = @($sec.Headers); Rows = $sumRows }
    }
    if ($metricSections.Count -eq 0) {
        $metricSections = @(@{ Title = $SheetMetrics; Headers = @($allSumHeaders); Rows = $sumRows })
    }
    $sheet3 = Build-StackedSheetXml -Sections $metricSections -StatusColName $H_Status -IntColNames $intCols

    $sst = New-Object System.Text.StringBuilder
    [void]$sst.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    [void]$sst.Append('<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="' + $shared.Count + '" uniqueCount="' + $shared.Count + '">')
    foreach ($s in $shared) {
        [void]$sst.Append('<si><t xml:space="preserve">' + (Escape-Xml $s) + '</t></si>')
    }
    [void]$sst.Append('</sst>')

    Write-Utf8NoBom (Join-Path $tmp "[Content_Types].xml") @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>
'@

    Write-Utf8NoBom (Join-Path $tmp "_rels\.rels") @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
'@

    Write-Utf8NoBom (Join-Path $tmp "xl\_rels\workbook.xml.rels") @'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>
'@

    $wbXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="$SheetSummary" sheetId="1" r:id="rId1"/>
    <sheet name="$SheetLinks" sheetId="2" r:id="rId2"/>
    <sheet name="$SheetMetrics" sheetId="3" r:id="rId3"/>
  </sheets>
</workbook>
"@
    Write-Utf8NoBom (Join-Path $tmp "xl\workbook.xml") $wbXml
    Write-Utf8NoBom (Join-Path $tmp "xl\styles.xml") $stylesXml
    Write-Utf8NoBom (Join-Path $tmp "xl\sharedStrings.xml") $sst.ToString()
    Write-Utf8NoBom (Join-Path $tmp "xl\worksheets\sheet1.xml") $sheet1
    Write-Utf8NoBom (Join-Path $tmp "xl\worksheets\sheet2.xml") $sheet2
    Write-Utf8NoBom (Join-Path $tmp "xl\worksheets\sheet3.xml") $sheet3

    if (Test-Path -LiteralPath $XlsxPath) {
        Remove-Item -LiteralPath $XlsxPath -Force
    }
    [System.IO.Compression.ZipFile]::CreateFromDirectory($tmp, $XlsxPath)
    Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue
}

# -------- main --------
if ([string]::IsNullOrWhiteSpace($CsvPath)) {
    $csvDir = Join-Path $env:USERPROFILE "Documents\doc\script\rbp\reports"
    if (!(Test-Path -LiteralPath $csvDir)) {
        New-Item -ItemType Directory -Path $csvDir -Force | Out-Null
    }
    $latestCsv = Get-ChildItem -LiteralPath $csvDir -Filter "health_check_*.csv" -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -notmatch '_links\.csv$' } |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1
    if (-not $latestCsv) {
        $latestCsv = Get-ChildItem -LiteralPath $csvDir -Filter "health_*.csv" -File -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -notmatch '_links\.csv$' } |
            Sort-Object LastWriteTime -Descending |
            Select-Object -First 1
    }
    if (-not $latestCsv) { throw "No health_check_*.csv found in $csvDir" }
    $CsvPath = $latestCsv.FullName
}

$CsvPath = Get-NativePath $CsvPath
if (!(Test-Path -LiteralPath $CsvPath)) { throw "CSV not found: $CsvPath" }

if ([string]::IsNullOrWhiteSpace($LinksCsvPath)) {
    $candidate = [System.IO.Path]::ChangeExtension($CsvPath, $null).TrimEnd('.') + "_links.csv"
    if (Test-Path -LiteralPath $candidate) { $LinksCsvPath = $candidate }
}
else {
    $LinksCsvPath = Get-NativePath $LinksCsvPath
}

$outDir = Get-XlsxOutputDir $OutputDir
if ([string]::IsNullOrWhiteSpace($outDir)) {
    $outDir = $DefaultXlsxDir
}
if (!(Test-Path -LiteralPath $outDir)) {
    New-Item -ItemType Directory -Path $outDir -Force | Out-Null
}

$xlsxName = [System.IO.Path]::GetFileNameWithoutExtension($CsvPath) + ".xlsx"
$xlsxPath = Join-Path $outDir $xlsxName
$logTarget = if ([string]::IsNullOrWhiteSpace($LogPath)) {
    Join-Path $outDir "health_check_xlsx_export.log"
} else { Get-NativePath $LogPath }

try {
    Write-LogLine "OpenXML export started. Csv=$CsvPath Xlsx=$xlsxPath" $logTarget
    $summary = ConvertFrom-SemicolonCsv -Path $CsvPath
    $links = $null
    if ($LinksCsvPath -and (Test-Path -LiteralPath $LinksCsvPath)) {
        $links = ConvertFrom-SemicolonCsv -Path $LinksCsvPath
    }
    New-OpenXmlXlsx -XlsxPath $xlsxPath -Summary $summary -Links $links
    Write-LogLine "OpenXML export success. Saved=$xlsxPath" $logTarget
    Write-Output "Created: $xlsxPath"
}
catch {
    Write-LogLine ("OpenXML export failed: " + $_.Exception.Message) $logTarget
    throw
}
