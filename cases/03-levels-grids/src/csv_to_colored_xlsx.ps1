param(
    [Parameter(Mandatory = $false)]
    [string]$CsvPath,
    [Parameter(Mandatory = $false)]
    [string]$DetailsCsvPath,
    [Parameter(Mandatory = $false)]
    [string]$OutputDir,
    [Parameter(Mandatory = $false)]
    [string]$LogPath
)

# ASCII-only source. Russian UI strings via Unicode code points.
$ErrorActionPreference = "Stop"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

function U {
    param([int[]]$Codes)
    return (-join ($Codes | ForEach-Object { [char]([int]$_) }))
}

$H_Status = U @(0x0421, 0x0442, 0x0430, 0x0442, 0x0443, 0x0441) # Status
$SheetSummary = U @(0x0421, 0x0432, 0x043E, 0x0434, 0x043A, 0x0430) # Svodka
$SheetDetails = U @(0x0414, 0x0435, 0x0442, 0x0430, 0x043B, 0x0438) # Detali

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
        return (Get-NativePath $Preferred)
    }
    $cfg = Join-Path $scriptDir "xlsx_out_path.cfg"
    if (Test-Path -LiteralPath $cfg) {
        $line = Get-Content -LiteralPath $cfg -ErrorAction SilentlyContinue |
            Where-Object { $_ -and -not $_.StartsWith("#") } |
            Select-Object -First 1
        if ($line) { return (Get-NativePath $line.Trim()) }
    }
    return $null
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

function New-OpenXmlXlsx {
    param(
        [string]$XlsxPath,
        [object]$Summary,
        [object]$Details
    )

    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $tmp = Join-Path $env:TEMP ("lg_openxml_" + [guid]::NewGuid().ToString("N"))
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
  <fills count="6">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFD9E1F2"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFC6EFCE"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFEB9C"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFC7CE"/></patternFill></fill>
  </fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="6">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
    <xf numFmtId="0" fontId="0" fillId="3" borderId="0" xfId="0" applyFill="1"/>
    <xf numFmtId="0" fontId="0" fillId="4" borderId="0" xfId="0" applyFill="1"/>
    <xf numFmtId="0" fontId="0" fillId="5" borderId="0" xfId="0" applyFill="1"/>
    <xf numFmtId="1" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
  </cellXfs>
</styleSheet>
'@

    function Build-SheetXml($Headers, $Rows, $StatusColName, $IntColNames) {
        $sb = New-Object System.Text.StringBuilder
        [void]$sb.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
        [void]$sb.Append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>')

        [void]$sb.Append('<row r="1">')
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
        $intSet = @{}
        foreach ($n in $IntColNames) { if ($n) { $intSet[$n] = $true } }

        $rNum = 2
        foreach ($row in $Rows) {
            [void]$sb.Append("<row r=`"$rNum`">")
            for ($c = 0; $c -lt $Headers.Count; $c++) {
                $h = $Headers[$c]
                $val = ""
                if ($row.PSObject.Properties.Name -contains $h) { $val = [string]$row.$h }
                $ref = (ColName ($c + 1)) + "$rNum"
                $style = 0
                if ($c -eq $statusIdx) {
                    switch ($val.ToUpperInvariant()) {
                        "GREEN"  { $style = 2 }
                        "YELLOW" { $style = 3 }
                        "RED"    { $style = 4 }
                        "ERROR"  { $style = 4 }
                    }
                }
                $asNum = $null
                if ($intSet.ContainsKey($h) -and $val -ne "" -and [double]::TryParse(($val.Replace(',', '.')), [ref]$asNum)) {
                    $num = [math]::Round($asNum)
                    [void]$sb.Append("<c r=`"$ref`" s=`"5`"><v>$num</v></c>")
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

        [void]$sb.Append('</sheetData></worksheet>')
        return $sb.ToString()
    }

    $sumHeaders = @($Summary.Headers)
    $sumRows = @($Summary.Rows)
    $detHeaders = @()
    $detRows = @()
    if ($Details -and $Details.Headers) {
        $detHeaders = @($Details.Headers)
        $detRows = @($Details.Rows)
    }
    if ($detHeaders.Count -eq 0) {
        $detHeaders = @("Host", "Status")
    }

    $sheet1 = Build-SheetXml $sumHeaders $sumRows $H_Status @()
    $sheet2 = Build-SheetXml $detHeaders $detRows $H_Status @()

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
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>
'@

    $wbXml = @"
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="$SheetSummary" sheetId="1" r:id="rId1"/>
    <sheet name="$SheetDetails" sheetId="2" r:id="rId2"/>
  </sheets>
</workbook>
"@
    Write-Utf8NoBom (Join-Path $tmp "xl\workbook.xml") $wbXml
    Write-Utf8NoBom (Join-Path $tmp "xl\styles.xml") $stylesXml
    Write-Utf8NoBom (Join-Path $tmp "xl\sharedStrings.xml") $sst.ToString()
    Write-Utf8NoBom (Join-Path $tmp "xl\worksheets\sheet1.xml") $sheet1
    Write-Utf8NoBom (Join-Path $tmp "xl\worksheets\sheet2.xml") $sheet2

    if (Test-Path -LiteralPath $XlsxPath) {
        Remove-Item -LiteralPath $XlsxPath -Force
    }
    [System.IO.Compression.ZipFile]::CreateFromDirectory($tmp, $XlsxPath)
    Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue
}

# -------- main --------
if ([string]::IsNullOrWhiteSpace($CsvPath)) {
    $csvDir = Join-Path $scriptDir "reports\cvc"
    $latestCsv = Get-ChildItem -LiteralPath $csvDir -Filter "lg_*.csv" -File -ErrorAction Stop |
        Where-Object { $_.Name -notmatch '_details\.csv$' } |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1
    if (-not $latestCsv) { throw "No lg_*.csv found" }
    $CsvPath = $latestCsv.FullName
}

$CsvPath = Get-NativePath $CsvPath
if (!(Test-Path -LiteralPath $CsvPath)) { throw "CSV not found: $CsvPath" }

if ([string]::IsNullOrWhiteSpace($DetailsCsvPath)) {
    $candidate = [System.IO.Path]::ChangeExtension($CsvPath, $null).TrimEnd('.') + "_details.csv"
    if (Test-Path -LiteralPath $candidate) { $DetailsCsvPath = $candidate }
}
else {
    $DetailsCsvPath = Get-NativePath $DetailsCsvPath
}

$outDir = Get-XlsxOutputDir $OutputDir
if ([string]::IsNullOrWhiteSpace($outDir)) {
    throw "XLSX output folder is not set. Run reset_report_session.cmd / choose_xlsx_path.cmd"
}
if (!(Test-Path -LiteralPath $outDir)) {
    New-Item -ItemType Directory -Path $outDir -Force | Out-Null
}

$xlsxName = [System.IO.Path]::GetFileNameWithoutExtension($CsvPath) + ".xlsx"
$xlsxPath = Join-Path $outDir $xlsxName
$logTarget = if ([string]::IsNullOrWhiteSpace($LogPath)) {
    Join-Path $outDir "xlsx_export.log"
} else { Get-NativePath $LogPath }

try {
    Write-LogLine "OpenXML export started. Csv=$CsvPath Xlsx=$xlsxPath" $logTarget
    $summary = ConvertFrom-SemicolonCsv -Path $CsvPath
    $details = $null
    if ($DetailsCsvPath -and (Test-Path -LiteralPath $DetailsCsvPath)) {
        $details = ConvertFrom-SemicolonCsv -Path $DetailsCsvPath
    }
    New-OpenXmlXlsx -XlsxPath $xlsxPath -Summary $summary -Details $details
    Write-LogLine "OpenXML export success. Saved=$xlsxPath" $logTarget
    Write-Output "Created: $xlsxPath"
}
catch {
    Write-LogLine ("OpenXML export failed: " + $_.Exception.Message) $logTarget
    throw
}
