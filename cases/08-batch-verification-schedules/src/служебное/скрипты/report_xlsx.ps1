# UTF-8 WITH BOM required (PowerShell 5.1).
# Отчёт прогона вставки (Сводка, Вставлено, Параметры, Пропуски) в .xlsx БЕЗ Excel: файл пишется напрямую как zip с XML (Open XML) средствами .NET
# (System.IO.Compression). Excel, COM и внешние модули не нужны, временные книги не создаются.
#
# Write-ReportXlsx -Path <файл.xlsx> -Sheets @( @{ Name; Headers; Rows; Widths; CountCols; ParentRows; Groups } , ... )
#   Name       — имя листа (до 31 символа);  Headers — заголовки столбцов;  Rows — массивы значений (число — число, остальное — текст)
#   Widths     — ширина столбцов (в символах);  CountCols — номера столбцов (с 1) с форматом числа «0»
#   ParentRows — номера строк листа (с 1, заголовок — строка 1) с выделением «строка модели»
#   Groups     — @{ Parent = номер строки; ChildCount = число строк под ней }: дочерние строки сворачиваются (outline), как «+» в Excel
# Первая строка закреплена, на заголовках автофильтр.

function Remove-XmlInvalidChars([string]$Text) {
    if ($null -eq $Text) { return "" }
    return ($Text -replace '[\x00-\x08\x0B\x0C\x0E-\x1F￾￿]', '')
}

function ConvertTo-XmlText([string]$Text) {
    return [System.Security.SecurityElement]::Escape((Remove-XmlInvalidChars $Text))
}

function Get-XlsxColumnName([int]$Index) {
    $n = $Index; $name = ""
    while ($n -gt 0) { $n--; $name = [string][char](65 + ($n % 26)) + $name; $n = [int][Math]::Floor($n / 26) }
    return $name
}

# styles.xml: 0 обычный; 1 заголовок (жирный, заливка D9E1F2, по центру); 2 строка модели (жирный, заливка F2F2F2); 3 число «0»; 4 строка модели + число
function Get-XlsxStylesXml {
    $ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="{0}">' -f $ns) +
        '<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts>' +
        '<fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>' +
        '<fill><patternFill patternType="solid"><fgColor rgb="FFD9E1F2"/><bgColor indexed="64"/></patternFill></fill>' +
        '<fill><patternFill patternType="solid"><fgColor rgb="FFF2F2F2"/><bgColor indexed="64"/></patternFill></fill></fills>' +
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
        '<cellXfs count="5">' +
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
        '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center"/></xf>' +
        '<xf numFmtId="0" fontId="1" fillId="3" borderId="0" xfId="0" applyFont="1" applyFill="1"/>' +
        '<xf numFmtId="1" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' +
        '<xf numFmtId="1" fontId="1" fillId="3" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1"/>' +
        '</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>'
}

function Test-XlsxNumber($Value) {
    return ($Value -is [int] -or $Value -is [long] -or $Value -is [double] -or $Value -is [decimal])
}

function Get-XlsxSheetXml($Sheet, [bool]$Selected) {
    $ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    $headers = @($Sheet.Headers)
    $cols = $headers.Count
    $rows = @($Sheet.Rows)
    $parents = @{}
    foreach ($p in @($Sheet.ParentRows)) { $parents[[int]$p] = $true }
    $countCols = @{}
    foreach ($c in @($Sheet.CountCols)) { $countCols[[int]$c] = $true }
    $child = @{}                 # номер строки -> $true: дочерняя строка группы (outline)
    $collapsedParent = @{}
    foreach ($g in @($Sheet.Groups)) {
        if ($null -eq $g) { continue }
        $collapsedParent[[int]$g.Parent] = $true
        for ($i = 1; $i -le [int]$g.ChildCount; $i++) { $child[[int]$g.Parent + $i] = $true }
    }
    $last = 1 + $rows.Count
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    [void]$sb.Append(('<worksheet xmlns="{0}">' -f $ns))
    if ($child.Count -gt 0) { [void]$sb.Append('<sheetPr><outlinePr summaryBelow="0"/></sheetPr>') }
    [void]$sb.Append(('<dimension ref="A1:{0}{1}"/>' -f (Get-XlsxColumnName ([Math]::Max(1, $cols))), $last))
    [void]$sb.Append(('<sheetViews><sheetView workbookViewId="0"{0}><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft"/></sheetView></sheetViews>' -f $(if ($Selected) { ' tabSelected="1"' } else { '' })))
    [void]$sb.Append(('<sheetFormatPr defaultRowHeight="15"{0}/>' -f $(if ($child.Count -gt 0) { ' outlineLevelRow="1"' } else { '' })))
    $widths = @($Sheet.Widths)
    if ($widths.Count -gt 0) {
        [void]$sb.Append('<cols>')
        for ($i = 0; $i -lt $widths.Count; $i++) {
            [void]$sb.Append(('<col min="{0}" max="{0}" width="{1}" customWidth="1"/>' -f ($i + 1), ([string]$widths[$i]).Replace(',', '.')))
        }
        [void]$sb.Append('</cols>')
    }
    [void]$sb.Append('<sheetData>')
    # строка 1 — заголовки
    [void]$sb.Append('<row r="1">')
    for ($c = 0; $c -lt $cols; $c++) {
        [void]$sb.Append(('<c r="{0}1" s="1" t="inlineStr"><is><t xml:space="preserve">{1}</t></is></c>' -f (Get-XlsxColumnName ($c + 1)), (ConvertTo-XmlText ([string]$headers[$c]))))
    }
    [void]$sb.Append('</row>')
    for ($r = 0; $r -lt $rows.Count; $r++) {
        $rn = $r + 2
        $row = @($rows[$r])
        $isParent = $parents.ContainsKey($rn)
        $attr = ""
        if ($child.ContainsKey($rn)) { $attr = ' outlineLevel="1" hidden="1"' }
        elseif ($collapsedParent.ContainsKey($rn)) { $attr = ' collapsed="1"' }
        [void]$sb.Append(('<row r="{0}"{1}>' -f $rn, $attr))
        for ($c = 0; $c -lt $cols; $c++) {
            $v = $null
            if ($c -lt $row.Count) { $v = $row[$c] }
            $ref = "{0}{1}" -f (Get-XlsxColumnName ($c + 1)), $rn
            $isNum = Test-XlsxNumber $v
            $style = 0
            if ($isParent -and $countCols.ContainsKey($c + 1)) { $style = 4 } elseif ($isParent) { $style = 2 } elseif ($countCols.ContainsKey($c + 1) -and $isNum) { $style = 3 }
            if ($null -eq $v -or ($v -is [string] -and $v -eq "")) {
                if ($style -ne 0) { [void]$sb.Append(('<c r="{0}" s="{1}"/>' -f $ref, $style)) }
            }
            elseif ($isNum) {
                [void]$sb.Append(('<c r="{0}" s="{1}"><v>{2}</v></c>' -f $ref, $style, ([System.Convert]::ToString($v, [System.Globalization.CultureInfo]::InvariantCulture))))
            }
            else {
                [void]$sb.Append(('<c r="{0}" s="{1}" t="inlineStr"><is><t xml:space="preserve">{2}</t></is></c>' -f $ref, $style, (ConvertTo-XmlText ([string]$v))))
            }
        }
        [void]$sb.Append('</row>')
    }
    [void]$sb.Append('</sheetData>')
    if ($cols -gt 0) { [void]$sb.Append(('<autoFilter ref="A1:{0}{1}"/>' -f (Get-XlsxColumnName $cols), $last)) }
    [void]$sb.Append('</worksheet>')
    return $sb.ToString()
}

function Write-ReportXlsx {
    param([string]$Path, $Sheets)
    $sheets = @($Sheets)
    if ($sheets.Count -lt 1) { throw "Нет листов для записи xlsx." }
    Add-Type -AssemblyName System.IO.Compression -ErrorAction SilentlyContinue
    $rel = 'http://schemas.openxmlformats.org/package/2006/relationships'
    $docRel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    $sheetOverrides = New-Object System.Text.StringBuilder
    $wbSheets = New-Object System.Text.StringBuilder
    $wbRels = New-Object System.Text.StringBuilder
    $names = @{}
    for ($i = 0; $i -lt $sheets.Count; $i++) {
        $n = $i + 1
        $name = (Remove-XmlInvalidChars ([string]$sheets[$i].Name)) -replace '[\\/\?\*\[\]:]', '_'
        if ($name.Length -gt 31) { $name = $name.Substring(0, 31) }
        if (-not $name -or $names.ContainsKey($name.ToLowerInvariant())) { $name = "Лист$n" }
        $names[$name.ToLowerInvariant()] = $true
        [void]$sheetOverrides.Append(('<Override PartName="/xl/worksheets/sheet{0}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' -f $n))
        [void]$wbSheets.Append(('<sheet name="{0}" sheetId="{1}" r:id="rId{1}"/>' -f (ConvertTo-XmlText $name), $n))
        [void]$wbRels.Append(('<Relationship Id="rId{0}" Type="{1}/worksheet" Target="worksheets/sheet{0}.xml"/>' -f $n, $docRel))
    }
    [void]$wbRels.Append(('<Relationship Id="rId{0}" Type="{1}/styles" Target="styles.xml"/>' -f ($sheets.Count + 1), $docRel))
    $parts = New-Object 'System.Collections.Generic.List[object]'
    $parts.Add(@('[Content_Types].xml', ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>{0}</Types>' -f $sheetOverrides.ToString())))
    $parts.Add(@('_rels/.rels', ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="{0}"><Relationship Id="rId1" Type="{1}/officeDocument" Target="xl/workbook.xml"/></Relationships>' -f $rel, $docRel)))
    $parts.Add(@('xl/workbook.xml', ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="{0}"><sheets>{1}</sheets></workbook>' -f $docRel, $wbSheets.ToString())))
    $parts.Add(@('xl/_rels/workbook.xml.rels', ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="{0}">{1}</Relationships>' -f $rel, $wbRels.ToString())))
    $parts.Add(@('xl/styles.xml', (Get-XlsxStylesXml)))
    for ($i = 0; $i -lt $sheets.Count; $i++) {
        $parts.Add(@(('xl/worksheets/sheet{0}.xml' -f ($i + 1)), (Get-XlsxSheetXml $sheets[$i] ($i -eq 0))))
    }
    $folder = Split-Path -Parent $Path
    if ($folder -and -not (Test-Path -LiteralPath $folder)) { New-Item -ItemType Directory -Path $folder -Force | Out-Null }
    $tmp = $Path + ".tmp"
    if (Test-Path -LiteralPath $tmp) { Remove-Item -LiteralPath $tmp -Force }
    $utf8 = New-Object System.Text.UTF8Encoding $false
    try {
        $fs = [System.IO.File]::Open($tmp, [System.IO.FileMode]::CreateNew)
        try {
            $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Create, $false)
            try {
                foreach ($part in $parts) {
                    $entry = $zip.CreateEntry([string]$part[0])
                    $es = $entry.Open()
                    try { $bytes = $utf8.GetBytes([string]$part[1]); $es.Write($bytes, 0, $bytes.Length) } finally { $es.Dispose() }
                }
            }
            finally { $zip.Dispose() }
        }
        finally { $fs.Dispose() }
        if (Test-Path -LiteralPath $Path) { [System.IO.File]::Delete($Path) }
        [System.IO.File]::Move($tmp, $Path)
    }
    catch {
        if (Test-Path -LiteralPath $tmp) { Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue }
        throw
    }
    return $Path
}
