# UTF-8 WITH BOM required (PowerShell 5.1).
# Чтение листов Excel COM → JSON для Google (без запуска push_google.ps1 как скрипта).

# Листы Google: «Сводка» и названия разделов в порядке разделы.txt. Нет файла — встроенный список.
try {
    if (-not (Get-Command Get-BvsGoogleSheetNames -ErrorAction SilentlyContinue)) {
        $__bvsSections = if ($PSScriptRoot) { Join-Path $PSScriptRoot "bvs_sections.ps1" } else { Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Path) "bvs_sections.ps1" }
        if (Test-Path -LiteralPath $__bvsSections) { . $__bvsSections }
    }
}
catch { Write-Warning ("разделы: " + $_.Exception.Message) }
if (Get-Command Get-BvsGoogleSheetNames -ErrorAction SilentlyContinue) {
    $script:BvsGooglePushSheetNames = @(Get-BvsGoogleSheetNames)
}
else {
    Write-Warning "встроенный список листов Google (bvs_sections.ps1 не подключён)"
    $script:BvsGooglePushSheetNames = @("Сводка", "ЭОМ", "СС", "ОВ", "ПД", "ПТ", "ВК", "ОТ", "ХС", "ИТП", "АВТ", "ХЦ")
}
$script:BvsGoogleHideSheetNames = @("Динамика")

function ConvertTo-PushRowArrays($Rows) {
    $out = New-Object System.Collections.Generic.List[object]
    if ($null -eq $Rows) { return ,@($out.ToArray()) }
    foreach ($row in @($Rows)) {
        if ($null -eq $row) {
            [void]$out.Add(@())
            continue
        }
        if ($row -is [string]) {
            [void]$out.Add(@([string]$row))
            continue
        }
        if ($row -is [System.Array]) {
            [void]$out.Add(@($row | ForEach-Object { if ($null -eq $_) { "" } else { [string]$_ } }))
            continue
        }
        $line = New-Object System.Collections.Generic.List[string]
        foreach ($p in $row.PSObject.Properties) {
            [void]$line.Add([string]$p.Value)
        }
        [void]$out.Add(@($line.ToArray()))
    }
    # запятая: иначе массив из одной строки «разворачивается» в плоский и в Google уходит rows=["текст"]
    return ,@($out.ToArray())
}

function Pad-PushRowsToRectangle {
    param([object[]]$Rows)
    $maxCols = 0
    foreach ($row in @($Rows)) {
        if ($null -eq $row) { continue }
        $n = @($row).Count
        if ($n -gt $maxCols) { $maxCols = $n }
    }
    if ($maxCols -lt 1) { return ,@($Rows) }
    $out = New-Object System.Collections.Generic.List[object]
    foreach ($row in @($Rows)) {
        $line = New-Object System.Collections.Generic.List[string]
        $arr = @($row)
        for ($c = 0; $c -lt $maxCols; $c++) {
            if ($c -lt $arr.Count -and $null -ne $arr[$c]) { [void]$line.Add([string]$arr[$c]) }
            else { [void]$line.Add("") }
        }
        [void]$out.Add(@($line.ToArray()))
    }
    return ,@($out.ToArray())
}

function Normalize-PayloadSheet($Sheet) {
    $rows = Pad-PushRowsToRectangle (ConvertTo-PushRowArrays $Sheet.rows)
    $fmts = @()
    if ($Sheet.formats) { $fmts = @($Sheet.formats) }
    $widths = @()
    if ($Sheet.columnWidths) { $widths = @($Sheet.columnWidths) }
    return [pscustomobject]@{
        name         = [string]$Sheet.name
        rows         = $rows
        columnWidths = $widths
        formats      = $fmts
        meta         = $Sheet.meta
    }
}

function Convert-CellText($Value) {
    if ($null -eq $Value) { return "" }
    return ([string]$Value)
}

function Get-SheetBounds($Sheet) {
    $used = $Sheet.UsedRange
    if ($null -eq $used) { return @{ Row1 = 1; Col1 = 1; Rows = 0; Cols = 0 } }
    return @{
        Row1 = 1
        Col1 = 1
        Rows = [int]$used.Row + [int]$used.Rows.Count - 1
        Cols = [int]$used.Column + [int]$used.Columns.Count - 1
    }
}

function Read-SheetRowsFromA1($Sheet) {
    $b = Get-SheetBounds $Sheet
    if ($b.Rows -lt 1 -or $b.Cols -lt 1) { return @() }
    $out = New-Object System.Collections.Generic.List[object]
    for ($r = 1; $r -le $b.Rows; $r++) {
        $line = New-Object System.Collections.Generic.List[string]
        for ($c = 1; $c -le $b.Cols; $c++) {
            $cell = $Sheet.Cells.Item($r, $c)
            $text = Convert-CellText $cell.Text
            if ($cell.MergeCells) {
                try {
                    $ma = $cell.MergeArea
                    if ([int]$ma.Column -eq $c -and [int]$ma.Row -eq $r) {
                        $text = Convert-CellText $ma.Cells.Item(1, 1).Text
                    }
                    elseif ($r -eq 1) {
                        $text = ""
                    }
                }
                catch { }
            }
            [void]$line.Add($text)
        }
        [void]$out.Add([object[]]@($line.ToArray()))
    }
    return $out.ToArray()
}

function Read-ColumnWidths($Sheet, [int]$Cols) {
    $w = New-Object System.Collections.Generic.List[int]
    for ($c = 1; $c -le $Cols; $c++) {
        $w.Add([int][Math]::Round([double]$Sheet.Columns.Item($c).ColumnWidth))
    }
    return $w.ToArray()
}

function Read-CellFormats($Sheet, [int]$Rows, [int]$Cols) {
    $formats = New-Object System.Collections.Generic.List[object]
    $white = 16777215
    for ($r = 1; $r -le $Rows; $r++) {
        for ($c = 1; $c -le $Cols; $c++) {
            $cell = $Sheet.Cells.Item($r, $c)
            if ($cell.MergeCells) {
                $ma = $cell.MergeArea
                if ([int]$ma.Row -ne $r -or [int]$ma.Column -ne $c) { continue }
            }
            $bg = [int]$cell.Interior.Color
            $bold = [bool]$cell.Font.Bold
            $wrap = [bool]$cell.WrapText
            $ha = [int]$cell.HorizontalAlignment
            $entry = @{ r = $r; c = $c }
            $has = $false
            if ($bg -ne $white -and $bg -ne 0) {
                $entry.bg = $bg
                $has = $true
            }
            if ($bold) {
                $entry.b = $true
                $has = $true
            }
            if ($wrap) {
                $entry.wrap = $true
                $has = $true
            }
            if ($ha -eq -4108) {
                $entry.ha = "center"
                $has = $true
            }
            if ($has) { [void]$formats.Add($entry) }
        }
    }
    return $formats.ToArray()
}

function Get-SummaryMeta($Sheet, [int]$Cols) {
    $dateCols = 0
    for ($c = 2; $c -le $Cols; $c++) {
        $h = [string]$Sheet.Cells.Item(2, $c).Text
        if ($h -match '^\d{4}-\d{2}-\d{2}$') { $dateCols++ } else { break }
    }
    return @{ dateCols = $dateCols }
}

function Build-GoogleSheetsPayloadFromWorkbook {
    param(
        $Workbook,
        [string[]]$SheetNames = $script:BvsGooglePushSheetNames
    )
    $sheetsPayload = New-Object System.Collections.Generic.List[object]
    foreach ($name in $SheetNames) {
        $sh = $null
        try { $sh = $Workbook.Worksheets.Item($name) } catch { $sh = $null }
        if ($null -eq $sh) { continue }
        $tableRows = @(Read-SheetRowsFromA1 $sh)
        $bounds = Get-SheetBounds $sh
        $meta = $null
        if ($name -eq "Сводка") { $meta = Get-SummaryMeta $sh $bounds.Cols }
        [void]$sheetsPayload.Add([pscustomobject]@{
            name         = $name
            rows         = $tableRows
            columnWidths = @(Read-ColumnWidths $sh $bounds.Cols)
            formats      = @(Read-CellFormats $sh $bounds.Rows $bounds.Cols)
            meta         = $meta
        })
    }
    return $sheetsPayload
}

function Write-GooglePushPayloadFile {
    param(
        $SheetsPayload,
        [string]$OutPath
    )
    if ($null -eq $SheetsPayload -or $SheetsPayload.Count -lt 1) {
        throw "Нет листов для google_push_payload.json."
    }
    $arr = @($SheetsPayload)
    if ($SheetsPayload -is [System.Collections.Generic.List[object]]) {
        $arr = @($SheetsPayload.ToArray())
    }
    $normalized = New-Object System.Collections.Generic.List[object]
    foreach ($sh in $arr) {
        [void]$normalized.Add((Normalize-PayloadSheet $sh))
    }
    $body = @{ sheets = @($normalized.ToArray()) }
    $json = $body | ConvertTo-Json -Depth 30 -Compress
    $utf8 = New-Object System.Text.UTF8Encoding $false
    $dir = Split-Path -Parent $OutPath
    if ($dir -and -not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
    [System.IO.File]::WriteAllText($OutPath, $json, $utf8)
}
