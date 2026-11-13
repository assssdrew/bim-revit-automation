# UTF-8 WITH BOM required (PowerShell 5.1).
# Запись xlsx через DocumentFormat.OpenXml (без Excel COM) + payload для Google.

$script:OpenXmlLoaded = $false
function Get-AssignmentModuleRoot {
    if ($script:BvsAssignmentDir) { return $script:BvsAssignmentDir }
    if ($PSScriptRoot) { return $PSScriptRoot }
    if ($MyInvocation.MyCommand.Path) { return (Split-Path -Parent $MyInvocation.MyCommand.Path) }
    throw "Не определена папка assignment_xlsx.ps1"
}
function Ensure-OpenXmlLoaded {
    if ($script:OpenXmlLoaded) { return }
    $dll = Join-Path (Join-Path (Split-Path -Parent (Get-AssignmentModuleRoot)) "lib") "DocumentFormat.OpenXml.dll"     # служебное\lib рядом со служебное\скрипты
    if (-not (Test-Path -LiteralPath $dll)) {
        throw "Нет $dll — нужен служебное\lib\DocumentFormat.OpenXml.dll"
    }
    try { Add-Type -Path $dll }
    catch {
        # сборка на сетевой папке (UNC) или помеченная Windows как «из интернета» не загружается напрямую: копия в профиль пользователя и загрузка оттуда
        $cache = Join-Path $env:LOCALAPPDATA "BatchRvt\batch_verification_schedules\lib"
        if (-not (Test-Path -LiteralPath $cache)) { New-Item -ItemType Directory -Path $cache -Force | Out-Null }
        $local = Join-Path $cache "DocumentFormat.OpenXml.dll"
        Copy-Item -LiteralPath $dll -Destination $local -Force
        try { Unblock-File -LiteralPath $local -ErrorAction SilentlyContinue } catch { }
        Add-Type -Path $local
    }
    $script:OpenXmlLoaded = $true
}

function Add-OpenXmlPart {
    param($ParentPart, [Type]$PartType)
    $mi = @($ParentPart.GetType().GetMethods()) | Where-Object {
        $_.Name -eq 'AddNewPart' -and $_.IsGenericMethodDefinition -and $_.GetParameters().Count -eq 0
    } | Select-Object -First 1
    if (-not $mi) { throw "AddNewPart не найден для $($ParentPart.GetType().FullName)" }
    $gm = $mi.MakeGenericMethod($PartType)
    return $gm.Invoke($ParentPart, @())
}

# Цвет в формате Excel/VBA (.Interior.Color): число R + G*256 + B*65536, то есть R — МЛАДШИЙ байт.
function Convert-BgrToRgbHex([int]$Bgr) {
    $r = $Bgr -band 0xFF
    $g = ($Bgr -shr 8) -band 0xFF
    $b = ($Bgr -shr 16) -band 0xFF
    return ("{0:X2}{1:X2}{2:X2}" -f $r, $g, $b)
}

# Символы, недопустимые в XML 1.0 (Revit иногда отдаёт управляющие коды в значениях параметров).
function Remove-InvalidXmlChars([string]$Text) {
    if ([string]::IsNullOrEmpty($Text)) { return $Text }
    return [regex]::Replace($Text, '[\x00-\x08\x0B\x0C\x0E-\x1F\uFFFE\uFFFF]', '')
}

function Pad-SheetRowsToRectangle {
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
    # запятая: иначе PowerShell «разворачивает» массив из одной строки в плоский массив
    return ,@($out.ToArray())
}

function New-XlsxStyleSheet {
    Ensure-OpenXmlLoaded
    $fills = New-Object DocumentFormat.OpenXml.Spreadsheet.Fills
    $noneFill = New-Object DocumentFormat.OpenXml.Spreadsheet.Fill
    $nonePat = New-Object DocumentFormat.OpenXml.Spreadsheet.PatternFill
    $nonePat.PatternType = [DocumentFormat.OpenXml.Spreadsheet.PatternValues]::None
    $noneFill.PatternFill = $nonePat
    $fills.AppendChild($noneFill) | Out-Null # 0 none
    $grayFill = New-Object DocumentFormat.OpenXml.Spreadsheet.Fill
    $grayPat = New-Object DocumentFormat.OpenXml.Spreadsheet.PatternFill
    $grayPat.PatternType = [DocumentFormat.OpenXml.Spreadsheet.PatternValues]::Gray125
    $grayFill.PatternFill = $grayPat
    $fills.AppendChild($grayFill) | Out-Null # 1 gray125
    $fonts = New-Object DocumentFormat.OpenXml.Spreadsheet.Fonts
    foreach ($isBold in @($false, $true)) {
        $f = New-Object DocumentFormat.OpenXml.Spreadsheet.Font
        if ($isBold) { $f.AppendChild((New-Object DocumentFormat.OpenXml.Spreadsheet.Bold)) | Out-Null }
        $fs = New-Object DocumentFormat.OpenXml.Spreadsheet.FontSize
        $fs.Val = 11.0
        $f.AppendChild($fs) | Out-Null
        $fn = New-Object DocumentFormat.OpenXml.Spreadsheet.FontName
        $fn.Val = "Calibri"
        $f.AppendChild($fn) | Out-Null
        $fonts.AppendChild($f) | Out-Null
    }
    # 2: серый шрифт — исправленные замечания
    $fg2 = New-Object DocumentFormat.OpenXml.Spreadsheet.Font
    $fs2 = New-Object DocumentFormat.OpenXml.Spreadsheet.FontSize
    $fs2.Val = 11.0
    $fg2.AppendChild($fs2) | Out-Null
    $fc2 = New-Object DocumentFormat.OpenXml.Spreadsheet.Color
    $fc2.Rgb = "FF808080"
    $fg2.AppendChild($fc2) | Out-Null
    $fn2 = New-Object DocumentFormat.OpenXml.Spreadsheet.FontName
    $fn2.Val = "Calibri"
    $fg2.AppendChild($fn2) | Out-Null
    $fonts.AppendChild($fg2) | Out-Null
    $borders = New-Object DocumentFormat.OpenXml.Spreadsheet.Borders
    $borders.AppendChild((New-Object DocumentFormat.OpenXml.Spreadsheet.Border)) | Out-Null # 0 без границ
    # 1: тонкая серая сетка со всех сторон (как на листах разделов раньше и как в Google)
    $gridBorder = New-Object DocumentFormat.OpenXml.Spreadsheet.Border
    foreach ($side in @('Left', 'Right', 'Top', 'Bottom')) {
        $b = New-Object ("DocumentFormat.OpenXml.Spreadsheet." + $side + "Border")
        $b.Style = [DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues]::Thin
        $bc = New-Object DocumentFormat.OpenXml.Spreadsheet.Color
        $bc.Rgb = "FF808080"
        $b.Color = $bc
        $gridBorder.AppendChild($b) | Out-Null
    }
    $borders.AppendChild($gridBorder) | Out-Null
    # 2: сетка без правой границы, 3: без левой — склейка шапки у пары «Пустые»/«Без значения»
    foreach ($skip in @('Right', 'Left')) {
        $b2 = New-Object DocumentFormat.OpenXml.Spreadsheet.Border
        foreach ($side in @('Left', 'Right', 'Top', 'Bottom')) {
            if ($side -eq $skip) { continue }
            $bs = New-Object ("DocumentFormat.OpenXml.Spreadsheet." + $side + "Border")
            $bs.Style = [DocumentFormat.OpenXml.Spreadsheet.BorderStyleValues]::Thin
            $bcc = New-Object DocumentFormat.OpenXml.Spreadsheet.Color
            $bcc.Rgb = "FF808080"
            $bs.Color = $bcc
            $b2.AppendChild($bs) | Out-Null
        }
        $borders.AppendChild($b2) | Out-Null
    }
    $cellStyleXfs = New-Object DocumentFormat.OpenXml.Spreadsheet.CellStyleFormats
    $cellStyleXfs.AppendChild((New-Object DocumentFormat.OpenXml.Spreadsheet.CellFormat)) | Out-Null
    $cellFormats = New-Object DocumentFormat.OpenXml.Spreadsheet.CellFormats
    $cf0 = New-Object DocumentFormat.OpenXml.Spreadsheet.CellFormat
    $cf0.FormatId = [UInt32]0
    $cf0.FontId = [UInt32]0
    $cf0.FillId = [UInt32]0
    $cf0.BorderId = [UInt32]0
    $cellFormats.AppendChild($cf0) | Out-Null # 0
    $cfBold = New-Object DocumentFormat.OpenXml.Spreadsheet.CellFormat
    $cfBold.FormatId = [UInt32]0
    $cfBold.FontId = [UInt32]1
    $cfBold.FillId = [UInt32]0
    $cfBold.BorderId = [UInt32]0
    $cfBold.ApplyFont = $true
    $cellFormats.AppendChild($cfBold) | Out-Null # 1
    $stylesheet = New-Object DocumentFormat.OpenXml.Spreadsheet.Stylesheet
    $stylesheet.Fills = $fills
    $stylesheet.Fonts = $fonts
    $stylesheet.Borders = $borders
    $stylesheet.CellStyleFormats = $cellStyleXfs
    $stylesheet.CellFormats = $cellFormats
    $stylesheet.Fills.Count = [UInt32]$fills.ChildElements.Count
    $stylesheet.Fonts.Count = [UInt32]$fonts.ChildElements.Count
    $stylesheet.Borders.Count = [UInt32]$borders.ChildElements.Count
    $stylesheet.CellFormats.Count = [UInt32]$cellFormats.ChildElements.Count
    $stylesheet.CellStyleFormats.Count = [UInt32]$cellStyleXfs.ChildElements.Count
    return @{
        Sheet          = $stylesheet
        FillMap        = @{}  # rgb hex -> fillId
        FormatMap      = @{}  # key -> cellFormatId
        NextFillId     = 2
        NextFormatId   = 2
        CellFormats    = $cellFormats
        Fills          = $fills
    }
}

function Get-OrAddFillStyle($StyleState, [string]$RgbHex) {
    if ($StyleState.FillMap.ContainsKey($RgbHex)) {
        return $StyleState.FillMap[$RgbHex]
    }
    $fill = New-Object DocumentFormat.OpenXml.Spreadsheet.Fill
    $p = New-Object DocumentFormat.OpenXml.Spreadsheet.PatternFill
    $p.PatternType = [DocumentFormat.OpenXml.Spreadsheet.PatternValues]::Solid
    $fg = New-Object DocumentFormat.OpenXml.Spreadsheet.ForegroundColor
    if ($RgbHex.Length -eq 6) { $fg.Rgb = "FF" + $RgbHex }
    else { $fg.Rgb = $RgbHex }
    $p.ForegroundColor = $fg
    $fill.PatternFill = $p
    $StyleState.Fills.AppendChild($fill) | Out-Null
    $id = $StyleState.NextFillId
    $StyleState.NextFillId++
    $StyleState.FillMap[$RgbHex] = $id
    $StyleState.Sheet.Fills.Count = [UInt32]$StyleState.Fills.ChildElements.Count
    return $id
}

function Get-CellFormatId($StyleState, [hashtable]$Opts) {
    $key = ($Opts.bg) + "|" + $Opts.bold + "|" + $Opts.border + "|" + $Opts.wrap + "|" + $Opts.center + "|" + $Opts.rot + "|" + $Opts.gray + "|" + $Opts.nbr + "|" + $Opts.nbl
    if ($StyleState.FormatMap.ContainsKey($key)) { return $StyleState.FormatMap[$key] }
    $cf = New-Object DocumentFormat.OpenXml.Spreadsheet.CellFormat
    $cf.FormatId = [UInt32]0
    $cf.FontId = [UInt32]0
    $cf.FillId = [UInt32]0
    $cf.BorderId = [UInt32]0
    if ($Opts.bold) {
        $cf.FontId = [UInt32]1
        $cf.ApplyFont = $true
    }
    elseif ($Opts.gray) {
        $cf.FontId = [UInt32]2
        $cf.ApplyFont = $true
    }
    if ($Opts.bg) {
        $cf.FillId = [UInt32](Get-OrAddFillStyle $StyleState $Opts.bg)
        $cf.ApplyFill = $true
    }
    if ($Opts.border) {
        $cf.BorderId = [UInt32]1
        if ($Opts.nbr) { $cf.BorderId = [UInt32]2 } elseif ($Opts.nbl) { $cf.BorderId = [UInt32]3 }
        $cf.ApplyBorder = $true
    }
    if ($Opts.rot) {
        # вертикальный текст (заголовки спецификаций в Сводке): снизу вверх, по центру колонки
        $cf.ApplyAlignment = $true
        $al = New-Object DocumentFormat.OpenXml.Spreadsheet.Alignment
        $al.TextRotation = [UInt32][int]$Opts.rot
        if ($Opts.wrap) { $al.WrapText = $true }
        $al.Horizontal = [DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues]::Center
        $al.Vertical = [DocumentFormat.OpenXml.Spreadsheet.VerticalAlignmentValues]::Bottom
        $cf.Alignment = $al
    }
    if ($Opts.wrap -and -not $Opts.rot) {
        $cf.ApplyAlignment = $true
        $al = New-Object DocumentFormat.OpenXml.Spreadsheet.Alignment
        $al.WrapText = $true
        $cf.Alignment = $al
    }
    if ($Opts.center) {
        $cf.ApplyAlignment = $true
        $al = if ($cf.Alignment) { $cf.Alignment } else { New-Object DocumentFormat.OpenXml.Spreadsheet.Alignment }
        $al.Horizontal = [DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues]::Center
        $cf.Alignment = $al
    }
    $StyleState.CellFormats.AppendChild($cf) | Out-Null
    $id = $StyleState.NextFormatId
    $StyleState.NextFormatId++
    $StyleState.FormatMap[$key] = $id
    $StyleState.Sheet.CellFormats.Count = [UInt32]$StyleState.CellFormats.ChildElements.Count
    return $id
}

function New-AssignmentXlsxWriter {
    param([string]$Path)
    Ensure-OpenXmlLoaded
    if (Test-Path -LiteralPath $Path) { Remove-Item -LiteralPath $Path -Force }
    $doc = [DocumentFormat.OpenXml.Packaging.SpreadsheetDocument]::Create($Path, [DocumentFormat.OpenXml.SpreadsheetDocumentType]::Workbook)
    $wbPart = $doc.AddWorkbookPart()
    $wbPart.Workbook = New-Object DocumentFormat.OpenXml.Spreadsheet.Workbook
    $stylesPart = Add-OpenXmlPart $wbPart ([Type][DocumentFormat.OpenXml.Packaging.WorkbookStylesPart])
    $styleState = New-XlsxStyleSheet
    $stylesPart.Stylesheet = $styleState.Sheet
    $sheets = New-Object DocumentFormat.OpenXml.Spreadsheet.Sheets
    $wbPart.Workbook.AppendChild($sheets) | Out-Null
    return [pscustomobject]@{
        Doc         = $doc
        Workbook    = $wbPart
        Sheets      = $sheets
        SheetIndex  = 0
        StyleState  = $styleState
        Payload     = New-Object System.Collections.Generic.List[object]
    }
}

function Close-AssignmentXlsxWriter($Writer) {
    $Writer.Workbook.Workbook.Save()
    $Writer.Doc.Save()
    $Writer.Doc.Dispose()
}

function Set-XlsxCellValue {
    param(
        [DocumentFormat.OpenXml.Spreadsheet.Cell]$Cell,
        [string]$Text
    )
    if ($null -eq $Text) { $Text = "" }
    $Text = Remove-InvalidXmlChars $Text
    # число — только целое без ведущих нулей ("007" остаётся текстом)
    if ($Text -match '^(0|[1-9]\d{0,10})$') {
        $Cell.DataType = [DocumentFormat.OpenXml.Spreadsheet.CellValues]::Number
        $cv = New-Object DocumentFormat.OpenXml.Spreadsheet.CellValue
        $cv.Text = $Text
        $Cell.AppendChild($cv) | Out-Null
        return
    }
    $Cell.DataType = [DocumentFormat.OpenXml.Spreadsheet.CellValues]::InlineString
    $inline = New-Object DocumentFormat.OpenXml.Spreadsheet.InlineString
    $t = New-Object DocumentFormat.OpenXml.Spreadsheet.Text
    $t.Text = $Text
    if ($Text -match '(^\s)|(\s$)') {
        $t.Space = [DocumentFormat.OpenXml.Spreadsheet.SpaceProcessingModeValues]::Preserve
    }
    $inline.AppendChild($t) | Out-Null
    $Cell.AppendChild($inline) | Out-Null
}

function Get-ColumnName([int]$Index) {
    $n = $Index
    $name = ""
    while ($n -gt 0) {
        $rem = ($n - 1) % 26
        $name = [char](65 + $rem) + $name
        $n = [int][Math]::Floor(($n - 1) / 26)  # [int] rounds 0.5+ up: column 15 became AO
    }
    return $name
}

function Add-XlsxSheetFromTable {
    param(
        $Writer,
        [string]$SheetName,
        [object[]]$Rows,
        [int[]]$ColumnWidths,
        [object[]]$Formats,
        [hashtable]$Meta,
        [int]$FreezeRows = 2,
        [hashtable[]]$Merges,
        [hashtable]$OutlineLevels,
        [int]$BorderFromRow = 0,
        [hashtable]$RowHeights,
        # что отправлять в Google вместо листа xlsx (например, только открытые замечания, 7 колонок)
        [hashtable]$PayloadData,
        [switch]$SkipPayload
    )
    $Rows = Pad-SheetRowsToRectangle $Rows
    $Writer.SheetIndex++
    $part = Add-OpenXmlPart $Writer.Workbook ([Type][DocumentFormat.OpenXml.Packaging.WorksheetPart])
    $partId = $Writer.Workbook.GetIdOfPart($part)
    $sheet = New-Object DocumentFormat.OpenXml.Spreadsheet.Sheet
    $sheet.Id = $partId
    $sheet.Name = $SheetName
    $sheet.SheetId = [UInt32]$Writer.SheetIndex
    $Writer.Sheets.AppendChild($sheet) | Out-Null
    $ws = New-Object DocumentFormat.OpenXml.Spreadsheet.Worksheet
    $sheetPr = New-Object DocumentFormat.OpenXml.Spreadsheet.SheetProperties
    $outlinePr = New-Object DocumentFormat.OpenXml.Spreadsheet.OutlineProperties
    $outlinePr.SummaryBelow = $false
    $sheetPr.OutlineProperties = $outlinePr
    $ws.SheetProperties = $sheetPr
    $sheetData = New-Object DocumentFormat.OpenXml.Spreadsheet.SheetData
    $formatLookup = @{}
    if ($Formats) {
        foreach ($f in $Formats) {
            $formatLookup[("{0},{1}" -f $f.r, $f.c)] = $f
        }
    }
    $rowCount = $Rows.Count
    $colCount = 0
    foreach ($rowItem in @($Rows)) {
        if ($null -ne $rowItem) {
            $n = @($rowItem).Count
            if ($n -gt $colCount) { $colCount = $n }
        }
    }
    $sheetFormatPr = New-Object DocumentFormat.OpenXml.Spreadsheet.SheetFormatProperties
    $sheetFormatPr.DefaultRowHeight = 15.0
    $sheetFormatPr.OutlineLevelRow = [Byte]2
    for ($r = 0; $r -lt $rowCount; $r++) {
        $excelRow = $r + 1
        $row = New-Object DocumentFormat.OpenXml.Spreadsheet.Row
        $row.RowIndex = [UInt32]$excelRow
        if ($RowHeights -and $RowHeights.ContainsKey($excelRow)) {
            $row.Height = [double]$RowHeights[$excelRow]
            $row.CustomHeight = $true
        }
        if ($OutlineLevels -and $OutlineLevels.ContainsKey($excelRow)) {
            $lvl = [int]$OutlineLevels[$excelRow]
            $row.OutlineLevel = [Byte]$lvl
            # вид по умолчанию «2 — спецификации»: уровень 1 виден, элементы (уровень 2) свёрнуты
            if ($lvl -ge 2) { $row.Hidden = $true }
        }
        # «+» у строки-итога свёрнутой группы (SummaryBelow=false: итог стоит над группой)
        $curLvl = 0
        if ($OutlineLevels -and $OutlineLevels.ContainsKey($excelRow)) { $curLvl = [int]$OutlineLevels[$excelRow] }
        if ($OutlineLevels -and $OutlineLevels.ContainsKey($excelRow + 1) -and [int]$OutlineLevels[$excelRow + 1] -gt $curLvl -and [int]$OutlineLevels[$excelRow + 1] -ge 2) {
            $row.Collapsed = $true
        }
        $line = $Rows[$r]
        if ($null -eq $line) { $line = @() }
        for ($c = 0; $c -lt $colCount; $c++) {
            $colName = Get-ColumnName ($c + 1)
            $cell = New-Object DocumentFormat.OpenXml.Spreadsheet.Cell
            $cell.CellReference = $colName + $excelRow
            $val = $line[$c]
            if ($null -eq $val) { $val = "" }
            $text = [string]$val
            Set-XlsxCellValue -Cell $cell -Text $text
            $fk = "{0},{1}" -f $excelRow, ($c + 1)
            $hasBorder = ($BorderFromRow -gt 0 -and $excelRow -ge $BorderFromRow)
            if ($formatLookup.ContainsKey($fk) -or $hasBorder) {
                $fo = $formatLookup[$fk]
                $opts = @{
                    bold   = $false
                    border = $hasBorder
                    wrap   = $false
                    center = $false
                    bg     = $null
                    rot    = 0
                    gray   = $false
                    nbr    = $false
                    nbl    = $false
                }
                if ($fo) {
                    $opts.bold = [bool]$fo.b
                    $opts.wrap = [bool]$fo.wrap
                    $opts.center = ($fo.ha -eq "center")
                    if ($fo.rot) { $opts.rot = [int]$fo.rot }
                    if ($fo.gray -or $fo.dim) { $opts.gray = $true }
                    if ($fo.nbr) { $opts.nbr = $true }
                    if ($fo.nbl) { $opts.nbl = $true }
                    if ($null -ne $fo.bg -and [int]$fo.bg -ne 0) {
                        $opts.bg = Convert-BgrToRgbHex ([int]$fo.bg)
                    }
                }
                $cell.StyleIndex = [UInt32](Get-CellFormatId $Writer.StyleState $opts)
            }
            $row.AppendChild($cell) | Out-Null
        }
        $sheetData.AppendChild($row) | Out-Null
    }
    $ws.AppendChild($sheetData) | Out-Null
    $viewsNode = $null
    $colsNode = $null
    if ($FreezeRows -gt 0) {
        $viewsNode = New-Object DocumentFormat.OpenXml.Spreadsheet.SheetViews
        $view = New-Object DocumentFormat.OpenXml.Spreadsheet.SheetView
        $view.WorkbookViewId = [UInt32]0
        $pane = New-Object DocumentFormat.OpenXml.Spreadsheet.Pane
        $pane.VerticalSplit = [double]$FreezeRows
        $pane.TopLeftCell = "A" + ($FreezeRows + 1)
        $pane.ActivePane = [DocumentFormat.OpenXml.Spreadsheet.PaneValues]::BottomLeft
        $pane.State = [DocumentFormat.OpenXml.Spreadsheet.PaneStateValues]::Frozen
        $view.AppendChild($pane) | Out-Null
        $viewsNode.AppendChild($view) | Out-Null
    }
    if ($ColumnWidths -and $ColumnWidths.Count -gt 0) {
        $colsNode = New-Object DocumentFormat.OpenXml.Spreadsheet.Columns
        for ($i = 0; $i -lt $ColumnWidths.Count; $i++) {
            $col = New-Object DocumentFormat.OpenXml.Spreadsheet.Column
            $col.Min = [UInt32]($i + 1)
            $col.Max = [UInt32]($i + 1)
            $col.Width = [double]$ColumnWidths[$i]
            $col.CustomWidth = $true
            $colsNode.AppendChild($col) | Out-Null
        }
    }
    if ($null -ne $viewsNode) {
        $ws.InsertBefore($viewsNode, $sheetData) | Out-Null
        $ws.InsertAfter($sheetFormatPr, $viewsNode) | Out-Null
    }
    else {
        $ws.InsertBefore($sheetFormatPr, $sheetData) | Out-Null
    }
    if ($null -ne $colsNode) {
        if ($null -ne $viewsNode) {
            $ws.InsertAfter($colsNode, $sheetFormatPr) | Out-Null
        }
        else {
            $ws.InsertAfter($colsNode, $sheetFormatPr) | Out-Null
        }
    }
    # объединение одной ячейки (A1:A1) Excel считает повреждением файла («восстановить?»), такие пропускаем
    $realMerges = @()
    if ($Merges) { $realMerges = @($Merges | Where-Object { $_.Reference -match '^([A-Z]+\d+):(?!\1$)' }) }
    if ($realMerges.Count -gt 0) {
        $mergeCells = New-Object DocumentFormat.OpenXml.Spreadsheet.MergeCells
        foreach ($m in $realMerges) {
            $mc = New-Object DocumentFormat.OpenXml.Spreadsheet.MergeCell
            $mc.Reference = $m.Reference
            $mergeCells.AppendChild($mc) | Out-Null
        }
        $mergeCells.Count = [UInt32]$mergeCells.ChildElements.Count
        $ws.AppendChild($mergeCells) | Out-Null
    }
    $part.Worksheet = $ws
    if ($SkipPayload) { return }
    $fmtList = @()
    if ($Formats) { $fmtList = @($Formats) }
    $payRows = $Rows
    $payWidths = $ColumnWidths
    if ($PayloadData) {
        $payRows = Pad-SheetRowsToRectangle $PayloadData.Rows
        $payWidths = $PayloadData.Widths
        $fmtList = @()
        if ($PayloadData.Formats) { $fmtList = @($PayloadData.Formats) }
    }
    [void]$Writer.Payload.Add([pscustomobject]@{
        name         = $SheetName
        rows         = $payRows
        columnWidths = $payWidths
        formats      = $fmtList
        meta         = $Meta
    })
}

# Ширина колонки по содержимому (единицы Excel ≈ символы Calibri 11; кириллица чуть шире).
# FirstRow — первая строка данных (1-based): строку-заметку (объединённую) в расчёт не берём.
function Get-AutoColumnWidths {
    param([object[]]$Rows, [int]$FirstRow = 2, [int]$Min = 10, [int]$Max = 90, [int[]]$MinPerColumn = @())
    $cols = 0
    foreach ($r in @($Rows)) { if ($null -ne $r -and @($r).Count -gt $cols) { $cols = @($r).Count } }
    $out = New-Object System.Collections.Generic.List[int]
    for ($c = 0; $c -lt $cols; $c++) {
        $len = 0
        for ($i = $FirstRow - 1; $i -lt @($Rows).Count; $i++) {
            $row = @($Rows)[$i]
            if ($null -eq $row -or $c -ge @($row).Count) { continue }
            $t = [string]@($row)[$c]
            $l = 0
            foreach ($line in ($t -split "\r?\n")) { if ($line.Length -gt $l) { $l = $line.Length } }
            if ($l -gt $len) { $len = $l }
        }
        $w = [int][Math]::Ceiling($len * 1.15) + 2
        $lo = $Min
        if ($c -lt $MinPerColumn.Count -and $MinPerColumn[$c] -gt $lo) { $lo = $MinPerColumn[$c] }
        if ($w -lt $lo) { $w = $lo }
        if ($w -gt $Max) { $w = $Max }
        $out.Add($w)
    }
    return ,@($out.ToArray())
}

function Build-OutlineLevelMap($Block, [int]$StartExcelRow) {
    $levels = @{}
    foreach ($g in $Block.Inners) {
        for ($r = [int]$g.First; $r -le [int]$g.Last; $r++) { $levels[$r] = 2 }
    }
    foreach ($g in $Block.Outers) {
        for ($r = [int]$g.First; $r -le [int]$g.Last; $r++) {
            if (-not $levels.ContainsKey($r)) { $levels[$r] = 1 }
        }
    }
    return $levels
}

# Заголовок колонки среза: «Замечания / дата / время сборки» (у старых срезов времени нет — только дата).
function Format-SliceHeader([string]$SliceId) {
    if ($SliceId -match '^(\d{4}-\d\d-\d\d) (\d\d:\d\d):\d\d$') {
        if ($Matches[2] -eq "00:00") { return ("Замечания" + "`n" + $Matches[1]) }
        return ("Замечания" + "`n" + $Matches[1] + "`n" + $Matches[2])
    }
    return ("Замечания" + "`n" + $SliceId)
}

# $Partial — только выборочная выгрузка (иначе $null и Сводка строится как раньше): @{ Checks; CheckedPairs; BuildId; StaleDays }.
# Добавляет последнюю колонку «Проверено» (самая старая проверка среди пар модели; красноватым — если старше StaleDays дней)
# и серый шрифт у пар, которые в этой сборке не проверялись.
function Build-SummarySheetData($State, $Prev, $DateCols, $Today, $Sep, $scriptByDateModelSum, $Deltas = $null, $Partial = $null) {
    # колонок срезов «Замечания / дата» (итог по модели) больше нет: замечание — строка спецификации, элементы не считаются
    $DateCols = @()
    $checks = @(Get-ChecksWithRemarksFromStream $State $Prev $Sep)
    $modelSet = @{}
    foreach ($k in @($State.Models.Keys)) { $modelSet[$k] = $true }
    foreach ($k in @($Prev.Keys)) { $modelSet[$k.Split($Sep)[0]] = $true }
    $models = @($modelSet.Keys | Sort-Object { Get-ModelSort $_ }, { $_ })
    $todayCount = $State.TodayCount
    $dateCount = @($DateCols).Count
    $extraCols = 0
    if ($null -ne $Partial) { $extraCols = 1 }
    $cols = 1 + $dateCount + $checks.Count + $extraCols
    $rows = New-Object System.Collections.Generic.List[object]
    $rows.Add(@("Сводка. Число строк спецификации с незаполненным параметром. Оранжевый — есть замечания. Зелёный ноль — чисто. Синий ноль — в прошлой выгрузке были, сейчас нет. В скобках — прошлое число. ↓ — исправлено, ↑ — новых с прошлой сборки."))
    $header = New-Object System.Collections.Generic.List[object]
    $header.Add("Модель")
    foreach ($d in $DateCols) { $header.Add((Format-SliceHeader $d)) }
    foreach ($c in $checks) { $header.Add($c) }
    if ($null -ne $Partial) { $header.Add("Проверено") }
    $rows.Add(@($header.ToArray()))
    # пары одной проверки («…_Пустые» и «…_Без значения») стоят рядом: в шапке между ними границы нет
    $pairCols = New-Object System.Collections.Generic.List[int]
    for ($i = 0; $i -lt ($checks.Count - 1); $i++) {
        $pa = Get-CheckNameParts $checks[$i]
        $pb = Get-CheckNameParts $checks[$i + 1]
        # склейка шапки только у пары «Пустые»/«Без значения»; «_999» рядом с «Пустые» остаётся обычной границей
        $pair = ($pa.Base -eq $pb.Base) -and (($pa.Kind -eq "empty" -and $pb.Kind -eq "novalue") -or ($pa.Kind -eq "novalue" -and $pb.Kind -eq "empty"))
        if ($pair) { $pairCols.Add(2 + $dateCount + $i) }
    }
    $formats = New-Object System.Collections.Generic.List[object]
    $formats.Add(@{ r = 1; c = 1; bg = 0x00FFFF; wrap = $true })
    for ($hc = 1; $hc -le $cols; $hc++) {
        if ($hc -gt 1) {
            # «Замечания / дата» и названия спецификаций — вертикально, чтобы при любом их числе таблица помещалась на экране
            $hf = @{ r = 2; c = $hc; bg = 0xF2E1D9; b = $true; rot = 90; wrap = $true }
            if ($pairCols.Contains($hc)) { $hf.nbr = $true }
            if ($pairCols.Contains($hc - 1)) { $hf.nbl = $true }
            $formats.Add($hf)
        }
        else {
            $formats.Add(@{ r = 2; c = $hc; bg = 0xF2E1D9; b = $true; wrap = $true })
        }
    }
    $longest = 0
    foreach ($ck in $checks) { if ([string]$ck -and ([string]$ck).Length -gt $longest) { $longest = ([string]$ck).Length } }
    # перенос названия в три строки: высота шапки под треть самого длинного названия (+25% на переносы по словам)
    $lineLen = [int][Math]::Ceiling($longest / 3.0 * 1.25)
    $headerHeight = [Math]::Min(400, [Math]::Max(60, [int][Math]::Ceiling($lineLen * 5.9) + 14))
    for ($r = 0; $r -lt $models.Count; $r++) {
        $model = $models[$r]
        $line = New-Object System.Collections.Generic.List[object]
        $line.Add($model)
        $oldest = ""
        for ($d = 0; $d -lt $dateCount; $d++) {
            $date = $DateCols[$d]
            $dm = $date + $Sep + $model
            $sum = 0
            if ($scriptByDateModelSum.ContainsKey($dm)) { $sum = [int]$scriptByDateModelSum[$dm] }
            $line.Add($sum)
        }
        for ($c = 0; $c -lt $checks.Count; $c++) {
            $key = $model + $Sep + $checks[$c]
            $now = 0
            if ($todayCount.ContainsKey($key)) { $now = [int]$todayCount[$key] }
            $was = $null
            if ($Prev.ContainsKey($key)) { $was = [int]$Prev[$key] }
            $cellText = Count-Text $now $was
            if ($Deltas -and $Deltas.ContainsKey($key)) {
                $dl = $Deltas[$key]
                if ($dl.Fixed -gt 0) { $cellText += (" " + [char]0x2193 + $dl.Fixed) }
                if ($dl.New -gt 0) { $cellText += (" " + [char]0x2191 + $dl.New) }
            }
            $line.Add($cellText)
            $excelR = 3 + $r
            $excelC = 2 + $dateCount + $c
            $wasNum = 0
            if ($null -ne $was) { $wasNum = [int]$was }
            $bg = 0xCEEFC6
            if ($now -gt 0) { $bg = 0x80C0FF }
            elseif ($wasNum -gt 0) { $bg = 0xF0D2B4 }
            $cf = @{ r = $excelR; c = $excelC; bg = $bg; ha = "center" }
            if ($null -ne $Partial -and -not $Partial.CheckedPairs.ContainsKey($key)) { $cf.dim = $true }
            $formats.Add($cf)
            if ($null -ne $Partial) {
                $when = Get-PairChecked $Partial.Checks $model ([string]$checks[$c])
                if ($when -and (-not $oldest -or $when -lt $oldest)) { $oldest = $when }
            }
        }
        if ($null -ne $Partial) {
            $line.Add($(if ($oldest) { Format-DayRu $oldest } else { "" }))
            $fcf = @{ r = (3 + $r); c = $cols; ha = "center" }
            if ($oldest) {
                $age = Get-DaysBetween $oldest $Partial.BuildId
                if ($age -ne "" -and [int]$age -ge [int]$Partial.StaleDays) { $fcf.bg = 0xC3C7F4 }
            }
            $formats.Add($fcf)
        }
        $rows.Add(@($line.ToArray()))
        $formats.Add(@{ r = (3 + $r); c = 1; b = $true })
    }
    # по содержимому; заголовки проверок длинные и переносятся (строка 2 высотой 90), поэтому для них минимум 14
    $minCols = @(10)
    # «Замечания / дата» повернута на 90°, ширина — под число (3–4 цифры)
    for ($i = 1; $i -lt $cols; $i++) { if ($i -le $dateCount) { $minCols += 8 } else { $minCols += 10 } }
    $dataRows = @($rows.ToArray())
    $dataRows[1] = @($dataRows[1] | ForEach-Object { "" })
    $widths = Get-AutoColumnWidths -Rows $dataRows -FirstRow 2 -Min 6 -Max 60 -MinPerColumn $minCols
    $merges = @(@{ Reference = ("A1:{0}1" -f (Get-ColumnName $cols)) })
    return @{
        Rows       = $rows.ToArray()
        Widths     = $widths
        Formats    = $formats.ToArray()
        Meta       = @{ dateCols = $dateCount; headerHeight = $headerHeight; pairCols = @($pairCols.ToArray()) }
        RowHeights = @{ 2 = $headerHeight }
        Merges     = $merges
    }
}

function Build-SectionSheetData($Items, [switch]$Wide, [int]$KeepFixedDays = 30, [string]$ExtraNote = "") {
    $note = "Вариант 1. Лист раздела: строка модели. Кнопка 1 — только модели, 2 — спецификации, 3 — элементы. Плюс слева раскрывает одну строку."
    $headers = @("Модель", "Спецификация", "Параметр", "Семейство и типоразмер", "Наименование", "Значение")
    $cols = 6
    if ($Wide) {
        $note = $note + (" Серым — исправленные за последние {0} дн. «Исправлено» в строках модели и спецификации — число исправленных." -f $KeepFixedDays)
        $headers = $headers + @("Статус", "Найдено", "Исправлено", "Дней")
        $cols = 10
    }
    if ($ExtraNote) { $note = $note + " " + $ExtraNote }
    $block = Build-SectionOutlineBlock $Items 3 -Wide:$Wide
    $rows = New-Object System.Collections.Generic.List[object]
    $rows.Add(@($note))
    $rows.Add($headers)
    for ($r = 0; $r -lt $block.Grid.GetLength(0); $r++) {
        $line = New-Object object[] $cols
        for ($c = 0; $c -lt $cols; $c++) { $line[$c] = $block.Grid[$r, $c] }
        $rows.Add($line)
    }
    $formats = New-Object System.Collections.Generic.List[object]
    $formats.Add(@{ r = 1; c = 1; bg = 0x00FFFF; wrap = $true })
    for ($hc = 1; $hc -le $cols; $hc++) { $formats.Add(@{ r = 2; c = $hc; bg = 0xF2E1D9; b = $true; wrap = $true }) }
    $last = 2 + $block.Grid.GetLength(0)
    foreach ($rr in $block.ModelRows) {
        for ($hc = 1; $hc -le $cols; $hc++) { $formats.Add(@{ r = $rr; c = $hc; bg = 0xF2F2F2; b = ($hc -eq 1) }) }
    }
    foreach ($rr in $block.CheckRows) {
        for ($hc = 1; $hc -le $cols; $hc++) { $formats.Add(@{ r = $rr; c = $hc; bg = 0xD1E8D1; b = ($hc -eq 2) }) }
    }
    if ($Wide) {
        foreach ($rr in $block.FixedRows) {
            for ($hc = 1; $hc -le $cols; $hc++) { $formats.Add(@{ r = $rr; c = $hc; gray = $true }) }
        }
        foreach ($rr in $block.NewRows) {
            $formats.Add(@{ r = $rr; c = 7; bg = 0xCCF2FF; b = $true })
        }
    }
    $outline = Build-OutlineLevelMap $block 3
    $widths = Get-AutoColumnWidths -Rows $rows.ToArray() -FirstRow 2 -Min 10 -Max 90
    $lastCol = Get-ColumnName $cols
    return @{
        Rows    = $rows.ToArray()
        Widths  = $widths
        Formats = $formats.ToArray()
        Block   = $block
        Merges  = @(@{ Reference = ("A1:{0}1" -f $lastCol) })
        Outline = $outline
    }
}

function Build-DynamicsSheetData($Items, [string]$PeriodText) {
    $note = "Отработанные замечания за период. $PeriodText Группировка как на листах разделов: 1 — модель, 2 — спецификация, 3 — каждое исправленное замечание (когда найдено, когда исправлено, сколько дней висело)."
    $headers = @("Модель", "Отработано", "Спецификация", "Параметр", "Семейство и типоразмер", "Наименование", "Значение", "Найдено", "Исправлено", "Дней")
    $sorted = @($Items | Sort-Object Section, { Get-ModelSort $_.Model }, Model, Check, Param)
    $block = Build-DynamicsOutlineBlock $sorted 3
    $cols = 10
    $rows = New-Object System.Collections.Generic.List[object]
    $rows.Add(@($note))
    $rows.Add($headers)
    if ($block.Grid.GetLength(0) -lt 1) {
        $rows.Add(@("За выбранный период отработанных замечаний не найдено. Нужны минимум две выгрузки в истории."))
    }
    else {
        for ($r = 0; $r -lt $block.Grid.GetLength(0); $r++) {
            $line = New-Object object[] $cols
            for ($c = 0; $c -lt $cols; $c++) { $line[$c] = $block.Grid[$r, $c] }
            $rows.Add($line)
        }
    }
    $formats = New-Object System.Collections.Generic.List[object]
    $formats.Add(@{ r = 1; c = 1; bg = 0x00FFFF; wrap = $true })
    for ($hc = 1; $hc -le $cols; $hc++) { $formats.Add(@{ r = 2; c = $hc; bg = 0xF2E1D9; b = $true; wrap = $true }) }
    $last = 2 + $block.Grid.GetLength(0)
    if ($block.Grid.GetLength(0) -lt 1) { $last = 3 }
    for ($rr = 3; $rr -le $last; $rr++) {
        $name = ""
        if ($rows.Count -ge $rr) { $name = [string]$rows[$rr - 1][0] }
        if (-not [string]::IsNullOrWhiteSpace($name)) {
            for ($hc = 1; $hc -le $cols; $hc++) { $formats.Add(@{ r = $rr; c = $hc; bg = 0xCEEFC6; b = ($hc -le 2) }) }
        }
        elseif ($rows.Count -ge $rr -and -not [string]::IsNullOrWhiteSpace([string]$rows[$rr - 1][1])) {
            $formats.Add(@{ r = $rr; c = 2; b = $true })
        }
    }
    $outline = if ($block.Grid.GetLength(0) -gt 0) { Build-OutlineLevelMap $block 3 } else { @{} }
    $widths = Get-AutoColumnWidths -Rows $rows.ToArray() -FirstRow 2 -Min 10 -Max 90
    return @{
        Rows    = $rows.ToArray()
        Widths  = $widths
        Formats = $formats.ToArray()
        Merges  = @(@{ Reference = "A1:J1" })
        Outline = $outline
    }
}

# Выборочная выгрузка: в заметке листа раздела — какие модели в этой сборке не (полностью) проверялись и когда проверялись последний раз.
function Get-StaleModelsNote($Items, $Partial) {
    if ($null -eq $Partial) { return "" }
    $byModel = @{}
    foreach ($it in @($Items)) {
        if (-not $byModel.ContainsKey($it.Model)) { $byModel[$it.Model] = @{} }
        $byModel[$it.Model][[string]$it.Check] = $true
    }
    $parts = New-Object System.Collections.Generic.List[string]
    foreach ($m in @($byModel.Keys | Sort-Object { Get-ModelSort $_ }, { $_ })) {
        $notChecked = 0
        $oldest = ""
        foreach ($c in @($byModel[$m].Keys)) {
            if ($Partial.CheckedPairs.ContainsKey(([string]$m + [char]0x1F + $c))) { continue }
            $notChecked++
            $when = Get-PairChecked $Partial.Checks $m $c
            if ($when -and (-not $oldest -or $when -lt $oldest)) { $oldest = $when }
        }
        if ($notChecked -lt 1) { continue }
        $txt = $m + ": не проверялось " + $notChecked + " из " + $byModel[$m].Count
        if ($oldest) { $txt += ", последняя проверка " + (Format-DayRu $oldest) }
        $parts.Add($txt)
    }
    if ($parts.Count -lt 1) { return "" }
    $shown = @($parts | Select-Object -First 8)
    $tail = ""
    if ($parts.Count -gt 8) { $tail = " … и ещё " + ($parts.Count - 8) }
    return ("Выборочная выгрузка. Не обновлялись в этой сборке (данные прошлых проверок): " + ($shown -join "; ") + $tail + ".")
}

function Save-GooglePushPayload($Writer, [string]$Path, $ClearSections = $null) {
    $body = @{
        sheets = @($Writer.Payload.ToArray())
    }
    if ($null -ne $ClearSections) {
        # выборочная выгрузка: в Google уходят только Сводка и затронутые разделы, очищаются только затронутые без замечаний
        $have = @{}
        foreach ($sh in @($Writer.Payload.ToArray())) { $have[[string]$sh.name] = $true }
        $body.partial = $true
        $body.clear = @($ClearSections | Where-Object { -not $have.ContainsKey([string]$_) })
    }
    $json = $body | ConvertTo-Json -Depth 30 -Compress
    $utf8 = New-Object System.Text.UTF8Encoding $false
    [System.IO.File]::WriteAllText($Path, $json, $utf8)
}

function Export-AssignmentWorkbookOpenXml {
    param(
        [string]$OutXlsx,
        [string]$ExportDir,
        [string]$PayloadPath,
        $Stream,
        [int]$IssueCount,
        $Prev,
        $DateCols,
        [string]$Today,
        $DynItems,
        $Period,
        [string[]]$SectionOrder,
        [hashtable]$SectionTitle,
        [char]$Sep,
        $ByDateModelSum,
        [switch]$DynamicsOnly,
        $Ledger = $null,
        $Deltas = $null,
        [string]$BuildId = "",
        [int]$KeepFixedDays = 30,
        # выборочная выгрузка: @{ Checks; CheckedPairs; BuildId; StaleDays; TouchedSections } (в режиме «Всё» не передаётся)
        $PartialInfo = $null
    )
    $writer = New-AssignmentXlsxWriter $OutXlsx
    try {
        if ($DynamicsOnly) {
            $data = Build-DynamicsSheetData $DynItems $Period.Label
            Add-XlsxSheetFromTable $writer "Динамика" $data.Rows $data.Widths $data.Formats $null 2 $data.Merges $data.Outline -BorderFromRow 2
        }
        else {
            $sections = @(Get-StreamSections $Stream $SectionOrder)
            if ($Ledger -and $BuildId) {
                # разделы, где сейчас замечаний нет, но есть недавно исправленные
                $withFixed = Get-SectionsWithFixed $Ledger $BuildId $KeepFixedDays
                $sections = @($SectionOrder | Where-Object { ($sections -contains $_) -or $withFixed.ContainsKey($_) })
            }
            if ($IssueCount -lt 1 -and $Prev.Count -lt 1) {
                $emptyNote = "Замечаний не осталось."
                if ($ExportDir) {
                    $resultText = (Get-BvsExportResult $ExportDir).ResultText          # export.txt [result] (прежде export_result.txt)
                    if ($resultText) {
                        if ($resultText -match 'exported=0') {
                            $emptyNote = "Строк не выгружено. Спецификации в сводной пустые: связи не попали в таблицы. Это не закрытые замечания."
                        }
                    }
                }
                Add-XlsxSheetFromTable $writer "Сводка" (,@($emptyNote)) @(48) @() $null 1 @()
            }
            else {
                $sumData = Build-SummarySheetData $Stream $Prev $DateCols $Today $Sep $ByDateModelSum $Deltas $PartialInfo
                Add-XlsxSheetFromTable $writer "Сводка" $sumData.Rows $sumData.Widths $sumData.Formats $sumData.Meta 2 $sumData.Merges -BorderFromRow 2 -RowHeights $sumData.RowHeights
            }
            foreach ($sec in $sections) {
                $title = $SectionTitle[$sec]
                if (-not $title) { $title = $sec }
                $items = @(Read-SectionIssuesFromStream $Stream.SectionDir $sec)
                if ($Ledger -and $BuildId) {
                    # xlsx: открытые + исправленные (серым), с колонками статуса; Google получает только открытые, 7 колонок
                    $full = @(Get-SectionItemsWithStatus $Ledger $items $sec $BuildId $KeepFixedDays)
                    $staleNote = Get-StaleModelsNote $items $PartialInfo
                    $data = Build-SectionSheetData $full -Wide -KeepFixedDays $KeepFixedDays -ExtraNote $staleNote
                    # в выборочной выгрузке листы разделов без проверенных моделей в Google не отправляются (экономия времени Apps Script)
                    $toGoogle = ($null -eq $PartialInfo) -or ($PartialInfo.TouchedSections -contains $sec)
                    if ($items.Count -gt 0 -and $toGoogle) {
                        $payData = Build-SectionSheetData $items -ExtraNote $staleNote
                        Add-XlsxSheetFromTable $writer $title $data.Rows $data.Widths $data.Formats $null 2 $data.Merges $data.Outline -BorderFromRow 2 -PayloadData $payData
                    }
                    else {
                        Add-XlsxSheetFromTable $writer $title $data.Rows $data.Widths $data.Formats $null 2 $data.Merges $data.Outline -BorderFromRow 2 -SkipPayload
                    }
                }
                else {
                    $data = Build-SectionSheetData $items
                    Add-XlsxSheetFromTable $writer $title $data.Rows $data.Widths $data.Formats $null 2 $data.Merges $data.Outline -BorderFromRow 2
                }
            }
            $dynData = Build-DynamicsSheetData $DynItems $Period.Label
            Add-XlsxSheetFromTable $writer "Динамика" $dynData.Rows $dynData.Widths $dynData.Formats $null 2 $dynData.Merges $dynData.Outline -BorderFromRow 2
        }
        if ($PayloadPath) {
            if ($null -ne $PartialInfo) { Save-GooglePushPayload $writer $PayloadPath @($PartialInfo.TouchedSections | ForEach-Object { $SectionTitle[$_] }) }
            else { Save-GooglePushPayload $writer $PayloadPath }
        }
    }
    finally {
        Close-AssignmentXlsxWriter $writer
    }
}
