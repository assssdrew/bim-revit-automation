# UTF-8 WITH BOM required (PowerShell 5.1).
# Проверка сборки задания и целостности xlsx.

if (-not $script:BvsGuardToolDir) {
    $script:BvsGuardToolDir = $PSScriptRoot
    if ([string]::IsNullOrWhiteSpace($script:BvsGuardToolDir) -and $MyInvocation.PSCommandPath) {
        $script:BvsGuardToolDir = Split-Path -Parent $MyInvocation.PSCommandPath
    }
}

function Test-BvsAssignmentBuilderReady {
    param([string]$ToolDir)
    if ([string]::IsNullOrWhiteSpace($ToolDir)) { return $false }
    $build = Join-Path $ToolDir "build_assignment.ps1"
    $xlsxMod = Join-Path $ToolDir "assignment_xlsx.ps1"
    $dll = Join-Path (Join-Path (Split-Path -Parent $ToolDir) "lib") "DocumentFormat.OpenXml.dll"     # служебное\lib рядом со служебное\скрипты
    $ledgerMod = Join-Path $ToolDir "assignment_ledger.ps1"
    if (-not (Test-Path -LiteralPath $ledgerMod)) { return $false }
    if (-not (Test-Path -LiteralPath $build)) { return $false }
    if (-not (Test-Path -LiteralPath $xlsxMod)) { return $false }
    if (-not (Test-Path -LiteralPath $dll)) { return $false }
    return $true
}

function Get-BvsAssignmentBuilderMessage {
    param([string]$ToolDir)
    if (-not (Test-BvsAssignmentBuilderReady $ToolDir)) {
        return @(
            "Папка инструмента устарела или неполная.",
            "Нужны: build_assignment.ps1, assignment_xlsx.ps1, assignment_ledger.ps1, ..\lib\DocumentFormat.OpenXml.dll",
            "Путь: " + $ToolDir
        ) -join "`r`n"
    }
    return ""
}

function Get-BvsOpenXmlDllPath {
    return Join-Path (Join-Path (Split-Path -Parent $script:BvsGuardToolDir) "lib") "DocumentFormat.OpenXml.dll"
}

function Test-BvsAssignmentXlsxOpenXmlValidator {
    param([string]$Path)
    $dll = Get-BvsOpenXmlDllPath
    if (-not (Test-Path -LiteralPath $dll)) { return $true }
    try {
        Add-Type -Path $dll -ErrorAction Stop
        $doc = [DocumentFormat.OpenXml.Packaging.SpreadsheetDocument]::Open($Path, $false)
        try {
            $validator = New-Object DocumentFormat.OpenXml.Validation.OpenXmlValidator
            foreach ($err in $validator.Validate($doc)) {
                return $false
            }
            return $true
        }
        finally {
            $doc.Dispose()
        }
    }
    catch {
        return $false
    }
}

function Move-BvsAssignmentXlsxToBad {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return $null }
    $dir = Split-Path -Parent $Path
    $base = [System.IO.Path]::GetFileNameWithoutExtension($Path)
    $bad = Join-Path $dir ($base + ".bad.xlsx")
    if (Test-Path -LiteralPath $bad) { Remove-Item -LiteralPath $bad -Force -ErrorAction SilentlyContinue }
    Move-Item -LiteralPath $Path -Destination $bad -Force
    return $bad
}

function Test-BvsAssignmentXlsx {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return $false }
    if (-not (Test-Path -LiteralPath $Path)) { return $false }
    $fi = Get-Item -LiteralPath $Path
    if ($fi.Length -lt 2048) { return $false }
    try {
        Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop
    }
    catch { }
    try {
        $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
        try {
            $names = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
            foreach ($e in $zip.Entries) { [void]$names.Add($e.FullName) }
            if (-not $names.Contains("[Content_Types].xml")) { return $false }
            if (-not $names.Contains("xl/workbook.xml")) { return $false }
            foreach ($entry in $zip.Entries) {
                if ($entry.FullName -notmatch '^xl/worksheets/sheet\d+\.xml$') { continue }
                $stream = $entry.Open()
                try {
                    $reader = New-Object System.IO.StreamReader($stream)
                    $xml = $reader.ReadToEnd()
                    $reader.Close()
                    if ($xml -match 't\s*=\s*"str"') { return $false }
                    if ($xml -match 'sheetViews' -and $xml -match 'sheetPr') {
                        $iv = $xml.IndexOf('sheetViews', [StringComparison]::OrdinalIgnoreCase)
                        $ip = $xml.IndexOf('sheetPr', [StringComparison]::OrdinalIgnoreCase)
                        if ($ip -gt $iv) { return $false }
                    }
                    if ($xml -notmatch '<(x:)?row[\s>]') { return $false }
                }
                finally {
                    $stream.Dispose()
                }
            }
            if (-not (Test-BvsAssignmentXlsxOpenXmlValidator $Path)) { return $false }
            return $true
        }
        finally {
            $zip.Dispose()
        }
    }
    catch {
        return $false
    }
}

function Assert-BvsAssignmentXlsx {
    param([string]$Path)
    if (-not (Test-BvsAssignmentXlsx $Path)) {
        $bad = Move-BvsAssignmentXlsxToBad $Path
        $hint = ""
        if ($bad) { $hint = " Сохранён как: " + $bad }
        throw ("Файл задания не прошёл проверку." + $hint + " Пересоберите задание или обновите batch_verification_schedules.")
    }
}
