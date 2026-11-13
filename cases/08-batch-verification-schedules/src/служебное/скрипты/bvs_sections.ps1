# UTF-8 WITH BOM required (PowerShell 5.1).
# Единый список разделов. Читает разделы.txt рядом с этим файлом.
# Нет файла или он не читается — встроенный список (тот же, что в разделы.txt) и предупреждение. Исключение наружу не бросает.

function Get-BvsSectionsDir {
    if ($PSScriptRoot) { return $PSScriptRoot }
    return (Split-Path -Parent $MyInvocation.MyCommand.Path)
}

# Порядок массива = порядок листов задания. Column = порядок столбцов окна.
function Get-BvsBuiltinSectionRows {
    return @(
        @{ Code = "EOM"; Title = "ЭОМ"; Names = @("ЭОМ", "EOM"); Column = 8 },
        @{ Code = "SS"; Title = "СС"; Names = @("СС", "SS"); Column = 9 },
        @{ Code = "OV"; Title = "ОВ"; Names = @("ОВ", "OV"); Column = 1 },
        @{ Code = "PD"; Title = "ПД"; Names = @("ПД", "PD"); Column = 2 },
        @{ Code = "PT"; Title = "ПТ"; Names = @("ПТ", "PT"); Column = 4 },
        @{ Code = "VK"; Title = "ВК"; Names = @("ВК", "VK"); Column = 3 },
        @{ Code = "OT"; Title = "ОТ"; Names = @("ОТ", "OT"); Column = 5 },
        @{ Code = "HS"; Title = "ХС"; Names = @("ХС", "HS"); Column = 6 },
        @{ Code = "ITP"; Title = "ИТП"; Names = @("ИТП", "ITP"); Column = 7 },
        @{ Code = "AVT"; Title = "АВТ"; Names = @("АВТ", "AVT"); Column = 10 },
        @{ Code = "HC"; Title = "ХЦ"; Names = @("ХЦ", "HC"); Column = 11 }
    )
}

function New-BvsSectionRow([string]$Code, [string]$Title, $Names, [int]$Column, [int]$Report) {
    $list = New-Object System.Collections.Generic.List[string]
    foreach ($n in @($Names)) { if ($n) { [void]$list.Add([string]$n) } }
    return [pscustomobject]@{ Code = $Code; Title = $Title; Names = @($list.ToArray()); Column = $Column; Report = $Report }
}

function Read-BvsSectionFile([string]$Path) {
    $text = [System.IO.File]::ReadAllText($Path, (New-Object System.Text.UTF8Encoding $true))
    $rows = New-Object System.Collections.Generic.List[object]
    $i = 0
    foreach ($raw in @($text -split "`r?`n")) {
        $line = $raw.Trim()
        if (-not $line -or $line.StartsWith("#")) { continue }
        $p = @($line.Split([char]"|"))
        if ($p.Count -lt 2) { throw ("строка без кода и названия: " + $line) }
        $code = $p[0].Trim()
        $title = $p[1].Trim()
        if (-not $code -or -not $title) { throw ("пустой код или название: " + $line) }
        $names = New-Object System.Collections.Generic.List[string]
        if ($p.Count -ge 3 -and $p[2].Trim()) {
            foreach ($n in @($p[2].Split(","))) { if ($n.Trim()) { [void]$names.Add($n.Trim()) } }
        }
        if ($names.Count -eq 0) { [void]$names.Add($title); [void]$names.Add($code) }
        $col = 0
        if ($p.Count -ge 4 -and $p[3].Trim()) { [void][int]::TryParse($p[3].Trim(), [ref]$col) }
        $i++
        if ($col -lt 1) { $col = 1000 + $i }
        [void]$rows.Add((New-BvsSectionRow $code $title $names.ToArray() $col $i))
    }
    if ($rows.Count -lt 1) { throw "в файле нет разделов" }
    return @($rows.ToArray())
}

function Get-BvsBuiltinSectionObjects {
    $rows = New-Object System.Collections.Generic.List[object]
    $i = 0
    foreach ($b in @(Get-BvsBuiltinSectionRows)) {
        $i++
        [void]$rows.Add((New-BvsSectionRow $b.Code $b.Title $b.Names ([int]$b.Column) $i))
    }
    return @($rows.ToArray())
}

# Пустой Path — файл рядом со скриптом (результат кэшируется). Явный путь кэшем не затирается.
function Read-BvsSections([string]$Path = "") {
    $explicit = -not [string]::IsNullOrWhiteSpace($Path)
    if (-not $explicit -and $script:BvsSectionsCache) { return @($script:BvsSectionsCache) }
    $file = $Path
    if (-not $explicit) { $file = Join-Path (Get-BvsSectionsDir) "разделы.txt" }
    try {
        if (-not (Test-Path -LiteralPath $file)) { throw ("нет файла " + $file) }
        $rows = @(Read-BvsSectionFile $file)
    }
    catch {
        Write-Warning ("разделы.txt не прочитан, встроенный список разделов. " + $_.Exception.Message)
        $rows = @(Get-BvsBuiltinSectionObjects)
    }
    if (-not $explicit) { $script:BvsSectionsCache = $rows }
    return @($rows)
}

function Get-BvsReportSections([string]$Path = "") { return @(Read-BvsSections $Path) }

function Get-BvsColumnSections([string]$Path = "") {
    return @(Read-BvsSections $Path | Sort-Object Column, Report)
}

function Get-BvsReportCodes([string]$Path = "") {
    return @(Get-BvsReportSections $Path | ForEach-Object { [string]$_.Code })
}

function Get-BvsColumnCodes([string]$Path = "") {
    return @(Get-BvsColumnSections $Path | ForEach-Object { [string]$_.Code })
}

function Get-BvsTitleMap([string]$Path = "") {
    $m = @{}
    foreach ($s in @(Get-BvsReportSections $Path)) { $m[[string]$s.Code] = [string]$s.Title }
    return $m
}

function Get-BvsGoogleSheetNames([string]$Path = "") {
    $names = New-Object System.Collections.Generic.List[string]
    [void]$names.Add("Сводка")
    foreach ($s in @(Get-BvsReportSections $Path)) { [void]$names.Add([string]$s.Title) }
    return @($names.ToArray())
}

function Get-BvsPcSections([string]$Path = "") {
    $list = New-Object System.Collections.Generic.List[object]
    foreach ($s in @(Get-BvsColumnSections $Path)) {
        $names = New-Object System.Collections.Generic.List[string]
        foreach ($n in @($s.Names)) { if ($n) { [void]$names.Add([string]$n) } }
        [void]$list.Add(@{ Code = [string]$s.Code; Title = [string]$s.Title; Names = @($names.ToArray()) })
    }
    return @($list.ToArray())
}
