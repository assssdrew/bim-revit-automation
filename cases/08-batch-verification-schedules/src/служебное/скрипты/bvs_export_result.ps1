# UTF-8 WITH BOM required (PowerShell 5.1).
# Итог выгрузки из сводной (пишет export_schedules.py в служебную папку прогона).
# Сейчас это один файл export.txt с разделами:
#   [result]   matched= / exported= / skipped= (и для «Выборочно» partial=1, pairs=, missing_specs=, unloaded_models=)
#   [manifest] строки «<файл>.tsv<TAB>спецификация<TAB>поля<TAB>поля фильтра», «*empty*<TAB>…», «*scope*<TAB>partial», «*pair*<TAB>модель<TAB>спецификация»
#   [catalog]  «spec<TAB>имя» и «model<TAB>имя<TAB>loaded|unloaded» (все спецификации и связи сводной)
# Файл пишется последним, одним разом: есть файл — выгрузка дошла до конца.
# Выгрузки прежних версий (три файла manifest.txt, export_result.txt, catalog.txt) читаются так же.
# Подключать через точку.

function Read-BvsExportResult([string]$Dir) {
    $r = [pscustomobject]@{
        Has        = $false          # выгрузка дошла до конца (есть export.txt или прежний manifest.txt)
        Format     = ""              # "export.txt" | "legacy"
        Manifest   = (New-Object System.Collections.Generic.List[string])
        ResultText = ""              # строки [result] как текст «ключ=значение» через CRLF (как прежний export_result.txt)
        Result     = @{}
        HasCatalog = $false          # каталог записан (прежний catalog.txt или раздел [catalog])
        Catalog    = (New-Object System.Collections.Generic.List[string])
        LegacyCatalogPath = ""
    }
    if (-not $Dir) { return $r }
    $enc = New-Object System.Text.UTF8Encoding $true
    $merged = Join-Path $Dir "export.txt"
    if (Test-Path -LiteralPath $merged) {
        $r.Has = $true
        $r.Format = "export.txt"
        $section = ""
        $res = New-Object System.Collections.Generic.List[string]
        foreach ($line in [System.IO.File]::ReadAllLines($merged, $enc)) {
            $t = $line.Trim()
            if ($t -eq "[result]" -or $t -eq "[manifest]" -or $t -eq "[catalog]") {
                $section = $t.Substring(1, $t.Length - 2)
                if ($section -eq "catalog") { $r.HasCatalog = $true }
                continue
            }
            if (-not $t) { continue }
            switch ($section) {
                "result" { $res.Add($t) }
                "manifest" { $r.Manifest.Add($line) }
                "catalog" { $r.Catalog.Add($line) }
            }
        }
        $r.ResultText = $(if ($res.Count) { ($res -join "`r`n") + "`r`n" } else { "" })
    }
    else {
        $manifest = Join-Path $Dir "manifest.txt"
        if (Test-Path -LiteralPath $manifest) {
            $r.Has = $true
            $r.Format = "legacy"
            foreach ($line in [System.IO.File]::ReadAllLines($manifest, $enc)) { $r.Manifest.Add($line) }
        }
        $resultFile = Join-Path $Dir "export_result.txt"
        if (Test-Path -LiteralPath $resultFile) { $r.ResultText = [System.IO.File]::ReadAllText($resultFile) }
        $cat = Join-Path $Dir "catalog.txt"
        if (Test-Path -LiteralPath $cat) {
            $r.HasCatalog = $true
            $r.LegacyCatalogPath = $cat
            foreach ($line in [System.IO.File]::ReadAllLines($cat, $enc)) { $r.Catalog.Add($line) }
        }
    }
    foreach ($kv in @($r.ResultText -split "`r?`n")) {
        if ($kv -match '^([A-Za-z_]+)=(.*)$') { $r.Result[$Matches[1]] = $Matches[2].Trim() }
    }
    return $r
}

# То же с запоминанием: за одну сборку итог читается один раз (поток, каталог, буфер, лист «Сводка», итог окна).
# Ключ — папка и время/размер файла итога: новая выгрузка в ту же папку читается заново.
$script:BvsExportCache = @{}
function Get-BvsExportResult([string]$Dir) {
    if (-not $Dir) { return (Read-BvsExportResult $Dir) }
    $stamp = ""
    foreach ($n in @("export.txt", "manifest.txt", "export_result.txt", "catalog.txt")) {
        $p = Join-Path $Dir $n
        if (Test-Path -LiteralPath $p) { $fi = Get-Item -LiteralPath $p; $stamp += ("{0}:{1}:{2};" -f $n, $fi.LastWriteTimeUtc.Ticks, $fi.Length) }
    }
    $key = [System.IO.Path]::GetFullPath($Dir) + "|" + $stamp
    if ($null -eq $script:BvsExportCache) { $script:BvsExportCache = @{} }
    $hit = $script:BvsExportCache[$key]
    if ($null -ne $hit) { return $hit }
    $r = Read-BvsExportResult $Dir
    if ($script:BvsExportCache.Count -gt 8) { $script:BvsExportCache.Clear() }
    $script:BvsExportCache[$key] = $r
    return $r
}

# Выгрузка дошла до конца: есть итог (export.txt; у прежних версий — manifest.txt)
function Test-BvsExportDone([string]$Dir) {
    if (-not $Dir) { return $false }
    return ((Test-Path -LiteralPath (Join-Path $Dir "export.txt")) -or (Test-Path -LiteralPath (Join-Path $Dir "manifest.txt")))
}

# Каталог выгрузки в постоянный файл истории (<история>_каталог.txt) — в прежнем формате catalog.txt (UTF-8 с BOM, CRLF)
function Save-BvsExportCatalog($Export, [string]$Destination) {
    if (-not $Export -or -not $Export.HasCatalog -or -not $Destination) { return $false }
    if ($Export.LegacyCatalogPath) { Copy-Item -LiteralPath $Export.LegacyCatalogPath -Destination $Destination -Force; return $true }
    $text = (@($Export.Catalog) -join "`r`n") + "`r`n"
    [System.IO.File]::WriteAllText($Destination, $text, (New-Object System.Text.UTF8Encoding $true))
    return $true
}
