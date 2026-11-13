# UTF-8 WITH BOM required (PowerShell 5.1).
# Сравнение двух выгрузок фильтров (формат: Спецификация, Поле фильтра, Условие, Значение — как фильтры_спецификаций.txt
# или результат dump_schedule_filters.ps1). Функции вынесены сюда, чтобы их можно было проверять тестами.

$script:FilterSep = [char]0x1F
$script:NoFiltersMark = "(нет фильтров)"

function Get-WS([string]$Text) { return ([regex]::Replace(([string]$Text), '\s+', ' ')).Trim() }

# Ключ спецификации: имя без префикса «П» (в любом написании, с пробелами любого вида), без регистра, «значния» = «значения»
function Get-SpecIdentity([string]$Name) {
    $t = Get-WS $Name
    $t = [regex]::Replace($t, '^[ПпΠπ]\s+', '')
    $number = ""
    $m = [regex]::Match($t, '^(\d+(?:\.\d+)*)_')
    if ($m.Success) { $number = $m.Groups[1].Value }
    $key = $t.ToLowerInvariant().Replace("значния", "значения").Replace(([string][char]0x63 + "антехприборы"), ([string][char]0x441 + "антехприборы"))
    return [pscustomobject]@{ Key = $key; Number = $number; Display = $t }
}

function Format-FilterText($F) {
    $s = ("{0} {1}" -f $F.Field, $F.Cond).Trim()
    if ($F.Value -ne "") { $s += (" «" + $F.Value + "»") }
    return $s
}

# Читает выгрузку: список спецификаций с их фильтрами в порядке файла
function Read-FilterDump([string]$Path) {
    $specs = New-Object System.Collections.Generic.List[object]
    $byKey = @{}
    $seen = @{}
    $first = $true
    foreach ($line in @(Get-Content -LiteralPath $Path -Encoding UTF8)) {
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        $p = @($line.Split([char]9))
        if ($first) { $first = $false; if ($p[0].Trim() -eq "Спецификация") { continue } }
        if ($p.Count -lt 2) { continue }
        $id = Get-SpecIdentity $p[0]
        $key = $id.Key
        # два листа с одинаковым именем — разные записи (#2, #3…), строки одной спецификации идут подряд
        $rowKey = $p[0]
        if (-not $seen.ContainsKey($rowKey)) {
            $n = 1
            if ($byKey.ContainsKey($key)) { $n = $byKey[$key] + 1 }
            $byKey[$key] = $n
            $uniq = $key
            if ($n -gt 1) { $uniq = $key + "#" + $n }
            $rec = [pscustomobject]@{ Key = $uniq; BaseKey = $key; Number = $id.Number; Name = $id.Display; Filters = (New-Object System.Collections.Generic.List[object]) }
            $seen[$rowKey] = $rec
            $specs.Add($rec)
        }
        $rec = $seen[$rowKey]
        $field = Get-WS $p[1]
        if ($field -eq $script:NoFiltersMark) { continue }
        $cond = ""; $val = ""
        if ($p.Count -ge 3) { $cond = (Get-WS $p[2]) }
        if ($p.Count -ge 4) { $val = ([string]$p[3]).Trim() }
        $rec.Filters.Add([pscustomobject]@{ Field = $field; Cond = $cond; Value = $val })
    }
    foreach ($r in $specs) { $r.Filters = $r.Filters.ToArray() }      # массив, не List: надёжнее при @(...) в PowerShell 5.1/7
    return $specs
}

function Get-NumberSortKey([string]$Number, [string]$Name) {
    $parts = @()
    if ($Number) { $parts = @($Number.Split('.') | ForEach-Object { $_.PadLeft(6, '0') }) } else { $parts = @("~") }
    return (($parts -join '.') + "|" + $Name.ToLowerInvariant())
}

# Различия фильтров одной спецификации: порядок фильтров не важен, повторы учитываются
function Compare-SpecFilters($OldFilters, $NewFilters) {
    $sep = $script:FilterSep
    $count = @{}
    foreach ($f in @($NewFilters)) {
        $k = $f.Field.ToLowerInvariant() + $sep + $f.Cond.ToLowerInvariant() + $sep + $f.Value
        if (-not $count.ContainsKey($k)) { $count[$k] = 0 }
        $count[$k]++
    }
    $oldLeft = New-Object System.Collections.Generic.List[object]
    foreach ($f in @($OldFilters)) {
        $k = $f.Field.ToLowerInvariant() + $sep + $f.Cond.ToLowerInvariant() + $sep + $f.Value
        if ($count.ContainsKey($k) -and $count[$k] -gt 0) { $count[$k]-- } else { $oldLeft.Add($f) }
    }
    $newLeft = New-Object System.Collections.Generic.List[object]
    # что осталось в «стало»: полный список «стало» минус совпавшие
    $matchedCount = @{}
    foreach ($f in @($OldFilters)) {
        $k = $f.Field.ToLowerInvariant() + $sep + $f.Cond.ToLowerInvariant() + $sep + $f.Value
        if (-not $matchedCount.ContainsKey($k)) { $matchedCount[$k] = 0 }
        $matchedCount[$k]++
    }
    $taken = @{}
    foreach ($f in @($NewFilters)) {
        $k = $f.Field.ToLowerInvariant() + $sep + $f.Cond.ToLowerInvariant() + $sep + $f.Value
        if (-not $taken.ContainsKey($k)) { $taken[$k] = 0 }
        $oldHas = 0
        if ($matchedCount.ContainsKey($k)) { $oldHas = $matchedCount[$k] }
        if ($taken[$k] -lt $oldHas) { $taken[$k]++ } else { $newLeft.Add($f) }
    }
    $out = New-Object System.Collections.Generic.List[object]
    $oldRest = New-Object System.Collections.Generic.List[object]
    foreach ($o in $oldLeft) { $oldRest.Add($o) }
    $newRest = New-Object System.Collections.Generic.List[object]
    foreach ($n in $newLeft) { $newRest.Add($n) }
    # 1) то же поле, другое условие и/или значение
    foreach ($o in @($oldRest.ToArray())) {
        $hit = $null
        foreach ($n in $newRest) { if ($n.Field.ToLowerInvariant() -eq $o.Field.ToLowerInvariant()) { $hit = $n; break } }
        if ($null -eq $hit) { continue }
        $what = @()
        if ($o.Cond.ToLowerInvariant() -ne $hit.Cond.ToLowerInvariant()) { $what += "условие" }
        if ($o.Value -ne $hit.Value) { $what += "значение" }
        $out.Add([pscustomobject]@{ Change = ("изменено: " + ($what -join " и ")); Was = (Format-FilterText $o); Became = (Format-FilterText $hit) })
        [void]$oldRest.Remove($o)
        [void]$newRest.Remove($hit)
    }
    # 2) то же условие и значение, другое поле
    foreach ($o in @($oldRest.ToArray())) {
        $hit = $null
        foreach ($n in $newRest) {
            if ($n.Cond.ToLowerInvariant() -eq $o.Cond.ToLowerInvariant() -and $n.Value -eq $o.Value) { $hit = $n; break }
        }
        if ($null -eq $hit) { continue }
        $out.Add([pscustomobject]@{ Change = "изменено: поле"; Was = (Format-FilterText $o); Became = (Format-FilterText $hit) })
        [void]$oldRest.Remove($o)
        [void]$newRest.Remove($hit)
    }
    foreach ($o in $oldRest) { $out.Add([pscustomobject]@{ Change = "удалено"; Was = (Format-FilterText $o); Became = "" }) }
    foreach ($n in $newRest) { $out.Add([pscustomobject]@{ Change = "добавлено"; Was = ""; Became = (Format-FilterText $n) }) }
    return $out.ToArray()
}

# Сопоставление: по имени без «П» (с номером); что не нашлось — по номеру, если на каждой стороне осталась ровно одна спека с ним
function Compare-FilterDumps($OldSpecs, $NewSpecs) {
    $newByKey = @{}
    foreach ($s in @($NewSpecs)) { $newByKey[$s.Key] = $s }
    $pairs = New-Object System.Collections.Generic.List[object]
    $oldUn = New-Object System.Collections.Generic.List[object]
    $matchedNew = @{}
    foreach ($o in @($OldSpecs)) {
        if ($newByKey.ContainsKey($o.Key)) {
            $pairs.Add([pscustomobject]@{ Old = $o; New = $newByKey[$o.Key]; Note = "" })
            $matchedNew[$o.Key] = $true
        }
        else { $oldUn.Add($o) }
    }
    $newUn = New-Object System.Collections.Generic.List[object]
    foreach ($n in @($NewSpecs)) { if (-not $matchedNew.ContainsKey($n.Key)) { $newUn.Add($n) } }
    $onlyOld = New-Object System.Collections.Generic.List[object]
    $onlyNew = New-Object System.Collections.Generic.List[object]
    $stillNew = New-Object System.Collections.Generic.List[object]
    foreach ($n in $newUn) { $stillNew.Add($n) }
    foreach ($o in $oldUn) {
        $cand = @($stillNew | Where-Object { $o.Number -and $_.Number -eq $o.Number })
        $oldSame = @($oldUn | Where-Object { $o.Number -and $_.Number -eq $o.Number })
        if ($cand.Count -eq 1 -and $oldSame.Count -eq 1) {
            $pairs.Add([pscustomobject]@{ Old = $o; New = $cand[0]; Note = ("сопоставлено по номеру; имя было «" + $o.Name + "»") })
            [void]$stillNew.Remove($cand[0])
        }
        else { $onlyOld.Add($o) }
    }
    foreach ($n in $stillNew) { $onlyNew.Add($n) }
    $rows = New-Object System.Collections.Generic.List[object]
    $unchanged = 0
    foreach ($p in $pairs) {
        $diffs = @(Compare-SpecFilters $p.Old.Filters $p.New.Filters)
        if ($diffs.Count -eq 0) { $unchanged++; continue }
        foreach ($d in $diffs) {
            $rows.Add([pscustomobject]@{
                Number = $p.New.Number; Spec = $p.New.Name; Change = $d.Change; Was = $d.Was; Became = $d.Became; Note = $p.Note
                Sort = (Get-NumberSortKey $p.New.Number $p.New.Name)
            })
        }
    }
    foreach ($o in $onlyOld) {
        $rows.Add([pscustomobject]@{
            Number = $o.Number; Spec = $o.Name; Change = "спека только в «было»"; Was = ("фильтров: " + $o.Filters.Count); Became = ""; Note = ""
            Sort = (Get-NumberSortKey $o.Number $o.Name)
        })
    }
    foreach ($n in $onlyNew) {
        $rows.Add([pscustomobject]@{
            Number = $n.Number; Spec = $n.Name; Change = "спека только в «стало»"; Was = ""; Became = ("фильтров: " + $n.Filters.Count); Note = ""
            Sort = (Get-NumberSortKey $n.Number $n.Name)
        })
    }
    $sorted = @($rows | Sort-Object Sort, Change)
    return [pscustomobject]@{
        Rows = $sorted; Matched = $pairs.Count; Unchanged = $unchanged
        OnlyOld = $onlyOld.Count; OnlyNew = $onlyNew.Count
        ChangedSpecs = @($sorted | Where-Object { $_.Change -notlike "спека только*" } | ForEach-Object { $_.Spec } | Select-Object -Unique).Count
    }
}

function ConvertTo-CsvCell([string]$Text) {
    $t = [string]$Text
    if ($t.Contains(';') -or $t.Contains('"') -or $t.Contains("`n")) { return ('"' + $t.Replace('"', '""') + '"') }
    return $t
}

function Write-CompareCsv([string]$Path, $Rows) {
    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add("Номер;Спецификация;Изменение;Было;Стало;Примечание")
    foreach ($r in @($Rows)) {
        $lines.Add((@($r.Number, $r.Spec, $r.Change, $r.Was, $r.Became, $r.Note) | ForEach-Object { ConvertTo-CsvCell $_ }) -join ';')
    }
    $dir = Split-Path -Parent $Path
    if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    [System.IO.File]::WriteAllLines($Path, $lines.ToArray(), (New-Object System.Text.UTF8Encoding $true))
}
