# UTF-8 WITH BOM required (PowerShell 5.1).
# Журнал замечаний: у каждого замечания (строка проверочной спецификации: модель + спецификация + семейство и
# типоразмер + наименование) есть дата первого обнаружения, статус (open/fixed) и дата исправления.
# Обновляется при каждой сборке задания. Файл: задание_журнал.txt рядом с историей.
#
# Колонки файла (TAB): Model, Section, Check, Family, Item, ParamName, Value, FirstSeen, LastSeen, Status, FixedAt, Reopened
# FirstSeen/LastSeen/FixedAt — идентификатор сборки «yyyy-MM-dd HH:mm:ss».

# Имя спецификации (см. assignment_export_stream.ps1); запасное определение, если файл подключён отдельно
if (-not (Get-Command ConvertTo-CanonicalSpecName -ErrorAction SilentlyContinue)) {
    function ConvertTo-CanonicalSpecName([string]$Name) { return ([string]$Name).Replace("значния", "значения").Replace(([string][char]0x43 + "антехприборы"), ([string][char]0x421 + "антехприборы")).Replace(([string][char]0x63 + "антехприборы"), ([string][char]0x441 + "антехприборы")) }
}

$script:LedgerKeySep = [char]0x1F
$script:LedgerRetainDays = 365

function Get-LedgerKey($Model, $Check, $Family, $Item) {
    $s = $script:LedgerKeySep
    return ([string]$Model + $s + [string]$Check + $s + [string]$Family + $s + [string]$Item)
}

function Read-IssueLedger([string]$Path) {
    $ledger = @{}
    if (-not (Test-Path -LiteralPath $Path)) { return $ledger }
    foreach ($line in @(Get-Content -LiteralPath $Path -Encoding UTF8)) {
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        $p = @($line.Split([char]9))
        if ($p.Count -lt 11) { continue }
        $reopened = 0
        if ($p.Count -ge 12) { [void][int]::TryParse($p[11], [ref]$reopened) }
        $rec = [pscustomobject]@{
            Model = $p[0]; Section = $p[1]; Check = (ConvertTo-CanonicalSpecName $p[2]); Family = $p[3]; Item = $p[4]
            ParamName = $p[5]; Value = $p[6]
            FirstSeen = $p[7]; LastSeen = $p[8]; Status = $p[9]; FixedAt = $p[10]; Reopened = $reopened
            Event = ""   # в памяти: что произошло в этой сборке (new / reopened / fixed)
        }
        $ledger[(Get-LedgerKey $rec.Model $rec.Check $rec.Family $rec.Item)] = $rec
    }
    return $ledger
}

function Write-IssueLedger([string]$Path, $Ledger, [string]$Today) {
    $limit = ([datetime]::Parse($Today)).AddDays(-$script:LedgerRetainDays).ToString("yyyy-MM-dd")
    $lines = New-Object System.Collections.Generic.List[string]
    foreach ($rec in @($Ledger.Values | Sort-Object Model, Check, Family, Item)) {
        # исправленные старше года из журнала убираем
        if ($rec.Status -eq "fixed" -and $rec.FixedAt -and $rec.FixedAt.Substring(0, 10) -lt $limit) { continue }
        $lines.Add(("{0}`t{1}`t{2}`t{3}`t{4}`t{5}`t{6}`t{7}`t{8}`t{9}`t{10}`t{11}" -f
                $rec.Model, $rec.Section, $rec.Check, $rec.Family, $rec.Item, $rec.ParamName, $rec.Value,
                $rec.FirstSeen, $rec.LastSeen, $rec.Status, $rec.FixedAt, $rec.Reopened))
    }
    $dir = Split-Path -Parent $Path
    if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    [System.IO.File]::WriteAllLines($Path, $lines.ToArray(), (New-Object System.Text.UTF8Encoding $true))
}

# Обновление журнала по текущей выгрузке.
#   $Issues      — все строки выгрузки этой сборки (Section, Model, Check, ParamName/Param, Family, Item, Value)
#   $CheckedSpecs — спецификации, проверенные в этой сборке (в т.ч. пустые). Замечание закрывается только если его
#                   спецификация проверена, иначе «не проверялось» и статус не меняется.
#   $Baseline     — журнал уже существовал: тогда новые замечания помечаются «Новое» (при первом запуске — нет)
#   $CheckedPairs — только выборочная выгрузка: проверенные пары «Модель<0x1F>Спецификация». Тогда $Issues — только строки
#                   этих пар, а закрывается лишь замечание проверенной пары (непроверенные модели остаются как были).
#                   В режиме «Всё» не передаётся и работает прежняя логика по спецификациям.
function Update-IssueLedger($Ledger, $Issues, $CheckedSpecs, [string]$BuildId, [bool]$Baseline, $CheckedPairs = $null) {
    $stats = @{ New = 0; Reopened = 0; Fixed = 0 }
    $present = @{}
    foreach ($it in @($Issues)) {
        $key = Get-LedgerKey $it.Model $it.Check $it.Family $it.Item
        if ($present.ContainsKey($key)) { continue }
        $present[$key] = $true
        $shown = [string]$it.ParamName
        if (-not $shown) { $shown = [string]$it.Param }
        if ($Ledger.ContainsKey($key)) {
            $rec = $Ledger[$key]
            $rec.LastSeen = $BuildId
            $rec.ParamName = $shown
            $rec.Value = [string]$it.Value
            if ($rec.Status -eq "fixed") {
                $rec.Status = "open"; $rec.FixedAt = ""; $rec.Reopened = [int]$rec.Reopened + 1
                $rec.Event = "reopened"; $stats.Reopened++
            }
        }
        else {
            $rec = [pscustomobject]@{
                Model = [string]$it.Model; Section = [string]$it.Section; Check = [string]$it.Check
                Family = [string]$it.Family; Item = [string]$it.Item; ParamName = $shown; Value = [string]$it.Value
                FirstSeen = $BuildId; LastSeen = $BuildId; Status = "open"; FixedAt = ""; Reopened = 0
                Event = $(if ($Baseline) { "new" } else { "" })
            }
            $Ledger[$key] = $rec
            if ($Baseline) { $stats.New++ }
        }
    }
    if ($null -ne $CheckedPairs) {
        if ($CheckedPairs.Count -gt 0) {
            foreach ($key in @($Ledger.Keys)) {
                $rec = $Ledger[$key]
                if ($rec.Status -ne "open" -or $present.ContainsKey($key)) { continue }
                if (-not $CheckedPairs.ContainsKey(([string]$rec.Model + $script:LedgerKeySep + [string]$rec.Check))) { continue }
                $rec.Status = "fixed"; $rec.FixedAt = $BuildId; $rec.Event = "fixed"; $stats.Fixed++
            }
        }
    }
    elseif ($CheckedSpecs -and $CheckedSpecs.Count -gt 0) {
        foreach ($key in @($Ledger.Keys)) {
            $rec = $Ledger[$key]
            if ($rec.Status -ne "open" -or $present.ContainsKey($key)) { continue }
            if (-not $CheckedSpecs.ContainsKey($rec.Check)) { continue }
            $rec.Status = "fixed"; $rec.FixedAt = $BuildId; $rec.Event = "fixed"; $stats.Fixed++
        }
    }
    return $stats
}

# Изменения этой сборки по паре «модель + спецификация»: Fixed / New (для пометок ↓ ↑ в Сводке)
function Get-LedgerDeltas($Ledger) {
    $d = @{}
    foreach ($rec in @($Ledger.Values)) {
        if (-not $rec.Event) { continue }
        $k = [string]$rec.Model + $script:LedgerKeySep + [string]$rec.Check
        if (-not $d.ContainsKey($k)) { $d[$k] = @{ Fixed = 0; New = 0 } }
        if ($rec.Event -eq "fixed") { $d[$k].Fixed++ } else { $d[$k].New++ }
    }
    return $d
}

function Get-DayOf([string]$BuildId) {
    if ([string]::IsNullOrEmpty($BuildId) -or $BuildId.Length -lt 10) { return "" }
    return $BuildId.Substring(0, 10)
}

function Format-DayRu([string]$BuildId) {
    $d = Get-DayOf $BuildId
    if (-not $d) { return "" }
    try { return ([datetime]::Parse($d)).ToString("dd.MM.yyyy") } catch { return $d }
}

function Get-DaysBetween([string]$FromBuild, [string]$ToBuild) {
    $a = Get-DayOf $FromBuild
    $b = Get-DayOf $ToBuild
    if (-not $a -or -not $b) { return "" }
    try { return [int]([datetime]::Parse($b) - [datetime]::Parse($a)).TotalDays } catch { return "" }
}

# Строки раздела для листа xlsx: открытые из выгрузки (все строки, как в Сводке) с данными журнала +
# исправленные за последние $KeepDays дней. Status: open / fixed.
function Get-SectionItemsWithStatus($Ledger, $OpenIssues, [string]$Section, [string]$BuildId, [int]$KeepDays) {
    $out = New-Object System.Collections.Generic.List[object]
    foreach ($it in @($OpenIssues)) {
        $rec = $Ledger[(Get-LedgerKey $it.Model $it.Check $it.Family $it.Item)]
        $first = $BuildId
        $isNew = $false
        if ($rec) { $first = $rec.FirstSeen; $isNew = ($rec.Event -eq "new" -or $rec.Event -eq "reopened") }
        $out.Add([pscustomobject]@{
            Section = $it.Section; Model = $it.Model; Check = $it.Check; Param = $it.Param; ParamName = $it.ParamName
            Family = $it.Family; Item = $it.Item; Value = $it.Value
            Status = "open"; IsNew = $isNew; FirstSeen = $first; FixedAt = ""
        })
    }
    $limit = ([datetime]::Parse((Get-DayOf $BuildId))).AddDays(-$KeepDays).ToString("yyyy-MM-dd")
    $fixed = @($Ledger.Values | Where-Object {
            $_.Status -eq "fixed" -and $_.Section -eq $Section -and $_.FixedAt -and (Get-DayOf $_.FixedAt) -ge $limit
        } | Sort-Object Model, Check, @{ Expression = { $_.FixedAt }; Descending = $true }, Family, Item)
    foreach ($rec in $fixed) {
        $out.Add([pscustomobject]@{
            Section = $rec.Section; Model = $rec.Model; Check = $rec.Check; Param = $rec.ParamName; ParamName = $rec.ParamName
            Family = $rec.Family; Item = $rec.Item; Value = $rec.Value
            Status = "fixed"; IsNew = $false; FirstSeen = $rec.FirstSeen; FixedAt = $rec.FixedAt
        })
    }
    return $out.ToArray()
}

# Коды разделов, у которых есть исправленные (в пределах $KeepDays) замечания
function Get-SectionsWithFixed($Ledger, [string]$BuildId, [int]$KeepDays) {
    $limit = ([datetime]::Parse((Get-DayOf $BuildId))).AddDays(-$KeepDays).ToString("yyyy-MM-dd")
    $set = @{}
    foreach ($rec in @($Ledger.Values)) {
        if ($rec.Status -eq "fixed" -and $rec.FixedAt -and (Get-DayOf $rec.FixedAt) -ge $limit) { $set[[string]$rec.Section] = $true }
    }
    return $set
}

# Исправленные в периоде [From..To] (yyyy-MM-dd) — для листа «Динамика»
function Get-LedgerFixedInPeriod($Ledger, [string]$From, [string]$To) {
    # сортировка — только ради детерминированного порядка строк на листе «Динамика» (перебор hashtable случаен)
    return @($Ledger.Values | Where-Object {
            $_.Status -eq "fixed" -and $_.FixedAt -and (Get-DayOf $_.FixedAt) -ge $From -and (Get-DayOf $_.FixedAt) -le $To
        } | Sort-Object Model, Check, Family, Item, FirstSeen)
}
