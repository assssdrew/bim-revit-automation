# Name: ui_logic.ps1
# Version: 3.1.0
# What it does: See the case README.
# Notes: Paths and project names in this portfolio copy are placeholders.

# UTF-8 WITH BOM required (PowerShell 5.1).
# Pure logic for the operations window (presets.ps1): no WinForms here, so it is
# covered by tests\ui_logic.tests.ps1. Dot-source AFTER job_lib.ps1 (uses
# Get-BasenameNoExt, Set-FilenameInPath, New-OpsFromSample, Invoke-NameOps).

# ---------------------------------------------------------------- paths

function ConvertTo-UiPathKey([string]$Path) {
    if ([string]::IsNullOrWhiteSpace($Path)) { return "" }
    return $Path.Trim().Trim('"').Replace("\", "/").TrimEnd("/").ToUpperInvariant()
}

# Folder part of a model path; keeps the separator style of the input.
function Get-PathFolder([string]$Path) {
    if ([string]::IsNullOrWhiteSpace($Path)) { return "" }
    $p = $Path.Trim().Trim('"')
    $i = [Math]::Max($p.LastIndexOf("\"), $p.LastIndexOf("/"))
    if ($i -le 0) { return "" }
    return $p.Substring(0, $i)
}

# Longest common folder of several folders, compared by whole segments and
# without regard to case. Returns "" if they share nothing. Separator style
# follows the first folder.
function Get-CommonPathPrefix([string[]]$Folders) {
    $list = @($Folders | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($list.Count -eq 0) { return "" }
    $first = $list[0].Trim().Trim('"')
    $useBack = $first.Contains("\") -and -not $first.StartsWith("RSN://", [StringComparison]::OrdinalIgnoreCase)
    $parts = @()
    foreach ($f in $list) {
        $n = $f.Trim().Trim('"').Replace("\", "/").TrimEnd("/")
        $parts += , @($n.Split("/"))
    }
    $ref = $parts[0]
    $common = New-Object System.Collections.Generic.List[string]
    for ($k = 0; $k -lt $ref.Count; $k++) {
        $same = $true
        foreach ($q in $parts) {
            if ($k -ge $q.Count -or ($q[$k] -ine $ref[$k])) { $same = $false; break }
        }
        if (-not $same) { break }
        $common.Add($ref[$k]) | Out-Null
    }
    # only roots/schemes in common ("", "RSN:", "") is not a folder
    $real = @($common | Where-Object { $_ -ne "" -and $_ -ne "RSN:" })
    if ($real.Count -eq 0) { return "" }
    $joined = $common -join "/"
    if ($useBack) { $joined = $joined.Replace("/", "\") }
    return $joined
}

# Text for "Поменять кусок пути" -> "было": common folder of the checked
# models (or of all models when none is checked); if they share nothing,
# the folder of the first model.
function Get-SegmentSuggestion([string[]]$CheckedPaths, [string[]]$AllPaths) {
    $paths = @($CheckedPaths | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($paths.Count -eq 0) { $paths = @($AllPaths | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }) }
    if ($paths.Count -eq 0) { return "" }
    $folders = @($paths | ForEach-Object { Get-PathFolder $_ })
    $common = Get-CommonPathPrefix $folders
    if ($common) { return $common }
    return $folders[0]
}

# (name and path validators: Test-ModelFileName / Test-NewModelPath live in job_lib.ps1, which uses them for the preview)

# Rows: objects with old_path, new_path, old_name. SelfOldPath = the row being edited.
function Test-PathConflict([string]$NewPath, [string]$SelfOldPath, [object[]]$Rows) {
    $nk = ConvertTo-UiPathKey $NewPath
    $sk = ConvertTo-UiPathKey $SelfOldPath
    foreach ($r in @($Rows)) {
        if ($null -eq $r) { continue }
        if ((ConvertTo-UiPathKey $r.old_path) -eq $sk) { continue }
        if ((ConvertTo-UiPathKey $r.new_path) -eq $nk) {
            return New-CheckResult $false "duplicate_new" ("Такой новый путь уже стоит у другой модели: {0}." -f $r.old_name)
        }
        if ((ConvertTo-UiPathKey $r.old_path) -eq $nk) {
            return New-CheckResult $false "other_old" ("Этот путь занят другой моделью из списка ({0}): она была бы перезаписана." -f $r.old_name)
        }
    }
    return New-CheckResult $true "" ""
}

# ------------------------------------------------------------ the rule

# State: empty | same | ok | fail. Ops/Got filled when a rule was built.
function Get-RuleStatus([string]$Sample, [string]$Target) {
    $a = ""
    $b = ""
    if ($null -ne $Sample) { $a = (Get-BasenameNoExt $Sample.Trim()) }
    if ($null -ne $Target) { $b = (Get-BasenameNoExt $Target.Trim()) }
    if (-not $a -and -not $b) {
        return [pscustomobject]@{ State = "empty"; Ops = @(); Got = ""; Message = "" }
    }
    if (-not $a -or -not $b) {
        return [pscustomobject]@{ State = "fail"; Ops = @(); Got = ""; Message = "Заполните оба поля: «Имя сейчас» и «Должно стать»." }
    }
    if ($a -ceq $b) {
        return [pscustomobject]@{ State = "same"; Ops = @(); Got = $a; Message = "Имена одинаковые - переименовывать нечего." }
    }
    $ops = @(New-OpsFromSample $a $b)
    if ($ops.Count -eq 0) {
        return [pscustomobject]@{ State = "fail"; Ops = @(); Got = ""; Message = "Не удалось выделить правило из этих имён." }
    }
    $got = Invoke-NameOps -Name $a -Ops $ops
    if ($got -cne $b) {
        return [pscustomobject]@{
            State = "fail"; Ops = $ops; Got = $got
            Message = ("Правило не воспроизводит пример: получилось «{0}», нужно «{1}»." -f $got, $b)
        }
    }
    return [pscustomobject]@{ State = "ok"; Ops = $ops; Got = $got; Message = "" }
}

# ------------------------------------------------------- manual edits
# A manual entry is @{ name = <string or $null>; path = <string or $null> }.
# Invariant: if path is set, name equals the file name inside it.

function Set-ManualName($Entry, [string]$NewName) {
    $e = @{ name = (Get-ManualField $Entry "name"); path = (Get-ManualField $Entry "path") }
    $e.name = $NewName
    if ($e.path) { $e.path = Set-FilenameInPath $e.path $NewName }
    return $e
}

function Set-ManualPath($Entry, [string]$NewPath) {
    $e = @{ name = (Get-ManualField $Entry "name"); path = (Get-ManualField $Entry "path") }
    $e.path = $NewPath
    $e.name = Get-BasenameNoExt $NewPath
    return $e
}

# ------------------------------------------------------------ clipboard

# Mode: both | old | new. One row per line (CRLF), columns separated by TAB.
function Format-RowsForClipboard([object[]]$Rows, [string]$Mode = "both") {
    $lines = New-Object System.Collections.Generic.List[string]
    foreach ($r in @($Rows)) {
        if ($null -eq $r) { continue }
        switch ($Mode) {
            "old" { $lines.Add([string]$r.old_path) | Out-Null }
            "new" { $lines.Add([string]$r.new_path) | Out-Null }
            default {
                if ([string]::IsNullOrEmpty([string]$r.new_path)) { $lines.Add([string]$r.old_path) | Out-Null }
                else { $lines.Add(("{0}`t{1}" -f $r.old_path, $r.new_path)) | Out-Null }
            }
        }
    }
    return ($lines -join "`r`n")
}

# Which manual edits are applied right now. Job flags (rename / move) follow the
# checkboxes, so an edit of a column that is switched off is remembered but not applied:
# a path edit needs "Перенести", a name edit needs "Переименовать". A path edit also
# changes the file name inside it, so with "Переименовать" alone only that name is used.
function Get-EffectiveManualRows($Manual, [bool]$Rename, [bool]$Move) {
    $out = @{}
    if ($null -eq $Manual) { return $out }
    foreach ($k in @($Manual.Keys)) {
        $e = $Manual[$k]
        $p = Get-ManualField $e "path"
        $n = Get-ManualField $e "name"
        if ($p -and $Move) { $out[$k] = @{ name = $n; path = $p }; continue }
        if ($n -and $Rename) { $out[$k] = @{ name = $n; path = $null } }
    }
    return $out
}

# "<model>: <what is wrong with its path>" for the models that cannot be run (at most Max lines).
function Get-PathErrorLines([object[]]$Rows, [int]$Max = 5) {
    $out = New-Object System.Collections.Generic.List[string]
    foreach ($r in @($Rows)) {
        if ($null -eq $r) { continue }
        $e = ""
        if ($r.PSObject.Properties["path_error"] -and $r.path_error) { $e = [string]$r.path_error }
        if (-not $e) { continue }
        if ($out.Count -ge $Max) { break }
        $out.Add(("{0}: {1}" -f $r.old_name, $e)) | Out-Null
    }
    return $out.ToArray()
}

# Tooltip text for a table cell (columns: 1 Станет, 2 Состояние, 4 Старый путь, 5 Новый путь).
function Get-CellTipText($Row, [int]$Col) {
    if ($null -eq $Row) { return "" }
    $nl = "`r`n"
    $err = ""
    if ($Row.PSObject.Properties["path_error"] -and $Row.path_error) { $err = [string]$Row.path_error }
    switch ($Col) {
        1 {
            $t = [string]$Row.new_name
            if ($Row.manual) { return ($t + $nl + "Изменено вручную. По правилу было бы: " + $Row.rule_name + $nl + "Двойной щелчок - править, пустое значение - вернуть правило.") }
            return ($t + $nl + "Двойной щелчок - править вручную.")
        }
        2 { return $err }
        4 {
            $t = [string]$Row.old_path
            if ($err -and $err.StartsWith("Старый путь")) { $t += $nl + $err }
            return $t
        }
        5 {
            $t = [string]$Row.new_path
            if ($err -and -not $err.StartsWith("Старый путь")) { $t += $nl + "ОШИБКА: " + $err }
            if ($Row.manual -eq "path") { return ($t + $nl + "Изменено вручную. По правилу было бы: " + $Row.rule_path + $nl + "Двойной щелчок - править, пустое значение - вернуть правило.") }
            return ($t + $nl + "Двойной щелчок - править вручную.")
        }
        default { return "" }
    }
}

# ---------------------------------------------------- "Подробности": blocks follow the main checkboxes
# The main window's checkboxes are the only switches of an operation; the dialog only holds parameters
# and never ticks them. A block is active only while its checkbox is on, otherwise it is grey with a hint.
function Get-DetailsBlockState([bool]$Rename, [bool]$Move, [bool]$Upgrade) {
    return [pscustomobject]@{
        RuleEnabled   = $Rename
        RuleHint      = $(if ($Rename) { "" } else { "Включите «Переименовать» в главном окне" })
        MoveEnabled   = $Move
        MoveHint      = $(if ($Move) { "" } else { "Включите «Перенести» в главном окне" })
        BufferEnabled = $Upgrade
        BufferHint    = $(if ($Upgrade) { "" } else { "Включите «Сменить версию Revit» в главном окне" })
    }
}

# "Перенести" is on but nothing says what to replace: the run would move nothing.
function Get-ModelListLines([string]$Path) {
    $utf8 = New-Object System.Text.UTF8Encoding $true
    $raw = [System.IO.File]::ReadAllLines($Path, $utf8)
    return @($raw)
}

# Saved model lists live in <tool>\списки. lists_dir.cfg remembers the last folder.
function Get-ModelListStore([string]$ToolDir) {
    $fallback = Join-Path $ToolDir "списки"
    $cfg = Join-Path $ToolDir "lists_dir.cfg"
    if (Test-Path -LiteralPath $cfg) {
        $lines = @(Get-ModelListLines $cfg)
        $text = ""
        if ($lines.Count -gt 0) { $text = [string]$lines[0] }
        if ($text) { $text = $text.Trim().TrimStart([char]0xFEFF) }
        if ($text -and (Test-Path -LiteralPath $text -PathType Container)) { return $text }
    }
    if (-not (Test-Path -LiteralPath $fallback)) {
        New-Item -ItemType Directory -Path $fallback -Force | Out-Null
    }
    return $fallback
}

function Save-ModelListFolder([string]$ToolDir, [string]$ChosenFile) {
    if ([string]::IsNullOrWhiteSpace($ChosenFile)) { return }
    $dir = Split-Path -Parent $ChosenFile
    if ([string]::IsNullOrWhiteSpace($dir) -or -not (Test-Path -LiteralPath $dir -PathType Container)) { return }
    $cfg = Join-Path $ToolDir "lists_dir.cfg"
    $utf8 = New-Object System.Text.UTF8Encoding $true
    [System.IO.File]::WriteAllLines($cfg, @($dir), $utf8)
}

function Read-ModelListPaths([string]$Path) {
    $out = New-Object System.Collections.Generic.List[string]
    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path)) { return $out.ToArray() }
    foreach ($line in @(Get-ModelListLines $Path)) {
        $t = [string]$line
        if ($t) { $t = $t.Trim().TrimStart([char]0xFEFF) }
        if (-not $t -or $t.StartsWith("#")) { continue }
        if ($t.ToLowerInvariant().EndsWith(".rvt") -or $t.StartsWith("RSN://", [StringComparison]::OrdinalIgnoreCase)) {
            [void]$out.Add($t)
        }
    }
    return $out.ToArray()
}

# Paths already on a tab, then paths copied from another tab. Same path is kept once.
function Merge-TabModelPaths($Target, $Source) {
    $out = New-Object System.Collections.Generic.List[string]
    $seen = @{}
    foreach ($batch in @($Target, $Source)) {
        foreach ($p in @($batch)) {
            $text = [string]$p
            if ([string]::IsNullOrWhiteSpace($text)) { continue }
            $key = ConvertTo-UiPathKey $text
            if ($seen.ContainsKey($key)) { continue }
            $seen[$key] = $true
            [void]$out.Add($text.Trim())
        }
    }
    return $out.ToArray()
}

function Test-MoveSettingsMissing([bool]$Move, [string]$Mode, [string]$OldSegment, [string]$OldPrefix) {
    if (-not $Move) { return $false }
    if ($Mode -eq "prefix") { return [string]::IsNullOrWhiteSpace($OldPrefix) }
    return [string]::IsNullOrWhiteSpace($OldSegment)
}

# ------------------------------------------- tab «Перенос и переименование»: fields, no checkboxes
# An empty field means the step is not done. "Имя" = example name now -> what it must become,
# "Папка" = what to replace in the path -> with what.

# prefix: every model path starts with "было" (on a folder boundary) - only the beginning is replaced;
# otherwise the text is replaced wherever it is in the path (segment).
function Get-AutoLocationMode([string[]]$Paths, [string]$Was) {
    if ([string]::IsNullOrWhiteSpace($Was)) { return "same" }
    $w = $Was.Trim().TrimEnd("\", "/")
    if (-not $w) { return "segment" }
    $any = $false
    foreach ($p in @($Paths)) {
        if ([string]::IsNullOrWhiteSpace($p)) { continue }
        $any = $true
        $t = $p.Trim()
        if (-not $t.StartsWith($w, [StringComparison]::OrdinalIgnoreCase)) { return "segment" }
        if ($t.Length -gt $w.Length) {
            $c = $t[$w.Length]
            if ($c -ne [char]"\" -and $c -ne [char]"/") { return "segment" }
        }
    }
    if (-not $any) { return "segment" }
    return "prefix"
}

# What the fields of the tab ask for. Rename/Move are switched on by filled fields only.
function Get-MoveFieldState {
    param([string]$RuleOld, [string]$RuleNew, [string]$FolderOld, [string]$FolderNew, [string[]]$Paths)
    $ops = @()
    $ruleMsg = ""
    $ruleState = "empty"
    $a = ""; $b = ""
    if ($null -ne $RuleOld) { $a = $RuleOld.Trim() }
    if ($null -ne $RuleNew) { $b = $RuleNew.Trim() }
    if ($b) {
        $st = Get-RuleStatus $a $b
        $ruleState = [string]$st.State
        if ($st.State -eq "ok") { $ops = @($st.Ops) }
        elseif ($st.State -eq "fail") { $ruleMsg = [string]$st.Message }
    }
    $was = ""; $became = ""
    if ($null -ne $FolderOld) { $was = $FolderOld.Trim() }
    if ($null -ne $FolderNew) { $became = $FolderNew.Trim() }
    $move = ($was -ne "" -and $became -ne "")
    $moveMsg = ""
    if ($became -and -not $was) { $moveMsg = "Папка: заполните «было» - что заменить в пути." }
    if ($move -and ((ConvertTo-UiPathKey $was) -eq (ConvertTo-UiPathKey $became))) {
        $move = $false
        $moveMsg = "Папка: «было» и «стало» совпадают."
    }
    $mode = "same"
    if ($move) { $mode = Get-AutoLocationMode $Paths $was }
    return [pscustomobject]@{
        Ops         = $ops
        RuleState   = $ruleState
        RuleMessage = $ruleMsg
        RuleWanted  = ($b -ne "")
        Rename      = (@($ops).Count -gt 0)
        Move        = $move
        Mode        = $mode
        Was         = $was
        Became      = $became
        MoveMessage = $moveMsg
    }
}

# ------------------------------------------------ tab «Смена версии Revit»: the version is detected
# Revit Server models: the year of their host in servers.cfg. Network (UNC) files are opened by the
# newer Revit directly, their year is not needed. One run = one source year.
function Get-UpgradeYearInfo {
    param([string[]]$Paths, $Servers)
    $rows = New-Object System.Collections.Generic.List[object]
    $years = New-Object System.Collections.Generic.List[string]
    $bad = 0
    foreach ($p in @($Paths)) {
        if ([string]::IsNullOrWhiteSpace($p)) { continue }
        $year = ""
        $reason = ""
        $source = ""
        if (Test-IsRsnPath $p) {
            $h = Get-RsnHost $p.Trim()
            $y = Get-YearForRsnHost $h $Servers
            if ($y) { $year = [string]$y; if (-not $years.Contains($year)) { [void]$years.Add($year) } }
            elseif ($h) { $reason = ("Год не определён: сервера {0} нет в servers.cfg" -f $h); $bad++ }
            else { $reason = "Год не определён: в пути нет сервера"; $bad++ }
            $source = "server"
        }
        else {
            # network file: the year written in its name (..._R22), if any; not needed to open it
            $ny = Get-YearFromName (Get-BasenameNoExt $p)
            if ($ny) { $year = [string]$ny; $source = "name"; if (-not $years.Contains($year)) { [void]$years.Add($year) } }
        }
        [void]$rows.Add([pscustomobject]@{ Path = $p; Year = $year; Reason = $reason; Source = $source })
    }
    $ok = $true
    $msg = ""
    $year = ""
    if ($bad -gt 0) {
        $ok = $false
        $msg = ("Не определён год Revit у моделей: {0}." -f $bad)
    }
    elseif ($years.Count -gt 1) {
        $ok = $false
        $msg = ("В списке модели разных лет Revit ({0}). Один запуск - один год." -f ($years.ToArray() -join ", "))
        foreach ($r in $rows) { if ($r.Year) { $r.Reason = "Разные годы в списке" } }
    }
    elseif ($years.Count -eq 1) { $year = $years[0] }
    return [pscustomobject]@{ Ok = $ok; Year = $year; Message = $msg; Rows = $rows.ToArray() }
}

# Target years offered: only newer than the detected one (all of them when it is unknown).
function Get-UpgradeTargetYears([string[]]$Years, [string]$Detected) {
    $out = New-Object System.Collections.Generic.List[string]
    $d = 0
    [void][int]::TryParse([string]$Detected, [ref]$d)
    foreach ($y in @($Years)) {
        $n = 0
        if (-not [int]::TryParse([string]$y, [ref]$n)) { continue }
        if ($d -gt 0 -and $n -le $d) { continue }
        [void]$out.Add([string]$y)
    }
    return $out.ToArray()
}

# Year for the read / apply passes of «Связи: добавить» and «Рабочие наборы»: the same detection
# as «Смена версии Revit» (servers.cfg for RSN, _Rxx in the file name for network files), but
# Revit has to be started here, so a network file without _Rxx is a stop too.
# Message is short (it sits in the button row), Detail is for the message box.
function Get-ReadYearInfo {
    param([string[]]$Paths, $Servers)
    $info = Get-UpgradeYearInfo $Paths $Servers
    $rows = @($info.Rows)
    $noYear = 0
    foreach ($r in $rows) {
        if (-not $r.Year -and -not $r.Reason) {
            $r.Reason = "Год не определён: в имени нет _Rxx (например _R22)"
            $noYear++
        }
    }
    $years = @($rows | Where-Object { $_.Year } | ForEach-Object { [string]$_.Year } | Select-Object -Unique)
    if ($rows.Count -eq 0) {
        return [pscustomobject]@{ Ok = $false; Year = ""; Message = "Год Revit не определён"; Detail = "Модели не выбраны."; Rows = $rows }
    }
    if ($years.Count -gt 1) {
        return [pscustomobject]@{ Ok = $false; Year = ""; Message = "Разные годы Revit в списке"
            Detail = ("В списке модели разных лет Revit ({0}). Один запуск - один год." -f ($years -join ", ")); Rows = $rows }
    }
    if (-not $info.Ok -or $noYear -gt 0 -or $years.Count -eq 0) {
        $bad = @($rows | Where-Object { $_.Reason })
        $lines = @($bad | Select-Object -First 5 | ForEach-Object { ("{0}: {1}" -f (Get-BasenameNoExt $_.Path), $_.Reason) })
        $detail = ("Год Revit не определён у моделей: {0}." -f $bad.Count)
        if ($lines.Count -gt 0) { $detail = $detail + [Environment]::NewLine + ($lines -join [Environment]::NewLine) }
        return [pscustomobject]@{ Ok = $false; Year = ""; Message = "Год Revit не определён"; Detail = $detail; Rows = $rows }
    }
    return [pscustomobject]@{ Ok = $true; Year = [string]$years[0]; Message = ""; Detail = ""; Rows = $rows }
}

function Get-ReadYearCaption($Info) {
    if ($null -eq $Info -or -not $Info.Ok -or -not $Info.Year) {
        if ($Info -and $Info.Message) { return [string]$Info.Message }
        return "Год Revit не определён"
    }
    return ("Revit {0} (по моделям)" -f $Info.Year)
}
