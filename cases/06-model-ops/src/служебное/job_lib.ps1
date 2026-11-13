# Name: job_lib.ps1
# Version: 3.1.0
# What it does: PowerShell job helpers: naming, path rewrite, mapping/list export for model ops.
# Inputs: job JSON/CSV from presets.
# Outputs: job_paths.csv, mapping.csv, lists.
# How to run: Dot-sourced from presets.ps1/start_ops.ps1.
# Notes: StrictMode enabled.
# Job helpers for batch_model_ops (Windows PowerShell 5.1).
# Naming ops, path rewrite, mapping/list export. ASCII source.

Set-StrictMode -Version Latest

function Get-ToolDir {
    if ($PSScriptRoot) { return $PSScriptRoot }
    return (Split-Path -Parent $MyInvocation.MyCommand.Path)
}

function Get-RbpRoot {
    $userProfile = $env:USERPROFILE
    if ([string]::IsNullOrWhiteSpace($userProfile)) { $userProfile = "C:\Users\Public" }
    return (Join-Path $userProfile "Documents\doc\script\rbp")
}

# Every report / log / CSV of this tool goes to <tool root>\отчёты.
# The old common folder rbp\reports is only read (older files).
function Get-OpsReportsDir([string]$Root = "") {
    if ([string]::IsNullOrWhiteSpace($Root)) {
        $v = Get-Variable -Name ToolDir -ValueOnly -ErrorAction SilentlyContinue
        if ($v) { $Root = [string]$v }
    }
    if ([string]::IsNullOrWhiteSpace($Root)) { return (Get-OpsLegacyReportsDir) }
    return (Join-Path $Root "отчёты")
}

function Get-OpsLegacyReportsDir {
    return (Join-Path (Get-RbpRoot) "reports")
}

# Creates the folder when missing and returns it.
function Initialize-OpsReportsDir([string]$Root = "") {
    $dir = Get-OpsReportsDir $Root
    if (-not (Test-Path -LiteralPath $dir)) {
        try { New-Item -ItemType Directory -Path $dir -Force -ErrorAction Stop | Out-Null } catch { }
    }
    return $dir
}

# Readers look in the tool folder first, then in the old one.
function Get-OpsReportsReadDirs([string]$Root = "") {
    $out = New-Object System.Collections.Generic.List[string]
    $a = Get-OpsReportsDir $Root
    $b = Get-OpsLegacyReportsDir
    [void]$out.Add($a)
    if ($b -ne $a) { [void]$out.Add($b) }
    return $out.ToArray()
}

function Get-DefaultBufferDir {
    return (Join-Path (Get-RbpRoot) "upgrade_buffer")
}

function Get-BasenameNoExt([string]$PathOrName) {
    if ([string]::IsNullOrWhiteSpace($PathOrName)) { return "" }
    $name = ($PathOrName -replace "/", "\").Split("\")[-1].Trim()
    if ($name.ToLower().EndsWith(".rvt")) {
        $name = $name.Substring(0, $name.Length - 4)
    }
    return $name
}

function Get-NameTokens([string]$Name) {
    $n = Get-BasenameNoExt $Name
    if ([string]::IsNullOrWhiteSpace($n)) { return @() }
    return @($n.Split([char]'_'))
}

function Test-IsRsnPath([string]$Path) {
    if ([string]::IsNullOrWhiteSpace($Path)) { return $false }
    return $Path.Trim().StartsWith("RSN://", [StringComparison]::OrdinalIgnoreCase)
}

function Get-RsnHost([string]$Path) {
    if ($Path -match '^(?i)RSN://([^/]+)/') { return $Matches[1].Trim() }
    return $null
}

function Get-FolderOfPath([string]$Path) {
    if ([string]::IsNullOrWhiteSpace($Path)) { return "" }
    $n = $Path.Replace("/", "\")
    $i = $n.LastIndexOf("\")
    if ($i -lt 0) {
        $i = $Path.LastIndexOf("/")
        if ($i -lt 0) { return "" }
        return $Path.Substring(0, $i)
    }
    # keep original separator style for RSN (forward slash)
    if ($Path.StartsWith("RSN://", [StringComparison]::OrdinalIgnoreCase) -or $Path.Contains("/")) {
        $slash = $Path.Replace("\", "/")
        $j = $slash.LastIndexOf("/")
        if ($j -ge 0) { return $slash.Substring(0, $j) }
    }
    return $n.Substring(0, $i)
}

function Join-FolderFile([string]$Folder, [string]$FileName) {
    if ([string]::IsNullOrWhiteSpace($Folder)) { return $FileName }
    if ($Folder.StartsWith("RSN://", [StringComparison]::OrdinalIgnoreCase) -or $Folder.Contains("/")) {
        $f = $Folder.TrimEnd("/", "\")
        return "$f/$FileName"
    }
    return (Join-Path $Folder $FileName)
}

function Set-FilenameInPath([string]$Path, [string]$NewBase) {
    $oldBase = Get-BasenameNoExt $Path
    $newFile = "$NewBase.rvt"
    if ([string]::IsNullOrWhiteSpace($oldBase)) {
        return (Join-FolderFile (Get-FolderOfPath $Path) $newFile)
    }
    $oldFile = "$oldBase.rvt"
    $idx = $Path.ToLower().LastIndexOf($oldFile.ToLower())
    if ($idx -ge 0) {
        return $Path.Substring(0, $idx) + $newFile + $Path.Substring($idx + $oldFile.Length)
    }
    $folder = Get-FolderOfPath $Path
    return (Join-FolderFile $folder $newFile)
}

function ConvertTo-CaseInsensitiveReplace([string]$Text, [string]$Old, [string]$New) {
    if ([string]::IsNullOrWhiteSpace($Text) -or [string]::IsNullOrWhiteSpace($Old)) { return $Text }
    $idx = $Text.IndexOf($Old, [StringComparison]::OrdinalIgnoreCase)
    if ($idx -lt 0) { return $Text }
    return $Text.Substring(0, $idx) + $New + $Text.Substring($idx + $Old.Length)
}

function Get-YearFromName([string]$Name) {
    $n = Get-BasenameNoExt $Name
    if ($n -match '_R(\d{2})$') {
        return (2000 + [int]$Matches[1]).ToString()
    }
    return $null
}

function Read-ServerTable([string]$CfgPath) {
    $rows = New-Object System.Collections.Generic.List[object]
    if (-not (Test-Path -LiteralPath $CfgPath)) { return @() }
    Get-Content -LiteralPath $CfgPath -Encoding UTF8 | ForEach-Object {
        $line = $_.Trim()
        if (-not $line -or $line.StartsWith("#")) { return }
        $parts = $line.Split("|")
        if ($parts.Count -lt 3) { return }
        $rows.Add([pscustomobject]@{
                Host  = $parts[0].Trim()
                Label = $parts[1].Trim()
                Year  = $parts[2].Trim()
            }) | Out-Null
    }
    if ($rows.Count -eq 0) { return @() }
    return $rows.ToArray()
}

function Get-YearForRsnHost([string]$HostName, $Servers) {
    if ([string]::IsNullOrWhiteSpace($HostName)) { return $null }
    foreach ($s in @($Servers)) {
        if ($s.Host -ieq $HostName) { return [string]$s.Year }
    }
    return $null
}

function New-NameOp {
    param(
        [Parameter(Mandatory = $true)][ValidateSet("delete", "replace", "insert")]
        [string]$Kind,
        [string]$Token = "",
        [string]$From = "",
        [string]$To = "",
        [string]$After = ""
    )
    return [pscustomobject]@{
        kind  = $Kind
        token = $Token
        from  = $From
        to    = $To
        after = $After
    }
}

function Get-OpLabel($Op) {
    switch ($Op.kind) {
        "delete" { return ("delete  {0}" -f $Op.token) }
        "replace" { return ("replace {0} -> {1}" -f $Op.from, $Op.to) }
        "insert" {
            if ($Op.after) { return ("insert  {0} after {1}" -f $Op.token, $Op.after) }
            return ("insert  {0}" -f $Op.token)
        }
        default { return ($Op.kind + "") }
    }
}

function Test-TokenInTail([string[]]$Arr, [int]$Start, [string]$Token) {
    for ($k = $Start; $k -lt $Arr.Count; $k++) {
        if ($Arr[$k] -ieq $Token) { return $true }
    }
    return $false
}

function New-OpsFromSample([string]$OldName, [string]$NewName) {
    $ops = New-Object System.Collections.Generic.List[object]
    $a = @(Get-NameTokens $OldName)
    $b = @(Get-NameTokens $NewName)
    if ($a.Count -eq 0 -and $b.Count -eq 0) { return @() }
    $i = 0
    $j = 0
    while ($i -lt $a.Count -or $j -lt $b.Count) {
        $haveA = ($i -lt $a.Count)
        $haveB = ($j -lt $b.Count)
        if ($haveA -and $haveB -and ($a[$i] -ieq $b[$j])) {
            $i++
            $j++
            continue
        }
        $aInB = $false
        $bInA = $false
        if ($haveA -and $haveB) {
            $aInB = Test-TokenInTail $b $j $a[$i]
            $bInA = Test-TokenInTail $a $i $b[$j]
        }
        # Neither token exists on the other side: one token was replaced in place.
        # (Delete + insert would anchor on the deleted token and push the new one to the end.)
        if ($haveA -and $haveB -and (-not $aInB) -and (-not $bInA)) {
            $ops.Add((New-NameOp -Kind replace -From $a[$i] -To $b[$j])) | Out-Null
            $i++
            $j++
            continue
        }
        if ($haveA -and ((-not $haveB) -or (-not $aInB))) {
            $ops.Add((New-NameOp -Kind delete -Token $a[$i])) | Out-Null
            $i++
            continue
        }
        if ($haveB -and ((-not $haveA) -or (-not $bInA))) {
            # anchor on the previous token of the NEW name: it exists after earlier ops ran
            $after = ""
            if ($j -gt 0) { $after = $b[$j - 1] }
            $ops.Add((New-NameOp -Kind insert -Token $b[$j] -After $after)) | Out-Null
            $j++
            continue
        }
        $ops.Add((New-NameOp -Kind replace -From $a[$i] -To $b[$j])) | Out-Null
        $i++
        $j++
    }
    if ($ops.Count -eq 0) { return @() }
    return $ops.ToArray()
}

function Invoke-NameOps {
    param(
        [string]$Name,
        [object[]]$Ops
    )
    $tokens = New-Object System.Collections.Generic.List[string]
    foreach ($t in @(Get-NameTokens $Name)) { $tokens.Add($t) | Out-Null }
    foreach ($op in @($Ops)) {
        if ($null -eq $op) { continue }
        switch ($op.kind) {
            "delete" {
                for ($i = $tokens.Count - 1; $i -ge 0; $i--) {
                    if ($tokens[$i] -ieq $op.token) { $tokens.RemoveAt($i) }
                }
            }
            "replace" {
                for ($i = 0; $i -lt $tokens.Count; $i++) {
                    if ($tokens[$i] -ieq $op.from) { $tokens[$i] = [string]$op.to }
                }
            }
            "insert" {
                $idx = -1
                if ($op.after) {
                    for ($i = 0; $i -lt $tokens.Count; $i++) {
                        if ($tokens[$i] -ieq $op.after) { $idx = $i; break }
                    }
                }
                if ($idx -ge 0) { $tokens.Insert($idx + 1, [string]$op.token) }
                else { $tokens.Add([string]$op.token) | Out-Null }
            }
        }
    }
    return ($tokens -join "_")
}

function ConvertTo-LocationPath {
    param(
        [string]$OldPath,
        [string]$NewBase,
        [string]$Mode,
        [string]$OldSegment,
        [string]$NewSegment,
        [string]$OldPrefix,
        [string]$NewPrefix
    )
    $withName = Set-FilenameInPath $OldPath $NewBase
    switch ($Mode) {
        "segment" {
            if ([string]::IsNullOrWhiteSpace($OldSegment)) { return $withName }
            return (ConvertTo-CaseInsensitiveReplace $withName $OldSegment $NewSegment)
        }
        "prefix" {
            if ([string]::IsNullOrWhiteSpace($OldPrefix)) { return $withName }
            if ($withName.StartsWith($OldPrefix, [StringComparison]::OrdinalIgnoreCase)) {
                return $NewPrefix + $withName.Substring($OldPrefix.Length)
            }
            return $withName
        }
        default { return $withName }
    }
}

function Test-NeedsYearBuffer {
    param(
        [string]$OldPath,
        [string]$SourceYear,
        [string]$TargetYear
    )
    if ([string]::IsNullOrWhiteSpace($TargetYear)) { return $false }
    if ($SourceYear -and ($SourceYear.Trim() -eq $TargetYear.Trim())) { return $false }
    # New Revit cannot open old Revit Server. UNC old files can be opened by newer Revit.
    return (Test-IsRsnPath $OldPath)
}

function Get-StagingPath([string]$BufferDir, [string]$NewBase) {
    if ([string]::IsNullOrWhiteSpace($BufferDir)) { $BufferDir = Get-DefaultBufferDir }
    # plain concatenation (same result as Join-Path on Windows, but also fine for a folder that does not exist)
    $sep = "\"
    if ($BufferDir.Contains("/") -and -not $BufferDir.Contains("\")) { $sep = "/" }
    return ($BufferDir.TrimEnd("\", "/") + $sep + ("{0}.rvt" -f $NewBase))
}

function Read-ModelList([string]$ListPath) {
    $out = New-Object System.Collections.Generic.List[string]
    if (-not (Test-Path -LiteralPath $ListPath)) { return @() }
    Get-Content -LiteralPath $ListPath -Encoding UTF8 | ForEach-Object {
        $line = $_.Trim()
        if (-not $line -or $line.StartsWith("#")) { return }
        $out.Add($line) | Out-Null
    }
    if ($out.Count -eq 0) { return @() }
    return $out.ToArray()
}

# ---------------------------------------------------------------- names

$script:ForeignExtensions = @("rfa", "rte", "rft", "dwg", "dxf", "ifc", "nwc", "nwd", "pdf", "txt", "bak", "csv", "xlsx", "docx", "zip")
$script:ReservedNames = @("CON", "PRN", "AUX", "NUL") + @(1..9 | ForEach-Object { "COM$_"; "LPT$_" })

function New-CheckResult([bool]$Ok, [string]$Code, [string]$Message) {
    return [pscustomobject]@{ Ok = $Ok; Code = $Code; Message = $Message; Value = "" }
}

# Name typed into "Станет": with or without ".rvt". Value = normalized name without extension.
function Test-ModelFileName([string]$Name) {
    $n = ""
    if ($null -ne $Name) { $n = $Name.Trim() }
    if ($n.ToLowerInvariant().EndsWith(".rvt")) { $n = $n.Substring(0, $n.Length - 4) }
    if ($n.Length -eq 0) {
        return New-CheckResult $false "empty" "Имя не может быть пустым."
    }
    $bad = New-Object System.Collections.Generic.List[string]
    foreach ($ch in $n.ToCharArray()) {
        if ("\/:*?`"<>|".IndexOf($ch) -ge 0 -or [int]$ch -lt 32) {
            $s = $(if ([int]$ch -lt 32) { "(управляющий)" } else { [string]$ch })
            if (-not $bad.Contains($s)) { $bad.Add($s) | Out-Null }
        }
    }
    if ($bad.Count -gt 0) {
        return New-CheckResult $false "chars" ("В имени есть недопустимые символы: {0}." -f ($bad -join " "))
    }
    if ($n.EndsWith(".") -or $n.EndsWith(" ")) {
        return New-CheckResult $false "trailing" "Имя не должно заканчиваться точкой или пробелом."
    }
    if ($n -match '\.([A-Za-z0-9]+)$') {
        $ext = $Matches[1].ToLowerInvariant()
        if ($script:ForeignExtensions -contains $ext) {
            return New-CheckResult $false "ext" ("Расширение «.{0}» не подходит: модель сохраняется как .rvt." -f $ext)
        }
    }
    $stem = $n.Split(".")[0].ToUpperInvariant()
    if ($script:ReservedNames -contains $stem) {
        return New-CheckResult $false "reserved" ("«{0}» - зарезервированное имя Windows." -f $stem)
    }
    if (($n.Length + 4) -gt 255) {
        return New-CheckResult $false "length" "Слишком длинное имя (максимум 255 знаков вместе с .rvt)."
    }
    $r = New-CheckResult $true "" ""
    $r.Value = $n
    return $r
}

# Path typed into "Новый путь": full path ending with .rvt. Value = trimmed path.
function Test-NewModelPath([string]$Path) {
    $p = ""
    if ($null -ne $Path) { $p = $Path.Trim().Trim('"') }
    if ($p.Length -eq 0) { return New-CheckResult $false "empty" "Путь не может быть пустым." }
    if (-not $p.ToLowerInvariant().EndsWith(".rvt")) {
        return New-CheckResult $false "ext" "Путь должен заканчиваться на .rvt."
    }
    $rest = $null
    if ($p -match '^(?i)RSN://[^/\\\s]+/(.+)$') { $rest = $Matches[1] }
    elseif ($p -match '^\\\\[^\\/]+[\\/][^\\/]+[\\/](.+)$') { $rest = $Matches[1] }
    elseif ($p -match '^[A-Za-z]:[\\/](.+)$') { $rest = $Matches[1] }
    if ($null -eq $rest) {
        return New-CheckResult $false "format" "Нужен полный путь: диск (C:\...), сеть (\\сервер\ресурс\...) или RSN://сервер/..."
    }
    foreach ($seg in $rest.Replace("\", "/").Split("/")) {
        if ($seg.Length -eq 0) {
            return New-CheckResult $false "segments" "В пути есть пустая папка (двойной слэш)."
        }
    }
    foreach ($ch in $rest.ToCharArray()) {
        if ("`:*?`"<>|".IndexOf($ch) -ge 0 -or [int]$ch -lt 32) {
            return New-CheckResult $false "chars" ("В пути есть недопустимый символ: {0}." -f $(if ([int]$ch -lt 32) { "(управляющий)" } else { [string]$ch }))
        }
    }
    $leaf = $rest.Replace("\", "/").Split("/")[-1]
    $nameCheck = Test-ModelFileName $leaf
    if (-not $nameCheck.Ok) { return $nameCheck }
    $r = New-CheckResult $true "" ""
    $r.Value = $p
    return $r
}

# Format only: a full path with a disk (D:\...), a network share (\\server\share\...) or RSN://host/...
function Test-AbsoluteModelPath([string]$Path) {
    if ([string]::IsNullOrWhiteSpace($Path)) { return $false }
    $p = $Path.Trim().Trim('"')
    return ($p -match '^(?i)RSN://[^/\\\s]+/.+$') -or ($p -match '^\\\\[^\\/]+[\\/][^\\/]+[\\/].+$') -or ($p -match '^[A-Za-z]:[\\/].+$')
}

# "" when the paths of a preview row are fine; otherwise a message a person can act on.
# The new path must be a full absolute path whatever the placement mode did to it.
function Get-PathProblem {
    param(
        [string]$OldPath, [string]$NewPath, [string]$StagingPath, [bool]$NeedsBuffer,
        [string]$Mode, [string]$OldSegment, [string]$NewSegment, [string]$OldPrefix, [string]$NewPrefix
    )
    if (-not (Test-AbsoluteModelPath $OldPath)) {
        return ("Старый путь не полный: «{0}». Нужен путь с диском (D:\...), сетевой (\\сервер\ресурс\...) или RSN://... Уберите модель из списка и добавьте заново." -f $OldPath)
    }
    $c = Test-NewModelPath $NewPath
    if (-not $c.Ok) {
        $what = $c.Message
        if ($c.Code -eq "empty") { $what = "Новый путь пуст." }
        elseif ($c.Code -eq "format") { $what = ("Новый путь не полный: «{0}» - нет диска, сервера или папки." -f $NewPath) }
        elseif ($c.Code -eq "segments") { $what = ("Новый путь не годится: «{0}» - в нём пустая папка (двойной слэш)." -f $NewPath) }
        else { $what = ("Новый путь не годится: {0} («{1}»)" -f $c.Message, $NewPath) }
        $hint = ""
        if ($c.Code -in @("empty", "format", "segments")) {
            if ($Mode -eq "segment") { $hint = (" После замены «куска пути» (было «{0}», стало «{1}») получился такой путь: проверьте оба поля в «Подробности»." -f $OldSegment, $NewSegment) }
            elseif ($Mode -eq "prefix") { $hint = (" После замены начала пути (было «{0}», стало «{1}») получился такой путь: проверьте оба поля в «Подробности»." -f $OldPrefix, $NewPrefix) }
        }
        return ($what + $hint)
    }
    if ($NeedsBuffer) {
        $s = Test-NewModelPath $StagingPath
        if (-not $s.Ok) {
            return ("Путь в буфере не полный: «{0}». Укажите папку буфера на этом ПК полным путём (например D:\buffer) в «Подробности»." -f $StagingPath)
        }
    }
    return ""
}

# Key of a manual edit: the OLD path, case/slash-insensitive.
function Get-ManualRowKey([string]$OldPath) {
    if ([string]::IsNullOrWhiteSpace($OldPath)) { return "" }
    return $OldPath.Trim().Trim('"').Replace("\", "/").TrimEnd("/").ToUpperInvariant()
}

# Strict-mode safe read of a manual entry field ("name" or "path"); $null when absent.
function Get-ManualField($Entry, [string]$Field) {
    if ($null -eq $Entry) { return $null }
    if ($Entry -is [System.Collections.IDictionary]) {
        if ($Entry.Contains($Field)) { return $Entry[$Field] }
        return $null
    }
    $p = $Entry.PSObject.Properties[$Field]
    if ($null -eq $p) { return $null }
    return $p.Value
}

# ManualRows: key = Get-ManualRowKey(old path), value = @{ name = ...; path = ... }
# (either may be $null). A manual edit beats the rule: path wins over name, and the
# new name is then the file name inside that path (mapping.csv and SaveAs must agree).
function New-JobPreviewRows {
    param(
        [string[]]$OldPaths,
        [object[]]$Ops,
        [bool]$DoRename,
        [string]$LocationMode,
        [string]$OldSegment,
        [string]$NewSegment,
        [string]$OldPrefix,
        [string]$NewPrefix,
        [string]$SourceYear,
        [string]$TargetYear,
        [string]$BufferDir,
        [hashtable]$NameOverrides,
        [hashtable]$ManualRows
    )
    $rows = New-Object System.Collections.Generic.List[object]
    $seenNew = @{}
    foreach ($oldPath in @($OldPaths)) {
        $oldName = Get-BasenameNoExt $oldPath
        $newName = $oldName
        if ($DoRename -and @($Ops).Count -gt 0) {
            $newName = Invoke-NameOps -Name $oldName -Ops $Ops
        }
        if ($NameOverrides -and $NameOverrides.ContainsKey($oldName.ToUpperInvariant())) {
            $newName = [string]$NameOverrides[$oldName.ToUpperInvariant()]
        }
        $srcYear = $SourceYear
        if ([string]::IsNullOrWhiteSpace($srcYear)) { $srcYear = Get-YearFromName $oldName }
        $ruleName = $newName
        $ruleLoc = @{
            Mode = $LocationMode; OldSegment = $OldSegment; NewSegment = $NewSegment
            OldPrefix = $OldPrefix; NewPrefix = $NewPrefix
        }
        $rulePath = ConvertTo-LocationPath -OldPath $oldPath -NewBase $ruleName @ruleLoc
        $manual = ""
        $newPath = $null
        $mk = Get-ManualRowKey $oldPath
        if ($ManualRows -and $ManualRows.ContainsKey($mk)) {
            $m = $ManualRows[$mk]
            $mPath = Get-ManualField $m "path"
            $mName = Get-ManualField $m "name"
            if ($mPath) {
                $newPath = [string]$mPath
                $newName = Get-BasenameNoExt $newPath
                $manual = "path"
            }
            elseif ($mName) {
                $newName = [string]$mName
                $manual = "name"
            }
        }
        if ($null -eq $newPath) {
            $newPath = ConvertTo-LocationPath -OldPath $oldPath -NewBase $newName @ruleLoc
        }
        $needBuf = Test-NeedsYearBuffer -OldPath $oldPath -SourceYear $srcYear -TargetYear $TargetYear
        $staging = ""
        if ($needBuf) { $staging = Get-StagingPath -BufferDir $BufferDir -NewBase $newName }
        $note = ""
        $key = $newPath.ToLowerInvariant()
        if ($seenNew.ContainsKey($key)) { $note = "duplicate new path" }
        else { $seenNew[$key] = $true }
        if ($oldPath.Replace("\", "/").ToLowerInvariant() -eq $newPath.Replace("\", "/").ToLowerInvariant()) {
            if (-not $needBuf) { $note = $(if ($note) { "$note; same path" } else { "same path" }) }
        }
        $hintMode = $LocationMode
        if ($manual -eq "path") { $hintMode = "manual" }
        $problem = Get-PathProblem -OldPath $oldPath -NewPath $newPath -StagingPath $staging -NeedsBuffer $needBuf `
            -Mode $hintMode -OldSegment $OldSegment -NewSegment $NewSegment -OldPrefix $OldPrefix -NewPrefix $NewPrefix
        $rows.Add([pscustomobject]@{
                path_error   = $problem
                manual       = $manual
                rule_name    = $ruleName
                rule_path    = $rulePath
                old_name     = $oldName
                new_name     = $newName
                old_path     = $oldPath
                new_path     = $newPath
                staging_path = $staging
                needs_buffer = $needBuf
                note         = $note
            }) | Out-Null
    }
    if ($rows.Count -eq 0) { return @() }
    return $rows.ToArray()
}

function ConvertTo-CsvEsc([string]$Val) {
    if ($null -eq $Val) { return "" }
    $s = [string]$Val
    $s = $s.Replace("`r", " ").Replace("`n", " ")
    if ($s.Contains(";") -or $s.Contains('"')) {
        return '"' + $s.Replace('"', '""') + '"'
    }
    return $s
}

function Get-JobProp($Job, [string]$Name) {
    if (-not $Job) { return "" }
    $p = $Job.PSObject.Properties[$Name]
    if (-not $p -or $null -eq $p.Value) { return "" }
    return ([string]$p.Value).Trim()
}

function ConvertTo-PrefixKey([string]$Path) {
    if ([string]::IsNullOrWhiteSpace($Path)) { return "" }
    return $Path.Trim().Replace("\", "/").TrimEnd("/").ToLowerInvariant()
}

# Common parent folder of the models' folders (guess for "Новая папка").
function Get-CommonModelFolder([string[]]$Paths) {
    $common = $null
    $dirs = @{}
    foreach ($p in @($Paths)) {
        if (-not $p) { continue }
        $q = $p.Trim().Replace("\", "/")
        $i = $q.LastIndexOf("/")
        if ($i -gt 0) { $dirs[$q.Substring(0, $i).ToLowerInvariant()] = $true }
    }
    if ($dirs.Count -lt 2) { return "" }
    foreach ($p in @($Paths)) {
        if ([string]::IsNullOrWhiteSpace($p)) { continue }
        $segs = @($p.Trim().Replace("\", "/").Split("/"))
        if ($segs.Count -lt 2) { continue }
        $segs = $segs[0..($segs.Count - 2)]
        if ($null -eq $common) { $common = $segs; continue }
        $n = [Math]::Min($common.Count, $segs.Count)
        $k = 0
        while ($k -lt $n -and ([string]$common[$k]).ToLowerInvariant() -eq ([string]$segs[$k]).ToLowerInvariant()) { $k++ }
        if ($k -eq 0) { return "" }
        $common = $common[0..($k - 1)]
    }
    if ($null -eq $common) { return "" }
    $s = (@($common) -join "/")
    if (@($common | Where-Object { $_ }).Count -lt 3) { return "" }
    # same slash style as the first path (UNC / drive: backslashes, RSN: forward slashes)
    $firstP = [string]@($Paths | Where-Object { $_ })[0]
    if ($firstP.Contains("\") -and -not $firstP.Trim().StartsWith("RSN://", [StringComparison]::OrdinalIgnoreCase)) {
        $s = $s.Replace("/", "\")
    }
    return $s
}

function Save-JobFiles {
    param(
        [string]$ToolDir,
        [object[]]$Rows,
        [object]$Job
    )
    if (-not (Test-Path -LiteralPath $ToolDir)) {
        New-Item -ItemType Directory -Path $ToolDir -Force | Out-Null
    }
    $mapping = Join-Path $ToolDir "mapping.csv"
    $paths = Join-Path $ToolDir "job_paths.csv"
    $oldList = Join-Path $ToolDir "rvt_list.txt"
    $newList = Join-Path $ToolDir "rvt_list_new.txt"
    $stgList = Join-Path $ToolDir "rvt_list_staging.txt"
    $jobJson = Join-Path $ToolDir "job.json"
    $check = Join-Path $ToolDir "rbp_checklist.txt"

    $mapLines = New-Object System.Collections.Generic.List[string]
    $mapLines.Add("old;new") | Out-Null
    $pathLines = New-Object System.Collections.Generic.List[string]
    $pathLines.Add("old_path;new_path;staging_path;old_name;new_name;needs_buffer") | Out-Null
    $olds = New-Object System.Collections.Generic.List[string]
    $news = New-Object System.Collections.Generic.List[string]
    $stgs = New-Object System.Collections.Generic.List[string]

    foreach ($r in @($Rows)) {
        $mapLines.Add(("{0};{1}" -f (ConvertTo-CsvEsc $r.old_name), (ConvertTo-CsvEsc $r.new_name))) | Out-Null
        $pathLines.Add(("{0};{1};{2};{3};{4};{5}" -f `
                    (ConvertTo-CsvEsc $r.old_path),
                (ConvertTo-CsvEsc $r.new_path),
                (ConvertTo-CsvEsc $r.staging_path),
                (ConvertTo-CsvEsc $r.old_name),
                (ConvertTo-CsvEsc $r.new_name),
                $(if ($r.needs_buffer) { "1" } else { "0" })
            )) | Out-Null
        $olds.Add($r.old_path) | Out-Null
        $news.Add($r.new_path) | Out-Null
        if ($r.needs_buffer -and $r.staging_path) { $stgs.Add($r.staging_path) | Out-Null }
    }

    $utf8 = New-Object System.Text.UTF8Encoding $true
    [System.IO.File]::WriteAllLines($mapping, $mapLines.ToArray(), $utf8)
    [System.IO.File]::WriteAllLines($paths, $pathLines.ToArray(), $utf8)
    [System.IO.File]::WriteAllLines($oldList, $olds.ToArray(), $utf8)
    [System.IO.File]::WriteAllLines($newList, $news.ToArray(), $utf8)
    if ($stgs.Count -gt 0) {
        [System.IO.File]::WriteAllLines($stgList, $stgs.ToArray(), $utf8)
    }
    elseif (Test-Path -LiteralPath $stgList) {
        Remove-Item -LiteralPath $stgList -Force
    }

    # Relink by folder prefix (relink-only). Absent file = old filename mapping.
    $relinkMap = Join-Path $ToolDir "relink_map.csv"
    $rOld = Get-JobProp $Job "relink_old"
    $rNew = Get-JobProp $Job "relink_new"
    if ($rOld -and $rNew -and (ConvertTo-PrefixKey $rOld) -ne (ConvertTo-PrefixKey $rNew)) {
        [System.IO.File]::WriteAllLines($relinkMap, [string[]]@(
                "old_prefix;new_prefix",
                ("{0};{1}" -f (ConvertTo-CsvEsc $rOld), (ConvertTo-CsvEsc $rNew))
            ), $utf8)
    }
    elseif (Test-Path -LiteralPath $relinkMap) {
        Remove-Item -LiteralPath $relinkMap -Force
    }

    $job | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $jobJson -Encoding UTF8

    $needBuf = @($Rows | Where-Object { $_.needs_buffer }).Count -gt 0
    $srcY = [string]$Job.source_year
    $dstY = [string]$Job.target_year
    $doRelink = [bool]$Job.relink
    $doCopy = [bool]$Job.copy
    $doDelete = [bool]$Job.delete_old
    $lines = New-Object System.Collections.Generic.List[string]
    $n = 1
    if ($needBuf) {
        $lines.Add(("{0}) RBP Revit {1}" -f $n, $(if ($srcY) { $srcY } else { "OLD" }))) | Out-Null
        $lines.Add("   Detach from Central = ON") | Out-Null
        $lines.Add("   Create New Local = OFF (detached file)") | Out-Null
        $lines.Add("   Task = detach_to_buffer.py") | Out-Null
        $lines.Add("   List = rvt_list.txt   (old RSN paths)") | Out-Null
        $n++
        $lines.Add(("{0}) RBP Revit {1}" -f $n, $(if ($dstY) { $dstY } else { "NEW" }))) | Out-Null
        $lines.Add("   Create New Local, Detach = OFF") | Out-Null
        $lines.Add("   Task = saveas_central.py") | Out-Null
        $lines.Add("   List = rvt_list_staging.txt") | Out-Null
        $n++
    }
    elseif ($Job.saveas_needed) {
        $ver = $dstY
        if (-not $ver) { $ver = $srcY }
        if (-not $ver) { $ver = "model year" }
        $lines.Add(("{0}) RBP Revit {1}" -f $n, $ver)) | Out-Null
        $lines.Add("   Create New Local, Detach = OFF") | Out-Null
        $lines.Add("   Task = saveas_central.py") | Out-Null
        $lines.Add("   List = rvt_list.txt") | Out-Null
        $n++
    }
    if ($Job.unc_move) {
        $lines.Add(("{0}) Windows Move: служебное\run_cascade.cmd -> Move UNC" -f $n)) | Out-Null
        $n++
    }
    if ($doRelink) {
        $ver = $dstY
        if (-not $ver) { $ver = $srcY }
        if (-not $ver) { $ver = "model year" }
        $lines.Add(("{0}) RBP Revit {1}" -f $n, $ver)) | Out-Null
        $lines.Add("   Create New Local, Detach = OFF") | Out-Null
        $lines.Add("   Task = update_rvt_links.py") | Out-Null
        $lines.Add("   List = rvt_list_new.txt") | Out-Null
        $n++
    }
    if ($doDelete -and -not $doCopy) {
        $lines.Add(("{0}) After green Pass 2: служебное\delete_old_models.cmd" -f $n)) | Out-Null
        $lines.Add("   dry-run, then YES: old UNC files go to _to_delete (restorable); RSN -> to_delete_rsn.txt -> Admin") | Out-Null
        if ($needBuf) {
            $lines.Add("   Staging on this PC is moved to _to_delete by the same script.") | Out-Null
        }
    }
    [System.IO.File]::WriteAllLines($check, $lines.ToArray(), $utf8)

    return [pscustomobject]@{
        mapping   = $mapping
        paths     = $paths
        old_list  = $oldList
        new_list  = $newList
        stg_list  = $stgList
        job       = $jobJson
        checklist = $check
        buffer    = $needBuf
    }
}

function Test-UncMoveCandidate {
    param(
        [object[]]$Rows,
        [bool]$Upgrade,
        [string]$LocationMode
    )
    if ($Upgrade) { return $false }
    if ($LocationMode -eq "same") { return $false }
    $any = $false
    foreach ($r in @($Rows)) {
        if (Test-IsRsnPath $r.old_path) { return $false }
        if ($r.old_path -ne $r.new_path) { $any = $true }
    }
    return $any
}

function New-OpsPassQueue {
    param(
        [Parameter(Mandatory = $true)]$Job,
        [Parameter(Mandatory = $true)][string]$ToolDir,
        [string]$CodeDir = "",
        [object[]]$Rows
    )
    # $ToolDir = tool root (lists, job_paths.csv, servers.cfg); $CodeDir = служебное\ (.py, .ps1)
    if ([string]::IsNullOrWhiteSpace($CodeDir)) { $CodeDir = Join-Path $ToolDir "служебное" }
    $passes = New-Object System.Collections.Generic.List[object]
    $needBuf = @($Rows | Where-Object { $_.needs_buffer }).Count -gt 0
    $srcY = [string]$Job.source_year
    $dstY = [string]$Job.target_year
    $allowOverwrite = $false
    $owProp = $Job.PSObject.Properties["allow_overwrite"]
    if ($owProp -and $owProp.Value) { $allowOverwrite = $true }
    # No year chosen: take it from servers.cfg by RSN host of the first RSN model.
    $autoY = ""
    $firstRsn = @($Rows | Where-Object { $_ -and (Test-IsRsnPath ([string]$_.old_path)) } | Select-Object -First 1)
    if ($firstRsn.Count -gt 0) {
        $cfgServers = @(Read-ServerTable (Join-Path $ToolDir "servers.cfg"))
        $y = Get-YearForRsnHost (Get-RsnHost ([string]$firstRsn[0].old_path)) $cfgServers
        if ($y) { $autoY = [string]$y }
    }

    if ($needBuf) {
        $passes.Add([pscustomobject]@{
                Kind   = "rbp"
                Name   = "detach"
                Task   = (Join-Path $CodeDir "detach_to_buffer.py")
                List   = (Join-Path $ToolDir "rvt_list.txt")
                Year   = $srcY
                Detach = $true
            }) | Out-Null
        $passes.Add([pscustomobject]@{
                Kind   = "rbp"
                Name   = "saveas"
                Task   = (Join-Path $CodeDir "saveas_central.py")
                List   = (Join-Path $ToolDir "rvt_list_staging.txt")
                Year   = $dstY
                Detach = $false
            }) | Out-Null
    }
    elseif ($Job.saveas_needed) {
        $ver = $dstY
        if (-not $ver) { $ver = $srcY }
        if (-not $ver) { $ver = $autoY }
        $passes.Add([pscustomobject]@{
                Kind   = "rbp"
                Name   = "saveas"
                Task   = (Join-Path $CodeDir "saveas_central.py")
                List   = (Join-Path $ToolDir "rvt_list.txt")
                Year   = $ver
                Detach = $false
            }) | Out-Null
    }
    if ($Job.unc_move) {
        $passes.Add([pscustomobject]@{
                Kind   = "move"
                Name   = "move"
                Task   = (Join-Path $CodeDir "move_centrals.ps1")
                List   = (Join-Path $ToolDir "job_paths.csv")
                Year   = ""
                Detach = $false
                Overwrite = $allowOverwrite
            }) | Out-Null
    }
    if ($Job.relink) {
        $ver = $dstY
        if (-not $ver) { $ver = $srcY }
        if (-not $ver) { $ver = $autoY }
        $passes.Add([pscustomobject]@{
                Kind   = "rbp"
                Name   = "relink"
                Task   = (Join-Path $CodeDir "update_rvt_links.py")
                List   = (Join-Path $ToolDir "rvt_list_new.txt")
                Year   = $ver
                Detach = $false
            }) | Out-Null
    }
    if ($Job.delete_old -and -not $Job.copy) {
        $passes.Add([pscustomobject]@{
                Kind   = "delete"
                Name   = "delete"
                Task   = (Join-Path $CodeDir "delete_old_models.ps1")
                List   = (Join-Path $ToolDir "job_paths.csv")
                Year   = ""
                Detach = $false
            }) | Out-Null
    }
    if ($passes.Count -eq 0) { return @() }
    return $passes.ToArray()
}

$script:OpsJobFileNames = @(
    "job.json", "job_paths.csv", "mapping.csv", "rvt_list.txt",
    "rvt_list_new.txt", "rvt_list_staging.txt", "relink_map.csv"
)

function Backup-OpsJobFiles([string]$ToolDir) {
    $dir = Join-Path $ToolDir ("_jobbak_" + [Guid]::NewGuid().ToString("N"))
    New-Item -ItemType Directory -Path $dir -Force | Out-Null
    $had = New-Object System.Collections.Generic.List[string]
    foreach ($name in $script:OpsJobFileNames) {
        $src = Join-Path $ToolDir $name
        if (Test-Path -LiteralPath $src) {
            Copy-Item -LiteralPath $src -Destination (Join-Path $dir $name) -Force
            [void]$had.Add($name)
        }
    }
    return [pscustomobject]@{ Dir = $dir; Names = $had.ToArray() }
}

function Restore-OpsJobFiles($Backup, [string]$ToolDir) {
    if (-not $Backup -or -not $Backup.Dir) { return }
    $keep = @{}
    foreach ($name in @($Backup.Names)) { $keep[[string]$name] = $true }
    foreach ($name in $script:OpsJobFileNames) {
        $dst = Join-Path $ToolDir $name
        if ($keep.ContainsKey($name)) {
            Copy-Item -LiteralPath (Join-Path $Backup.Dir $name) -Destination $dst -Force
        }
        elseif (Test-Path -LiteralPath $dst) {
            Remove-Item -LiteralPath $dst -Force
        }
    }
    if (Test-Path -LiteralPath $Backup.Dir) {
        Remove-Item -LiteralPath $Backup.Dir -Recurse -Force
    }
}
