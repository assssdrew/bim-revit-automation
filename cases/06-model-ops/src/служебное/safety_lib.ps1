# Name: safety_lib.ps1
# Version: 3.1.0
# What it does: See the case README.
# Notes: Paths and project names in this portfolio copy are placeholders.

# Safety helpers for destructive batch_model_ops steps (delete old / move).
# Windows PowerShell 5.1 and PowerShell 7. ASCII source, UTF-8 WITH BOM.
# Dotted into delete_old_models.ps1 and move_centrals.ps1.
#
# Rules implemented here:
#   - job_paths.csv is read as real CSV (paths may contain ';').
#   - an old model is "verified" only when the Pass 1/2 reports say OK and
#     the new file exists and is not empty.
#   - nothing is deleted: files are MOVED to <folder>\_to_delete\<stamp>\
#     with a manifest.csv next to them (restore = move back).

function ConvertTo-PathKey([string]$Path) {
    if ([string]::IsNullOrWhiteSpace($Path)) { return "" }
    return $Path.Trim().Trim('"').Replace("\", "/").TrimEnd("/").ToUpperInvariant()
}

function Get-NameKey([string]$PathOrName) {
    if ([string]::IsNullOrWhiteSpace($PathOrName)) { return "" }
    $n = $PathOrName.Trim().Trim('"').Replace("\", "/").Split("/")[-1].Trim()
    if ($n.ToLowerInvariant().EndsWith(".rvt")) { $n = $n.Substring(0, $n.Length - 4) }
    return $n.ToUpperInvariant()
}

function Test-RsnPath([string]$Path) {
    if ([string]::IsNullOrWhiteSpace($Path)) { return $false }
    return $Path.Trim().StartsWith("RSN://", [StringComparison]::OrdinalIgnoreCase)
}

function Format-CsvField([string]$Val) {
    if ($null -eq $Val) { return "" }
    $s = ([string]$Val).Replace("`r", " ").Replace("`n", " ")
    if ($s.Contains(";") -or $s.Contains('"')) {
        return '"' + $s.Replace('"', '""') + '"'
    }
    return $s
}

function Get-RunStamp {
    return (Get-Date -Format "yyyy-MM-dd_HHmmss")
}

function Confirm-Yes([string]$Prompt) {
    $answer = Read-Host $Prompt
    return ($answer -ceq "YES")
}

function Get-CsvProp($Row, [string]$Name) {
    $p = $Row.PSObject.Properties[$Name]
    if ($null -eq $p -or $null -eq $p.Value) { return "" }
    return ([string]$p.Value).Trim()
}

# ---------------------------------------------------------------- job files

function Read-JobPaths([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "job_paths.csv not found: $Path"
    }
    $raw = @(Import-Csv -LiteralPath $Path -Delimiter ";" -Encoding UTF8)
    $rows = New-Object System.Collections.Generic.List[object]
    if ($raw.Count -eq 0) { return @() }
    $names = @($raw[0].PSObject.Properties.Name)
    if (($names -notcontains "old_path") -or ($names -notcontains "new_path")) {
        throw "job_paths.csv: header must contain old_path;new_path (got: $($names -join ';'))"
    }
    foreach ($r in $raw) {
        $old = (Get-CsvProp $r "old_path").Trim('"')
        if (-not $old -or $old.StartsWith("#")) { continue }
        $new = (Get-CsvProp $r "new_path").Trim('"')
        $oldName = Get-CsvProp $r "old_name"
        $newName = Get-CsvProp $r "new_name"
        if (-not $oldName) { $oldName = Get-NameKey $old }
        if (-not $newName) { $newName = Get-NameKey $new }
        $nb = Get-CsvProp $r "needs_buffer"
        $rows.Add([pscustomobject]@{
                OldPath      = $old
                NewPath      = $new
                StagingPath  = (Get-CsvProp $r "staging_path").Trim('"')
                OldName      = $oldName
                NewName      = $newName
                NeedsBuffer  = ($nb -eq "1" -or $nb -ieq "true")
            }) | Out-Null
    }
    if ($rows.Count -eq 0) { return @() }
    return $rows.ToArray()
}

function Read-JobJson([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "job.json not found: $Path (needed to know which passes ran)"
    }
    $text = Get-Content -LiteralPath $Path -Raw -Encoding UTF8
    return ($text | ConvertFrom-Json)
}

function Get-JobFlag($Job, [string]$Name) {
    $p = $Job.PSObject.Properties[$Name]
    if ($null -eq $p -or $null -eq $p.Value) { return $false }
    return [bool]$p.Value
}

# ------------------------------------------------------------- pass reports

$script:ReportKinds = @{
    saveas = @{ Marker = "_active_saveas_central.txt"; Prefix = "saveas_central_" }
    detach = @{ Marker = "_active_detach_to_buffer.txt"; Prefix = "detach_to_buffer_" }
    relink = @{ Marker = "_active_update_rvt_links.txt"; Prefix = "update_rvt_links_" }
}

function Find-PassReport {
    param(
        [Parameter(Mandatory = $true)][string]$ReportsDir,
        [Parameter(Mandatory = $true)][ValidateSet("saveas", "detach", "relink")][string]$Kind,
        [Parameter(Mandatory = $true)][datetime]$NotOlderThan
    )
    $k = $script:ReportKinds[$Kind]
    $found = $null
    $marker = Join-Path $ReportsDir $k.Marker
    if (Test-Path -LiteralPath $marker -PathType Leaf) {
        $p = (Get-Content -LiteralPath $marker -Raw -Encoding UTF8)
        if ($p) { $p = $p.Trim() }
        if ($p -and (Test-Path -LiteralPath $p -PathType Leaf)) { $found = Get-Item -LiteralPath $p }
    }
    if (-not $found -and (Test-Path -LiteralPath $ReportsDir -PathType Container)) {
        $found = Get-ChildItem -LiteralPath $ReportsDir -File -Filter ($k.Prefix + "*.csv") -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -notmatch '_details\.csv$' } |
            Sort-Object LastWriteTime -Descending |
            Select-Object -First 1
    }
    if (-not $found) {
        return [pscustomobject]@{ Path = $null; Reason = "no $Kind report in $ReportsDir" }
    }
    if ($found.LastWriteTime -lt $NotOlderThan.AddSeconds(-2)) {
        return [pscustomobject]@{
            Path   = $null
            Reason = "$Kind report is older than the saved job (stale): $($found.Name)"
        }
    }
    return [pscustomobject]@{ Path = $found.FullName; Reason = "" }
}

function Add-StatusEntry($Index, [string]$Key, [string]$Status) {
    if (-not $Key) { return }
    if (-not $Index.ContainsKey($Key)) {
        $Index[$Key] = [pscustomobject]@{ Last = $Status; EverOk = $false }
    }
    $e = $Index[$Key]
    $e.Last = $Status
    if ($Status -eq "OK") { $e.EverOk = $true }
}

# Keys: saveas/detach -> path key of old_path; relink -> "N:<NAME>" (host name) and
# "P:<path key>" when host (or a host_path column) holds a full path.
function Read-PassReport {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][ValidateSet("saveas", "detach", "relink")][string]$Kind
    )
    $index = @{}
    $rows = @(Import-Csv -LiteralPath $Path -Delimiter ";" -Encoding UTF8)
    foreach ($r in $rows) {
        $st = (Get-CsvProp $r "status").ToUpperInvariant()
        if (-not $st) { continue }
        if ($Kind -eq "relink") {
            # update_rvt_links.py writes the full host path into "host";
            # older/other reports may carry only the model name.
            $hostRaw = Get-CsvProp $r "host"
            $hostName = Get-NameKey $hostRaw
            if ($hostName) { Add-StatusEntry $index ("N:" + $hostName) $st }
            if ($hostRaw.Contains("/") -or $hostRaw.Contains("\")) {
                Add-StatusEntry $index ("P:" + (ConvertTo-PathKey $hostRaw)) $st
            }
            $hp = ConvertTo-PathKey (Get-CsvProp $r "host_path")
            if ($hp) { Add-StatusEntry $index ("P:" + $hp) $st }
        }
        else {
            Add-StatusEntry $index (ConvertTo-PathKey (Get-CsvProp $r "old_path")) $st
        }
    }
    return $index
}

function Get-EffectiveStatus($Entry) {
    if ($null -eq $Entry) { return $null }
    if ($Entry.Last -eq "OK") { return "OK" }
    # a re-run that found the target already there is fine if an earlier run wrote it
    if ($Entry.Last -eq "SKIP_EXISTS" -and $Entry.EverOk) { return "OK" }
    return $Entry.Last
}

function New-PassContext {
    param(
        [Parameter(Mandatory = $true)]$Job,
        [Parameter(Mandatory = $true)][object[]]$Rows,
        [Parameter(Mandatory = $true)][string]$ReportsDir,
        [Parameter(Mandatory = $true)][datetime]$JobTime
    )
    $anyBuffer = @($Rows | Where-Object { $_.NeedsBuffer }).Count -gt 0
    $needSaveAs = $anyBuffer -or (Get-JobFlag $Job "saveas_needed")
    $relink = Get-JobFlag $Job "relink"

    $ctx = [pscustomobject]@{
        NeedSaveAs = $needSaveAs
        Relink     = $relink
        Reports    = @{}
        NameCounts = @{}
    }
    $want = New-Object System.Collections.Generic.List[string]
    if ($anyBuffer) { $want.Add("detach") | Out-Null }
    if ($needSaveAs) { $want.Add("saveas") | Out-Null }
    if ($relink) { $want.Add("relink") | Out-Null }
    foreach ($kind in $want) {
        $rep = Find-PassReport -ReportsDir $ReportsDir -Kind $kind -NotOlderThan $JobTime
        $idx = @{}
        if ($rep.Path) { $idx = Read-PassReport -Path $rep.Path -Kind $kind }
        $ctx.Reports[$kind] = [pscustomobject]@{ Path = $rep.Path; Reason = $rep.Reason; Index = $idx }
    }
    foreach ($r in $Rows) {
        $nk = Get-NameKey $r.NewName
        if (-not $nk) { continue }
        if ($ctx.NameCounts.ContainsKey($nk)) { $ctx.NameCounts[$nk]++ } else { $ctx.NameCounts[$nk] = 1 }
    }
    return $ctx
}

# -------------------------------------------------------------- file checks

function Test-NewFile([string]$Path) {
    if ([string]::IsNullOrWhiteSpace($Path)) {
        return [pscustomobject]@{ State = "missing"; Size = 0 }
    }
    if (Test-RsnPath $Path) {
        return [pscustomobject]@{ State = "unverifiable"; Size = 0 }
    }
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        return [pscustomobject]@{ State = "missing"; Size = 0 }
    }
    $len = (Get-Item -LiteralPath $Path).Length
    if ($len -le 0) { return [pscustomobject]@{ State = "empty"; Size = 0 } }
    return [pscustomobject]@{ State = "ok"; Size = $len }
}

function Get-PassStatusReason($Ctx, [string]$Kind, [string]$Key, [string]$Label) {
    $rep = $Ctx.Reports[$Kind]
    if (-not $rep.Path) { return ("{0}: {1}" -f $Label, $rep.Reason) }
    $st = Get-EffectiveStatus $rep.Index[$Key]
    if ($st -eq "OK") { return $null }
    if (-not $st) { return ("{0}: no record in report for this model" -f $Label) }
    return ("{0}: status {1}" -f $Label, $st)
}

# Returns Ok = $true only when every pass the job ran reports OK for this row
# and the new file is present and non-empty (RSN new file: cannot be checked).
function Test-ModelVerified {
    param(
        [Parameter(Mandatory = $true)]$Row,
        [Parameter(Mandatory = $true)]$Ctx
    )
    $reasons = New-Object System.Collections.Generic.List[string]
    $checked = 0

    if ($Row.NeedsBuffer) {
        $checked++
        $r1 = Get-PassStatusReason $Ctx "detach" (ConvertTo-PathKey $Row.OldPath) "detach"
        if ($r1) { $reasons.Add($r1) | Out-Null }
        $r2 = Get-PassStatusReason $Ctx "saveas" (ConvertTo-PathKey $Row.StagingPath) "saveas(buffer)"
        if ($r2) { $reasons.Add($r2) | Out-Null }
    }
    elseif ($Ctx.NeedSaveAs) {
        $checked++
        $r1 = Get-PassStatusReason $Ctx "saveas" (ConvertTo-PathKey $Row.OldPath) "saveas"
        if ($r1) { $reasons.Add($r1) | Out-Null }
    }

    if ($Ctx.Relink) {
        $checked++
        $rep = $Ctx.Reports["relink"]
        $nameKey = Get-NameKey $Row.NewName
        $pathKey = "P:" + (ConvertTo-PathKey $Row.NewPath)
        $label = "relink"
        if ($rep.Path -and $rep.Index.ContainsKey($pathKey)) {
            $st = Get-EffectiveStatus $rep.Index[$pathKey]
            if (($st -ne "OK") -and ($st -ne "NO_CHANGES")) {
                $reasons.Add(("{0}: status {1}" -f $label, $st)) | Out-Null
            }
        }
        elseif ($rep.Path -and $Ctx.NameCounts[$nameKey] -gt 1) {
            $reasons.Add("relink: model name is not unique in the job, report has no host_path") | Out-Null
        }
        else {
            $r3 = Get-PassStatusReason $Ctx "relink" ("N:" + $nameKey) $label
            if ($r3) {
                # NO_CHANGES = host had no links to update; that is a success
                if ($r3 -notmatch 'status NO_CHANGES$') { $reasons.Add($r3) | Out-Null }
            }
        }
    }

    if ($checked -eq 0) {
        $reasons.Add("job has no saveas/relink pass to verify against") | Out-Null
    }

    $nf = Test-NewFile $Row.NewPath
    switch ($nf.State) {
        "missing" { $reasons.Add("new file not found: $($Row.NewPath)") | Out-Null }
        "empty" { $reasons.Add("new file is empty (0 bytes): $($Row.NewPath)") | Out-Null }
    }

    return [pscustomobject]@{
        Ok      = ($reasons.Count -eq 0)
        Reasons = $reasons.ToArray()
        NewFile = $nf
    }
}

# --------------------------------------------------------------- quarantine

# Moves a file to <its folder>\_to_delete\<Stamp>\ (same volume -> instant) and
# appends to manifest.csv there. Never overwrites; name clashes get _1, _2, ...
function Move-ToQuarantine {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Stamp,
        [string]$Reason = ""
    )
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "not a file: $Path" }
    $dir = Split-Path -Parent $Path
    $leaf = Split-Path -Leaf $Path
    $qdir = Join-Path (Join-Path $dir "_to_delete") $Stamp
    if (-not (Test-Path -LiteralPath $qdir)) {
        New-Item -ItemType Directory -Path $qdir -Force | Out-Null
    }
    $dest = Join-Path $qdir $leaf
    $i = 0
    $base = [System.IO.Path]::GetFileNameWithoutExtension($leaf)
    $ext = [System.IO.Path]::GetExtension($leaf)
    while (Test-Path -LiteralPath $dest) {
        $i++
        $dest = Join-Path $qdir ("{0}_{1}{2}" -f $base, $i, $ext)
    }
    $size = (Get-Item -LiteralPath $Path).Length
    Move-Item -LiteralPath $Path -Destination $dest
    if ((Test-Path -LiteralPath $Path) -or -not (Test-Path -LiteralPath $dest -PathType Leaf) `
            -or ((Get-Item -LiteralPath $dest).Length -ne $size)) {
        throw "quarantine move not verified: $Path -> $dest"
    }
    $manifest = Join-Path $qdir "manifest.csv"
    if (-not (Test-Path -LiteralPath $manifest)) {
        Set-Content -LiteralPath $manifest -Encoding UTF8 -Value "original_path;quarantine_path;moved_at;reason"
    }
    $line = "{0};{1};{2};{3}" -f (Format-CsvField $Path), (Format-CsvField $dest),
    (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), (Format-CsvField $Reason)
    Add-Content -LiteralPath $manifest -Encoding UTF8 -Value $line
    return $dest
}
