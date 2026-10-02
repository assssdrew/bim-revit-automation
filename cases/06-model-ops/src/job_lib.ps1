# Job helpers for batch_model_ops (Windows PowerShell 5.1).
# Naming ops, path rewrite, mapping/list export. ASCII source.

Set-StrictMode -Version Latest

function Get-ToolDir {
    if ($PSScriptRoot) { return $PSScriptRoot }
    return (Split-Path -Parent $MyInvocation.MyCommand.Path)
}

function Get-RbpRoot {
    $userProfile = $env:USERPROFILE
    if ([string]::IsNullOrWhiteSpace($userProfile)) { $userProfile = $env:PUBLIC }
    return (Join-Path $userProfile "Documents\doc\script\rbp")
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
        if ($i -lt $a.Count -and $j -lt $b.Count -and ($a[$i] -ieq $b[$j])) {
            $i++
            $j++
            continue
        }
        $aInB = $false
        $bInA = $false
        if ($i -lt $a.Count -and $j -lt $b.Count) {
            $aInB = Test-TokenInTail $b $j $a[$i]
            $bInA = Test-TokenInTail $a $i $b[$j]
        }
        if ($i -lt $a.Count -and (($j -ge $b.Count) -or (-not $aInB))) {
            $ops.Add((New-NameOp -Kind delete -Token $a[$i])) | Out-Null
            $i++
            continue
        }
        if ($j -lt $b.Count -and (($i -ge $a.Count) -or (-not $bInA))) {
            $after = ""
            if ($i -gt 0) { $after = $a[$i - 1] }
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
    return (Join-Path $BufferDir ("{0}.rvt" -f $NewBase))
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
        [hashtable]$NameOverrides
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
        $newPath = ConvertTo-LocationPath -OldPath $oldPath -NewBase $newName `
            -Mode $LocationMode -OldSegment $OldSegment -NewSegment $NewSegment `
            -OldPrefix $OldPrefix -NewPrefix $NewPrefix
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
        $rows.Add([pscustomobject]@{
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
        $lines.Add(("{0}) Windows Move: run_cascade.cmd -> Move UNC" -f $n)) | Out-Null
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
        $lines.Add(("{0}) After green Pass 2: delete_old_models.cmd" -f $n)) | Out-Null
        $lines.Add("   UNC auto-delete; RSN -> to_delete_rsn.txt -> Admin") | Out-Null
        if ($needBuf) {
            $lines.Add("   Staging on this PC is deleted by the same script.") | Out-Null
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
        [object[]]$Rows
    )
    $passes = New-Object System.Collections.Generic.List[object]
    $needBuf = @($Rows | Where-Object { $_.needs_buffer }).Count -gt 0
    $srcY = [string]$Job.source_year
    $dstY = [string]$Job.target_year

    if ($needBuf) {
        $passes.Add([pscustomobject]@{
                Kind   = "rbp"
                Name   = "detach"
                Task   = (Join-Path $ToolDir "detach_to_buffer.py")
                List   = (Join-Path $ToolDir "rvt_list.txt")
                Year   = $srcY
                Detach = $true
            }) | Out-Null
        $passes.Add([pscustomobject]@{
                Kind   = "rbp"
                Name   = "saveas"
                Task   = (Join-Path $ToolDir "saveas_central.py")
                List   = (Join-Path $ToolDir "rvt_list_staging.txt")
                Year   = $dstY
                Detach = $false
            }) | Out-Null
    }
    elseif ($Job.saveas_needed) {
        $ver = $dstY
        if (-not $ver) { $ver = $srcY }
        $passes.Add([pscustomobject]@{
                Kind   = "rbp"
                Name   = "saveas"
                Task   = (Join-Path $ToolDir "saveas_central.py")
                List   = (Join-Path $ToolDir "rvt_list.txt")
                Year   = $ver
                Detach = $false
            }) | Out-Null
    }
    if ($Job.unc_move) {
        $passes.Add([pscustomobject]@{
                Kind   = "move"
                Name   = "move"
                Task   = (Join-Path $ToolDir "move_centrals.ps1")
                List   = (Join-Path $ToolDir "job_paths.csv")
                Year   = ""
                Detach = $false
            }) | Out-Null
    }
    if ($Job.relink) {
        $ver = $dstY
        if (-not $ver) { $ver = $srcY }
        $passes.Add([pscustomobject]@{
                Kind   = "rbp"
                Name   = "relink"
                Task   = (Join-Path $ToolDir "update_rvt_links.py")
                List   = (Join-Path $ToolDir "rvt_list_new.txt")
                Year   = $ver
                Detach = $false
            }) | Out-Null
    }
    if ($Job.delete_old -and -not $Job.copy) {
        $passes.Add([pscustomobject]@{
                Kind   = "delete"
                Name   = "delete"
                Task   = (Join-Path $ToolDir "delete_old_models.ps1")
                List   = (Join-Path $ToolDir "job_paths.csv")
                Year   = ""
                Detach = $false
            }) | Out-Null
    }
    if ($passes.Count -eq 0) { return @() }
    return $passes.ToArray()
}
