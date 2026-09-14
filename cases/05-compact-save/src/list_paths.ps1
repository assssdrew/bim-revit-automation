# Shared paths for rvt_list.txt: share when writable, else per-user local copy.
# Optional sidecar: scripts also embed these helpers if this file is missing on the share.

function Remove-LeftoverWriteProbes {
    param(
        [string]$Folder,
        [int]$Depth = 0
    )

    if ([string]::IsNullOrWhiteSpace($Folder)) { return }
    if (-not (Test-Path -LiteralPath $Folder -PathType Container)) { return }
    Get-ChildItem -LiteralPath $Folder -Force -File -ErrorAction SilentlyContinue |
        Where-Object {
            $_.Name -like '.write_probe*' -or $_.Name -like '.compact_write_probe*'
        } |
        ForEach-Object {
            try { Remove-Item -LiteralPath $_.FullName -Force -ErrorAction Stop } catch { }
        }
    if ($Depth -ge 2) { return }
    Get-ChildItem -LiteralPath $Folder -Force -Directory -ErrorAction SilentlyContinue |
        ForEach-Object {
            Remove-LeftoverWriteProbes -Folder $_.FullName -Depth ($Depth + 1)
        }
}

function Test-DirWritable {
    param([string]$Folder)

    if ([string]::IsNullOrWhiteSpace($Folder)) { return $false }
    if (-not (Test-Path -LiteralPath $Folder -PathType Container)) { return $false }

    # Always use an absolute path — Join-Path with a bad/empty base can write into the current dir (e.g. Desktop).
    try {
        $fullDir = [System.IO.Path]::GetFullPath($Folder)
    }
    catch {
        return $false
    }
    if ([string]::IsNullOrWhiteSpace($fullDir)) { return $false }

    # Clean only this folder (not recursive) before the probe.
    Get-ChildItem -LiteralPath $fullDir -Force -File -ErrorAction SilentlyContinue |
        Where-Object {
            $_.Name -like '.write_probe*' -or $_.Name -like '.compact_write_probe*'
        } |
        ForEach-Object {
            try { Remove-Item -LiteralPath $_.FullName -Force -ErrorAction Stop } catch { }
        }

    # Fixed name + DeleteOnClose: file should vanish when the stream closes (no GUID leftovers).
    $probe = [System.IO.Path]::Combine($fullDir, ".write_probe_tmp")
    $fs = $null
    try {
        $fs = New-Object System.IO.FileStream(
            $probe,
            [System.IO.FileMode]::Create,
            [System.IO.FileAccess]::Write,
            [System.IO.FileShare]::None,
            4096,
            [System.IO.FileOptions]::DeleteOnClose
        )
        $bytes = [System.Text.Encoding]::ASCII.GetBytes("ok")
        $fs.Write($bytes, 0, $bytes.Length)
        return $true
    }
    catch {
        return $false
    }
    finally {
        if ($fs) {
            try { $fs.Dispose() } catch { }
        }
        try {
            if ([System.IO.File]::Exists($probe)) {
                [System.IO.File]::Delete($probe)
            }
        }
        catch { }
    }
}

function Get-LocalBatchCompactDir {
    $candidates = @()
    if ($env:LOCALAPPDATA) {
        $candidates += (Join-Path $env:LOCALAPPDATA "BatchRvt\batch_compact_save")
    }
    if ($env:USERPROFILE) {
        $candidates += (Join-Path $env:USERPROFILE "Documents\doc\script\rbp\batch_compact_save")
    }

    foreach ($dir in $candidates) {
        if ([string]::IsNullOrWhiteSpace($dir)) { continue }
        try {
            if (-not (Test-Path -LiteralPath $dir)) {
                New-Item -ItemType Directory -Path $dir -Force | Out-Null
            }
            if (Test-DirWritable $dir) {
                return $dir
            }
        }
        catch { }
    }

    $fallback = Join-Path ([System.IO.Path]::GetTempPath()) "batch_compact_save"
    if (-not (Test-Path -LiteralPath $fallback)) {
        New-Item -ItemType Directory -Path $fallback -Force | Out-Null
    }
    return $fallback
}

function Get-ActiveListPointerPath {
    $dir = Join-Path $env:LOCALAPPDATA "BatchRvt\batch_compact_save"
    if (-not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
    return (Join-Path $dir "active_list.path")
}

function Save-ActiveListPointer {
    param([string]$ListPath)

    if ([string]::IsNullOrWhiteSpace($ListPath)) { return }
    try {
        $ptr = Get-ActiveListPointerPath
        Set-Content -LiteralPath $ptr -Value $ListPath.Trim() -Encoding UTF8 -ErrorAction Stop
    }
    catch { }
}

function Read-ActiveListPointer {
    $ptr = Get-ActiveListPointerPath
    if (-not (Test-Path -LiteralPath $ptr)) { return $null }
    try {
        $raw = (Get-Content -LiteralPath $ptr -Raw -ErrorAction Stop)
        if ([string]::IsNullOrWhiteSpace($raw)) { return $null }
        return $raw.Trim()
    }
    catch {
        return $null
    }
}

function Test-RvtListHasModels {
    param([string]$ListPath)

    if ([string]::IsNullOrWhiteSpace($ListPath)) { return $false }
    if (-not (Test-Path -LiteralPath $ListPath)) { return $false }
    $n = @(
        Get-Content -LiteralPath $ListPath -ErrorAction SilentlyContinue |
            ForEach-Object { $_.Trim() } |
            Where-Object {
                $_ -and -not $_.StartsWith("#") -and (
                    $_.ToLower().EndsWith(".rvt") -or $_ -match '^(?i)RSN://'
                )
            }
    ).Count
    return ($n -gt 0)
}

function New-ListPathInfo {
    param(
        [string]$OutList,
        [string]$CfgDir,
        [bool]$IsLocalList,
        [string]$ShareList
    )
    return @{
        OutList     = $OutList
        CfgDir      = $CfgDir
        IsLocalList = $IsLocalList
        ShareList   = $ShareList
    }
}

function Resolve-RvtListPaths {
    param(
        [string]$ToolDir,
        [switch]$ForLaunch
    )

    Remove-LeftoverWriteProbes -Folder $ToolDir

    $shareList = Join-Path $ToolDir "rvt_list.txt"
    $localDir = Get-LocalBatchCompactDir
    $localList = Join-Path $localDir "rvt_list.txt"
    $shareWritable = Test-DirWritable $ToolDir

    if ($ForLaunch) {
        $pointed = Read-ActiveListPointer
        if ($pointed -and (Test-RvtListHasModels $pointed)) {
            $isLocal = -not [string]::Equals($pointed, $shareList, [System.StringComparison]::OrdinalIgnoreCase)
            return (New-ListPathInfo -OutList $pointed -CfgDir (Split-Path -Parent $pointed) `
                    -IsLocalList:$isLocal -ShareList $shareList)
        }
        if ($shareWritable -and (Test-RvtListHasModels $shareList)) {
            return (New-ListPathInfo -OutList $shareList -CfgDir $ToolDir `
                    -IsLocalList:$false -ShareList $shareList)
        }
        if (Test-RvtListHasModels $localList) {
            return (New-ListPathInfo -OutList $localList -CfgDir $localDir `
                    -IsLocalList:$true -ShareList $shareList)
        }
        if (Test-RvtListHasModels $shareList) {
            return (New-ListPathInfo -OutList $shareList -CfgDir $ToolDir `
                    -IsLocalList:$false -ShareList $shareList)
        }
        if ($shareWritable) {
            return (New-ListPathInfo -OutList $shareList -CfgDir $ToolDir `
                    -IsLocalList:$false -ShareList $shareList)
        }
        return (New-ListPathInfo -OutList $localList -CfgDir $localDir `
                -IsLocalList:$true -ShareList $shareList)
    }

    if ($shareWritable) {
        return (New-ListPathInfo -OutList $shareList -CfgDir $ToolDir `
                -IsLocalList:$false -ShareList $shareList)
    }
    return (New-ListPathInfo -OutList $localList -CfgDir $localDir `
            -IsLocalList:$true -ShareList $shareList)
}

function Write-TextFileLines {
    param(
        [string]$Path,
        [string[]]$Lines
    )

    $dir = Split-Path -Parent $Path
    if ($dir -and -not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
    if (Test-Path -LiteralPath $Path) {
        try {
            $item = Get-Item -LiteralPath $Path -Force
            if ($item.IsReadOnly) { $item.IsReadOnly = $false }
        }
        catch { }
    }
    $Lines | Set-Content -LiteralPath $Path -Encoding UTF8 -ErrorAction Stop
}

function Write-CfgContent {
    param(
        [string]$Path,
        [string]$Value
    )

    try {
        $dir = Split-Path -Parent $Path
        if ($dir -and -not (Test-Path -LiteralPath $dir)) {
            New-Item -ItemType Directory -Path $dir -Force | Out-Null
        }
        Set-Content -LiteralPath $Path -Value $Value -Encoding UTF8 -ErrorAction Stop
    }
    catch {
        Write-Host ("WARN: could not save {0}: {1}" -f (Split-Path -Leaf $Path), $_.Exception.Message)
    }
}
