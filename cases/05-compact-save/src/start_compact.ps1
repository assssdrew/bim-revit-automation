# Launch Revit Batch Processor: compact save, Create New Local, no extra checkboxes.
# ASCII-friendly messages for Windows PowerShell 5.1.

$ErrorActionPreference = "Stop"
$ToolDir = Split-Path -Parent $MyInvocation.MyCommand.Path

$_listPathsFile = Join-Path $ToolDir "list_paths.ps1"
$_listPathsLoaded = $false
if (Test-Path -LiteralPath $_listPathsFile) {
    try {
        . $_listPathsFile
        $_listPathsLoaded = (
            [bool](Get-Command Resolve-RvtListPaths -ErrorAction SilentlyContinue) -and
            [bool](Get-Command Save-ActiveListPointer -ErrorAction SilentlyContinue) -and
            [bool](Get-Command Read-ActiveListPointer -ErrorAction SilentlyContinue) -and
            [bool](Get-Command Remove-LeftoverWriteProbes -ErrorAction SilentlyContinue)
        )
    }
    catch {
        $_listPathsLoaded = $false
    }
}
if (-not $_listPathsLoaded) {
    function Remove-LeftoverWriteProbes {
        param(
            [string]$Folder,
            [int]$Depth = 0
        )
        if ([string]::IsNullOrWhiteSpace($Folder)) { return }
        if (-not (Test-Path -LiteralPath $Folder -PathType Container)) { return }
        Get-ChildItem -LiteralPath $Folder -Force -File -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -like '.write_probe*' -or $_.Name -like '.compact_write_probe*' } |
            ForEach-Object { try { Remove-Item -LiteralPath $_.FullName -Force -ErrorAction Stop } catch { } }
        if ($Depth -ge 2) { return }
        Get-ChildItem -LiteralPath $Folder -Force -Directory -ErrorAction SilentlyContinue |
            ForEach-Object { Remove-LeftoverWriteProbes -Folder $_.FullName -Depth ($Depth + 1) }
    }
    function Test-DirWritable {
        param([string]$Folder)
        if ([string]::IsNullOrWhiteSpace($Folder)) { return $false }
        if (-not (Test-Path -LiteralPath $Folder -PathType Container)) { return $false }
        try { $fullDir = [System.IO.Path]::GetFullPath($Folder) } catch { return $false }
        if ([string]::IsNullOrWhiteSpace($fullDir)) { return $false }
        Get-ChildItem -LiteralPath $fullDir -Force -File -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -like '.write_probe*' -or $_.Name -like '.compact_write_probe*' } |
            ForEach-Object { try { Remove-Item -LiteralPath $_.FullName -Force -ErrorAction Stop } catch { } }
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
        catch { return $false }
        finally {
            if ($fs) { try { $fs.Dispose() } catch { } }
            try { if ([System.IO.File]::Exists($probe)) { [System.IO.File]::Delete($probe) } } catch { }
        }
    }
    function Get-LocalBatchCompactDir {
        $candidates = @()
        if ($env:LOCALAPPDATA) { $candidates += (Join-Path $env:LOCALAPPDATA "BatchRvt\batch_compact_save") }
        if ($env:USERPROFILE) { $candidates += (Join-Path $env:USERPROFILE "Documents\doc\script\rbp\batch_compact_save") }
        foreach ($dir in $candidates) {
            if ([string]::IsNullOrWhiteSpace($dir)) { continue }
            try {
                if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
                if (Test-DirWritable $dir) { return $dir }
            }
            catch { }
        }
        $fallback = Join-Path ([System.IO.Path]::GetTempPath()) "batch_compact_save"
        if (-not (Test-Path -LiteralPath $fallback)) { New-Item -ItemType Directory -Path $fallback -Force | Out-Null }
        return $fallback
    }
    function Get-ActiveListPointerPath {
        $dir = Join-Path $env:LOCALAPPDATA "BatchRvt\batch_compact_save"
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        return (Join-Path $dir "active_list.path")
    }
    function Save-ActiveListPointer {
        param([string]$ListPath)
        if ([string]::IsNullOrWhiteSpace($ListPath)) { return }
        try { Set-Content -LiteralPath (Get-ActiveListPointerPath) -Value $ListPath.Trim() -Encoding UTF8 -ErrorAction Stop } catch { }
    }
    function Read-ActiveListPointer {
        $ptr = Get-ActiveListPointerPath
        if (-not (Test-Path -LiteralPath $ptr)) { return $null }
        try {
            $raw = (Get-Content -LiteralPath $ptr -Raw -ErrorAction Stop)
            if ([string]::IsNullOrWhiteSpace($raw)) { return $null }
            return $raw.Trim()
        }
        catch { return $null }
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
        param([string]$OutList, [string]$CfgDir, [bool]$IsLocalList, [string]$ShareList)
        return @{ OutList = $OutList; CfgDir = $CfgDir; IsLocalList = $IsLocalList; ShareList = $ShareList }
    }
    function Resolve-RvtListPaths {
        param([string]$ToolDir, [switch]$ForLaunch)
        Remove-LeftoverWriteProbes -Folder $ToolDir
        $shareList = Join-Path $ToolDir "rvt_list.txt"
        $localDir = Get-LocalBatchCompactDir
        $localList = Join-Path $localDir "rvt_list.txt"
        $shareWritable = Test-DirWritable $ToolDir
        if ($ForLaunch) {
            $pointed = Read-ActiveListPointer
            if ($pointed -and (Test-RvtListHasModels $pointed)) {
                $isLocal = -not [string]::Equals($pointed, $shareList, [System.StringComparison]::OrdinalIgnoreCase)
                return (New-ListPathInfo -OutList $pointed -CfgDir (Split-Path -Parent $pointed) -IsLocalList:$isLocal -ShareList $shareList)
            }
            if ($shareWritable -and (Test-RvtListHasModels $shareList)) {
                return (New-ListPathInfo -OutList $shareList -CfgDir $ToolDir -IsLocalList:$false -ShareList $shareList)
            }
            if (Test-RvtListHasModels $localList) {
                return (New-ListPathInfo -OutList $localList -CfgDir $localDir -IsLocalList:$true -ShareList $shareList)
            }
            if (Test-RvtListHasModels $shareList) {
                return (New-ListPathInfo -OutList $shareList -CfgDir $ToolDir -IsLocalList:$false -ShareList $shareList)
            }
            if ($shareWritable) {
                return (New-ListPathInfo -OutList $shareList -CfgDir $ToolDir -IsLocalList:$false -ShareList $shareList)
            }
            return (New-ListPathInfo -OutList $localList -CfgDir $localDir -IsLocalList:$true -ShareList $shareList)
        }
        if ($shareWritable) {
            return (New-ListPathInfo -OutList $shareList -CfgDir $ToolDir -IsLocalList:$false -ShareList $shareList)
        }
        return (New-ListPathInfo -OutList $localList -CfgDir $localDir -IsLocalList:$true -ShareList $shareList)
    }
}

$listPaths = Resolve-RvtListPaths -ToolDir $ToolDir -ForLaunch
if (Get-Command Remove-LeftoverWriteProbes -ErrorAction SilentlyContinue) {
    Remove-LeftoverWriteProbes -Folder $ToolDir
}
$TaskScript = Join-Path $ToolDir "compact_save.py"
$FileList = $listPaths.OutList

function Find-BatchRvt {
    $names = @("BatchRvt.exe", "BatchRvtGUI.exe")
    $dirs = New-Object System.Collections.Generic.List[string]

    Get-Process -ErrorAction SilentlyContinue |
        Where-Object { $_.ProcessName -match '^(?i)BatchRvt(GUI)?$|RevitBatchProcessor$' } |
        ForEach-Object {
            try {
                $exeDir = Split-Path -Parent $_.MainModule.FileName
                if ($exeDir) { $dirs.Add($exeDir) }
            }
            catch { }
        }

    foreach ($root in @($env:LOCALAPPDATA, $env:ProgramFiles, ${env:ProgramFiles(x86)}, $env:ProgramW6432)) {
        if ([string]::IsNullOrWhiteSpace($root)) { continue }
        foreach ($rel in @(
            "RevitBatchProcessor",
            "BatchRvt",
            "Revit Batch Processor",
            "Programs\Revit Batch Processor"
        )) {
            $dirs.Add((Join-Path $root $rel))
        }
    }

    if ($env:LOCALAPPDATA -and (Test-Path -LiteralPath $env:LOCALAPPDATA)) {
        Get-ChildItem -LiteralPath $env:LOCALAPPDATA -Filter "BatchRvt.exe" -Recurse -ErrorAction SilentlyContinue -Depth 5 |
            Select-Object -First 8 |
            ForEach-Object { $dirs.Add($_.DirectoryName) }
    }

    $seen = @{}
    foreach ($dir in $dirs) {
        if ([string]::IsNullOrWhiteSpace($dir)) { continue }
        try { $full = [System.IO.Path]::GetFullPath($dir) } catch { continue }
        if ($seen.ContainsKey($full)) { continue }
        $seen[$full] = $true
        foreach ($name in $names) {
            $exe = Join-Path $full $name
            if (Test-Path -LiteralPath $exe) { return $exe }
        }
    }
    return $null
}

function Read-ModelList {
    if (-not (Test-Path -LiteralPath $FileList)) { return @() }
    @(
        Get-Content -LiteralPath $FileList -ErrorAction SilentlyContinue |
            ForEach-Object { $_.Trim() } |
            Where-Object {
                $_ -and -not $_.StartsWith("#") -and (
                    $_.ToLower().EndsWith(".rvt") -or $_ -match '^(?i)RSN://'
                )
            }
    )
}

function Get-RvtYearFromFile([string]$Path) {
    if ([string]::IsNullOrWhiteSpace($Path)) { return $null }
    if ($Path -match '^(?i)RSN://') { return $null }
    try {
        $fs = [System.IO.File]::Open(
            $Path,
            [System.IO.FileMode]::Open,
            [System.IO.FileAccess]::Read,
            [System.IO.FileShare]::ReadWrite
        )
        try {
            $len = [Math]::Min(32768, [int]$fs.Length)
            $buf = New-Object byte[] $len
            $n = $fs.Read($buf, 0, $len)
        }
        finally {
            $fs.Close()
        }
        $uni = [System.Text.Encoding]::Unicode.GetString($buf, 0, $n)
        if ($uni -match 'Autodesk Revit 20(\d{2})') {
            return ("20" + $Matches[1])
        }
        $ascii = [System.Text.Encoding]::ASCII.GetString($buf, 0, $n)
        if ($ascii -match 'Autodesk Revit 20(\d{2})') {
            return ("20" + $Matches[1])
        }
    }
    catch { }
    return $null
}

function Guess-Year([string]$Path) {
    if ($Path -match '(?i)_R(20|21|22|23|24|25|26)\b') {
        return ("20" + $Matches[1])
    }
    if ($Path -match '(?i)RSN://revit-server-2021\.example\.local\b') { return "2021" }
    if ($Path -match '(?i)RSN://revit-server-2022\.example\.local\b') { return "2022" }
    if ($Path -match '(?i)RSN://revit-server-2023\.example\.local\b') { return "2023" }
    if ($Path -match '(?i)RSN://revit-server-2024\.example\.local\b') { return "2024" }
    if ($Path -match '(?i)RSN://revit-server-2021') { return "2021" }
    if ($Path -match '(?i)RSN://revit-server-2022') { return "2022" }
    if ($Path -match '(?i)RSN://revit-server-2023') { return "2023" }
    if ($Path -match '(?i)RSN://revit-server-2024') { return "2024" }
    if ($Path -match '(?i)Revit Server 202([1-6])') { return ("202" + $Matches[1]) }
    return (Get-RvtYearFromFile $Path)
}

function Get-LiveStatusPath {
    $dir = Join-Path $env:LOCALAPPDATA "BatchRvt\batch_compact_save"
    if (-not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
    return (Join-Path $dir "live_status.txt")
}

function Save-ActiveModePointer {
    param([string]$Mode)
    $env:BATCH_COMPACT_MODE = $Mode
    $dir = Join-Path $env:LOCALAPPDATA "BatchRvt\batch_compact_save"
    if (-not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
    try {
        Set-Content -LiteralPath (Join-Path $dir "active_mode.txt") -Value $Mode -Encoding UTF8 -ErrorAction Stop
    }
    catch { }
}

function Get-CompactRunIdPath {
    $dir = Join-Path $env:LOCALAPPDATA "BatchRvt\batch_compact_save"
    if (-not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
    return (Join-Path $dir "active_run_id.txt")
}

function Save-CompactRunId {
    $id = [guid]::NewGuid().ToString("N")
    $env:BATCH_COMPACT_RUN_ID = $id
    try {
        Set-Content -LiteralPath (Get-CompactRunIdPath) -Value $id -Encoding UTF8 -ErrorAction Stop
    }
    catch { }
    return $id
}

function Get-CompactModeArgs([string]$Mode) {
    if ($Mode -eq "deep") { return @("--audit") }
    return @("--create_new_local")
}

function Group-ModelsByYear {
    param([string[]]$Models)
    $byYear = @{}
    foreach ($m in @($Models)) {
        $y = Guess-Year $m
        if (-not $y) { $y = "" }
        if (-not $byYear.ContainsKey($y)) {
            $byYear[$y] = New-Object System.Collections.Generic.List[string]
        }
        [void]$byYear[$y].Add($m)
    }
    $ordered = @($byYear.Keys | Where-Object { $_ } | Sort-Object)
    if ($byYear.ContainsKey("")) { $ordered += "" }
    $groups = New-Object System.Collections.Generic.List[object]
    foreach ($y in $ordered) {
        $label = if ($y) { [string]$y } else { "авто" }
        [void]$groups.Add(@{
            Year   = [string]$y
            Models = @($byYear[$y])
            Label  = $label
        })
    }
    return $groups
}

function Get-YearFileListPath([string]$Year) {
    $dir = Join-Path $env:LOCALAPPDATA "BatchRvt\batch_compact_save\year_lists"
    if (-not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
    $tag = if ($Year) { $Year } else { "auto" }
    return (Join-Path $dir ("rvt_list_{0}.txt" -f $tag))
}

function Write-RvtListSafe {
    param(
        [string[]]$Paths,
        [string]$ListPath = ""
    )
    if ([string]::IsNullOrWhiteSpace($ListPath)) { $ListPath = $FileList }
    $unique = @(
        $Paths |
            ForEach-Object { $_.Trim() } |
            Where-Object { $_ } |
            Sort-Object -Unique
    )
    if ($unique.Count -eq 0) { return $false }
    try {
        if (Get-Command Write-TextFileLines -ErrorAction SilentlyContinue) {
            Write-TextFileLines -Path $ListPath -Lines $unique
        }
        else {
            $dir = Split-Path -Parent $ListPath
            if ($dir -and -not (Test-Path -LiteralPath $dir)) {
                New-Item -ItemType Directory -Path $dir -Force | Out-Null
            }
            $unique | Set-Content -LiteralPath $ListPath -Encoding UTF8 -ErrorAction Stop
        }
        return $true
    }
    catch {
        return $false
    }
}

function Find-RbpRsnPatch {
    $candidates = @(
        (Join-Path $ToolDir "служебное\rbp_rsn_patch\apply_rsn_support.ps1")
        (Join-Path $ToolDir "rbp_rsn_patch\apply_rsn_support.ps1")
    )
    $toolsRoot = Split-Path -Parent $ToolDir
    $candidates += (Join-Path $toolsRoot "batch_set_project_units\rbp_rsn_patch\apply_rsn_support.ps1")
    foreach ($p in $candidates) {
        if (Test-Path -LiteralPath $p) { return $p }
    }
    return $null
}

function Ensure-RbpReady {
    param(
        [string]$Mode,
        [string[]]$Models,
        [string]$Exe
    )
    $hasRsn = @($Models | Where-Object { $_ -match '^(?i)RSN://' }).Count -gt 0
    if (-not ($hasRsn -or ($Mode -eq "deep"))) { return $true }
    $patchPs = Find-RbpRsnPatch
    if (-not $patchPs) { return $true }
    $scriptsDir = Join-Path (Split-Path -Parent $Exe) "Scripts"
    if (Test-Path -LiteralPath (Join-Path $scriptsDir "revit_file_list.py")) {
        & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $patchPs -ScriptsPath $scriptsDir -Quiet
    }
    else {
        & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $patchPs -Quiet
    }
    return ($LASTEXITCODE -eq 0)
}

function Start-CompactYearBatch {
    param(
        [Parameter(Mandatory = $true)]
        $Group,
        [Parameter(Mandatory = $true)]
        [string]$Mode,
        [Parameter(Mandatory = $true)]
        [string]$Exe,
        [Parameter(Mandatory = $true)]
        [string]$ExeDir,
        [switch]$NoWait
    )

    $year = [string]$Group.Year
    $models = @($Group.Models)
    $listPath = Get-YearFileListPath -Year $year
    if (-not (Write-RvtListSafe -Paths $models -ListPath $listPath)) {
        return @{
            Ok      = $false
            Message = ("Не удалось записать список для {0}." -f $Group.Label)
            Code    = 1
            Process = $null
        }
    }

    $rbpArgs = @(
        "--task_script", $TaskScript,
        "--file_list", $listPath
    ) + (Get-CompactModeArgs $Mode) + @(
        "--worksets", "close_all"
    )
    if ($year) {
        $rbpArgs += @("--revit_version", $year)
    }

    if ($NoWait) {
        $argLine = @()
        foreach ($a in $rbpArgs) {
            if ($a -match '\s') { $argLine += ('"{0}"' -f $a) } else { $argLine += $a }
        }
        $proc = Start-Process -FilePath $Exe -ArgumentList ($argLine -join " ") -WorkingDirectory $ExeDir -WindowStyle Hidden -PassThru
        return @{
            Ok      = $true
            Message = ("Запущен Revit {0}." -f $Group.Label)
            Code    = 0
            Process = $proc
        }
    }

    & $Exe @rbpArgs
    $code = $LASTEXITCODE
    if ($code -eq 0) {
        return @{ Ok = $true; Message = ("Revit {0} завершил работу." -f $Group.Label); Code = 0; Process = $null }
    }
    return @{
        Ok      = $false
        Message = ("Revit {0}, код выхода: {1}" -f $Group.Label, $code)
        Code    = $code
        Process = $null
    }
}

function Invoke-CompactLaunch {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet("quick", "deep")]
        [string]$Mode,
        [string[]]$Models,
        [switch]$NoWait
    )

    Save-ActiveModePointer -Mode $Mode
    $env:BATCH_COMPACT_MODE = $Mode
    [void](Save-CompactRunId)

    if (-not (Test-Path -LiteralPath $TaskScript)) {
        return @{ Ok = $false; Message = ("Нет файла задания: {0}" -f $TaskScript); Code = 1 }
    }

    $useModels = @($Models | Where-Object { $_ })
    if ($useModels.Count -eq 0) {
        $useModels = @(Read-ModelList)
    }
    else {
        if (-not (Write-RvtListSafe -Paths $useModels -ListPath $FileList)) {
            return @{ Ok = $false; Message = "Не удалось записать список моделей."; Code = 1 }
        }
        Save-ActiveListPointer -ListPath $FileList
    }

    if ($useModels.Count -eq 0) {
        return @{ Ok = $false; Message = "Список моделей пуст."; Code = 1 }
    }

    $exe = Find-BatchRvt
    if (-not $exe) {
        return @{
            Ok      = $false
            Message = "На этом ПК не найден Revit Batch Processor (BatchRvt.exe)."
            Code    = 1
        }
    }

    if (-not (Ensure-RbpReady -Mode $Mode -Models $useModels -Exe $exe)) {
        return @{
            Ok      = $false
            Message = "Не удалось подготовить Revit Batch Processor для этого режима. Запустите RBP один раз и повторите сжатие."
            Code    = 1
        }
    }

    $exeDir = Split-Path -Parent $exe
    $batchExe = Join-Path $exeDir "BatchRvt.exe"
    if (Test-Path -LiteralPath $batchExe) {
        $exe = $batchExe
        $exeDir = Split-Path -Parent $exe
    }

    $groups = @(Group-ModelsByYear -Models $useModels)
    if ($groups.Count -eq 0) {
        return @{ Ok = $false; Message = "Список моделей пуст."; Code = 1 }
    }

    $labels = @($groups | ForEach-Object { $_.Label })
    $yearNote = if ($groups.Count -eq 1) {
        ("Revit {0}" -f $groups[0].Label)
    }
    else {
        ("годы по очереди: {0}" -f ($labels -join " → "))
    }

    if ($NoWait) {
        $first = $groups[0]
        $rest = @()
        if ($groups.Count -gt 1) {
            $rest = @($groups[1..($groups.Count - 1)])
        }
        $started = Start-CompactYearBatch -Group $first -Mode $Mode -Exe $exe -ExeDir $exeDir -NoWait
        if (-not $started.Ok -or -not $started.Process) {
            return $started
        }
        return @{
            Ok         = $true
            Message    = ("Запущено: {0}." -f $yearNote)
            Code       = 0
            Process    = $started.Process
            YearQueue  = $rest
            YearLabel  = $first.Label
            YearIndex  = 1
            YearTotal  = $groups.Count
            Exe        = $exe
            ExeDir     = $exeDir
        }
    }

    $fail = @()
    $idx = 0
    foreach ($g in $groups) {
        $idx++
        Write-Host ("[{0}/{1}] Revit {2} — {3} модел(ей)" -f $idx, $groups.Count, $g.Label, @($g.Models).Count)
        $one = Start-CompactYearBatch -Group $g -Mode $Mode -Exe $exe -ExeDir $exeDir
        if (-not $one.Ok) { $fail += ("{0}: код {1}" -f $g.Label, $one.Code) }
        if ($idx -lt $groups.Count) { Start-Sleep -Seconds 3 }
    }
    if ($fail.Count -gt 0) {
        return @{
            Ok      = $false
            Message = ("Готово с ошибками ({0}). {1}" -f $yearNote, ($fail -join "; "))
            Code    = 1
        }
    }
    return @{
        Ok      = $true
        Message = ("Revit Batch Processor завершил работу ({0})." -f $yearNote)
        Code    = 0
    }
}

if ($CompactLaunchLibrary -or $script:CompactLaunchLibrary) {
    return
}

Write-Host ""
Write-Host "============================================"
Write-Host "  Compact save  -  Revit Batch Processor"
Write-Host "  This .py does NOT run without RBP / Revit"
Write-Host "============================================"
Write-Host ""
if ($listPaths.IsLocalList) {
    Write-Host "NOTE: rvt_list.txt is local (no write access to share tool folder)."
    Write-Host ("      {0}" -f $FileList)
    Write-Host ""
}
Write-Host "RBP fields (set these if you start the GUI by hand):"
Write-Host ("  Task script : {0}" -f $TaskScript)
Write-Host ("  File list   : {0}" -f $FileList)
Write-Host "  Central     : Create New Local"
Write-Host "  Detach      : OFF"
Write-Host "  Worksets    : Close All"
Write-Host "  Timeout     : 60-90 min"
Write-Host "  Compact     : already ON in the script (no extra checkbox)"
Write-Host ""

if (-not (Test-Path -LiteralPath $TaskScript)) {
    Write-Host ("ERROR: missing script: {0}" -f $TaskScript)
    exit 1
}

$models = @(Read-ModelList)
if ($models.Count -eq 0) {
    Write-Host "ERROR: rvt_list.txt is empty."
    Write-Host "Run 1_choose_models (choose_models_path.cmd) first."
    exit 1
}

$exe = Find-BatchRvt
if (-not $exe) {
    Write-Host "ERROR: BatchRvt.exe not found on this PC."
    Write-Host "Install Revit Batch Processor, then either:"
    Write-Host "  - keep RBP open and run this again, or"
    Write-Host "  - open RBP GUI manually:"
    Write-Host ("      Task script = {0}" -f $TaskScript)
    Write-Host ("      File list   = {0}" -f $FileList)
    Write-Host "      Create New Local, Detach = OFF"
    exit 1
}

$yearGuesses = @($models | ForEach-Object { Guess-Year $_ })
$years = @($yearGuesses | Where-Object { $_ } | Sort-Object -Unique)
$unknownYearCount = @($yearGuesses | Where-Object { -not $_ }).Count
$hasRsn = @($models | Where-Object { $_ -match '^(?i)RSN://' }).Count -gt 0

Write-Host ("Models: {0}" -f $models.Count)
Write-Host ("List:   {0}" -f $FileList)
Write-Host ("Script: {0}" -f $TaskScript)
Write-Host ("RBP:    {0}" -f $exe)
if ($years.Count -gt 0) {
    Write-Host ("Years:  {0}" -f ($years -join ", "))
}
Write-Host ""
Write-Host "This rewrites models (central Compact for shared/RSN)."
Write-Host "Do this after backup, preferably when nobody has the files open."
Write-Host ""
Write-Host "Y = start RBP now with these fields"
Write-Host "Enter = cancel (open Revit Batch Processor GUI yourself)"
$ok = Read-Host "Choice"
if ($ok -notmatch '^(?i)y(es)?$') {
    Write-Host "Cancelled."
    exit 0
}

if ($hasRsn) {
    $patchPs = Join-Path $ToolDir "rbp_rsn_patch\apply_rsn_support.ps1"
    if (-not (Test-Path -LiteralPath $patchPs)) {
        $toolsRoot = Split-Path -Parent $ToolDir
        $patchPs = Join-Path $toolsRoot "batch_set_project_units\rbp_rsn_patch\apply_rsn_support.ps1"
    }
    if (Test-Path -LiteralPath $patchPs) {
        Write-Host ""
        Write-Host "List has RSN:// - applying RBP RSN patch (File exists check)..."
        $scriptsDir = Join-Path (Split-Path -Parent $exe) "Scripts"
        if (Test-Path -LiteralPath (Join-Path $scriptsDir "revit_file_list.py")) {
            & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $patchPs -ScriptsPath $scriptsDir
        }
        else {
            & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $patchPs
        }
        if ($LASTEXITCODE -ne 0) {
            Write-Host ""
            Write-Host "ERROR: RSN patch failed. Without it BatchRvt reports File exists: NO for RSN://."
            Write-Host "       Keep RBP open and run: rbp_rsn_patch\установить_поддержку_RSN_в_RBP.cmd"
            exit 1
        }
        Write-Host "RBP RSN patch: OK"
    }
    else {
        Write-Host "WARN: rbp_rsn_patch missing - RSN:// may fail with File exists: NO"
    }
}

$rbpArgs = @(
    "--task_script", $TaskScript,
    "--file_list", $FileList,
    "--create_new_local",
    "--worksets", "close_all"
)

# Lock Revit year only when EVERY file is the same year.
# Mixed list or a file without _Rxx in the name: let RBP detect per file.
if (($years.Count -eq 1) -and ($unknownYearCount -eq 0)) {
    $rbpArgs += @("--revit_version", $years[0])
    Write-Host ("Revit version: {0}" -f $years[0])
}
elseif ($years.Count -gt 1 -or $unknownYearCount -gt 0) {
    Write-Host "Mixed / unknown years - do NOT set Revit version in the GUI."
    Write-Host "RBP will open each file in its own Revit year."
}
else {
    Write-Host "Could not detect Revit year from the list."
    Write-Host "Leave Revit version empty in the GUI (auto)."
}

if ($hasRsn) {
    Write-Host "List has RSN:// - Create New Local is already set."
    Write-Host "  (RSN patch is applied automatically before BatchRvt starts)"
}

Write-Host ""
Write-Host "Starting BatchRvt..."
Write-Host ""

$exeDir = Split-Path -Parent $exe
$batchExe = Join-Path $exeDir "BatchRvt.exe"
if (Test-Path -LiteralPath $batchExe) {
    $exe = $batchExe
}

& $exe @rbpArgs
$code = $LASTEXITCODE
Write-Host ""
if ($code -eq 0) {
    Write-Host "BatchRvt finished (exit 0). Check log for OK / ERROR lines."
    Write-Host "Reports: share ...\batch_compact_save\отчеты  (or local %LOCALAPPDATA%\BatchRvt\batch_compact_save\отчеты if no share write)"
    Write-Host "Open .xlsx there (dated name). Path also in BatchRvt log: Report xlsx: ..."
    Write-Host "CSV backup: отчеты\csv\"
}
else {
    Write-Host ("BatchRvt exit code: {0}" -f $code)
}
exit $code
