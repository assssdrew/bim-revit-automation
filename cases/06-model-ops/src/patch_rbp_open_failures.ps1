# Same installer as tools/rbp_patches/apply_open_fail.ps1
# Kept here so a standalone copy of this toolkit still patches RBP.
param(
    [string]$ScriptsPath = ""
)

$ErrorActionPreference = "Stop"
$script:Marker = "OPEN_FAIL_PATCH_v2"
$here = $PSScriptRoot
$parent = Split-Path -Parent $here

function Write-Status([string]$Text) { Write-Host $Text }
function Write-Fail([string]$Text) {
    [Console]::Error.WriteLine($Text)
    Write-Host $Text
}

function Find-OpenFailSource {
    $candidates = @(
        (Join-Path $here "rbp_open_fail_patch"),
        (Join-Path $parent "rbp_patches\open_fail"),
        (Join-Path $here "open_fail")
    )
    foreach ($dir in $candidates) {
        $a = Join-Path $dir "revit_failure_handling.py"
        $b = Join-Path $dir "revit_dialog_util.py"
        if ((Test-Path -LiteralPath $a -PathType Leaf) -and (Test-Path -LiteralPath $b -PathType Leaf)) {
            return $dir
        }
    }
    return $null
}

function Test-RbpScripts([string]$Path) {
    if ([string]::IsNullOrWhiteSpace($Path)) { return $false }
    return (
        (Test-Path -LiteralPath (Join-Path $Path "revit_failure_handling.py") -PathType Leaf) -and
        (Test-Path -LiteralPath (Join-Path $Path "revit_dialog_util.py") -PathType Leaf)
    )
}

function Find-RbpScripts {
    $candidates = New-Object System.Collections.Generic.List[string]
    Get-Process -ErrorAction SilentlyContinue |
        Where-Object { $_.ProcessName -match '^(?i)BatchRvt(GUI)?$|RevitBatchProcessor$' } |
        ForEach-Object {
            try {
                $exeDir = Split-Path -Parent $_.MainModule.FileName
                if ($exeDir) { $candidates.Add((Join-Path $exeDir "Scripts")) }
            }
            catch { }
        }
    foreach ($root in @($env:LOCALAPPDATA, $env:ProgramFiles, ${env:ProgramFiles(x86)})) {
        if ([string]::IsNullOrWhiteSpace($root)) { continue }
        foreach ($relative in @(
            "RevitBatchProcessor\Scripts",
            "BatchRvt\Scripts"
        )) {
            $candidates.Add((Join-Path $root $relative))
        }
    }
    $seen = @{}
    foreach ($candidate in $candidates) {
        if ([string]::IsNullOrWhiteSpace($candidate)) { continue }
        try { $full = [System.IO.Path]::GetFullPath($candidate) } catch { continue }
        if ($seen.ContainsKey($full)) { continue }
        $seen[$full] = $true
        if (Test-RbpScripts $full) { return $full }
    }
    return $null
}

function Install-PatchedScript {
    param(
        [string]$SrcDir,
        [string]$DestDir,
        [string]$FileName
    )
    $src = Join-Path $SrcDir $FileName
    $dst = Join-Path $DestDir $FileName
    $backup = $dst + ".bak_before_open_fail"
    if (-not (Test-Path -LiteralPath $backup)) {
        Copy-Item -LiteralPath $dst -Destination $backup -Force
        Write-Status ("Backup: " + $backup)
    }
    Copy-Item -LiteralPath $src -Destination $dst -Force
    $check = [System.IO.File]::ReadAllText($dst)
    if (-not $check.Contains($script:Marker)) {
        throw ("Installed file does not contain " + $script:Marker + ": " + $dst)
    }
    Write-Status ("Installed: " + $FileName)
}

try {
    $srcDir = Find-OpenFailSource
    if (-not $srcDir) {
        throw "Patch source missing (open_fail / rbp_open_fail_patch)."
    }
    if ([string]::IsNullOrWhiteSpace($ScriptsPath)) {
        $ScriptsPath = Find-RbpScripts
    }
    if (-not (Test-RbpScripts $ScriptsPath)) {
        Write-Fail "BatchRvt Scripts folder not found."
        Write-Fail "Run on the PC that runs Revit Batch Processor."
        Write-Fail "Typical path: %LOCALAPPDATA%\RevitBatchProcessor\Scripts"
        exit 2
    }
    Write-Status ("RBP Scripts: " + $ScriptsPath)
    Write-Status ("Patch source: " + $srcDir)
    Install-PatchedScript $srcDir $ScriptsPath "revit_failure_handling.py"
    Install-PatchedScript $srcDir $ScriptsPath "revit_dialog_util.py"
    Write-Status ("RBP open-failure patch: OK (" + $script:Marker + ")")
    exit 0
}
catch {
    Write-Fail ("Open-failure patch FAILED: " + $_.Exception.Message)
    if ($ScriptsPath) { Write-Fail ("ScriptsPath was: " + $ScriptsPath) }
    exit 1
}
