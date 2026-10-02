# Collects target .rvt under discipline folders for Revit Batch Processor.
# Finds folders named RVT (any depth), takes any .rvt in that folder
# (ExampleProject names like PROJ_ANNEX_K01.rvt — no *_R## required).
# Skips Backup / _backup / Revit_temp / Reserve / Families / .0001.rvt
#
# Usage:
#   powershell -ExecutionPolicy Bypass -File make_rvt_list.ps1 -Root "\\share\...\Models" -Out ".\rvt_list.txt"
#   powershell -ExecutionPolicy Bypass -File make_rvt_list.ps1 -Roots @("...\5_1_EOM","...\5_2_VK") -Out ".\rvt_list.txt"

param(
    [Parameter(Mandatory = $false)]
    [string]$Root,

    [Parameter(Mandatory = $false)]
    [string[]]$Roots,

    [string]$Out = ".\rvt_list.txt",

    [switch]$SkipBackups = $true
)

function Get-NativePath([string]$Path) {
    if ([string]::IsNullOrWhiteSpace($Path)) { return $null }
    $p = $Path.Trim().Trim('"')
    $marker = "FileSystem::"
    $idx = $p.IndexOf($marker, [StringComparison]::OrdinalIgnoreCase)
    if ($idx -ge 0) {
        $p = $p.Substring($idx + $marker.Length)
    }
    if ($p.StartsWith("\\")) { return $p }
    try {
        return (New-Object System.IO.FileInfo($p)).FullName
    }
    catch {
        return $p
    }
}

function Get-U {
    param([int[]]$Codes)
    return (-join ($Codes | ForEach-Object { [char]([int]$_) }))
}

function Test-ExcludedModelPath([string]$Path) {
    if ([string]::IsNullOrWhiteSpace($Path)) { return $true }
    $p = $Path.Replace("/", "\")
    if ($p -match '(?i)\\Backup\\') { return $true }
    if ($p -match '(?i)\\Revit_temp\\') { return $true }
    if ($p -match '(?i)_backup(\\|$)') { return $true }
    $rezerv = Get-U @(0x0420, 0x0435, 0x0437, 0x0435, 0x0440, 0x0432)
    $sem = Get-U @(0x0421, 0x0435, 0x043C, 0x0435, 0x0439, 0x0441, 0x0442, 0x0432, 0x0430)
    if ($p -match ("(?i)\\" + [regex]::Escape($rezerv) + "\\")) { return $true }
    if ($p -match ("(?i)\\" + [regex]::Escape($sem) + "\\")) { return $true }
    return $false
}

function Test-IsTargetRvtName([string]$Name) {
    if ([string]::IsNullOrWhiteSpace($Name)) { return $false }
    if ($Name -match '\.\d{4}\.rvt$') { return $false }
    if ($Name -match '(?i)_backup') { return $false }
    return ($Name -match '(?i)\.rvt$')
}

$scanRoots = @()
if ($Roots -and $Roots.Count -gt 0) {
    $scanRoots = @($Roots | ForEach-Object { Get-NativePath $_ } | Where-Object { $_ })
}
elseif (-not [string]::IsNullOrWhiteSpace($Root)) {
    $scanRoots = @(Get-NativePath $Root)
}
else {
    Write-Error "Provide -Root or -Roots"
    exit 1
}

$found = New-Object System.Collections.Generic.List[string]
foreach ($folder in $scanRoots) {
    if (-not (Test-Path -LiteralPath $folder)) {
        Write-Warning "Root not found: $folder"
        continue
    }
    $rvtDirs = @(
        Get-ChildItem -LiteralPath $folder -Recurse -Directory -Force -ErrorAction SilentlyContinue |
            Where-Object {
                ($_.Name -ieq "RVT") -and
                ((-not $SkipBackups) -or (-not (Test-ExcludedModelPath $_.FullName)))
            }
    )
    foreach ($rvtDir in $rvtDirs) {
        Get-ChildItem -LiteralPath $rvtDir.FullName -File -Force -ErrorAction SilentlyContinue |
            Where-Object {
                ($_.Extension -ieq ".rvt") -and
                (Test-IsTargetRvtName $_.Name)
            } |
            ForEach-Object {
                $native = Get-NativePath $_.FullName
                if ($native) { [void]$found.Add($native) }
            }
    }
}

$files = @($found | Sort-Object -Unique)
$files | Set-Content -LiteralPath $Out -Encoding UTF8

Write-Host "Found $($files.Count) files"
Write-Host "List saved: $Out"
if ($files.Count -gt 0) {
    Write-Host "Example path: $($files[0])"
}
