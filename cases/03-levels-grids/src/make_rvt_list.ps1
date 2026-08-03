# Collects all .rvt paths under a root folder into a text list for Revit Batch Processor.
param(
    [Parameter(Mandatory = $true)]
    [string]$Root,

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
    try {
        return (New-Object System.IO.FileInfo($p)).FullName
    }
    catch {
        return $p
    }
}

$Root = Get-NativePath $Root

if (-not (Test-Path -LiteralPath $Root)) {
    Write-Error "Root not found: $Root"
    exit 1
}

$files = Get-ChildItem -LiteralPath $Root -Recurse -Filter *.rvt -File |
    Where-Object {
        if (-not $SkipBackups) { return $true }
        $_.Name -notmatch '\.\d{4}\.rvt$'
    } |
    ForEach-Object { Get-NativePath $_.FullName } |
    Sort-Object -Unique

$files | Set-Content -LiteralPath $Out -Encoding UTF8

Write-Host "Found $($files.Count) files"
Write-Host "List saved: $Out"
if ($files.Count -gt 0) {
    Write-Host "Example path: $($files[0])"
}
exit 0
