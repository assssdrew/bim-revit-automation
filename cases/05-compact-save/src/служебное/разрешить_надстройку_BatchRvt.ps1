# Name: разрешить_надстройку_BatchRvt.ps1
# Version: 1.0
# What it does: Add CodeSigning trust DWORDs for BatchRvt across Revit add-in folders.
# Inputs: Optional -Year; when empty, every year with a BatchRvt .addin.
# Outputs: HKCU CodeSigning registry values per add-in GUID.
# How to run: разрешить_надстройку_BatchRvt.cmd or powershell -File разрешить_надстройку_BatchRvt.ps1
# Notes: Trusts unsigned BatchRvt for the current Windows user.

param(
    [string]$Year = ""
)

function Trust-BatchRvtAddin([string]$RevitYear) {
    $addinsRoot = Join-Path $env:APPDATA ("Autodesk\Revit\Addins\" + $RevitYear)
    if (-not (Test-Path -LiteralPath $addinsRoot)) {
        return @{
            Status = "skip"
            Year   = $RevitYear
            Detail = "no Addins folder (RBP add-in not installed for this year)"
        }
    }

    $files = @(
        Get-ChildItem -LiteralPath $addinsRoot -Recurse -Filter *.addin -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -match 'BatchRvt|BatchRvtAddin' }
    )
    if ($files.Count -eq 0) {
        return @{
            Status = "skip"
            Year   = $RevitYear
            Detail = "BatchRvt .addin not found"
        }
    }

    $guid = $null
    $addinFile = $null
    foreach ($f in $files) {
        $txt = Get-Content -LiteralPath $f.FullName -Raw
        if ($txt -match '<AddInId>\s*\{?([0-9A-Fa-f\-]{36})\}?\s*</AddInId>') {
            $guid = $Matches[1].ToUpper()
            $addinFile = $f.FullName
            break
        }
    }
    if (-not $guid) {
        return @{
            Status = "fail"
            Year   = $RevitYear
            Detail = "AddInId missing in .addin"
        }
    }

    $key = "HKCU:\SOFTWARE\Autodesk\Revit\Autodesk Revit $RevitYear\CodeSigning"
    if (-not (Test-Path $key)) {
        New-Item -Path $key -Force | Out-Null
    }
    New-ItemProperty -Path $key -Name $guid -PropertyType DWord -Value 1 -Force | Out-Null

    return @{
        Status = "ok"
        Year   = $RevitYear
        Detail = ("{0}  DWORD {1} = 1" -f $key, $guid)
        File   = $addinFile
    }
}

if ($Year -match '^20(1[8-9]|2[0-9])$') {
    $targets = @($Year.Trim())
}
else {
    $targets = @(2018..2026 | ForEach-Object { "$_" })
}

Write-Host "Trust BatchRvt add-in for this Windows user."
Write-Host ("Years: {0}" -f ($targets -join ", "))
Write-Host ""

$okCount = 0
$skipCount = 0
$failCount = 0
foreach ($y in $targets) {
    $r = Trust-BatchRvtAddin $y
    if ($r.Status -eq "ok") {
        $okCount++
        Write-Host ("OK   Revit {0}" -f $r.Year)
        Write-Host ("     {0}" -f $r.Detail)
        if ($r.File) { Write-Host ("     {0}" -f $r.File) }
    }
    elseif ($r.Status -eq "skip") {
        $skipCount++
        Write-Host ("skip Revit {0}: {1}" -f $r.Year, $r.Detail)
    }
    else {
        $failCount++
        Write-Host ("FAIL Revit {0}: {1}" -f $r.Year, $r.Detail)
    }
}

Write-Host ""
Write-Host ("Done. Trusted: {0}  skipped: {1}  failed: {2}" -f $okCount, $skipCount, $failCount)
Write-Host "Close Revit completely, then start RBP."
if ($okCount -eq 0) { exit 1 }
exit 0
