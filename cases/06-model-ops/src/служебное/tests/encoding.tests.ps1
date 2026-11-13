# Name: encoding.tests.ps1
# Version: 3.1.0
# What it does: See the case README.
# Notes: Paths and project names in this portfolio copy are placeholders.

# Every .ps1 of the tool must be UTF-8 WITH BOM (Windows PowerShell 5.1 reads a
# BOM-less file as ANSI and breaks Russian text) and must parse.
# Under 5.1 this also checks the files against the 5.1 parser.
#   powershell -NoProfile -ExecutionPolicy Bypass -File tests\encoding.tests.ps1

$ToolSrc = Split-Path -Parent $PSScriptRoot
$script:Fail = 0
$script:Pass = 0
function Check([bool]$Cond, [string]$Msg) {
    if ($Cond) { $script:Pass++ }
    else { $script:Fail++; Write-Host "  FAIL: $Msg" -ForegroundColor Red }
}

$files = @(Get-ChildItem -LiteralPath $ToolSrc -Recurse -File -Filter *.ps1)
Check ($files.Count -gt 5) "found the tool's .ps1 files ($($files.Count))"
foreach ($f in $files) {
    $rel = $f.FullName.Substring($ToolSrc.Length).TrimStart("\", "/")
    $b = [System.IO.File]::ReadAllBytes($f.FullName)
    $bom = ($b.Length -ge 3 -and $b[0] -eq 0xEF -and $b[1] -eq 0xBB -and $b[2] -eq 0xBF)
    Check $bom "UTF-8 BOM present: $rel"
    $errors = $null
    [void][System.Management.Automation.Language.Parser]::ParseFile($f.FullName, [ref]$null, [ref]$errors)
    Check (@($errors).Count -eq 0) ("parses without errors: {0} {1}" -f $rel, (@($errors | ForEach-Object { $_.Message }) -join "; "))
}

Write-Host ""
Write-Host ("Checks passed: {0}, failed: {1}" -f $script:Pass, $script:Fail)
if ($script:Fail -gt 0) { exit 1 }
exit 0
