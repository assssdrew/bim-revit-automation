# Name: psquirks.tests.ps1
# Version: 3.1.0
# What it does: See the case README.
# Notes: Paths and project names in this portfolio copy are placeholders.

# Static guard against a PowerShell trap that made the "Подробности" dialog fail silently:
#   $l = New-Object System.Collections.Generic.List[object]
#   @($l)          # throws "Argument types do not match" (Windows PowerShell 5.1 and PowerShell 7)
# Use $l.ToArray() (or a plain array) instead. The check parses every .ps1 of the tool.
#   powershell -NoProfile -ExecutionPolicy Bypass -File tests\psquirks.tests.ps1

$ToolSrc = Split-Path -Parent $PSScriptRoot
$script:Fail = 0
$script:Pass = 0
function Check([bool]$Cond, [string]$Msg) {
    if ($Cond) { $script:Pass++ }
    else { $script:Fail++; Write-Host "  FAIL: $Msg" -ForegroundColor Red }
}

# name of the function a node lives in ("" = script level)
function Get-Owner($Node) {
    $p = $Node.Parent
    while ($p) {
        if ($p -is [System.Management.Automation.Language.FunctionDefinitionAst]) { return $p.Name }
        $p = $p.Parent
    }
    return ""
}

# findings for one file: @(...) applied to a variable that is a New-Object List[object] in the same
# function ($script:/$global: variables count for the whole file)
function Find-ListArrayWrap([string]$Path) {
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$null, [ref]$null)
    $lists = @{}
    $assigns = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.AssignmentStatementAst] }, $true)
    foreach ($a in $assigns) {
        $t = $a.Right.Extent.Text
        if ($t -match '^New-Object\s+(-TypeName\s+)?\[?System\.Collections\.Generic\.List\[object\]' -and $a.Left -is [System.Management.Automation.Language.VariableExpressionAst]) {
            $vp = $a.Left.VariablePath
            $scope = $(if ($vp.IsScript -or $vp.IsGlobal) { "*" } else { Get-Owner $a })
            $lists[($scope + "|" + $vp.UserPath.ToLowerInvariant().Replace("script:", "").Replace("global:", ""))] = $true
        }
    }
    $found = @()
    $wraps = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.ArrayExpressionAst] }, $true)
    foreach ($w in $wraps) {
        $inner = $w.SubExpression.Statements
        if ($inner.Count -ne 1) { continue }
        $e = $inner[0].Extent.Text.Trim()
        if ($e -match '^\$(script:|global:)?([A-Za-z_][A-Za-z0-9_]*)$') {
            $name = $Matches[2].ToLowerInvariant()
            $scoped = [bool]$Matches[1]
            $key = $(if ($scoped) { "*|" + $name } else { (Get-Owner $w) + "|" + $name })
            if ($lists.ContainsKey($key)) { $found += ("{0}:{1}  @({2})" -f (Split-Path -Leaf $Path), $w.Extent.StartLineNumber, $e) }
        }
    }
    return $found
}

$files = @(Get-ChildItem -LiteralPath $ToolSrc -Recurse -File -Filter *.ps1 | Where-Object { $_.FullName -notmatch '[\\/]tests[\\/]' })
Check ($files.Count -gt 5) "found the tool's .ps1 files ($($files.Count))"
foreach ($f in $files) {
    $hits = @(Find-ListArrayWrap $f.FullName)
    Check ($hits.Count -eq 0) ("no @(...) over a New-Object List[object] in {0}: {1}" -f $f.Name, ($hits -join "; "))
}

# the detector itself: it must see the trap in a sample
$tmp = Join-Path ([System.IO.Path]::GetTempPath()) ("quirk_" + [Guid]::NewGuid().ToString("N").Substring(0, 8) + ".ps1")
$lines = @(
    '$l = New-Object System.Collections.Generic.List[object]; $script:d = New-Object System.Collections.Generic.List[object]',
    '$a = @($l); $b = @($script:d)',
    '$s = New-Object System.Collections.Generic.List[string]; $c = @($s); $x = @($l | Sort-Object)',
    'function Other { param($l) return @($l) }',                        # same name, other scope: fine
    'function Mine { $m = New-Object System.Collections.Generic.List[object]; return @($m) }')
Set-Content -LiteralPath $tmp -Value $lines -Encoding UTF8
$sample = @(Find-ListArrayWrap $tmp)
Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
Check ($sample.Count -eq 3) "detector: finds the 3 traps, ignores List[string], pipelines and a same-named variable of another function (found $($sample.Count): $($sample -join '; '))"

Write-Host ""
Write-Host ("Checks passed: {0}, failed: {1}" -f $script:Pass, $script:Fail)
if ($script:Fail -gt 0) { exit 1 }
exit 0
