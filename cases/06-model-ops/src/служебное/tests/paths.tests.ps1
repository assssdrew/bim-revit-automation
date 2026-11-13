# Name: paths.tests.ps1
# Version: 3.1.0
# What it does: See the case README.
# Notes: Paths and project names in this portfolio copy are placeholders.

# Every local script/file a launcher or script of the tool refers to must exist.
# Layout: the tool root keeps only the launcher, docs, servers.cfg and the user's
# data; all code (and tests\) is in служебное\.
# .vbs:  finds служебное\presets.ps1 by scanning its subfolders (no Cyrillic in .vbs)
# .cmd:  "%~dp0x.cmd", "%~dp0..\x.ps1"  -> relative to the .cmd folder
# .ps1:  Join-Path $CodeDir|$PSScriptRoot|$here "x.ps1" -> the script's folder,
#        Join-Path $ToolDir "x"         -> the tool root (parent of служебное),
#        Join-Path $toolsRoot "x\y.ps1" -> tools\ ; a list of candidates @(...)
#        passes when at least one of them exists (fallback chains).
#        Reassigning the same variable in one function is a fallback chain too.
# Only *.ps1, *.cmd, *.py, *.vbs names are checked.
#   powershell -NoProfile -ExecutionPolicy Bypass -File tests\paths.tests.ps1

$CodeSrc = Split-Path -Parent $PSScriptRoot        # служебное
$ToolSrc = Split-Path -Parent $CodeSrc             # tool root
$ToolsRoot = Split-Path -Parent $ToolSrc           # tools
$CodeName = Split-Path -Leaf $CodeSrc
$script:Fail = 0
$script:Pass = 0
function Check([bool]$Cond, [string]$Msg) {
    if ($Cond) { $script:Pass++ }
    else { $script:Fail++; Write-Host "  FAIL: $Msg" -ForegroundColor Red }
}
function Rel([string]$Path) { return $Path.Substring($ToolSrc.Length).TrimStart("\", "/") }
function Resolve-Ref([string]$Base, [string]$Name) {
    $p = $Base
    foreach ($part in ($Name -split '[\\/]')) {
        if ($part -eq "" -or $part -eq ".") { continue }
        if ($part -eq "..") { $p = Split-Path -Parent $p } else { $p = Join-Path $p $part }
    }
    return $p
}
$ExtRx = '\.(ps1|cmd|py|vbs)$'

$all = @(Get-ChildItem -LiteralPath $ToolSrc -Recurse -File | Where-Object { $_.FullName -notmatch '[\\/]tests[\\/]' })
$vbs = @($all | Where-Object { $_.Extension -ieq ".vbs" })
$cmds = @($all | Where-Object { $_.Extension -ieq ".cmd" })
$ps1s = @($all | Where-Object { $_.Extension -ieq ".ps1" })

# layout: the root holds only the allowed tracked items, code is in служебное
Check ($CodeName -eq "служебное") "tests live in служебное\tests (code folder: $CodeName)"
$rootVbs = @($vbs | Where-Object { $_.DirectoryName -eq $ToolSrc })
Check ($rootVbs.Count -eq 1) ("one .vbs launcher in the root (found {0})" -f $rootVbs.Count)
Check (@($cmds | Where-Object { $_.DirectoryName -eq $ToolSrc }).Count -eq 0) "no .cmd in the root (they live in служебное)"
$Allowed = @("Операции с моделями.vbs", "ИНСТРУКЦИЯ.txt", "README.md", "servers.cfg", ".gitignore", "служебное")
# user data the window writes next to the launcher (ignored by git) may also be there
$DataNames = @("rvt_list.txt", "rvt_list_new.txt", "rvt_list_staging.txt", "rvt_list_move.txt", "rvt_list_retry.txt",
    "mapping.csv", "job_paths.csv", "job.json", "rbp_checklist.txt", "to_delete_rsn.txt", "relink_map.csv",
    "models_path.cfg", "last_selection.cfg", "reports", "__pycache__",
    "section_codes.csv", "name_mask.cfg", "workset_plan.csv", "workset_job.json", "rvt_list_worksets.txt")
foreach ($it in @(Get-ChildItem -LiteralPath $ToolSrc -Force)) {
    $ok = ($Allowed -contains $it.Name) -or ($DataNames -contains $it.Name) -or ($it.Name -like "*.bak_*")
    Check $ok ("root holds only the launcher, docs, servers.cfg, служебное and user data: unexpected '{0}'" -f $it.Name)
    if (-not $it.PSIsContainer) {
        Check ($it.Extension -notmatch '^\.(ps1|py|cmd|bat)$') ("no code in the root: {0}" -f $it.Name)
    }
}
foreach ($n in $Allowed) { Check (Test-Path -LiteralPath (Join-Path $ToolSrc $n)) "root has $n" }
$git = Get-Command git -ErrorAction SilentlyContinue
if ($git) {
    # tracked files only (a user's local data must not count); UTF-8 so Cyrillic names survive 5.1
    $oldEnc = [Console]::OutputEncoding
    try {
        [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
        $tracked = @(& git -C $ToolSrc -c core.quotepath=off ls-files -- . 2>$null)
    }
    finally { [Console]::OutputEncoding = $oldEnc }
    if ($LASTEXITCODE -eq 0 -and $tracked.Count -gt 0) {
        $top = @($tracked | ForEach-Object { ($_ -split '/')[0] } | Sort-Object -Unique)
        $extra = @($top | Where-Object { $Allowed -notcontains $_ })
        Check ($extra.Count -eq 0) ("tracked root items are only the allowed ones (extra: {0})" -f ($extra -join ", "))
    }
}
foreach ($n in @("presets.ps1", "job_lib.ps1", "ui_logic.ps1", "start_ops.ps1", "choose_models.ps1", "safety_lib.ps1",
        "delete_old_models.ps1", "move_centrals.ps1", "update_rvt_links.py", "saveas_central.py", "detach_to_buffer.py", "ops_lib.py", "relink_prefix.py",
        "read_worksets.py", "workset_plan.py", "apply_workset_links.py", "workset_logic.ps1", "workset_ui.ps1")) {
    Check (Test-Path -LiteralPath (Join-Path $CodeSrc $n) -PathType Leaf) "code is in служебное: $n"
}
Check ($cmds.Count -ge 7) ("found the .cmd files ({0})" -f $cmds.Count)
Check ($ps1s.Count -gt 5) ("found the .ps1 files ({0})" -f $ps1s.Count)

$refs = 0
foreach ($f in $vbs) {
    $text = [System.IO.File]::ReadAllText($f.FullName)
    Check ($text -notmatch '[^\x00-\x7F]') ("{0}: ASCII only (WSH reads .vbs as ANSI)" -f (Rel $f.FullName))
    # emulate the launcher: first subfolder with presets.ps1, else next to the .vbs
    Check ($text -match 'For Each \w+ In fso\.GetFolder\(tool\)\.SubFolders') ("{0}: scans subfolders for presets.ps1" -f (Rel $f.FullName))
    $found = @(Get-ChildItem -LiteralPath $f.DirectoryName -Directory | Where-Object { Test-Path -LiteralPath (Join-Path $_.FullName "presets.ps1") -PathType Leaf })
    $refs++
    Check ($found.Count -eq 1 -and $found[0].Name -eq "служебное") ("{0}: launches служебное\presets.ps1 (found: {1})" -f (Rel $f.FullName), (@($found | ForEach-Object { $_.Name }) -join ", "))
}

foreach ($f in $cmds) {
    $text = [System.IO.File]::ReadAllText($f.FullName)
    foreach ($m in [regex]::Matches($text, '%~dp0([^"%\s]+)')) {
        $n = $m.Groups[1].Value
        if ($n -notmatch $ExtRx) { continue }
        $refs++
        $p = Resolve-Ref $f.DirectoryName $n
        Check (Test-Path -LiteralPath $p -PathType Leaf) ("{0}: %~dp0{1} -> {2}" -f (Rel $f.FullName), $n, $p)
    }
    foreach ($m in [regex]::Matches($text, '(?im)^\s*pushd\s+"%~dp0([^"]*)"')) {
        $p = Resolve-Ref $f.DirectoryName $m.Groups[1].Value
        # data (rvt_list.txt, rbp_checklist.txt) is read/written relative to the tool root
        Check ($p -eq $ToolSrc -and (Test-Path -LiteralPath (Join-Path $p "servers.cfg") -PathType Leaf)) ("{0}: pushd goes to the tool root ({1})" -f (Rel $f.FullName), $p)
    }
}

$SelfVars = @("codedir", "psscriptroot", "here")
foreach ($f in $ps1s) {
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($f.FullName, [ref]$null, [ref]$null)
    $joins = $ast.FindAll({ param($n)
            $n -is [System.Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq "Join-Path" }, $true)
    $groups = @{}
    foreach ($j in $joins) {
        $el = @($j.CommandElements)
        if ($el.Count -ne 3) { continue }
        if (-not ($el[1] -is [System.Management.Automation.Language.VariableExpressionAst])) { continue }
        if (-not ($el[2] -is [System.Management.Automation.Language.StringConstantExpressionAst])) { continue }
        $n = $el[2].Value
        if ($n -notmatch $ExtRx) { continue }
        $v = $el[1].VariablePath.UserPath.ToLowerInvariant()
        if ($SelfVars -contains $v) { $base = $f.DirectoryName }
        elseif ($v -eq "tooldir") { $base = $ToolSrc }
        elseif ($v -eq "toolsroot") { $base = $ToolsRoot }
        else { continue }
        $p = Resolve-Ref $base $n
        $owner = $j.Parent
        while ($owner -and -not ($owner -is [System.Management.Automation.Language.ArrayExpressionAst]) -and -not ($owner -is [System.Management.Automation.Language.ScriptBlockAst])) { $owner = $owner.Parent }
        if ($owner -is [System.Management.Automation.Language.ArrayExpressionAst]) { $key = "list@" + $owner.Extent.StartOffset }
        else {
            # $x = Join-Path ...; if (-not (Test-Path $x)) { $x = Join-Path ... }  -> one fallback chain
            $a = $j.Parent
            while ($a -and -not ($a -is [System.Management.Automation.Language.AssignmentStatementAst]) -and -not ($a -is [System.Management.Automation.Language.StatementBlockAst])) { $a = $a.Parent }
            $fn = $j.Parent
            while ($fn -and -not ($fn -is [System.Management.Automation.Language.FunctionDefinitionAst])) { $fn = $fn.Parent }
            $fnKey = $(if ($fn) { $fn.Extent.StartOffset } else { 0 })
            if ($a -is [System.Management.Automation.Language.AssignmentStatementAst]) { $key = "var@" + $fnKey + "|" + $a.Left.Extent.Text.ToLowerInvariant() }
            else { $key = "one@" + $j.Extent.StartOffset }
        }
        if (-not $groups.ContainsKey($key)) { $groups[$key] = New-Object System.Collections.ArrayList }
        [void]$groups[$key].Add([pscustomobject]@{ Name = $n; Path = $p; Line = $j.Extent.StartLineNumber })
    }
    foreach ($k in $groups.Keys) {
        $g = @($groups[$k].ToArray())
        $refs += $g.Count
        $ok = @($g | Where-Object { Test-Path -LiteralPath $_.Path -PathType Leaf }).Count -gt 0
        Check $ok ("{0}:{1}: {2}" -f (Rel $f.FullName), $g[0].Line, (@($g | ForEach-Object { $_.Name + " -> " + $_.Path }) -join " | "))
    }
}
Check ($refs -gt 20) "found the references ($refs)"

# job_lib: RBP / PowerShell passes run code from служебное but read the lists from the root
$jl = [System.IO.File]::ReadAllText((Join-Path $CodeSrc "job_lib.ps1"))
$tasks = [regex]::Matches($jl, 'Task\s*=\s*\(Join-Path \$(\w+) "')
$lists = [regex]::Matches($jl, 'List\s*=\s*\(Join-Path \$(\w+) "')
Check ($tasks.Count -ge 6) "job_lib: found the pass Task paths ($($tasks.Count))"
Check (@($tasks | Where-Object { $_.Groups[1].Value -ne "CodeDir" }).Count -eq 0) "job_lib: every pass Task is under `$CodeDir (служебное)"
Check (@($lists | Where-Object { $_.Groups[1].Value -ne "ToolDir" }).Count -eq 0) "job_lib: every pass List is under `$ToolDir (tool root)"
foreach ($n in @("delete_old_models.ps1", "move_centrals.ps1")) {
    $t = [System.IO.File]::ReadAllText((Join-Path $CodeSrc $n))
    Check ($t.Contains('$ToolDir = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)')) "$n defaults -ToolDir to the tool root"
}
$so = [System.IO.File]::ReadAllText((Join-Path $CodeSrc "start_ops.ps1"))
Check (([regex]::Matches($so, '-ToolDir `"\$ToolDir`"')).Count -ge 2) "start_ops passes -ToolDir to move/delete (dry-run and run)"
foreach ($n in @("presets.ps1", "choose_models.ps1")) {
    $t = [System.IO.File]::ReadAllText((Join-Path $CodeSrc $n))
    Check ($t.Contains('$ToolDir = Split-Path -Parent $CodeDir')) "${n}: data root = parent of the code folder"
}

# the moved helpers must not be left in the root
foreach ($n in @("presets.cmd", "run_cascade.cmd", "patch_rbp_open_failures.ps1", "show_ops.vbs")) {
    Check (-not (Test-Path -LiteralPath (Join-Path $ToolSrc $n))) "not in the root any more: $n"
}

Write-Host ""
Write-Host ("Checks passed: {0}, failed: {1}" -f $script:Pass, $script:Fail)
if ($script:Fail -gt 0) { exit 1 }
exit 0
