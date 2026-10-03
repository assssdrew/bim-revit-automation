# Name: apply_rsn_support.ps1
# Version: 2.0
# What it does: Copy RSN-aware helpers into the installed Revit Batch Processor Scripts tree.
# Inputs: Optional -ScriptsPath; auto-detects default RBP location when empty.
# Outputs: Modified revit_script_util and related files under RBP Scripts.
# How to run: powershell -ExecutionPolicy Bypass -File apply_rsn_support.ps1
# Notes: Safe to re-run; v2 avoids IronPython ascii issues on Cyrillic RSN paths.
[CmdletBinding()]
param(
    [string]$ScriptsPath = "",
    [switch]$Quiet
)

$ErrorActionPreference = "Stop"
$script:Marker = "RSN_PATCH_PS_v2"

function Write-Status([string]$Text) {
    Write-Host $Text
}

function Write-Fail([string]$Text) {
    [Console]::Error.WriteLine($Text)
    Write-Host $Text
}

function Test-RbpScripts([string]$Path) {
    if ([string]::IsNullOrWhiteSpace($Path)) { return $false }
    return (
        (Test-Path -LiteralPath (Join-Path $Path "revit_file_list.py") -PathType Leaf) -and
        (Test-Path -LiteralPath (Join-Path $Path "revit_script_host.py") -PathType Leaf)
    )
}

function Find-RbpScripts {
    $candidates = New-Object System.Collections.Generic.List[string]

    Get-Process -ErrorAction SilentlyContinue |
        Where-Object { $_.ProcessName -match '^(?i)BatchRvt(GUI)?$|RevitBatchProcessor$' } |
        ForEach-Object {
            try {
                $exeDir = Split-Path -Parent $_.MainModule.FileName
                if ($exeDir) {
                    $candidates.Add((Join-Path $exeDir "Scripts"))
                    Write-Status ("Found RBP process: " + $_.MainModule.FileName)
                }
            }
            catch {
                Write-Status ("Process " + $_.ProcessName + " found, but path is not readable.")
            }
        }

    foreach ($root in @($env:LOCALAPPDATA, $env:ProgramFiles, ${env:ProgramFiles(x86)}, $env:ProgramW6432)) {
        if ([string]::IsNullOrWhiteSpace($root)) { continue }
        foreach ($relative in @(
            "RevitBatchProcessor\Scripts",
            "BatchRvt\Scripts",
            "Revit Batch Processor\Scripts",
            "Programs\Revit Batch Processor\Scripts"
        )) {
            $candidates.Add((Join-Path $root $relative))
        }
    }

    # Known path from potapov logs
    if ($env:LOCALAPPDATA) {
        $candidates.Add((Join-Path $env:LOCALAPPDATA "RevitBatchProcessor\Scripts"))
    }

    $la = $env:LOCALAPPDATA
    if ($la -and (Test-Path -LiteralPath $la)) {
        Get-ChildItem -LiteralPath $la -Filter "BatchRvt.exe" -Recurse -ErrorAction SilentlyContinue -Depth 5 |
            Select-Object -First 8 |
            ForEach-Object { $candidates.Add((Join-Path $_.DirectoryName "Scripts")) }
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

function Restore-FromBackupIfPatched([string]$Path) {
    $text = [System.IO.File]::ReadAllText($Path)
    if ($text.Contains($script:Marker)) {
        return $false  # already latest
    }
    if ($text -match 'RSN_PATCH_PS_v\d+') {
        $bak = $Path + ".bak_before_rsn"
        if (Test-Path -LiteralPath $bak) {
            Copy-Item -LiteralPath $bak -Destination $Path -Force
            Write-Status ("Restored older patch from backup: " + [IO.Path]::GetFileName($Path))
            return $true
        }
        throw ("Found old RSN patch in {0}, but .bak_before_rsn is missing. Reinstall RBP Scripts or restore manually." -f $Path)
    }
    return $false
}

function Save-PatchedFile([string]$Path, [string]$Text) {
    $backup = $Path + ".bak_before_rsn"
    if (-not (Test-Path -LiteralPath $backup)) {
        # Backup only pristine (unpatched) content once.
        $current = [System.IO.File]::ReadAllText($Path)
        if ($current -notmatch 'RSN_PATCH_PS_v\d+') {
            Copy-Item -LiteralPath $Path -Destination $backup -Force
            Write-Status ("Backup: " + $backup)
        }
    }
    $normalized = $Text.Replace("`r`n", "`n").Replace("`r", "`n")
    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($Path, $normalized, $utf8NoBom)
}

function Patch-ByRegex {
    param(
        [string]$Text,
        [string]$Pattern,
        [string]$Replacement,
        [string]$Label
    )
    $options = [System.Text.RegularExpressions.RegexOptions]::Multiline
    $rx = New-Object System.Text.RegularExpressions.Regex($Pattern, $options)
    if (-not $rx.IsMatch($Text)) {
        throw ("RBP version mismatch: block '{0}' was not found." -f $Label)
    }
    return $rx.Replace($Text, $Replacement, 1)
}

function Get-IndentedBlock {
    param([string[]]$Lines)
    $sb = New-Object System.Text.StringBuilder
    foreach ($line in $Lines) {
        if ($line -eq "") {
            [void]$sb.AppendLine("")
        }
        else {
            [void]$sb.AppendLine('${1}' + $line)
        }
    }
    return $sb.ToString().TrimEnd("`r", "`n")
}

function Patch-RevitFileList([string]$Path) {
    [void](Restore-FromBackupIfPatched $Path)
    $text = [System.IO.File]::ReadAllText($Path)
    if ($text.Contains($script:Marker)) {
        Write-Status "Already patched: revit_file_list.py"
        return
    }

    $m = $script:Marker

    $text = Patch-ByRegex -Text $text -Label "RevitFileInfo.IsCloudModel" `
        -Pattern '(?m)^([ \t]*)def IsCloudModel\(self\):\r?\n\1[ \t]+return self\.GetRevitCloudModelInfo\(\)\.IsValid\(\)\s*$' `
        -Replacement (Get-IndentedBlock @(
            'def IsCloudModel(self):'
            '    return self.GetRevitCloudModelInfo().IsValid()'
            ''
            'def IsServerModel(self):'
            ('    # ' + $m + ' IronPython-safe: never str() on Cyrillic paths')
            '    p = self.revitFilePath'
            '    if p is None:'
            '        return False'
            '    try:'
            '        return p.StartsWith("RSN://", System.StringComparison.OrdinalIgnoreCase)'
            '    except Exception:'
            '        try:'
            '            u = unicode(p)'
            '            return u.upper().startswith(u"RSN://")'
            '        except Exception:'
            '            return False'
        ))

    $text = Patch-ByRegex -Text $text -Label "RevitFileInfo.GetFileSize" `
        -Pattern '(?m)^([ \t]*)def GetFileSize\(self\):\r?\n\1[ \t]+return path_util\.GetFileSize\(self\.revitFilePath\)\s*$' `
        -Replacement (Get-IndentedBlock @(
            'def GetFileSize(self):'
            ('    # ' + $m)
            '    if self.IsServerModel():'
            '        return 0'
            '    return path_util.GetFileSize(self.revitFilePath)'
        ))

    $text = Patch-ByRegex -Text $text -Label "RevitFileInfo.Exists" `
        -Pattern '(?m)^([ \t]*)def Exists\(self\):\r?\n\1[ \t]+return path_util\.FileExists\(self\.revitFilePath\)\s*$' `
        -Replacement (Get-IndentedBlock @(
            'def Exists(self):'
            ('    # ' + $m)
            '    if self.IsServerModel():'
            '        return True'
            '    return path_util.FileExists(self.revitFilePath)'
        ))

    $text = Patch-ByRegex -Text $text -Label "SupportedRevitFileInfo version assign" `
        -Pattern '(?m)^([ \t]*)self\.revitVersionText = revitVersionText\r?\n\1self\.revitVersionNumber = revitVersionNumber\r?\n\1return\s*$' `
        -Replacement (Get-IndentedBlock @(
            ('# ' + $m + ' guess year from *_R22.rvt (no str() on full path)')
            'if revitVersionNumber is None and self.revitFileInfo.IsServerModel():'
            '    full = self.revitFileInfo.GetFullPath()'
            '    name = u""'
            '    try:'
            '        if full is not None:'
            '            s = full.Replace("\\", "/")'
            '            parts = s.Split("/")'
            '            name = unicode(parts[parts.Length - 1]).upper()'
            '    except Exception:'
            '        try:'
            '            name = unicode(full).replace(u"\\", u"/").split(u"/")[-1].upper()'
            '        except Exception:'
            '            name = u""'
            '    for year in range(2026, 2014, -1):'
            '        yy = str(year)[-2:]'
            '        if (u"_R" + unicode(year)) in name or (u"_R" + unicode(yy) + u".") in name or (u"_R" + unicode(yy) + u"_") in name:'
            '            if RevitVersion.IsSupportedRevitVersionNumber(str(year)):'
            '                revitVersionNumber = RevitVersion.GetSupportedRevitVersion(str(year))'
            '                revitVersionText = str(year)'
            '            break'
            'self.revitVersionText = revitVersionText'
            'self.revitVersionNumber = revitVersionNumber'
            'return'
        ))

    Save-PatchedFile $Path $text
    Write-Status "Patched: revit_file_list.py"
}

function Patch-RevitScriptHost([string]$Path) {
    [void](Restore-FromBackupIfPatched $Path)
    $text = [System.IO.File]::ReadAllText($Path)
    if ($text.Contains($script:Marker)) {
        Write-Status "Already patched: revit_script_host.py"
        return
    }

    $m = $script:Marker

    # Helper injected once near top of RunBatchTaskScript usage area - as local check via .NET
    $rsnCheck = 'centralFilePath is not None and hasattr(centralFilePath, "StartsWith") and centralFilePath.StartsWith("RSN://", System.StringComparison.OrdinalIgnoreCase)'

    $text = Patch-ByRegex -Text $text -Label "script-host File.Exists gate" `
        -Pattern '(?m)^([ \t]*)elif not isCloudModel and not path_util\.FileExists\(centralFilePath\):\r?\n\1[ \t]+output\(\)\r?\n\1[ \t]+output\("ERROR: Revit project file does not exist!"\)' `
        -Replacement (Get-IndentedBlock @(
            'elif ('
            '        not isCloudModel'
            ('        and not (' + $rsnCheck + ')  # ' + $m)
            '        and not path_util.FileExists(centralFilePath)'
            '    ):'
            '    output()'
            '    output("ERROR: Revit project file does not exist!")'
        ))

    $text = Patch-ByRegex -Text $text -Label "script-host RSN classification" `
        -Pattern '(?m)^([ \t]*)if isCloudModel:\r?\n\1[ \t]+output\(\)\r?\n\1[ \t]+output\("The file is a Cloud Model\."\)\r?\n\1elif revit_file_util\.IsWorkshared\(centralFilePath\):' `
        -Replacement (Get-IndentedBlock @(
            'if isCloudModel:'
            '    output()'
            '    output("The file is a Cloud Model.")'
            ('elif (' + $rsnCheck + '):  # ' + $m)
            '    output()'
            '    output("The file is a Revit Server model (RSN).")'
            '    openCreateNewLocal = True'
            '    if centralFileOpenOption != BatchRvt.CentralFileOpenOption.CreateNewLocal:'
            '        output()'
            '        output("WARNING: RSN forces Create New Local; Detach is not used.")'
            'elif revit_file_util.IsWorkshared(centralFilePath):'
        ))

    Save-PatchedFile $Path $text
    Write-Status "Patched: revit_script_host.py"
}

try {
    if ([string]::IsNullOrWhiteSpace($ScriptsPath)) {
        $ScriptsPath = Find-RbpScripts
    }
    if (-not (Test-RbpScripts $ScriptsPath)) {
        Write-Fail "BatchRvt Scripts folder not found."
        Write-Fail "Keep Revit Batch Processor open, then run choose_models_path.cmd again (menu 3)."
        exit 2
    }

    Write-Status ("RBP Scripts: " + $ScriptsPath)
    try {
        $probe = Join-Path $ScriptsPath "_rsn_patch_write_test.tmp"
        "ok" | Set-Content -LiteralPath $probe -Encoding ASCII
        Remove-Item -LiteralPath $probe -Force
    }
    catch {
        Write-Fail ("Cannot write to Scripts folder (permission): " + $ScriptsPath)
        Write-Fail $_.Exception.Message
        exit 1
    }

    Patch-RevitFileList (Join-Path $ScriptsPath "revit_file_list.py")
    Patch-RevitScriptHost (Join-Path $ScriptsPath "revit_script_host.py")
    Write-Status "RBP RSN patch: OK (v2, Cyrillic-safe)"
    exit 0
}
catch {
    Write-Fail ("RSN patch FAILED: " + $_.Exception.Message)
    if ($ScriptsPath) { Write-Fail ("ScriptsPath was: " + $ScriptsPath) }
    exit 1
}
