# Adds RSN:// support to the installed Revit Batch Processor scripts.
# No Python required. Safe to run repeatedly.
# v2: avoid IronPython str()/ascii on Cyrillic RSN paths.
# v3: RSN + Detach allowed (superseded for deep).
# v4: deep = open CENTRAL with Audit only (no Create New Local, no Detach).
#     RBP stock always routes workshared centrals to RunDetachedDocumentAction;
#     for BATCH_COMPACT_MODE=deep we force RunDocumentAction (DoNotDetach).

[CmdletBinding()]
param(
    [string]$ScriptsPath = "",
    [switch]$Quiet
)

$ErrorActionPreference = "Stop"
$script:Marker = "RSN_PATCH_PS_v4"

function Write-Status([string]$Text) {
    if ($Quiet) { return }
    Write-Host $Text
}

function Write-Fail([string]$Text) {
    if ($Quiet) {
        [Console]::Error.WriteLine($Text)
        return
    }
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
        Get-ChildItem -LiteralPath $root -Directory -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -match '(?i)BatchRvt|RevitBatchProcessor|Revit Batch Processor' } |
            ForEach-Object {
                $scripts = Join-Path $_.FullName "Scripts"
                if (Test-RbpScripts $scripts) { $candidates.Add($scripts) }
            }
        Get-ChildItem -LiteralPath $root -Directory -ErrorAction SilentlyContinue |
            ForEach-Object {
                $scripts = Join-Path $_.FullName "Scripts"
                if (Test-RbpScripts $scripts) { $candidates.Add($scripts) }
            }
    }

    foreach ($c in $candidates) {
        if (Test-RbpScripts $c) { return $c }
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
            'try:'
            '    _p = self.revitFilePath'
            '    _name = None'
            '    try:'
            '        _name = System.IO.Path.GetFileName(_p)'
            '    except Exception:'
            '        _name = None'
            '    if _name:'
            '        import re as _re_rsn'
            '        _m = _re_rsn.search(r"(?i)_R(2[1-6])(?:\D|$)", _name)'
            '        if _m:'
            '            revitVersionNumber = int("20" + _m.group(1))'
            '            revitVersionText = str(revitVersionNumber)'
            'except Exception:'
            '    pass'
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

    # RSN classification + deep/quick open policy
    $text = Patch-ByRegex -Text $text -Label "script-host RSN classification" `
        -Pattern '(?m)^([ \t]*)if isCloudModel:\r?\n\1[ \t]+output\(\)\r?\n\1[ \t]+output\("The file is a Cloud Model\."\)\r?\n\1elif revit_file_util\.IsWorkshared\(centralFilePath\):' `
        -Replacement (Get-IndentedBlock @(
            'if isCloudModel:'
            '    output()'
            '    output("The file is a Cloud Model.")'
            ('elif (' + $rsnCheck + '):  # ' + $m)
            '    output()'
            '    output("The file is a Revit Server model (RSN).")'
            '    _bcm = ""'
            '    try:'
            '        import os as _os_bcm'
            '        _bcm = (_os_bcm.environ.get("BATCH_COMPACT_MODE") or "").strip().lower()'
            '    except Exception:'
            '        _bcm = ""'
            '    if _bcm == "deep":'
            '        openCreateNewLocal = False'
            '        output()'
            '        output("BATCH_COMPACT deep (RSN): open CENTRAL + Audit, no local, no Detach.")'
            '    else:'
            '        openCreateNewLocal = True'
            '        if centralFileOpenOption != BatchRvt.CentralFileOpenOption.CreateNewLocal:'
            '            output()'
            '            output("WARNING: RSN quick path forces Create New Local.")'
            'elif revit_file_util.IsWorkshared(centralFilePath):'
        ))

    # After CreateNewLocal flag from stock RBP, force deep = no local
    $text = Patch-ByRegex -Text $text -Label "script-host deep force no-local" `
        -Pattern '(?m)^([ \t]*)if centralFileOpenOption == BatchRvt\.CentralFileOpenOption\.CreateNewLocal:\r?\n\1[ \t]+openCreateNewLocal = True' `
        -Replacement (Get-IndentedBlock @(
            'if centralFileOpenOption == BatchRvt.CentralFileOpenOption.CreateNewLocal:'
            '    openCreateNewLocal = True'
            ('# ' + $m + ' deep: never Create New Local / never Detach')
            '_bcm2 = ""'
            'try:'
            '    import os as _os_bcm2'
            '    _bcm2 = (_os_bcm2.environ.get("BATCH_COMPACT_MODE") or "").strip().lower()'
            'except Exception:'
            '    _bcm2 = ""'
            'if _bcm2 == "deep":'
            '    openCreateNewLocal = False'
        ))

    # Workshared central without local: stock RBP always Detach — deep opens central instead
    $text = Patch-ByRegex -Text $text -Label "script-host deep open central not detach" `
        -Pattern '(?m)^([ \t]*)elif isCentralModel or isLocalModel:\r?\n\1[ \t]+result = revit_script_util\.RunDetachedDocumentAction\(' `
        -Replacement (Get-IndentedBlock @(
            'elif isCentralModel or isLocalModel:'
            ('    # ' + $m)
            '    _bcm3 = ""'
            '    try:'
            '        import os as _os_bcm3'
            '        _bcm3 = (_os_bcm3.environ.get("BATCH_COMPACT_MODE") or "").strip().lower()'
            '    except Exception:'
            '        _bcm3 = ""'
            '    if _bcm3 == "deep":'
            '        output()'
            '        output("BATCH_COMPACT deep: RunDocumentAction = open CENTRAL (Audit, DoNotDetach).")'
            '        result = revit_script_util.RunDocumentAction(uiapp, openInUI, centralFilePath, auditOnOpening, processDocument, output)'
            '    else:'
            '        result = revit_script_util.RunDetachedDocumentAction('
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
        Write-Fail "Keep Revit Batch Processor open, then run Сжатие.vbs again."
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
    Write-Status "RBP RSN patch: OK (v4, deep=open central+Audit, no Detach)"
    exit 0
}
catch {
    Write-Fail ("RSN patch FAILED: " + $_.Exception.Message)
    if ($ScriptsPath) { Write-Fail ("ScriptsPath was: " + $ScriptsPath) }
    exit 1
}
