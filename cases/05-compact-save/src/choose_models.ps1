# Unified model picker for Revit Batch Processor.
# Sources: local folder / local files / Revit Server (RSN://).
# Writes rvt_list.txt - can replace or append.
# Encoding: ASCII-friendly messages for Windows PowerShell 5.1.

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Drawing

$ToolDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# Path helpers: try list_paths.ps1 on share; if missing, use built-in copy (partial deploy OK).
$_listPathsFile = Join-Path $ToolDir "list_paths.ps1"
$_listPathsLoaded = $false
if (Test-Path -LiteralPath $_listPathsFile) {
    try {
        . $_listPathsFile
        $_listPathsLoaded = (
            [bool](Get-Command Resolve-RvtListPaths -ErrorAction SilentlyContinue) -and
            [bool](Get-Command Save-ActiveListPointer -ErrorAction SilentlyContinue) -and
            [bool](Get-Command Write-TextFileLines -ErrorAction SilentlyContinue) -and
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
    function Write-TextFileLines {
        param([string]$Path, [string[]]$Lines)
        $dir = Split-Path -Parent $Path
        if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        if (Test-Path -LiteralPath $Path) {
            try {
                $item = Get-Item -LiteralPath $Path -Force
                if ($item.IsReadOnly) { $item.IsReadOnly = $false }
            }
            catch { }
        }
        $Lines | Set-Content -LiteralPath $Path -Encoding UTF8 -ErrorAction Stop
    }
    function Write-CfgContent {
        param([string]$Path, [string]$Value)
        try {
            $dir = Split-Path -Parent $Path
            if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
            Set-Content -LiteralPath $Path -Value $Value -Encoding UTF8 -ErrorAction Stop
        }
        catch {
            Write-Host ("WARN: could not save {0}: {1}" -f (Split-Path -Leaf $Path), $_.Exception.Message)
        }
    }
}

$listPaths = Resolve-RvtListPaths -ToolDir $ToolDir
if (Get-Command Remove-LeftoverWriteProbes -ErrorAction SilentlyContinue) {
    Remove-LeftoverWriteProbes -Folder $ToolDir
}
$OutList = $listPaths.OutList
$ShareOutList = $listPaths.ShareList
$OutListIsLocal = $listPaths.IsLocalList
$PathCfg = Join-Path $listPaths.CfgDir "models_path.cfg"
$LastCfg = Join-Path $listPaths.CfgDir "last_selection.cfg"
$ServersCfg = Join-Path $ToolDir "servers.cfg"

try {
    if ($ToolDir -and (Test-Path -LiteralPath $ToolDir)) {
        Set-Location -LiteralPath $ToolDir
    }
}
catch { }

if ($OutListIsLocal) {
    Write-Host ""
    Write-Host "NOTE: no write access to share tool folder."
    Write-Host ("      List will be saved locally: {0}" -f $OutList)
    Write-Host "      Сжатие.cmd подхватит этот список."
    Write-Host ""
}

function Show-OwnedDialog($Dialog) {
    if ($script:UiOwner -and -not $script:UiOwner.IsDisposed) {
        return $Dialog.ShowDialog($script:UiOwner)
    }
    $owner = New-Object System.Windows.Forms.Form
    $owner.Text = "batch_compact_save"
    $owner.TopMost = $true
    $owner.ShowInTaskbar = $false
    $owner.StartPosition = "CenterScreen"
    $owner.Size = New-Object System.Drawing.Size(1, 1)
    $owner.Opacity = 0
    $owner.Show()
    $owner.Activate()
    try {
        return $Dialog.ShowDialog($owner)
    }
    finally {
        $owner.Hide()
        $owner.Close()
        $owner.Dispose()
    }
}

try {
    if ($ToolDir -and (Test-Path -LiteralPath $ToolDir)) {
        Set-Location -LiteralPath $ToolDir
    }
}
catch { }

function Show-OwnedDialog($Dialog) {
    $owner = New-Object System.Windows.Forms.Form
    $owner.Text = "batch_compact_save"
    $owner.TopMost = $true
    $owner.ShowInTaskbar = $false
    $owner.StartPosition = "CenterScreen"
    $owner.Size = New-Object System.Drawing.Size(1, 1)
    $owner.Opacity = 0
    $owner.Show()
    $owner.Activate()
    try {
        return $Dialog.ShowDialog($owner)
    }
    finally {
        $owner.Hide()
        $owner.Close()
        $owner.Dispose()
    }
}

$DefaultServers = @(
    @{ Host = "revit-server-2021.example.local"; Label = "Revit Server 2021"; Year = "2021" }
    @{ Host = "revit-server-2022.example.local"; Label = "Revit Server 2022"; Year = "2022" }
    @{ Host = "revit-server-2023.example.local"; Label = "Revit Server 2023"; Year = "2023" }
    @{ Host = "revit-server-2024.example.local"; Label = "Revit Server 2024"; Year = "2024" }
)

function Get-NativePath([string]$Path) {
    if ([string]::IsNullOrWhiteSpace($Path)) { return $null }
    $p = $Path.Trim().Trim('"')
    $marker = "FileSystem::"
    $idx = $p.IndexOf($marker, [StringComparison]::OrdinalIgnoreCase)
    if ($idx -ge 0) {
        $p = $p.Substring($idx + $marker.Length)
    }
    # Keep UNC as-is (FileInfo.FullName can rewrite oddly on some hosts)
    if ($p.StartsWith("\\")) { return $p }
    try {
        return (New-Object System.IO.FileInfo($p)).FullName
    }
    catch {
        return $p
    }
}

function Write-DefaultServersFile {
    $lines = @(
        "# Format: host|display name|REST year"
        "# host without RSN://"
    )
    foreach ($s in $DefaultServers) {
        $lines += ("{0}|{1}|{2}" -f $s.Host, $s.Label, $s.Year)
    }
    try {
        Write-TextFileLines -Path $ServersCfg -Lines $lines
    }
    catch {
        Write-Host ("WARN: could not write servers.cfg: {0}" -f $_.Exception.Message)
    }
}

function Ensure-ServersFile {
    if (-not (Test-Path -LiteralPath $ServersCfg)) {
        Write-DefaultServersFile
        return
    }
    $any = $false
    Get-Content -LiteralPath $ServersCfg -ErrorAction SilentlyContinue |
        ForEach-Object {
            if (Parse-ServerLine $_) { $any = $true }
        }
    if (-not $any) {
        Write-DefaultServersFile
    }
}

function Parse-ServerLine([string]$Line) {
    $raw = $Line.Trim()
    if (-not $raw -or $raw.StartsWith("#")) { return $null }

    $hostName = $null
    $label = $null
    $year = $null

    if ($raw.Contains("|")) {
        $parts = $raw.Split("|")
        $hostName = $parts[0].Trim()
        if ($parts.Count -ge 2) { $label = $parts[1].Trim() }
        if ($parts.Count -ge 3) { $year = $parts[2].Trim() }
    }
    else {
        $hostName = $raw
    }

    if ([string]::IsNullOrWhiteSpace($hostName)) { return $null }
    if ([string]::IsNullOrWhiteSpace($label)) { $label = $hostName }
    if ([string]::IsNullOrWhiteSpace($year)) {
        # guess from known defaults
        $hit = $DefaultServers | Where-Object { $_.Host -eq $hostName } | Select-Object -First 1
        if ($hit) { $year = $hit.Year } else { $year = "2024" }
    }

    return [pscustomobject]@{
        Host  = $hostName
        Label = $label
        Year  = $year
    }
}

function Get-Servers {
    Ensure-ServersFile
    $result = @()
    Get-Content -LiteralPath $ServersCfg -ErrorAction SilentlyContinue |
        ForEach-Object {
            $item = Parse-ServerLine $_
            if ($item) { $result += $item }
        }
    return $result
}

function Save-Server {
    param(
        [string]$HostName,
        [string]$Label = "",
        [string]$Year = "2024"
    )
    Ensure-ServersFile
    $list = @(Get-Servers)
    if ($list | Where-Object { $_.Host -eq $HostName }) { return }
    if ([string]::IsNullOrWhiteSpace($Label)) { $Label = $HostName }
    Add-Content -LiteralPath $ServersCfg -Value ("{0}|{1}|{2}" -f $HostName, $Label, $Year) -Encoding UTF8
}

function Normalize-RsnPath([string]$Server, [string]$Path) {
    $p = $Path.Trim().Trim('"').Replace("\", "/")
    while ($p.StartsWith("/")) { $p = $p.Substring(1) }

    if ($p -match '^(?i)RSN://') {
        return $p
    }

    if (-not $p.ToLower().EndsWith(".rvt")) {
        throw "Path must end with .rvt: $p"
    }

    return ("RSN://{0}/{1}" -f $Server, $p)
}

function Read-ExistingList {
    if (-not (Test-Path -LiteralPath $OutList)) { return @() }
    @(
        Get-Content -LiteralPath $OutList -ErrorAction SilentlyContinue |
            ForEach-Object { $_.Trim() } |
            Where-Object {
                $_ -and -not $_.StartsWith("#") -and (
                    $_.ToLower().EndsWith(".rvt") -or $_ -match '^(?i)RSN://'
                )
            }
    )
}

function Write-RvtList {
    param(
        [string[]]$Paths,
        [switch]$Append
    )

    $incoming = @()
    foreach ($raw in $Paths) {
        if ([string]::IsNullOrWhiteSpace($raw)) { continue }
        $p = $raw.Trim()
        if ($p -match '^(?i)RSN://') {
            $incoming += $p
        }
        else {
            $native = Get-NativePath $p
            if ($native) { $incoming += $native }
        }
    }

    if ($Append) {
        $incoming = @(Read-ExistingList) + $incoming
    }

    $unique = @($incoming | Where-Object { $_ } | Sort-Object -Unique)

    if ($unique.Count -eq 0) {
        Write-Host "No models selected. List not updated."
        return $false
    }

    try {
        Write-TextFileLines -Path $OutList -Lines $unique
    }
    catch {
        if (-not $OutListIsLocal) {
            Write-Host ""
            Write-Host ("WARN: cannot write share list: {0}" -f $_.Exception.Message)
            $localDir = Get-LocalBatchCompactDir
            $script:OutList = Join-Path $localDir "rvt_list.txt"
            $script:PathCfg = Join-Path $localDir "models_path.cfg"
            $script:LastCfg = Join-Path $localDir "last_selection.cfg"
            $script:OutListIsLocal = $true
            Write-Host ("      Falling back to local list: {0}" -f $script:OutList)
            try {
                Write-TextFileLines -Path $script:OutList -Lines $unique
            }
            catch {
                Write-Host ""
                Write-Host ("ERROR: cannot write rvt_list.txt: {0}" -f $_.Exception.Message)
                return $false
            }
        }
        else {
            Write-Host ""
            Write-Host ("ERROR: cannot write rvt_list.txt: {0}" -f $_.Exception.Message)
            return $false
        }
    }

    Write-Host ""
    Write-Host ("Models in list: {0}" -f $unique.Count)
    Write-Host ("List: {0}" -f $OutList)
    Save-ActiveListPointer -ListPath $OutList
    if ($OutListIsLocal) {
        Write-Host "Saved for this PC (no share write). Next: Сжатие.cmd"
    }
    $localN = @($unique | Where-Object { $_ -notmatch '^(?i)RSN://' }).Count
    $rsnN = @($unique | Where-Object { $_ -match '^(?i)RSN://' }).Count
    Write-Host ("  Local/UNC: {0}" -f $localN)
    Write-Host ("  Revit Server (RSN): {0}" -f $rsnN)
    Write-Host ("Example: {0}" -f $unique[0])
    return $true
}

function Get-InnermostMessage($ErrorRecord) {
    $ex = $null
    if ($ErrorRecord -is [System.Management.Automation.ErrorRecord]) {
        $ex = $ErrorRecord.Exception
    }
    else {
        $ex = $ErrorRecord
    }
    $msg = ""
    while ($ex) {
        if ($ex.Message) { $msg = [string]$ex.Message }
        $ex = $ex.InnerException
    }
    return $msg
}

function Test-IsServerUnreachable([string]$Text) {
    if ([string]::IsNullOrWhiteSpace($Text)) { return $false }
    return ($Text -match 'Unable to connect|No connection could be made|actively refused|The operation has timed out|timed out|The remote name could not be resolved|network path was not found')
}

function Format-RevitServerError {
    param(
        [string]$Raw,
        [string]$Server,
        [string]$Year
    )
    if (Test-IsServerUnreachable $Raw) {
        return (
            "Нет связи с Revit Server {0} (год {1})." + [Environment]::NewLine + [Environment]::NewLine +
            "Проверьте VPN/сеть и что служба Revit Server запущена." + [Environment]::NewLine +
            "Окно можно закрыть и выбрать другой сервер."
        ) -f $Server, $Year
    }
    return $Raw
}

function Invoke-RevitServerGet {
    param([string]$Url)

    $request = [System.Net.HttpWebRequest]::Create($Url)
    $request.Method = "GET"
    $request.Timeout = 8000
    $request.ReadWriteTimeout = 8000
    $request.AutomaticDecompression = [System.Net.DecompressionMethods]::GZip -bor [System.Net.DecompressionMethods]::Deflate
    $request.Accept = "application/json"
    $request.Headers.Add("User-Name", $env:USERNAME)
    $request.Headers.Add("User-Machine-Name", $env:COMPUTERNAME)
    $request.Headers.Add("Operation-GUID", [guid]::NewGuid().ToString())

    $response = $null
    try {
        $response = $request.GetResponse()
        $stream = $response.GetResponseStream()
        $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8)
        $json = $reader.ReadToEnd()
        $reader.Close()
        if ([string]::IsNullOrWhiteSpace($json)) {
            throw "Empty response body"
        }
        return ($json | ConvertFrom-Json)
    }
    catch {
        $inner = Get-InnermostMessage $_
        if ($inner) { throw $inner }
        throw
    }
    finally {
        if ($response) { $response.Close() }
    }
}

function Encode-RevitServerSegment {
    param([string]$Text)
    # UTF-8 percent-encoding (same idea as Uri.EscapeDataString, explicit)
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($Text)
    $sb = New-Object System.Text.StringBuilder
    foreach ($b in $bytes) {
        $c = [char]$b
        $isUnreserved =
            ($b -ge 48 -and $b -le 57) -or
            ($b -ge 65 -and $b -le 90) -or
            ($b -ge 97 -and $b -le 122) -or
            $b -eq 45 -or $b -eq 46 -or $b -eq 95 -or $b -eq 126
        if ($isUnreserved) {
            [void]$sb.Append($c)
        }
        else {
            [void]$sb.Append('%')
            [void]$sb.Append($b.ToString("X2"))
        }
    }
    return $sb.ToString()
}

function Get-RevitServerContents {
    param(
        [string]$Server,
        [string]$Folder,
        [string]$Year
    )

    $inputFolder = ""
    if ($null -ne $Folder) { $inputFolder = [string]$Folder }
    $inputFolder = $inputFolder.Trim().Trim("/\").Replace("\", "/").Trim("/")
    Write-Host ("RS contents: Server={0} Year={1} Folder=[{2}]" -f $Server, $Year, $inputFolder)

    $segments = @()
    if (-not [string]::IsNullOrWhiteSpace($inputFolder)) {
        $segments = @(
            $inputFolder.Split(@("/", "|"), [System.StringSplitOptions]::RemoveEmptyEntries)
        )
    }

    $base = ("http://{0}/RevitServerAdminRESTService{1}/AdminRESTService.svc/" -f $Server, $Year)

    $pathVariants = @()
    if ($segments.Count -eq 0) {
        $pathVariants += "%7C/contents"
        $pathVariants += "%7C/Contents"
    }
    else {
        $enc = @($segments | ForEach-Object { Encode-RevitServerSegment $_ })
        $joined = ($enc -join "%7C")
        $rawJoined = ($segments -join "|")

        # Common working patterns for folder contents
        $pathVariants += ("%7C{0}/contents" -f $joined)
        $pathVariants += ("%7C{0}%7C/contents" -f $joined)   # trailing pipe
        $pathVariants += ("%7C{0}/Contents" -f $joined)
        $pathVariants += ("%7C{0}%7C/Contents" -f $joined)
        $pathVariants += ("|{0}/contents" -f $rawJoined)
        $pathVariants += ("|{0}|/contents" -f $rawJoined)
    }

    $result = [pscustomobject]@{
        Folders = @()
        Models  = @()
        Error   = $null
        Url     = $null
        Path    = $null
    }

    $resp = $null
    $tried = @()
    $lastError = $null

    foreach ($svcPath in $pathVariants) {
        $url = $base + $svcPath
        $result.Url = $url
        $tried += $url
        try {
            $resp = Invoke-RevitServerGet -Url $url
            $lastError = $null
            break
        }
        catch {
            $lastError = Get-InnermostMessage $_
            $resp = $null
            if (Test-IsServerUnreachable $lastError) {
                break
            }
        }
    }

    if (-not $resp) {
        $result.Error = ("{0} | last URL: {1}" -f $lastError, $result.Url)
        return $result
    }

    $respPath = ""
    if ($resp.PSObject.Properties.Name -contains "Path" -and $null -ne $resp.Path) {
        $respPath = [string]$resp.Path
    }
    $result.Path = $respPath

    $folderItems = @()
    if ($null -ne $resp.Folders) {
        $folderItems = @($resp.Folders)
    }

    foreach ($f in $folderItems) {
        $name = $null
        if ($f.Name) { $name = [string]$f.Name }
        elseif ($f.name) { $name = [string]$f.name }
        if ($name) { $result.Folders += $name }
    }

    $modelItems = @()
    if ($null -ne $resp.Models) {
        $modelItems = @($resp.Models)
    }

    foreach ($m in $modelItems) {
        $name = $null
        if ($m.Name) { $name = [string]$m.Name }
        elseif ($m.name) { $name = [string]$m.name }
        elseif ($m -is [string]) { $name = [string]$m }
        if (-not $name) { continue }
        if (-not $name.ToLower().EndsWith(".rvt")) { $name = ($name + ".rvt") }
        $result.Models += $name
    }

    $result.Folders = @($result.Folders | Sort-Object -Unique)
    $result.Models = @($result.Models | Sort-Object -Unique)

    # Guard: requesting a subfolder but got root listing back (Path empty + leaf visible as child)
    if ($segments.Count -gt 0) {
        $leaf = $segments[$segments.Count - 1]
        $normResp = $respPath.Replace("|", "/").Replace("\", "/").Trim("/")
        $gotRoot = [string]::IsNullOrWhiteSpace($normResp)
        $leafSeenAsChild = ($result.Folders -contains $leaf)
        if ($gotRoot -and $leafSeenAsChild) {
            $result.Folders = @()
            $result.Models = @()
            $result.Error = (
                "Folder did not open: server returned root listing instead of '{0}'. " +
                "Tried URLs like: {1}"
            ) -f $leaf, $tried[0]
            return $result
        }
    }

    return $result
}

function ConvertTo-RsnPath {
    param(
        [string]$Server,
        [string]$FolderPath,
        [string]$ModelName
    )
    # IMPORTANT: do NOT use $folder — in PowerShell $Folder/$folder are the same variable
    # and "$folder = ''" would wipe the -FolderPath/-Folder argument.
    $normalizedFolder = ""
    if (-not [string]::IsNullOrWhiteSpace($FolderPath)) {
        $normalizedFolder = $FolderPath.Trim().Trim("/\").Replace("\", "/").Trim("/")
    }
    $model = [string]$ModelName
    if (-not [string]::IsNullOrWhiteSpace($model) -and -not $model.ToLower().EndsWith(".rvt")) {
        $model = $model + ".rvt"
    }
    if ([string]::IsNullOrWhiteSpace($normalizedFolder)) {
        return ("RSN://{0}/{1}" -f $Server, $model)
    }
    return ("RSN://{0}/{1}/{2}" -f $Server, $normalizedFolder, $model)
}

function Get-RevitServerModelsRecursive {
    param(
        [string]$Server,
        [string]$Folder,
        [string]$Year,
        [int]$Depth = 0
    )
    if ($Depth -gt 40) { return @() }

    $contents = Get-RevitServerContents -Server $Server -Folder $Folder -Year $Year
    if ($contents.Error) {
        Write-Host ("REST warn [{0}]: {1}" -f $Folder, $contents.Error)
        return @()
    }

    $found = @()
    foreach ($m in $contents.Models) {
        $found += (ConvertTo-RsnPath -Server $Server -FolderPath $Folder -ModelName $m)
    }
    foreach ($sub in $contents.Folders) {
        if ([string]::IsNullOrWhiteSpace($Folder)) {
            $child = $sub
        }
        else {
            $child = ($Folder.Trim("/\") + "/" + $sub)
        }
        $found += @(Get-RevitServerModelsRecursive -Server $Server -Folder $child -Year $Year -Depth ($Depth + 1))
    }
    return $found
}

function Show-RevitServerBrowser {
    param($ServerObj)

    Add-Type -AssemblyName System.Drawing | Out-Null

    $script:RsBrowserState = @{
        CurrentFolder = ""
        Selected      = @()
        Server        = [string]$ServerObj.Host
        Year          = [string]$ServerObj.Year
        Form          = $null
        List          = $null
        PathBox       = $null
        Status        = $null
    }

    $form = New-Object System.Windows.Forms.Form
    $form.Text = ("Revit Server: {0} ({1})" -f $ServerObj.Label, $ServerObj.Host)
    $form.Size = New-Object System.Drawing.Size(920, 640)
    $form.StartPosition = "CenterScreen"
    $form.TopMost = $true
    $form.MinimizeBox = $false
    $form.MaximizeBox = $true

    $lblPath = New-Object System.Windows.Forms.Label
    $lblPath.Text = "Folder path - select Folder then Open / double-click"
    $lblPath.Location = New-Object System.Drawing.Point(12, 12)
    $lblPath.AutoSize = $true

    $txtPath = New-Object System.Windows.Forms.TextBox
    $txtPath.Location = New-Object System.Drawing.Point(12, 36)
    $txtPath.Size = New-Object System.Drawing.Size(600, 25)

    $btnGo = New-Object System.Windows.Forms.Button
    $btnGo.Text = "Go"
    $btnGo.Location = New-Object System.Drawing.Point(620, 34)
    $btnGo.Size = New-Object System.Drawing.Size(70, 28)

    $btnUp = New-Object System.Windows.Forms.Button
    $btnUp.Text = "Up"
    $btnUp.Location = New-Object System.Drawing.Point(698, 34)
    $btnUp.Size = New-Object System.Drawing.Size(70, 28)

    $btnOpen = New-Object System.Windows.Forms.Button
    $btnOpen.Text = "Open"
    $btnOpen.Location = New-Object System.Drawing.Point(776, 34)
    $btnOpen.Size = New-Object System.Drawing.Size(90, 28)

    $list = New-Object System.Windows.Forms.ListView
    $list.Location = New-Object System.Drawing.Point(12, 72)
    $list.Size = New-Object System.Drawing.Size(860, 420)
    $list.View = [System.Windows.Forms.View]::Details
    $list.FullRowSelect = $true
    $list.MultiSelect = $true
    $list.HideSelection = $false
    $list.Activation = [System.Windows.Forms.ItemActivation]::Standard
    [void]$list.Columns.Add("Name", 640)
    [void]$list.Columns.Add("Type", 180)

    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Location = New-Object System.Drawing.Point(12, 500)
    $lblStatus.Size = New-Object System.Drawing.Size(860, 36)
    $lblStatus.Text = "Loading..."

    $btnAddSelected = New-Object System.Windows.Forms.Button
    $btnAddSelected.Text = "Add selected (models / folders)"
    $btnAddSelected.Location = New-Object System.Drawing.Point(12, 545)
    $btnAddSelected.Size = New-Object System.Drawing.Size(220, 32)

    $btnAddFolder = New-Object System.Windows.Forms.Button
    $btnAddFolder.Text = "Add ALL under current path"
    $btnAddFolder.Location = New-Object System.Drawing.Point(240, 545)
    $btnAddFolder.Size = New-Object System.Drawing.Size(240, 32)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel"
    $btnCancel.Location = New-Object System.Drawing.Point(792, 545)
    $btnCancel.Size = New-Object System.Drawing.Size(80, 32)
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

    $script:RsBrowserState.Form = $form
    $script:RsBrowserState.List = $list
    $script:RsBrowserState.PathBox = $txtPath
    $script:RsBrowserState.Status = $lblStatus

    # Use .Invoke(folder, showError) - named params on scriptblocks are unreliable in PS 5.1
    $script:RsRefresh = {
        param($Folder, $ShowError)

        $st = $script:RsBrowserState
        $requestFolder = ""
        if ($null -ne $Folder) {
            $requestFolder = [string]$Folder
        }
        $st.CurrentFolder = $requestFolder
        $st.PathBox.Text = $requestFolder

        $st.Form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        $st.Status.Text = ("Loading [{0}] ..." -f $requestFolder)
        $st.List.Items.Clear()
        [System.Windows.Forms.Application]::DoEvents()

        $contents = Get-RevitServerContents -Server ([string]$st.Server) -Folder $requestFolder -Year ([string]$st.Year)

        # Detect the exact bug from screenshots: asked for folder, got root URL
        if (-not [string]::IsNullOrWhiteSpace($requestFolder)) {
            if ($contents.Url -match 'AdminRESTService\.svc/%7C/contents$' -or
                $contents.Url -match 'AdminRESTService\.svc/%7C/Contents$') {
                $contents.Error = (
                    "Internal bug: requested folder '{0}' but REST URL is root: {1}" -f
                    $requestFolder, $contents.Url
                )
                $contents.Folders = @()
                $contents.Models = @()
            }
        }

        if ($contents.Error) {
            $friendly = Format-RevitServerError -Raw $contents.Error -Server ([string]$st.Server) -Year ([string]$st.Year)
            $st.Status.Text = $friendly.Replace([Environment]::NewLine, " | ")
            $st.Form.Cursor = [System.Windows.Forms.Cursors]::Default
            if ($ShowError -or (Test-IsServerUnreachable $contents.Error)) {
                [System.Windows.Forms.MessageBox]::Show(
                    $st.Form,
                    $friendly,
                    "Revit Server",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Warning
                ) | Out-Null
            }
            return $false
        }

        foreach ($f in @($contents.Folders)) {
            $item = New-Object System.Windows.Forms.ListViewItem([string]$f)
            [void]$item.SubItems.Add('Folder')
            # Hashtable Tag is reliable in WinForms from PS 5.1 (PSCustomObject Tag can lose props)
            $item.Tag = @{
                Kind   = 'folder'
                Name   = [string]$f
                Folder = [string]$requestFolder
            }
            [void]$st.List.Items.Add($item)
        }
        foreach ($m in @($contents.Models)) {
            $fullRsn = ConvertTo-RsnPath -Server ([string]$st.Server) -FolderPath $requestFolder -ModelName ([string]$m)
            $item = New-Object System.Windows.Forms.ListViewItem([string]$m)
            [void]$item.SubItems.Add('Model (.rvt)')
            $item.Tag = @{
                Kind    = 'model'
                Name    = [string]$m
                Folder  = [string]$requestFolder
                FullRsn = [string]$fullRsn
            }
            [void]$st.List.Items.Add($item)
            Write-Host ("LIST model: {0}" -f $fullRsn)
        }

        $fc = @($contents.Folders).Count
        $mc = @($contents.Models).Count
        $pathShow = if ([string]::IsNullOrWhiteSpace($requestFolder)) { "/" } else { ("/" + $requestFolder) }
        $st.Status.Text = (
            "req=[{0}]  folders={1} models={2}{3}URL={4}" -f
            $requestFolder, $fc, $mc, [Environment]::NewLine, $contents.Url
        )
        $st.Form.Cursor = [System.Windows.Forms.Cursors]::Default
        return $true
    }

    $script:RsOpenSelected = {
        $st = $script:RsBrowserState
        $lv = $st.List
        $item = $null
        if ($lv.SelectedItems.Count -gt 0) { $item = $lv.SelectedItems[0] }
        elseif ($lv.FocusedItem) { $item = $lv.FocusedItem }

        if (-not $item) {
            [System.Windows.Forms.MessageBox]::Show("Select a Folder row first.", "Revit Server") | Out-Null
            return
        }

        $meta = $item.Tag
        if (-not $meta -or [string]$meta['Kind'] -ne 'folder') {
            [System.Windows.Forms.MessageBox]::Show("Selected row is not a Folder.", "Revit Server") | Out-Null
            return
        }

        $name = [string]$meta['Name']
        $prev = [string]$st.CurrentFolder
        if ([string]::IsNullOrWhiteSpace($prev)) {
            $newFolder = $name
        }
        else {
            $newFolder = ($prev.Trim("/") + "/" + $name)
        }

        $ok = [bool]$script:RsRefresh.Invoke($newFolder, $true)
        if (-not $ok) {
            [void]$script:RsRefresh.Invoke($prev, $false)
        }
    }

    $btnGo.Add_Click({
        $st = $script:RsBrowserState
        $typed = $st.PathBox.Text.Trim().Trim("/\").Replace("\", "/")
        [void]$script:RsRefresh.Invoke($typed, $true)
    })

    $btnUp.Add_Click({
        $st = $script:RsBrowserState
        $cur = [string]$st.CurrentFolder
        if ([string]::IsNullOrWhiteSpace($cur)) { return }
        $parts = @($cur.Split("/") | Where-Object { $_ })
        if ($parts.Count -le 1) { $parent = "" }
        else { $parent = ($parts[0..($parts.Count - 2)] -join "/") }
        [void]$script:RsRefresh.Invoke($parent, $false)
    })

    $btnOpen.Add_Click({ & $script:RsOpenSelected })
    $list.Add_ItemActivate({ & $script:RsOpenSelected })

    $btnAddSelected.Add_Click({
        $st = $script:RsBrowserState
        $picked = New-Object System.Collections.Generic.List[string]
        $selFolders = @()
        $selModels = 0

        $uiFolder = [string]$st.CurrentFolder
        if ([string]::IsNullOrWhiteSpace($uiFolder)) {
            $uiFolder = [string]$st.PathBox.Text.Trim().Trim("/\").Replace("\", "/")
        }

        foreach ($it in @($st.List.SelectedItems)) {
            $meta = $it.Tag
            if (-not $meta) { continue }
            $kind = [string]$meta['Kind']

            if ($kind -eq 'model') {
                $selModels++
                $fullRsn = [string]$meta['FullRsn']
                if ([string]::IsNullOrWhiteSpace($fullRsn)) {
                    $modelFolder = [string]$meta['Folder']
                    if ([string]::IsNullOrWhiteSpace($modelFolder)) { $modelFolder = $uiFolder }
                    $fullRsn = ConvertTo-RsnPath -Server ([string]$st.Server) -FolderPath $modelFolder -ModelName ([string]$meta['Name'])
                }
                if (-not [string]::IsNullOrWhiteSpace($uiFolder)) {
                    $expectedPrefix = ("RSN://{0}/{1}/" -f [string]$st.Server, $uiFolder)
                    if (-not $fullRsn.StartsWith($expectedPrefix, [System.StringComparison]::OrdinalIgnoreCase)) {
                        $fullRsn = ConvertTo-RsnPath -Server ([string]$st.Server) -FolderPath $uiFolder -ModelName ([string]$meta['Name'])
                    }
                }
                if (-not [string]::IsNullOrWhiteSpace($fullRsn)) {
                    Write-Host ("ADD selected model: {0}" -f $fullRsn)
                    if (-not $picked.Contains($fullRsn)) { [void]$picked.Add($fullRsn) }
                }
                continue
            }

            if ($kind -eq 'folder') {
                $name = [string]$meta['Name']
                if ([string]::IsNullOrWhiteSpace($uiFolder)) {
                    $folderPath = $name
                }
                else {
                    $folderPath = ($uiFolder.Trim("/") + "/" + $name)
                }
                $selFolders += $folderPath
            }
        }

        if ($selFolders.Count -gt 0) {
            $st.Form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            $st.Status.Text = ("Collecting models from {0} selected folder(s)..." -f $selFolders.Count)
            [System.Windows.Forms.Application]::DoEvents()
            foreach ($folderPath in $selFolders) {
                Write-Host ("ADD selected folder (recursive): {0}" -f $folderPath)
                $fromFolder = @(Get-RevitServerModelsRecursive -Server ([string]$st.Server) -Folder $folderPath -Year ([string]$st.Year))
                foreach ($p in $fromFolder) {
                    if ($p -and -not $picked.Contains($p)) { [void]$picked.Add($p) }
                }
            }
            $st.Form.Cursor = [System.Windows.Forms.Cursors]::Default
        }

        if ($picked.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show(
                "Select .rvt model(s) and/or Folder row(s) (Ctrl+click), then Add selected.`n`n" +
                "Highlighted folders only are collected (not siblings like АР/БФ).`n" +
                "To take the entire current path, use 'Add ALL under current path'.",
                "Revit Server"
            ) | Out-Null
            return
        }

        Write-Host ("Selected rows -> models={0} folders={1} total RSN={2}" -f $selModels, $selFolders.Count, $picked.Count)
        $st.Selected = @($picked)
        $st.Form.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $st.Form.Close()
    })

    $btnAddFolder.Add_Click({
        $st = $script:RsBrowserState
        $folderForRecursive = [string]$st.CurrentFolder
        if ([string]::IsNullOrWhiteSpace($folderForRecursive)) {
            $folderForRecursive = [string]$st.PathBox.Text.Trim().Trim("/\").Replace("\", "/")
        }
        if ([string]::IsNullOrWhiteSpace($folderForRecursive)) {
            [System.Windows.Forms.MessageBox]::Show(
                "Current path is empty. Open the target folder (Open / Go) first.",
                "Revit Server"
            ) | Out-Null
            return
        }

        # If user highlighted folders, do NOT silently take siblings — offer the safe action.
        $highlightedFolders = @()
        foreach ($it in @($st.List.SelectedItems)) {
            $meta = $it.Tag
            if ($meta -and [string]$meta['Kind'] -eq 'folder') {
                $name = [string]$meta['Name']
                $highlightedFolders += ($folderForRecursive.Trim("/") + "/" + $name)
            }
        }
        if ($highlightedFolders.Count -gt 0) {
            $ans = [System.Windows.Forms.MessageBox]::Show(
                ("You highlighted {0} folder(s).`n`n" +
                 "YES = only those folders`n" +
                 "NO  = ALL models under current path (including АР/БФ/...):`n{1}`n`n" +
                 "Cancel = abort") -f $highlightedFolders.Count, $folderForRecursive,
                "Revit Server",
                [System.Windows.Forms.MessageBoxButtons]::YesNoCancel,
                [System.Windows.Forms.MessageBoxIcon]::Question
            )
            if ($ans -eq [System.Windows.Forms.DialogResult]::Cancel) { return }
            if ($ans -eq [System.Windows.Forms.DialogResult]::Yes) {
                $st.Form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
                $picked = New-Object System.Collections.Generic.List[string]
                foreach ($fp in $highlightedFolders) {
                    Write-Host ("ADD highlighted folder: {0}" -f $fp)
                    foreach ($p in @(Get-RevitServerModelsRecursive -Server ([string]$st.Server) -Folder $fp -Year ([string]$st.Year))) {
                        if ($p -and -not $picked.Contains($p)) { [void]$picked.Add($p) }
                    }
                }
                $st.Form.Cursor = [System.Windows.Forms.Cursors]::Default
                if ($picked.Count -eq 0) {
                    [System.Windows.Forms.MessageBox]::Show("No models found in highlighted folders.", "Revit Server") | Out-Null
                    return
                }
                $st.Selected = @($picked)
                $st.Form.DialogResult = [System.Windows.Forms.DialogResult]::OK
                $st.Form.Close()
                return
            }
            # NO -> fall through to full current path
        }
        else {
            $ans = [System.Windows.Forms.MessageBox]::Show(
                ("Add ALL models under:`n{0}`n`n(This ignores which rows are highlighted.)" -f $folderForRecursive),
                "Revit Server",
                [System.Windows.Forms.MessageBoxButtons]::OKCancel,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            if ($ans -ne [System.Windows.Forms.DialogResult]::OK) { return }
        }

        $st.Form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        $st.Status.Text = "Collecting ALL models under current path..."
        [System.Windows.Forms.Application]::DoEvents()
        Write-Host ("ADD ALL under: {0}" -f $folderForRecursive)
        $picked = @(Get-RevitServerModelsRecursive -Server ([string]$st.Server) -Folder $folderForRecursive -Year ([string]$st.Year))
        $st.Form.Cursor = [System.Windows.Forms.Cursors]::Default
        if ($picked.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("No models found in this folder.", "Revit Server") | Out-Null
            [void]$script:RsRefresh.Invoke([string]$folderForRecursive, $false)
            return
        }
        $st.Selected = $picked
        $st.Form.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $st.Form.Close()
    })

    $form.Controls.AddRange(@(
        $lblPath, $txtPath, $btnGo, $btnUp, $btnOpen, $list, $lblStatus,
        $btnAddSelected, $btnAddFolder, $btnCancel
    ))
    $form.CancelButton = $btnCancel
    $form.Add_Shown({ [void]$script:RsRefresh.Invoke("", $false) })

    $null = $form.ShowDialog()
    if ($form.DialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
        return @($script:RsBrowserState.Selected)
    }
    return @()
}

function Select-ServerHost {
    Ensure-ServersFile
    $servers = @(Get-Servers)

    Write-Host ""
    Write-Host "Revit Server:"
    for ($i = 0; $i -lt $servers.Count; $i++) {
        $s = $servers[$i]
        Write-Host ("  {0} - {1}  ({2}, year {3})" -f ($i + 1), $s.Label, $s.Host, $s.Year)
    }
    Write-Host ("  {0} - Add new server" -f ($servers.Count + 1))
    Write-Host "  0 - Back"
    Write-Host ""
    $sel = Read-Host "Number"
    if ($sel -eq "0") { return $null }

    $selNum = 0
    [void][int]::TryParse($sel, [ref]$selNum)
    if ($selNum -ge 1 -and $selNum -le $servers.Count) {
        return $servers[$selNum - 1]
    }
    elseif ($selNum -eq ($servers.Count + 1)) {
        $newHost = Read-Host "Host name or IP"
        if ([string]::IsNullOrWhiteSpace($newHost)) { return $null }
        $newLabel = Read-Host "Display name (optional)"
        $newYear = Read-Host "REST year (e.g. 2024)"
        if ([string]::IsNullOrWhiteSpace($newYear)) { $newYear = "2024" }
        if ([string]::IsNullOrWhiteSpace($newLabel)) { $newLabel = $newHost.Trim() }
        Save-Server -HostName $newHost.Trim() -Label $newLabel.Trim() -Year $newYear.Trim()
        return [pscustomobject]@{
            Host  = $newHost.Trim()
            Label = $newLabel.Trim()
            Year  = $newYear.Trim()
        }
    }
    Write-Host "Invalid choice."
    return $null
}

function Ensure-RbpRsnPatch {
    # Auto-apply RBP Scripts patch so RSN:// is accepted by Batch Processor.
    $patchPs = Join-Path $ToolDir "служебное\rbp_rsn_patch\apply_rsn_support.ps1"
    if (-not (Test-Path -LiteralPath $patchPs)) {
        $patchPs = Join-Path $ToolDir "rbp_rsn_patch\apply_rsn_support.ps1"
    }
    if (-not (Test-Path -LiteralPath $patchPs)) {
        $toolsRoot = Split-Path -Parent $ToolDir
        $patchPs = Join-Path $toolsRoot "batch_set_project_units\rbp_rsn_patch\apply_rsn_support.ps1"
    }
    if (-not (Test-Path -LiteralPath $patchPs)) {
        Write-Host "WARN: RSN support script missing."
        return $false
    }

    Write-Host ""
    Write-Host "Ensuring RBP supports RSN:// (built-in PowerShell patch)..."
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $patchPs
    $code = $LASTEXITCODE
    if ($code -eq 0) {
        Write-Host "RBP RSN support: OK"
        return $true
    }
    if ($code -eq 2) {
        Write-Host "WARN: RBP Scripts not found on this PC."
        Write-Host "      Keep Revit Batch Processor open, then select Server again."
        return $false
    }
    Write-Host ("WARN: RSN support finished with code {0}." -f $code)
    return $false
}

function Select-FromRevitServer {
    $serverObj = Select-ServerHost
    if (-not $serverObj) { return @() }

    Write-Host ("Opening browser for {0} ({1})..." -f $serverObj.Label, $serverObj.Host)
    $selected = @(Show-RevitServerBrowser -ServerObj $serverObj)

    if ($selected.Count -gt 0) {
        try {
            $lastLines = @(
                "server=$($serverObj.Host)"
                "label=$($serverObj.Label)"
                "year=$($serverObj.Year)"
                "count=$($selected.Count)"
                "updated=$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
            ) + $selected
            Write-TextFileLines -Path $LastCfg -Lines $lastLines
        }
        catch {
            Write-Host ("WARN: could not save last server selection: {0}" -f $_.Exception.Message)
        }
    }

    return $selected
}

function Show-FolderPathDialog {
    # Same Explorer-style dialog as file Open (address bar, tree, UNC).
    # Trick: ValidateNames/CheckFileExists off + dummy FileName = pick folder.
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Title = "Папка с моделями — вставьте путь в адресную строку и нажмите «Открыть»"
    $dlg.Filter = "Folders|*.nevermatch|All files (*.*)|*.*"
    $dlg.FilterIndex = 1
    $dlg.CheckFileExists = $false
    $dlg.CheckPathExists = $false
    $dlg.ValidateNames = $false
    $dlg.Multiselect = $false
    $dlg.DereferenceLinks = $true
    $dlg.FileName = "Select this folder"
    $dlg.ValidateNames = $false
    $dlg.Multiselect = $false
    $dlg.DereferenceLinks = $true
    $dlg.FileName = "Select this folder"
    $dlg.CheckPathExists = $false

    if (Test-Path -LiteralPath $PathCfg) {
        $prev = (
            Get-Content -LiteralPath $PathCfg -ErrorAction SilentlyContinue |
                Where-Object { $_ -and -not $_.StartsWith("#") -and -not $_.StartsWith("FILES:") } |
                Select-Object -First 1
        )
        if ($prev) {
            $prev = Get-NativePath $prev.Trim()
            if ($prev -and (Test-Path -LiteralPath $prev -PathType Container)) {
                try { $dlg.InitialDirectory = $prev } catch { }
            }
        }
    }
    if ([string]::IsNullOrWhiteSpace($dlg.InitialDirectory)) {
        try { $dlg.InitialDirectory = $ToolDir } catch { }
    }

    if (-not $CompactUiLibrary) {
        Write-Host "A folder window should appear on top. If not, look behind this console."
        Write-Host "Paste the UNC path into the address bar, then Open."
    }
    if ((Show-OwnedDialog $dlg) -ne [System.Windows.Forms.DialogResult]::OK) {
        if ($CompactUiLibrary) { return $null }
        $pasted = Read-Host "No window / Cancel. Paste parent folder path, or Enter to abort"
        if ([string]::IsNullOrWhiteSpace($pasted)) { return $null }
        return (Get-NativePath $pasted.Trim().Trim('"'))
    }

    $picked = $dlg.FileName
    if ([string]::IsNullOrWhiteSpace($picked)) { return $null }

    if (Test-Path -LiteralPath $picked -PathType Leaf) {
        return (Split-Path -Parent $picked)
    }
    if (Test-Path -LiteralPath $picked -PathType Container) {
        return $picked
    }

    $dir = Split-Path -Parent $picked
    if ($dir -and (Test-Path -LiteralPath $dir -PathType Container)) {
        return $dir
    }
    return $null
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
    $rezerv = Get-U @(0x0420, 0x0435, 0x0437, 0x0435, 0x0440, 0x0432) # Резерв
    $sem = Get-U @(0x0421, 0x0435, 0x043C, 0x0435, 0x0439, 0x0441, 0x0442, 0x0432, 0x0430) # Семейства
    if ($p -match ("(?i)\\" + [regex]::Escape($rezerv) + "\\")) { return $true }
    if ($p -match ("(?i)\\" + [regex]::Escape($sem) + "\\")) { return $true }
    return $false
}

function Test-IsTargetRvtName([string]$Name) {
    # Any .rvt in RVT\. Skip .0001.rvt and *_backup.
    if ([string]::IsNullOrWhiteSpace($Name)) { return $false }
    if ($Name -match '\.\d{4}\.rvt$') { return $false }
    if ($Name -match '(?i)_backup') { return $false }
    return ($Name -match '(?i)\.rvt$')
}

function Get-TargetRvtsFromDisciplineFolders {
    param([string[]]$Folders)

    $found = New-Object System.Collections.Generic.List[string]
    foreach ($folder in @($Folders)) {
        if ([string]::IsNullOrWhiteSpace($folder)) { continue }
        if (-not (Test-Path -LiteralPath $folder -PathType Container)) { continue }
        if (Test-ExcludedModelPath $folder) { continue }

        $rvtDirs = @(
            Get-ChildItem -LiteralPath $folder -Recurse -Directory -Force -Depth 8 -ErrorAction SilentlyContinue |
                Where-Object {
                    ($_.Name -ieq "RVT") -and
                    (-not (Test-ExcludedModelPath $_.FullName))
                }
        )
        foreach ($rvtDir in $rvtDirs) {
            $files = @(
                Get-ChildItem -LiteralPath $rvtDir.FullName -File -Force -ErrorAction SilentlyContinue |
                    Where-Object {
                        ($_.Extension -ieq ".rvt") -and
                        (Test-IsTargetRvtName $_.Name)
                    }
            )
            foreach ($f in $files) {
                $native = Get-NativePath $f.FullName
                if ($native -and -not (Test-ExcludedModelPath $native)) {
                    [void]$found.Add($native)
                }
            }
        }
    }
    return @($found | Sort-Object -Unique)
}

function Show-DisciplineFolderChecklist {
    param([string]$ParentFolder)

    $children = @(
        Get-ChildItem -LiteralPath $ParentFolder -Directory -Force -ErrorAction SilentlyContinue |
            Sort-Object Name
    )
    if ($children.Count -eq 0) {
        Write-Host ("No subfolders in: {0}" -f $ParentFolder)
        return @()
    }

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Select discipline folders (Ctrl+click / Space)"
    $form.Size = New-Object System.Drawing.Size(760, 560)
    $form.StartPosition = "CenterScreen"
    $form.TopMost = $true
    $form.MinimizeBox = $false
    $form.MaximizeBox = $true
    $form.MinimumSize = New-Object System.Drawing.Size(560, 400)

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Location = New-Object System.Drawing.Point(12, 10)
    $lbl.Size = New-Object System.Drawing.Size(720, 40)
    $lbl.Anchor = "Top,Left,Right"
    $lbl.Text = ("Parent: {0}{1}Check folders with models (e.g. 5_2_VS, 9_PT). Script finds any .rvt in RVT\ (no *_R## filter)." -f $ParentFolder, [Environment]::NewLine)

    $list = New-Object System.Windows.Forms.CheckedListBox
    $list.Location = New-Object System.Drawing.Point(12, 58)
    $list.Size = New-Object System.Drawing.Size(720, 390)
    $list.Anchor = "Top,Bottom,Left,Right"
    $list.CheckOnClick = $true
    $list.IntegralHeight = $false

    foreach ($d in $children) {
        [void]$list.Items.Add($d.Name, $false)
    }

    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Location = New-Object System.Drawing.Point(12, 456)
    $lblStatus.Size = New-Object System.Drawing.Size(420, 22)
    $lblStatus.Anchor = "Bottom,Left"
    $lblStatus.Text = ("Folders: {0}" -f $children.Count)

    $btnAll = New-Object System.Windows.Forms.Button
    $btnAll.Text = "Check all"
    $btnAll.Location = New-Object System.Drawing.Point(12, 484)
    $btnAll.Size = New-Object System.Drawing.Size(100, 30)
    $btnAll.Anchor = "Bottom,Left"
    $btnAll.Add_Click({
        for ($i = 0; $i -lt $list.Items.Count; $i++) { $list.SetItemChecked($i, $true) }
    })

    $btnNone = New-Object System.Windows.Forms.Button
    $btnNone.Text = "Uncheck all"
    $btnNone.Location = New-Object System.Drawing.Point(120, 484)
    $btnNone.Size = New-Object System.Drawing.Size(100, 30)
    $btnNone.Anchor = "Bottom,Left"
    $btnNone.Add_Click({
        for ($i = 0; $i -lt $list.Items.Count; $i++) { $list.SetItemChecked($i, $false) }
    })

    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = "OK"
    $btnOk.Location = New-Object System.Drawing.Point(532, 484)
    $btnOk.Size = New-Object System.Drawing.Size(90, 30)
    $btnOk.Anchor = "Bottom,Right"
    $btnOk.Add_Click({
        $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.Close()
    })

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel"
    $btnCancel.Location = New-Object System.Drawing.Point(630, 484)
    $btnCancel.Size = New-Object System.Drawing.Size(90, 30)
    $btnCancel.Anchor = "Bottom,Right"
    $btnCancel.Add_Click({
        $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $form.Close()
    })

    $form.AcceptButton = $btnOk
    $form.CancelButton = $btnCancel
    $form.Controls.AddRange(@($lbl, $list, $lblStatus, $btnAll, $btnNone, $btnOk, $btnCancel))
    $null = $form.ShowDialog()
    if ($form.DialogResult -ne [System.Windows.Forms.DialogResult]::OK) {
        return @()
    }

    $selected = @()
    for ($i = 0; $i -lt $list.Items.Count; $i++) {
        if ($list.GetItemChecked($i)) {
            $selected += (Join-Path $ParentFolder ([string]$list.Items[$i]))
        }
    }
    return @($selected)
}

function Select-LocalFolder {
    Write-Host ""
    Write-Host "Step 1/2: open PARENT folder (e.g. ...\4_Razrabotka proekta)"
    $typed = Show-FolderPathDialog
    if ([string]::IsNullOrWhiteSpace($typed)) {
        return @()
    }

    $root = Get-NativePath $typed
    if (-not (Test-Path -LiteralPath $root -PathType Container)) {
        Write-Host ("Folder not found: {0}" -f $root)
        return @()
    }

    Write-CfgContent -Path $PathCfg -Value $root
    Write-Host ("Parent: {0}" -f $root)

    Write-Host "Step 2/2: check discipline folders (Ctrl / Space), then OK"
    $disciplineFolders = @(Show-DisciplineFolderChecklist -ParentFolder $root)
    if ($disciplineFolders.Count -eq 0) {
        Write-Host "No discipline folders selected."
        return @()
    }

    Write-Host ("Scanning {0} folder(s) for RVT\*.rvt ..." -f $disciplineFolders.Count)
    foreach ($d in $disciplineFolders) {
        Write-Host ("  - {0}" -f (Split-Path -Leaf $d))
    }

    $files = @(Get-TargetRvtsFromDisciplineFolders -Folders $disciplineFolders)
    Write-Host ("Found {0} target model(s)." -f $files.Count)
    if ($files.Count -gt 0) {
        Write-Host ("Example: {0}" -f $files[0])
    }
    return $files
}

function Select-LocalFiles {
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Title = "Модели Revit (*.rvt) — можно вставить путь в адресную строку"
    $dlg.Filter = "Revit (*.rvt)|*.rvt|All files (*.*)|*.*"
    $dlg.Multiselect = $true
    $dlg.CheckFileExists = $true
    try { $dlg.InitialDirectory = Get-LocalDialogFolder } catch { }

    if ((Show-OwnedDialog $dlg) -ne [System.Windows.Forms.DialogResult]::OK) {
        return @()
    }

    $files = @($dlg.FileNames | ForEach-Object { Get-NativePath $_ })
    $kept = @()
    $skipped = @()
    foreach ($f in $files) {
        if (Test-ExcludedModelPath $f) {
            $skipped += $f
        }
        else {
            $kept += $f
        }
    }
    if ($skipped.Count -gt 0) {
        Write-Host ("Skipped {0} backup/copy path(s) (Резерв / Backup):" -f $skipped.Count)
        foreach ($s in $skipped) {
            Write-Host ("  skip: {0}" -f $s)
        }
    }
    $cfgText = "FILES:" + [Environment]::NewLine + ($kept -join [Environment]::NewLine)
    Write-CfgContent -Path $PathCfg -Value $cfgText
    return $kept
}

# ---------------- console menu (skipped when dotted into the GUI) ----------------
if (-not $CompactUiLibrary) {

Write-Host ""
Write-Host "============================================"
Write-Host "  Select Revit models -> rvt_list.txt"
Write-Host "  Local disk / UNC  +  Revit Server (RSN)"
Write-Host "============================================"
Write-Host ""

$existing = @(Read-ExistingList)
if ($existing.Count -gt 0) {
    Write-Host ("Current list already has {0} model(s)." -f $existing.Count)
    Write-Host ""
}

Write-Host "Source:"
Write-Host "  1 - Parent folder -> pick discipline folders (Ctrl) -> RVT\*.rvt"
Write-Host "  2 - Individual files on this PC / share"
Write-Host "  3 - Revit Server (RSN://)"
Write-Host "  4 - Show current rvt_list.txt"
Write-Host "  5 - Clear rvt_list.txt"
Write-Host "  0 - Exit"
Write-Host ""
$choice = Read-Host "Number"

if ($choice -eq "0") { exit 0 }

if ($choice -eq "4") {
    if ($existing.Count -eq 0) {
        Write-Host "List is empty."
    }
    else {
        Write-Host ""
        for ($i = 0; $i -lt $existing.Count; $i++) {
            Write-Host ("  {0}. {1}" -f ($i + 1), $existing[$i])
        }
    }
    exit 0
}

if ($choice -eq "5") {
    if (Test-Path -LiteralPath $OutList) {
        Remove-Item -LiteralPath $OutList -Force
    }
    Write-Host "List cleared."
    exit 0
}

$picked = @()
if ($choice -eq "1") {
    $picked = @(Select-LocalFolder)
}
elseif ($choice -eq "2") {
    $picked = @(Select-LocalFiles)
}
elseif ($choice -eq "3") {
    $picked = @(Select-FromRevitServer)
}
else {
    Write-Host "Invalid choice."
    exit 1
}

if ($picked.Count -eq 0) {
    Write-Host "Nothing selected."
    exit 0
}

$doAppend = $false
if ($existing.Count -gt 0) {
    Write-Host ""
    Write-Host "List already has models. What to do?"
    Write-Host "  1 - Replace list"
    Write-Host "  2 - Add to list (append)"
    Write-Host "  0 - Cancel"
    Write-Host ""
    $mode = Read-Host "Number"
    if ($mode -eq "0") { exit 0 }
    elseif ($mode -eq "2") { $doAppend = $true }
    elseif ($mode -ne "1") {
        Write-Host "Invalid choice."
        exit 1
    }
}

if ($doAppend) {
    if (-not (Write-RvtList -Paths $picked -Append)) { exit 1 }
}
else {
    if (-not (Write-RvtList -Paths $picked)) { exit 1 }
}

$finalList = @(Read-ExistingList)
$hasRsn = @($finalList | Where-Object { $_ -match '^(?i)RSN://' }).Count -gt 0
if ($hasRsn) {
    [void](Ensure-RbpRsnPatch)
}

Write-Host ""
Write-Host "Next:"
Write-Host "  Сжатие.cmd  (окно: список и запуск)"
Write-Host "  Or open RBP manually:"
Write-Host "       Task script = compact_save.py (on share)"
Write-Host ("       File list   = {0}" -f $OutList)
Write-Host "       Create New Local, Detach = OFF"
Write-Host ""
if ($hasRsn) {
    Write-Host "IMPORTANT (list has RSN://):"
    Write-Host "  Central processing = Create New Local"
    Write-Host "  Detach from Central = OFF"
    Write-Host "  Revit version = model year (*_R22 -> 2022)"
    Write-Host "  Support for RSN:// is applied automatically when you start compact."
}
else {
    Write-Host "IMPORTANT if shared centrals:"
    Write-Host "  Central processing = Create New Local, Detach = OFF"
}
Write-Host ""

}
