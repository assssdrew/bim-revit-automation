# Unified model picker for Revit Batch Processor (same UX as other batch_* toolkits).
# Sources: local folder / local files / Revit Server (RSN://).
# Writes rvt_list.txt (default) — for Pass 1 use OLD ...IOS... paths.
# Any .rvt in RVT\ (no *_R## required). Skip backups.
# Encoding: ASCII-friendly messages for Windows PowerShell 5.1.

param(
    [string]$OutFile = "rvt_list.txt"
)

Add-Type -AssemblyName System.Windows.Forms

$ToolDir = Split-Path -Parent $MyInvocation.MyCommand.Path
if ([string]::IsNullOrWhiteSpace($OutFile)) { $OutFile = "rvt_list.txt" }
$OutList = Join-Path $ToolDir $OutFile
$IsMoveList = ($OutFile -ieq "rvt_list_move.txt")
$PathCfg = Join-Path $ToolDir "models_path.cfg"
$ServersCfg = Join-Path $ToolDir "servers.cfg"
$LastCfg = Join-Path $ToolDir "last_selection.cfg"
$MakeList = Join-Path $ToolDir "make_rvt_list.ps1"

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
    $lines | Set-Content -LiteralPath $ServersCfg -Encoding UTF8
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

    $unique | Set-Content -LiteralPath $OutList -Encoding UTF8
    Write-Host ""
    Write-Host ("Models in list: {0}" -f $unique.Count)
    Write-Host ("List: {0}" -f $OutList)
    $localN = @($unique | Where-Object { $_ -notmatch '^(?i)RSN://' }).Count
    $rsnN = @($unique | Where-Object { $_ -match '^(?i)RSN://' }).Count
    Write-Host ("  Local/UNC: {0}" -f $localN)
    Write-Host ("  Revit Server (RSN): {0}" -f $rsnN)
    Write-Host ("Example: {0}" -f $unique[0])
    return $true
}

function Invoke-RevitServerGet {
    param([string]$Url)

    $request = [System.Net.HttpWebRequest]::Create($Url)
    $request.Method = "GET"
    $request.Timeout = 60000
    $request.ReadWriteTimeout = 60000
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
            $lastError = $_.Exception.Message
            $resp = $null
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
    $btnAddSelected.Text = "Add selected models"
    $btnAddSelected.Location = New-Object System.Drawing.Point(12, 545)
    $btnAddSelected.Size = New-Object System.Drawing.Size(180, 32)

    $btnAddFolder = New-Object System.Windows.Forms.Button
    $btnAddFolder.Text = "Add ALL .rvt in this folder"
    $btnAddFolder.Location = New-Object System.Drawing.Point(200, 545)
    $btnAddFolder.Size = New-Object System.Drawing.Size(280, 32)

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
            $st.Status.Text = ("REST error: {0}" -f $contents.Error)
            $st.Form.Cursor = [System.Windows.Forms.Cursors]::Default
            if ($ShowError) {
                [System.Windows.Forms.MessageBox]::Show(
                    ("Cannot open folder '{0}'." -f $requestFolder) + [Environment]::NewLine + [Environment]::NewLine + $contents.Error,
                    "Revit Server"
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
        $picked = @()
        foreach ($it in $st.List.SelectedItems) {
            $meta = $it.Tag
            if (-not $meta) { continue }
            if ([string]$meta['Kind'] -ne 'model') { continue }

            $fullRsn = [string]$meta['FullRsn']
            if ([string]::IsNullOrWhiteSpace($fullRsn)) {
                $modelFolder = [string]$meta['Folder']
                if ([string]::IsNullOrWhiteSpace($modelFolder)) { $modelFolder = [string]$st.CurrentFolder }
                if ([string]::IsNullOrWhiteSpace($modelFolder)) {
                    $modelFolder = [string]$st.PathBox.Text.Trim().Trim("/\").Replace("\", "/")
                }
                $fullRsn = ConvertTo-RsnPath -Server ([string]$st.Server) -FolderPath $modelFolder -ModelName ([string]$meta['Name'])
            }

            # Guard: never allow root-only RSN when UI shows a nested folder
            $uiFolder = [string]$st.CurrentFolder
            if ([string]::IsNullOrWhiteSpace($uiFolder)) {
                $uiFolder = [string]$st.PathBox.Text.Trim().Trim("/\").Replace("\", "/")
            }
            if (-not [string]::IsNullOrWhiteSpace($uiFolder)) {
                $expectedPrefix = ("RSN://{0}/{1}/" -f [string]$st.Server, $uiFolder)
                if (-not $fullRsn.StartsWith($expectedPrefix, [System.StringComparison]::OrdinalIgnoreCase)) {
                    $fullRsn = ConvertTo-RsnPath -Server ([string]$st.Server) -FolderPath $uiFolder -ModelName ([string]$meta['Name'])
                }
            }

            if ($fullRsn -match '(?i)^RSN://[^/]+/[^/]+\.rvt$' -and -not [string]::IsNullOrWhiteSpace($uiFolder)) {
                [System.Windows.Forms.MessageBox]::Show(
                    ("Ошибка пути (без папок):{0}{1}{0}Ожидалась папка: {2}" -f [Environment]::NewLine, $fullRsn, $uiFolder),
                    "Revit Server"
                ) | Out-Null
                return
            }

            Write-Host ("ADD selected: {0}" -f $fullRsn)
            $picked += $fullRsn
        }
        if ($picked.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show(
                "Select .rvt models (Ctrl+click), or open a Folder with Open.",
                "Revit Server"
            ) | Out-Null
            return
        }
        $st.Selected = $picked
        $st.Form.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $st.Form.Close()
    })

    $btnAddFolder.Add_Click({
        $st = $script:RsBrowserState
        $folderHere = [string]$st.CurrentFolder
        if ([string]::IsNullOrWhiteSpace($folderHere)) {
            $folderHere = [string]$st.PathBox.Text.Trim().Trim("/\").Replace("\", "/")
        }
        if ([string]::IsNullOrWhiteSpace($folderHere)) {
            [System.Windows.Forms.MessageBox]::Show(
                "Open the target folder first (Open / Go), then Add ALL .rvt.",
                "Revit Server"
            ) | Out-Null
            return
        }

        # Same idea as local *.rvt in one folder: take every model listed here.
        $picked = @()
        foreach ($it in $st.List.Items) {
            $meta = $it.Tag
            if (-not $meta) { continue }
            if ([string]$meta['Kind'] -ne 'model') { continue }
            $fullRsn = [string]$meta['FullRsn']
            if ([string]::IsNullOrWhiteSpace($fullRsn)) {
                $fullRsn = ConvertTo-RsnPath -Server ([string]$st.Server) -FolderPath $folderHere -ModelName ([string]$meta['Name'])
            }
            if (-not [string]::IsNullOrWhiteSpace($fullRsn)) {
                $picked += $fullRsn
            }
        }
        $picked = @($picked | Select-Object -Unique)

        # Parent folder with only subfolders (e.g. Стадия ПП) -> recurse like local parent scan
        if ($picked.Count -eq 0) {
            $st.Form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            $st.Status.Text = "Collecting models in folder (including subfolders)..."
            [System.Windows.Forms.Application]::DoEvents()
            $picked = @(Get-RevitServerModelsRecursive -Server ([string]$st.Server) -Folder $folderHere -Year ([string]$st.Year))
            $st.Form.Cursor = [System.Windows.Forms.Cursors]::Default
        }

        if ($picked.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("No models found in this folder.", "Revit Server") | Out-Null
            [void]$script:RsRefresh.Invoke([string]$folderHere, $false)
            return
        }
        Write-Host ("ADD folder ({0}): {1} model(s)" -f $folderHere, $picked.Count)
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
    # Local/UNC picking does not need this; called only for Server (or when list has RSN).
    $patchPs = Join-Path $ToolDir "rbp_rsn_patch\apply_rsn_support.ps1"
    if (-not (Test-Path -LiteralPath $patchPs)) {
        $toolsRoot = Split-Path -Parent $ToolDir
        $patchPs = Join-Path $toolsRoot "batch_set_project_units\rbp_rsn_patch\apply_rsn_support.ps1"
    }
    if (-not (Test-Path -LiteralPath $patchPs)) {
        Write-Host "WARN: RSN patch script missing: rbp_rsn_patch\apply_rsn_support.ps1"
        return $false
    }

    Write-Host ""
    Write-Host "Ensuring RBP supports RSN:// (built-in PowerShell patch)..."
    # Do not use -Quiet: we need the real error text if patch fails.
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $patchPs
    $code = $LASTEXITCODE
    if ($code -eq 0) {
        Write-Host "RBP RSN patch: OK"
        return $true
    }
    if ($code -eq 2) {
        Write-Host "WARN: RBP Scripts not found on this PC."
        Write-Host "      Keep Revit Batch Processor open, then select Server again."
        return $false
    }
    Write-Host ("WARN: RSN patch finished with code {0}." -f $code)
    Write-Host "      See messages above / rbp_rsn_patch\README.md"
    return $false
}

function Ensure-RbpOpenFailPatch {
    # Stock RBP cancels Open on Cyrillic Revit warnings. Apply on every list write.
    $toolsRoot = Split-Path -Parent $ToolDir
    $candidates = @(
        (Join-Path $toolsRoot "rbp_patches\apply_open_fail.ps1"),
        (Join-Path $ToolDir "patch_rbp_open_failures.ps1"),
        (Join-Path $toolsRoot "batch_model_ops\patch_rbp_open_failures.ps1")
    )
    $patchPs = $null
    foreach ($c in $candidates) {
        if (Test-Path -LiteralPath $c -PathType Leaf) { $patchPs = $c; break }
    }
    if (-not $patchPs) {
        Write-Host "WARN: RBP open-fail patch script missing (tools\rbp_patches)."
        return $false
    }
    Write-Host ""
    Write-Host "Ensuring RBP does not cancel Open on Cyrillic warnings..."
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $patchPs
    $code = $LASTEXITCODE
    if ($code -eq 0) {
        Write-Host "RBP open-fail patch: OK"
        return $true
    }
    if ($code -eq 2) {
        Write-Host "WARN: RBP Scripts not found on this PC."
        Write-Host "      Keep Revit Batch Processor open, then pick models again."
        return $false
    }
    Write-Host ("WARN: open-fail patch finished with code {0}." -f $code)
    Write-Host "      See tools\rbp_patches\README.md"
    return $false
}

function Ensure-RbpProgressMonitor {
    $toolsRoot = Split-Path -Parent $ToolDir
    $ensurePs = Join-Path $toolsRoot "rbp_progress_monitor\ensure_running.ps1"
    if (-not (Test-Path -LiteralPath $ensurePs -PathType Leaf)) {
        Write-Host "WARN: rbp_progress_monitor missing (optional)."
        return $false
    }
    Write-Host ""
    Write-Host "RBP progress monitor: tray (window opens when BatchRvt starts)..."
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $ensurePs
    return ($LASTEXITCODE -eq 0)
}

function Select-FromRevitServer {
    # Local/UNC (menu 1-2) unchanged. Server needs RBP patch for File exists check.
    [void](Ensure-RbpRsnPatch)

    $serverObj = Select-ServerHost
    if (-not $serverObj) { return @() }

    Write-Host ("Opening browser for {0} ({1})..." -f $serverObj.Label, $serverObj.Host)
    $selected = @(Show-RevitServerBrowser -ServerObj $serverObj)

    if ($selected.Count -gt 0) {
        @(
            "server=$($serverObj.Host)"
            "label=$($serverObj.Label)"
            "year=$($serverObj.Year)"
            "count=$($selected.Count)"
            "updated=$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        ) + $selected | Set-Content -LiteralPath $LastCfg -Encoding UTF8
    }

    return $selected
}

function Show-FolderPathDialog {
    # Same Explorer-style dialog as file Open (address bar, tree, UNC).
    # Trick: ValidateNames/CheckFileExists off + dummy FileName = pick folder.
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Title = "Parent folder (e.g. Models): paste path, then Open"
    $dlg.Filter = "Folders|*.nevermatch|All files (*.*)|*.*"
    $dlg.FilterIndex = 1
    $dlg.CheckFileExists = $false
    $dlg.CheckPathExists = $true
    $dlg.ValidateNames = $false
    $dlg.Multiselect = $false
    $dlg.DereferenceLinks = $true
    $dlg.FileName = "Select this folder"

    if (Test-Path -LiteralPath $PathCfg) {
        $prev = (
            Get-Content -LiteralPath $PathCfg -ErrorAction SilentlyContinue |
                Where-Object { $_ -and -not $_.StartsWith("#") -and -not $_.StartsWith("FILES:") } |
                Select-Object -First 1
        )
        if ($prev) {
            $prev = Get-NativePath $prev.Trim()
            if ($prev -and (Test-Path -LiteralPath $prev -PathType Container)) {
                $dlg.InitialDirectory = $prev
            }
        }
    }

    if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        return $null
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
    # ExampleProject (and others): any .rvt in RVT\. Skip .0001.rvt and *_backup.
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
            Get-ChildItem -LiteralPath $folder -Recurse -Directory -Force -ErrorAction SilentlyContinue |
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
    Write-Host "Step 1/2: open PARENT folder (e.g. ...\ExampleProject)"
    $typed = Show-FolderPathDialog
    if ([string]::IsNullOrWhiteSpace($typed)) {
        return @()
    }

    $root = Get-NativePath $typed
    if (-not (Test-Path -LiteralPath $root -PathType Container)) {
        Write-Host ("Folder not found: {0}" -f $root)
        return @()
    }

    Set-Content -LiteralPath $PathCfg -Value $root -Encoding UTF8
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
    $dlg.Title = "Select Revit model files (you can paste a path in the address bar)"
    $dlg.Filter = "Revit (*.rvt)|*.rvt|All files (*.*)|*.*"
    $dlg.Multiselect = $true
    $dlg.CheckFileExists = $true

    if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        return @()
    }

    $files = @($dlg.FileNames | ForEach-Object { Get-NativePath $_ })
    $cfgText = "FILES:" + [Environment]::NewLine + ($files -join [Environment]::NewLine)
    Set-Content -LiteralPath $PathCfg -Value $cfgText -Encoding UTF8
    return $files
}

function Test-LikelyUnwantedModel([string]$Path) {
    if ([string]::IsNullOrWhiteSpace($Path)) { return $true }
    $name = ""
    try { $name = [System.IO.Path]::GetFileName($Path) } catch { $name = $Path }
    $full = $Path.Replace("/", "\")
    if ($name -match '\.\d{4}\.rvt$') { return $true }
    if ($full -match '(?i)\\Backup\\') { return $true }
    $rezerv = Get-U @(0x0440, 0x0435, 0x0437, 0x0435, 0x0440, 0x0432)
    $kopiya = Get-U @(0x043A, 0x043E, 0x043F, 0x0438, 0x044F)
    $low = $name.ToLowerInvariant()
    if ($low -match 'backup|_bak|\.old|_copy') { return $true }
    if ($low.Contains($rezerv.ToLowerInvariant())) { return $true }
    if ($low.Contains($kopiya.ToLowerInvariant())) { return $true }
    return $false
}

function Get-ModelPickLabel([string]$Path) {
    if ([string]::IsNullOrWhiteSpace($Path)) { return "" }
    if ($Path -match '^(?i)RSN://') {
        $rest = $Path -replace '^(?i)RSN://', ''
        $parts = @($rest.Split(@('/'), [System.StringSplitOptions]::RemoveEmptyEntries))
        $file = if ($parts.Count -gt 0) { $parts[$parts.Count - 1] } else { $Path }
        $folder = if ($parts.Count -gt 1) { $parts[$parts.Count - 2] } else { "" }
        if ($folder) { return ("{0}    [{1}]" -f $file, $folder) }
        return $file
    }
    $file = [System.IO.Path]::GetFileName($Path)
    $folder = [System.IO.Path]::GetFileName([System.IO.Path]::GetDirectoryName($Path))
    if ($folder) { return ("{0}    [{1}]" -f $file, $folder) }
    return $file
}

function Show-ModelChecklist {
    param([string[]]$Paths)

    $src = @($Paths | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
    if ($src.Count -le 1) { return @($src) }

    # Plain Object[] — PS 5.1 @() / foreach on List[object] throws "Argument types do not match"
    $items = @(
        foreach ($p in $src) {
            $unwanted = Test-LikelyUnwantedModel $p
            [pscustomobject]@{
                Path     = $p
                Label    = Get-ModelPickLabel $p
                Checked  = (-not $unwanted)
                Unwanted = $unwanted
            }
        }
    )

    $form = New-Object System.Windows.Forms.Form
    $form.Text = Get-U @(0x0412, 0x044B, 0x0431, 0x043E, 0x0440, 0x0020, 0x043C, 0x043E, 0x0434, 0x0435, 0x043B, 0x0435, 0x0439, 0x0020, 0x0434, 0x043B, 0x044F, 0x0020, 0x0052, 0x0042, 0x0050)
    $form.Size = New-Object System.Drawing.Size(920, 640)
    $form.StartPosition = "CenterScreen"
    $form.MinimizeBox = $false
    $form.MaximizeBox = $true
    $form.MinimumSize = New-Object System.Drawing.Size(720, 480)

    $lblHint = New-Object System.Windows.Forms.Label
    $lblHint.Location = New-Object System.Drawing.Point(12, 10)
    $lblHint.Size = New-Object System.Drawing.Size(880, 36)
    $lblHint.Text = (
        (Get-U @(0x041E, 0x0442, 0x043C, 0x0435, 0x0442, 0x044C, 0x0442, 0x0435, 0x0020, 0x043B, 0x0438, 0x0448, 0x043D, 0x0435, 0x0435, 0x0020, 0x043D, 0x0435, 0x043D, 0x0443, 0x0436, 0x043D, 0x044B, 0x0435, 0x0020, 0x0028, 0x0440, 0x0435, 0x0437, 0x0435, 0x0440, 0x0432, 0x002C, 0x0020, 0x0062, 0x0061, 0x0063, 0x006B, 0x0075, 0x0070, 0x002C, 0x0020, 0x043A, 0x043E, 0x043F, 0x0438, 0x0438, 0x0029, 0x002E, 0x0020)) +
        (Get-U @(0x0424, 0x0438, 0x043B, 0x044C, 0x0442, 0x0440, 0x0020, 0x043F, 0x043E, 0x0020, 0x0438, 0x043C, 0x0435, 0x043D, 0x0438, 0x0020, 0x0438, 0x043B, 0x0438, 0x0020, 0x043F, 0x0430, 0x043F, 0x043A, 0x0435, 0x002E))
    )

    $txtFilter = New-Object System.Windows.Forms.TextBox
    $txtFilter.Location = New-Object System.Drawing.Point(12, 50)
    $txtFilter.Size = New-Object System.Drawing.Size(880, 24)

    $list = New-Object System.Windows.Forms.CheckedListBox
    $list.Location = New-Object System.Drawing.Point(12, 82)
    $list.Size = New-Object System.Drawing.Size(880, 430)
    $list.CheckOnClick = $true
    $list.IntegralHeight = $false
    $list.Anchor = "Top,Bottom,Left,Right"
    $txtFilter.Anchor = "Top,Left,Right"
    $lblHint.Anchor = "Top,Left,Right"

    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Location = New-Object System.Drawing.Point(12, 518)
    $lblStatus.Size = New-Object System.Drawing.Size(500, 22)
    $lblStatus.Anchor = "Bottom,Left"

    $script:ModelPickState = @{
        Items     = $items
        List      = $list
        Status    = $lblStatus
        Filter    = $txtFilter
        Visible   = @()
        Refreshing = $false
    }

    $script:ModelPickRefresh = {
        $st = $script:ModelPickState
        $filter = ([string]$st.Filter.Text).Trim().ToLowerInvariant()
        $st.Refreshing = $true
        $st.List.BeginUpdate()
        try {
            $st.List.Items.Clear()
            $vis = New-Object System.Collections.ArrayList
            foreach ($it in $st.Items) {
                $hay = (($it.Label + " " + $it.Path).ToLowerInvariant())
                if ($filter -and $hay.IndexOf($filter) -lt 0) { continue }
                $idx = $st.List.Items.Add($it.Label)
                $st.List.SetItemChecked($idx, [bool]$it.Checked)
                [void]$vis.Add($it)
            }
            $st.Visible = @($vis.ToArray())
        }
        finally {
            $st.List.EndUpdate()
            $st.Refreshing = $false
        }
        $checkedN = @($st.Items | Where-Object { $_.Checked }).Count
        $skipN = @($st.Items | Where-Object { $_.Unwanted }).Count
        $st.Status.Text = ("Selected {0} / {1}   (auto-unchecked extras: {2})" -f $checkedN, $st.Items.Count, $skipN)
    }

    $list.Add_ItemCheck({
        param($sender, $e)
        $st = $script:ModelPickState
        if ($st.Refreshing) { return }
        $vis = @($st.Visible)
        if ($e.Index -ge 0 -and $e.Index -lt $vis.Count) {
            $vis[$e.Index].Checked = ($e.NewValue -eq [System.Windows.Forms.CheckState]::Checked)
            $checkedN = @($st.Items | Where-Object { $_.Checked }).Count
            $skipN = @($st.Items | Where-Object { $_.Unwanted }).Count
            $st.Status.Text = ("Selected {0} / {1}   (auto-unchecked extras: {2})" -f $checkedN, $st.Items.Count, $skipN)
        }
    })
    $txtFilter.Add_TextChanged({ & $script:ModelPickRefresh })

    $btnAll = New-Object System.Windows.Forms.Button
    $btnAll.Text = "Check visible"
    $btnAll.Location = New-Object System.Drawing.Point(12, 548)
    $btnAll.Size = New-Object System.Drawing.Size(120, 32)
    $btnAll.Anchor = "Bottom,Left"
    $btnAll.Add_Click({
        foreach ($it in @($script:ModelPickState.Visible)) { $it.Checked = $true }
        & $script:ModelPickRefresh
    })

    $btnNone = New-Object System.Windows.Forms.Button
    $btnNone.Text = "Uncheck visible"
    $btnNone.Location = New-Object System.Drawing.Point(140, 548)
    $btnNone.Size = New-Object System.Drawing.Size(130, 32)
    $btnNone.Anchor = "Bottom,Left"
    $btnNone.Add_Click({
        foreach ($it in @($script:ModelPickState.Visible)) { $it.Checked = $false }
        & $script:ModelPickRefresh
    })

    $btnJunk = New-Object System.Windows.Forms.Button
    $btnJunk.Text = "Uncheck extras"
    $btnJunk.Location = New-Object System.Drawing.Point(278, 548)
    $btnJunk.Size = New-Object System.Drawing.Size(120, 32)
    $btnJunk.Anchor = "Bottom,Left"
    $btnJunk.Add_Click({
        foreach ($it in $script:ModelPickState.Items) {
            if ($it.Unwanted) { $it.Checked = $false }
        }
        & $script:ModelPickRefresh
    })

    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = "OK"
    $btnOk.Location = New-Object System.Drawing.Point(692, 548)
    $btnOk.Size = New-Object System.Drawing.Size(90, 32)
    $btnOk.Anchor = "Bottom,Right"
    $btnOk.Add_Click({
        $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.Close()
    })

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel"
    $btnCancel.Location = New-Object System.Drawing.Point(790, 548)
    $btnCancel.Size = New-Object System.Drawing.Size(90, 32)
    $btnCancel.Anchor = "Bottom,Right"
    $btnCancel.Add_Click({
        $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $form.Close()
    })

    $form.AcceptButton = $btnOk
    $form.CancelButton = $btnCancel
    $form.Controls.AddRange(@(
        $lblHint, $txtFilter, $list, $lblStatus, $btnAll, $btnNone, $btnJunk, $btnOk, $btnCancel
    ))
    $txtFilter.Anchor = "Top,Left,Right"
    $list.Anchor = "Top,Bottom,Left,Right"

    $form.Add_Shown({ & $script:ModelPickRefresh })
    $null = $form.ShowDialog()
    if ($form.DialogResult -ne [System.Windows.Forms.DialogResult]::OK) {
        return @()
    }
    return @($script:ModelPickState.Items | Where-Object { $_.Checked } | ForEach-Object { $_.Path })
}

# ---------------- UI ----------------
Write-Host ""
Write-Host "============================================"
Write-Host ("  Select Revit models -> {0}" -f $OutFile)
Write-Host "  Local disk / UNC  +  Revit Server (RSN)"
Write-Host "  Pass 1 list = OLD names (...IOS...) for SaveAs"
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
Write-Host "  3 - Revit Server: open folder, Add ALL .rvt (no extra checkbox)"
Write-Host ("  4 - Show current {0}" -f $OutFile)
Write-Host ("  5 - Clear {0}" -f $OutFile)
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
$skipChecklist = $false
if ($choice -eq "1") {
    $picked = @(Select-LocalFolder)
    $skipChecklist = $true
}
elseif ($choice -eq "2") {
    $picked = @(Select-LocalFiles)
}
elseif ($choice -eq "3") {
    $picked = @(Select-FromRevitServer)
    $skipChecklist = $true
}
else {
    Write-Host "Invalid choice."
    exit 1
}

if ($picked.Count -eq 0) {
    Write-Host "Nothing selected."
    exit 0
}

if ((-not $skipChecklist) -and ($picked.Count -gt 1)) {
    Write-Host ("Found {0} model(s). Opening checkbox to exclude extras..." -f $picked.Count)
    $picked = @(Show-ModelChecklist -Paths $picked)
    if ($picked.Count -eq 0) {
        Write-Host "Nothing selected."
        exit 0
    }
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
[void](Ensure-RbpOpenFailPatch)
[void](Ensure-RbpProgressMonitor)
if ($hasRsn) {
    [void](Ensure-RbpRsnPatch)
}

Write-Host ""
Write-Host "Next:"
Write-Host "  Open presets.ps1 (run_cascade.cmd item 1) -> Save job"
Write-Host "  Then follow rbp_checklist.txt"
Write-Host ""
if ($hasRsn) {
    Write-Host "IMPORTANT (list has RSN://):"
    Write-Host "  Central processing = Create New Local"
    Write-Host "  Detach from Central = OFF"
    Write-Host "  Revit version = model year"
}
else {
    Write-Host "IMPORTANT if shared centrals:"
    Write-Host "  Central processing = Create New Local, Detach = OFF"
}
Write-Host ""
