# Name: Watch-FtpModels.ps1
# Version: 1.0
# What it does: Poll FTP folders for new/changed .rvt files and send Telegram or ntfy alerts.
# Inputs: config.json, credentials.xml, state.json.
# Outputs: Push notifications; updated state.json.
# How to run: powershell -File Watch-FtpModels.ps1 or via scheduled task.
# Notes: Requires Save-FtpCredentials.ps1 first.
param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot "config.json")
)

$ErrorActionPreference = "Stop"

# Telegram/HTTPS often need TLS 1.2 on older Windows PowerShell
try {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls
} catch {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
}

if (-not (Test-Path $ConfigPath)) {
    throw "Missing config.json"
}

$config = Get-Content $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
$credPath = Join-Path $PSScriptRoot "credentials.xml"
$statePath = Join-Path $PSScriptRoot "state.json"

function ConvertTo-StrictBool {
    param([object]$Value)
    if ($Value -is [bool]) { return $Value }
    if ($null -eq $Value) { return $false }
    $s = ([string]$Value).Trim().ToLowerInvariant()
    return @("1", "true", "yes", "on") -contains $s
}

if (-not (Test-Path $credPath)) {
    throw "Missing credentials.xml. Run Save-FtpCredentials.ps1 first."
}

$cred = Import-Clixml -Path $credPath
$user = $cred.UserName
$pass = $cred.GetNetworkCredential().Password
$useSsl = ConvertTo-StrictBool $config.UseSsl

function Get-WatchPaths {
    $paths = New-Object System.Collections.Generic.List[string]
    if ($null -ne $config.RemotePaths) {
        foreach ($p in @($config.RemotePaths)) {
            if (-not [string]::IsNullOrWhiteSpace([string]$p)) {
                $paths.Add(([string]$p).Trim())
            }
        }
    }
    if ($paths.Count -eq 0 -and -not [string]::IsNullOrWhiteSpace([string]$config.RemotePath)) {
        $paths.Add(([string]$config.RemotePath).Trim())
    }
    if ($paths.Count -eq 0) {
        throw "Set RemotePaths (array) or RemotePath in config.json"
    }
    return $paths
}

$watchPaths = Get-WatchPaths
$provider = "telegram"
if ($config.NotifyProvider) {
    $provider = ([string]$config.NotifyProvider).Trim().ToLowerInvariant()
}
Write-Host "Config: Host=$($config.Host) Port=$($config.Port) UseSsl=$useSsl Provider=$provider Paths=$($watchPaths.Count)"
foreach ($p in $watchPaths) { Write-Host "  - $p" }

function Normalize-FtpPath {
    param([string]$RemotePath)
    $path = if ($null -eq $RemotePath) { "/" } else { $RemotePath.TrimEnd("/") }
    if ([string]::IsNullOrWhiteSpace($path)) { $path = "/" }
    if (-not $path.StartsWith("/")) { $path = "/$path" }
    return $path
}

function Normalize-NtfyTopic {
    param([string]$Topic)
    $t = $Topic.Trim()
    $t = $t -replace '^https?://ntfy\.sh/', ''
    $t = $t -replace '^ntfy\.sh/', ''
    $t = $t.Trim('/')
    return $t
}

function Get-FtpRawListing {
    param(
        [string]$HostName,
        [int]$Port,
        [string]$RemotePath,
        [string]$Username,
        [string]$Password,
        [bool]$UseSsl
    )

    $path = Normalize-FtpPath $RemotePath
    $uri = "ftp://${HostName}:${Port}${path}"
    $request = [System.Net.FtpWebRequest]::Create($uri)
    $request.Method = [System.Net.WebRequestMethods+Ftp]::ListDirectoryDetails
    $request.Credentials = New-Object System.Net.NetworkCredential($Username, $Password)
    if ($UseSsl) {
        $request.EnableSsl = $true
        [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }
    }
    else {
        $request.EnableSsl = $false
    }
    $request.UseBinary = $true
    $request.UsePassive = $true
    $request.KeepAlive = $false
    $request.Timeout = 60000
    $request.ReadWriteTimeout = 60000

    $response = $request.GetResponse()
    try {
        $reader = New-Object System.IO.StreamReader($response.GetResponseStream())
        $raw = $reader.ReadToEnd()
        $reader.Close()
        return $raw
    }
    finally {
        $response.Close()
    }
}

function Parse-FtpEntries {
    param([string]$Raw)

    $entries = New-Object System.Collections.Generic.List[object]

    foreach ($line in ($Raw -split "`r?`n")) {
        if ([string]::IsNullOrWhiteSpace($line)) { continue }

        $name = $null
        $size = "0"
        $stamp = ""
        $isDir = $false

        if ($line -match '^(?<perm>[-d][rwx-]{9})\s+\d+\s+\S+\s+\S+\s+(?<size>\d+)\s+(?<stamp>\S+\s+\d+\s+[\d:]+)\s+(?<name>.+)$') {
            $isDir = $Matches.perm.StartsWith("d")
            $name = $Matches.name.Trim()
            $size = $Matches.size
            $stamp = $Matches.stamp
        }
        elseif ($line -match '^(?<stamp>\d{2}-\d{2}-\d{2}\s+\d{1,2}:\d{2}(?:AM|PM)?)\s+<DIR>\s+(?<name>.+)$') {
            $isDir = $true
            $name = $Matches.name.Trim()
            $stamp = $Matches.stamp
            $size = "DIR"
        }
        elseif ($line -match '^(?<stamp>\d{2}-\d{2}-\d{2}\s+\d{1,2}:\d{2}(?:AM|PM)?)\s+(?<size>\d+)\s+(?<name>.+)$') {
            $isDir = $false
            $name = $Matches.name.Trim()
            $size = $Matches.size
            $stamp = $Matches.stamp
        }
        else {
            continue
        }

        if ($name -eq "." -or $name -eq ".." -or $name -eq "Archive") { continue }

        $entries.Add([pscustomobject]@{
            Name  = $name
            IsDir = $isDir
            Size  = $size
            Stamp = $stamp
        })
    }

    return $entries
}

function Test-WatchedFile {
    param([string]$Name)

    foreach ($ext in $config.WatchExtensions) {
        if ($Name.ToLowerInvariant().EndsWith($ext.ToLowerInvariant())) {
            return $true
        }
    }
    return $false
}

function Get-FtpSnapshot {
    param(
        [string]$HostName,
        [int]$Port,
        [string]$RemotePath,
        [string]$Username,
        [string]$Password,
        [bool]$UseSsl,
        [int]$Depth
    )

    $snapshot = @{}
    $queue = New-Object System.Collections.Generic.Queue[object]
    $queue.Enqueue([pscustomobject]@{ Path = (Normalize-FtpPath $RemotePath); Depth = 0 })

    while ($queue.Count -gt 0) {
        $item = $queue.Dequeue()
        $raw = Get-FtpRawListing -HostName $HostName -Port $Port -RemotePath $item.Path `
            -Username $Username -Password $Password -UseSsl $UseSsl
        $entries = Parse-FtpEntries -Raw $raw

        foreach ($e in $entries) {
            $relBase = $item.Path.TrimEnd("/")
            $full = if ($relBase -eq "" -or $relBase -eq "/") {
                $e.Name
            } else {
                "$relBase/$($e.Name)"
            }
            $full = $full.TrimStart("/")

            if ($e.IsDir) {
                if ($config.WatchFolders) {
                    $snapshot["DIR:$full"] = "$($e.Size)|$($e.Stamp)"
                }
                if ($item.Depth -lt $Depth) {
                    $childPath = if ($item.Path -eq "/") { "/$($e.Name)" } else { "$($item.Path.TrimEnd('/'))/$($e.Name)" }
                    $queue.Enqueue([pscustomobject]@{ Path = $childPath; Depth = ($item.Depth + 1) })
                }
            }
            else {
                if (Test-WatchedFile -Name $e.Name) {
                    $snapshot["FILE:$full"] = "$($e.Size)|$($e.Stamp)"
                }
            }
        }
    }

    return $snapshot
}

function Send-Ntfy {
    param(
        [string]$Topic,
        [string]$Title,
        [string]$Message,
        [string]$Token
    )

    $topic = Normalize-NtfyTopic $Topic
    if ([string]::IsNullOrWhiteSpace($topic)) {
        throw "NtfyTopic is empty"
    }

    $payload = @{
        topic    = $topic
        title    = $Title
        message  = $Message
        priority = 3
        tags     = @("file_folder")
    } | ConvertTo-Json -Compress

    $headers = @{ }
    if ($Token) {
        $headers["Authorization"] = "Bearer $Token"
    }

    $utf8 = New-Object System.Text.UTF8Encoding $false
    $bytes = $utf8.GetBytes($payload)

    Invoke-RestMethod -Method Post -Uri "https://ntfy.sh" -Headers $headers -ContentType "application/json; charset=utf-8" -Body $bytes | Out-Null
}

function Send-Telegram {
    param(
        [string]$BotToken,
        [string]$ChatId,
        [string]$Title,
        [string]$Message
    )

    if ([string]::IsNullOrWhiteSpace($BotToken) -or [string]::IsNullOrWhiteSpace($ChatId)) {
        throw "Set TelegramBotToken and TelegramChatId in config.json"
    }

    $text = "$Title`n`n$Message"
    if ($text.Length -gt 4000) {
        $text = $text.Substring(0, 3990) + "`n..."
    }

    $uri = "https://api.telegram.org/bot$BotToken/sendMessage"
    $payload = @{
        chat_id                  = $ChatId
        text                     = $text
        disable_web_page_preview = $true
    } | ConvertTo-Json -Compress

    $utf8 = New-Object System.Text.UTF8Encoding $false
    $bytes = $utf8.GetBytes($payload)

    Invoke-RestMethod -Method Post -Uri $uri -ContentType "application/json; charset=utf-8" -Body $bytes | Out-Null
}

function Send-Alert {
    param(
        [string]$Title,
        [string]$Message
    )

    switch ($provider) {
        "telegram" {
            Send-Telegram -BotToken ([string]$config.TelegramBotToken) -ChatId ([string]$config.TelegramChatId) -Title $Title -Message $Message
        }
        "ntfy" {
            Send-Ntfy -Topic $config.NtfyTopic -Title $Title -Message $Message -Token $config.NtfyToken
        }
        default {
            throw "Unknown NotifyProvider: $provider (use telegram or ntfy)"
        }
    }
}

$depth = 1
if ($null -ne $config.RecursiveDepth) { $depth = [int]$config.RecursiveDepth }
if ($null -eq $config.WatchFolders) { $config | Add-Member -NotePropertyName WatchFolders -NotePropertyValue $true -Force }

$current = @{}
foreach ($watchPath in $watchPaths) {
    Write-Host "Scanning $watchPath ..."
    $part = Get-FtpSnapshot `
        -HostName $config.Host `
        -Port ([int]$config.Port) `
        -RemotePath $watchPath `
        -Username $user `
        -Password $pass `
        -UseSsl $useSsl `
        -Depth $depth
    foreach ($key in $part.Keys) {
        $current[$key] = $part[$key]
    }
}

$previous = @{}
if (Test-Path $statePath) {
    $prevObj = Get-Content $statePath -Raw -Encoding UTF8 | ConvertFrom-Json
    $prevObj.PSObject.Properties | ForEach-Object { $previous[$_.Name] = $_.Value }
}

function Get-FtpStamp {
    param([string]$StateValue)
    if ([string]::IsNullOrWhiteSpace($StateValue)) { return "" }
    $parts = $StateValue -split "\|", 2
    if ($parts.Count -ge 2) { return $parts[1].Trim() }
    return $StateValue.Trim()
}

function Format-ChangeLine {
    param(
        [string]$Kind,
        [string]$NiceName,
        [string]$StateValue = ""
    )
    $stamp = Get-FtpStamp $StateValue
    if ($stamp) {
        return "$Kind`: $NiceName`n  FTP time: $stamp"
    }
    return "$Kind`: $NiceName"
}

$changes = New-Object System.Collections.Generic.List[string]
$seenLocal = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

foreach ($name in $current.Keys) {
    $nice = $name -replace '^(DIR|FILE):', ''
    if (-not $previous.ContainsKey($name)) {
        $changes.Add((Format-ChangeLine -Kind "NEW" -NiceName $nice -StateValue $current[$name]))
    }
    elseif ($previous[$name] -ne $current[$name]) {
        $oldStamp = Get-FtpStamp $previous[$name]
        $newStamp = Get-FtpStamp $current[$name]
        $line = "UPD: $nice"
        if ($newStamp) { $line += "`n  FTP time: $newStamp" }
        if ($oldStamp -and $oldStamp -ne $newStamp) { $line += "`n  was: $oldStamp" }
        $changes.Add($line)
    }
}

foreach ($name in $previous.Keys) {
    if (-not $current.ContainsKey($name)) {
        $nice = $name -replace '^(DIR|FILE):', ''
        $changes.Add((Format-ChangeLine -Kind "DEL" -NiceName $nice -StateValue $previous[$name]))
    }
}

$current | ConvertTo-Json -Compress | Set-Content -Path $statePath -Encoding UTF8

if ($changes.Count -eq 0) {
    Write-Host "$(Get-Date -Format o) no changes ($($current.Count) entries)"
    exit 0
}

if ($previous.Count -eq 0 -and -not $config.NotifyOnFirstRun) {
    Write-Host "$(Get-Date -Format o) baseline saved ($($current.Count) entries), no notify"
    exit 0
}

$body = "Checked: $seenLocal`n`n" + (($changes | Select-Object -First 20) -join "`n`n")
if ($changes.Count -gt 20) {
    $body += "`n`n... +$($changes.Count - 20) more"
}

Send-Alert -Title $config.NtfyTitle -Message $body
Write-Host "$(Get-Date -Format o) notified: $($changes.Count) change(s)"
