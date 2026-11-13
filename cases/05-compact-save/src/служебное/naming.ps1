# Exported list and report file names. UTF-8 WITH BOM required (PowerShell 5.1).
# Standard: yyyy-MM-dd_HHmm_<Ryy|Rmix|Rlocal>_<project>_<models|report>[_quick|_deep].ext
# Project token: Revit Server root folder, or the folder above the discipline next to RVT.

function ConvertTo-CompactFileToken {
    param(
        [string]$Text,
        [string]$Fallback = "na"
    )
    if ([string]::IsNullOrWhiteSpace($Text)) { return $Fallback }
    $badChars = [char[]]@('\', '/', ':', '*', '?', '"', '<', '>', '|', ' ')
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $Text.Trim().ToCharArray()) {
        $code = [int][char]$ch
        $bad = ($code -lt 32) -or ($badChars -contains $ch)
        if ($bad) { [void]$sb.Append('_') } else { [void]$sb.Append($ch) }
    }
    $token = $sb.ToString().Trim('._')
    if ($token.Length -gt 40) { $token = $token.Substring(0, 40).Trim('._') }
    if ([string]::IsNullOrWhiteSpace($token)) { return $Fallback }
    return $token
}

function Get-CompactProjectName {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return "" }
    $text = $Path.Trim()
    if ($text -match '^(?i)RSN://([^/]+)(?:/(.*))?$') {
        $hostName = [string]$Matches[1]
        $rest = [string]$Matches[2]
        $parts = @($rest.Split('/') | Where-Object { $_ })
        if ($parts.Count -ge 2) { return [string]$parts[0] }
        return $hostName
    }
    $norm = $text.Replace('/', '\')
    $parts = @($norm.Split('\') | Where-Object { $_ -and ($_ -ne '.') })
    for ($i = 0; $i -lt $parts.Count; $i++) {
        if ($parts[$i] -eq 'RVT' -or $parts[$i] -eq 'rvt') {
            if ($i -ge 2) { return [string]$parts[$i - 2] }
            if ($i -ge 1) { return [string]$parts[$i - 1] }
        }
    }
    if ($parts.Count -ge 2) { return [string]$parts[$parts.Count - 2] }
    return ""
}

function Get-CompactPathYear {
    param([string]$Path)
    if (Get-Command Guess-Year -ErrorAction SilentlyContinue) {
        $fromGuess = Guess-Year $Path
        if ($fromGuess) { return [string]$fromGuess }
    }
    if ($Path -match '(?i)_R(20|21|22|23|24|25|26)\b') { return ("20" + $Matches[1]) }
    if ($Path -match '(?i)RSN://revit-server-2021\.example\.local\b') { return "2021" }
    if ($Path -match '(?i)RSN://revit-server-2022\.example\.local\b') { return "2022" }
    if ($Path -match '(?i)RSN://revit-server-2023\.example\.local\b') { return "2023" }
    if ($Path -match '(?i)RSN://revit-server-2024\.example\.local\b') { return "2024" }
    if ($Path -match '(?i)RSN://revit-server-202([1-6])') { return ("202" + $Matches[1]) }
    if ($Path -match '(?i)Revit Server 202([1-6])') { return ("202" + $Matches[1]) }
    return ""
}

function Get-CompactRevitToken {
    param([string[]]$Paths)
    $years = New-Object System.Collections.Generic.List[string]
    $unknown = $false
    foreach ($p in @($Paths)) {
        if ([string]::IsNullOrWhiteSpace($p)) { continue }
        $y = Get-CompactPathYear $p
        if ($y) { [void]$years.Add($y) } else { $unknown = $true }
    }
    $uniq = @($years | Sort-Object -Unique)
    if ($uniq.Count -eq 1 -and -not $unknown) {
        return ("R" + $uniq[0].Substring(2))
    }
    if ($uniq.Count -eq 0) { return "Rlocal" }
    return "Rmix"
}

function Get-CompactProjectToken {
    param([string[]]$Paths)
    $names = New-Object System.Collections.Generic.List[string]
    foreach ($p in @($Paths)) {
        if ([string]::IsNullOrWhiteSpace($p)) { continue }
        $name = Get-CompactProjectName $p
        if (-not [string]::IsNullOrWhiteSpace($name)) { [void]$names.Add($name) }
    }
    $uniq = @($names | Sort-Object -Unique)
    if ($uniq.Count -eq 1) { return (ConvertTo-CompactFileToken -Text $uniq[0] -Fallback "na") }
    if ($uniq.Count -eq 0) { return "na" }
    return "mixed"
}

function Get-CompactStandardFileName {
    param(
        [string[]]$Paths,
        [ValidateSet("models", "report")]
        [string]$Kind = "models",
        [string]$Mode = "",
        [datetime]$When = (Get-Date)
    )
    $stamp = $When.ToString("yyyy-MM-dd_HHmm")
    $revit = Get-CompactRevitToken -Paths $Paths
    $project = Get-CompactProjectToken -Paths $Paths
    $name = "{0}_{1}_{2}_{3}" -f $stamp, $revit, $project, $Kind
    if ($Kind -eq "report" -and -not [string]::IsNullOrWhiteSpace($Mode)) {
        $modeTag = "deep"
        if ($Mode -match '^(?i)quick|fast|быстр') { $modeTag = "quick" }
        $name = "{0}_{1}" -f $name, $modeTag
    }
    $ext = ".txt"
    if ($Kind -eq "report") { $ext = ".xlsx" }
    return ($name + $ext)
}
