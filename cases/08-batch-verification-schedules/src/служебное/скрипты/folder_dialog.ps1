# UTF-8 WITH BOM required (PowerShell 5.1).
# Единственное окно выбора папки: OpenFileDialog с адресной строкой (подходят серверные и сетевые пути \\сервер\папка).
# Системный диалог выбора папки (из .NET) не используется нигде. Подключается из ui_schedules.ps1 и build_release.ps1.
# Нужен Add-Type -AssemblyName System.Windows.Forms; $form (если есть) — владелец окна.

function Get-NativePath([string]$Path) {
    if ([string]::IsNullOrWhiteSpace($Path)) { return $null }
    $p = $Path.Trim().Trim('"')
    if ($p.StartsWith("\\")) { return $p }
    try { return (New-Object System.IO.FileInfo($p)).FullName }
    catch { return $p }
}

function Resolve-PickedFolderPath([string]$Picked) {
    if ([string]::IsNullOrWhiteSpace($Picked)) { return $null }
    $p = $Picked.Trim().Trim('"')
    for ($i = 0; $i -lt 3; $i++) {
        $leaf = ""
        try { $leaf = [System.IO.Path]::GetFileName($p.TrimEnd("\/")) } catch { $leaf = "" }
        $isDummy = (
            $leaf -match '^(?i)Select this folder' -or
            $leaf -match '(?i)\.nevermatch$' -or
            $leaf -eq "." -or
            $leaf -eq "Folder"
        )
        if (-not $isDummy) { break }
        try { $p = [System.IO.Path]::GetDirectoryName($p) } catch { break }
        if ([string]::IsNullOrWhiteSpace($p)) { break }
    }
    $ext = ""
    try { $ext = [System.IO.Path]::GetExtension($p) } catch { $ext = "" }
    if ($ext -and $ext.ToLower() -eq ".rvt") {
        try { return [System.IO.Path]::GetDirectoryName($p) } catch { return $p }
    }
    return $p
}

# Начальная папка окна «Обзор…»: папка из строки (файл — его папка). RSN:// и несуществующие пути — без изменений
function Get-DialogStartDir([string]$Path) {
    $p = ([string]$Path).Trim().Trim('"')
    if (-not $p -or $p -match '^(?i)RSN://') { return "" }
    try {
        if (Test-Path -LiteralPath $p -PathType Container) { return $p }
        $d = Split-Path -Parent $p
        if ($d -and (Test-Path -LiteralPath $d -PathType Container)) { return $d }
    } catch { }
    return ""
}

function Show-FolderPathDialog([string]$Title = "Папка с моделями — зайдите внутрь нужной папки и нажмите «Открыть»", [string]$StartPath = "") {
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Title = $Title
    $dlg.Filter = "Все файлы (*.*)|*.*"
    $dlg.FilterIndex = 1
    $dlg.CheckFileExists = $false
    $dlg.CheckPathExists = $false
    $dlg.ValidateNames = $false
    $dlg.Multiselect = $false
    $dlg.DereferenceLinks = $true
    $dlg.FileName = "Select this folder"
    $start = Get-DialogStartDir $StartPath
    if ($start) { $dlg.InitialDirectory = $start }
    $owner = $form
    $result = if ($owner -and -not $owner.IsDisposed) { $dlg.ShowDialog($owner) } else { $dlg.ShowDialog() }
    if ($result -ne [System.Windows.Forms.DialogResult]::OK) { return $null }
    $resolved = Resolve-PickedFolderPath $dlg.FileName
    return (Get-NativePath $resolved)
}
