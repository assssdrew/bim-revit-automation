# UTF-8 WITH BOM required (PowerShell 5.1).
# Отдельный экземпляр Excel для автоматизации: не цепляется к открытому у пользователя и Quit() его не закрывает.

if (-not $script:BvsExcelModuleDir) {
    if ($PSScriptRoot) {
        $script:BvsExcelModuleDir = $PSScriptRoot
    }
    elseif ($MyInvocation.MyCommand.Path) {
        $script:BvsExcelModuleDir = Split-Path -Parent $MyInvocation.MyCommand.Path
    }
    elseif ($script:BvsAssignmentDir) {
        $script:BvsExcelModuleDir = $script:BvsAssignmentDir
    }
}

function Get-BvsExcelModuleDir {
    if ($script:BvsExcelModuleDir) { return $script:BvsExcelModuleDir }
    if ($script:BvsAssignmentDir) { return $script:BvsAssignmentDir }
    throw "Не определена папка excel_isolated.ps1. Запускайте сборку через build_assignment.ps1 из batch_verification_schedules."
}

if (-not ('BvsExcelRot' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

public static class BvsExcelRot
{
    [DllImport("ole32.dll")]
    private static extern int GetRunningObjectTable(int reserved, out IRunningObjectTable prot);

    [DllImport("ole32.dll")]
    private static extern int CreateBindCtx(int reserved, out IBindCtx ppbc);

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

    public static List<object> GetExcelApplications()
    {
        var list = new List<object>();
        IRunningObjectTable rot;
        if (GetRunningObjectTable(0, out rot) != 0) { return list; }
        IEnumMoniker enumMoniker;
        rot.EnumRunning(out enumMoniker);
        if (enumMoniker == null) { return list; }
        enumMoniker.Reset();
        IMoniker[] monikers = new IMoniker[1];
        IntPtr fetched = IntPtr.Zero;
        while (enumMoniker.Next(1, monikers, fetched) == 0)
        {
            IBindCtx ctx;
            if (CreateBindCtx(0, out ctx) != 0) { continue; }
            string displayName;
            monikers[0].GetDisplayName(ctx, null, out displayName);
            if (displayName == null || displayName.IndexOf("Excel", StringComparison.OrdinalIgnoreCase) < 0) { continue; }
            object obj;
            if (rot.GetObject(monikers[0], out obj) != 0 || obj == null) { continue; }
            try
            {
                var t = obj.GetType();
                var wb = t.GetProperty("Workbooks");
                if (wb == null) { continue; }
                list.Add(obj);
            }
            catch { }
        }
        return list;
    }

    public static int GetExcelProcessId(object excelApp)
    {
        if (excelApp == null) { return 0; }
        try
        {
            var t = excelApp.GetType();
            var hwndProp = t.GetProperty("Hwnd");
            if (hwndProp == null) { return 0; }
            int hwnd = (int)hwndProp.GetValue(excelApp, null);
            if (hwnd == 0) { return 0; }
            uint pid;
            GetWindowThreadProcessId(new IntPtr(hwnd), out pid);
            return (int)pid;
        }
        catch { return 0; }
    }

    public static object GetApplicationByProcessId(int processId)
    {
        var apps = GetExcelApplications();
        object fallback = null;
        foreach (var obj in apps)
        {
            int pid = GetExcelProcessId(obj);
            if (pid == processId) { return obj; }
            if (pid == 0) { fallback = obj; }
        }
        if (apps.Count == 1 && fallback != null) { return fallback; }
        return null;
    }
}
'@
}

function Get-ExcelExePath {
    $keys = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\excel.exe'
    )
    foreach ($k in $keys) {
        if (-not (Test-Path -LiteralPath $k)) { continue }
        $p = (Get-ItemProperty -LiteralPath $k -ErrorAction SilentlyContinue).'(default)'
        if ($p -and (Test-Path -LiteralPath $p)) { return $p }
    }
    $candidates = @(
        "${env:ProgramFiles}\Microsoft Office\root\Office16\EXCEL.EXE",
        "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16\EXCEL.EXE"
    )
    foreach ($c in $candidates) {
        if (Test-Path -LiteralPath $c) { return $c }
    }
    return $null
}

# Пустая книга-«затравка» для запуска Excel с ключом /x (чтобы экземпляр зарегистрировался в ROT).
# Раньше брали служебное\excel_blank.xlsx: в репозитории его нет, а локальная копия была битой,
# из-за чего Excel показывал окно «восстановить файл?» и COM не подключался.
# Теперь минимальный валидный xlsx создаётся заново во временной папке и удаляется после закрытия Excel.
function New-BvsExcelSeedFile {
    Add-Type -AssemblyName System.IO.Compression -ErrorAction SilentlyContinue
    $path = Join-Path ([System.IO.Path]::GetTempPath()) ("bvs_excel_seed_{0}.xlsx" -f [Guid]::NewGuid().ToString("N"))
    $ns = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    $parts = [ordered]@{
        '[Content_Types].xml' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>'
        '_rels/.rels' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
        'xl/workbook.xml' = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="{0}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>' -f $ns)
        'xl/_rels/workbook.xml.rels' = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'
        'xl/styles.xml' = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="{0}"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>' -f $ns)
        'xl/worksheets/sheet1.xml' = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="{0}"><sheetData/></worksheet>' -f $ns)
    }
    $utf8 = New-Object System.Text.UTF8Encoding $false
    $fs = [System.IO.File]::Open($path, [System.IO.FileMode]::CreateNew)
    try {
        $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Create, $false)
        try {
            foreach ($name in $parts.Keys) {
                $entry = $zip.CreateEntry($name)
                $es = $entry.Open()
                try {
                    $bytes = $utf8.GetBytes([string]$parts[$name])
                    $es.Write($bytes, 0, $bytes.Length)
                }
                finally { $es.Dispose() }
            }
        }
        finally { $zip.Dispose() }
    }
    finally { $fs.Dispose() }
    return $path
}

# Запуск своего Excel: PID берётся из Start-Process -PassThru (процесс, который запустили мы), а не из «новых процессов EXCEL»:
# Excel пользователя, появившийся во время прогона, за свой не принимается и никогда не закрывается. Не получили процесс — 0.
function Start-IsolatedExcelProcess {
    param(
        [string]$Exe,
        [string]$SeedPath
    )
    # путь в кавычках: во временной папке пользователя бывают пробелы и кириллица
    $proc = Start-Process -FilePath $Exe -ArgumentList ('/x "{0}"' -f $SeedPath) -WindowStyle Hidden -PassThru
    if ($null -eq $proc) { return 0 }
    return [int]$proc.Id
}

function Test-ProcessAlive([int]$ProcessId) {
    return [bool](Get-Process -Id $ProcessId -ErrorAction SilentlyContinue)
}

# Закрывает ТОЛЬКО указанный нами процесс Excel (PID свой: из запуска или по hwnd своего COM-объекта); имя процесса проверяется
function Stop-OwnExcelProcess {
    param([int]$ProcessId)
    if ($ProcessId -le 0) { return }
    $p = Get-Process -Id $ProcessId -ErrorAction SilentlyContinue
    if ($p -and $p.ProcessName -match '^(?i)EXCEL$') { Stop-Process -Id $ProcessId -Force -ErrorAction SilentlyContinue }
}

function Wait-ExcelComForProcessId {
    param(
        [int]$ProcessId,
        [int]$TimeoutSec = 180
    )
    $deadline = (Get-Date).AddSeconds($TimeoutSec)
    while ((Get-Date) -lt $deadline) {
        if (-not (Test-ProcessAlive $ProcessId)) { return $null }        # запущенный нами процесс завершился — ждать нечего
        $excel = [BvsExcelRot]::GetApplicationByProcessId($ProcessId)
        if ($null -ne $excel) {
            $owner = [BvsExcelRot]::GetExcelProcessId($excel)             # по hwnd COM-объекта: чужой Excel не принимаем
            if ($owner -eq 0 -or $owner -eq $ProcessId) { return $excel }
        }
        Start-Sleep -Milliseconds 400
    }
    return $null
}

# Excel запускается скрытно (/x, окно Hidden). Закрывается только свой процесс: PID из запуска (или по hwnd своего COM-объекта).
# Свою seed-книгу удаляем всегда. Если свой PID определить не удалось — НИЧЕГО не закрываем, ставим $script:ExcelPidUnknown и бросаем
# исключение «Excel не закрыт, PID неизвестен». Ожидание COM — 2 попытки по 60 с.
$script:ExcelPidUnknown = $false
function New-IsolatedExcelApplication {
    $exe = Get-ExcelExePath
    if (-not $exe) {
        throw "Не найден EXCEL.EXE. Сборка xlsx через COM невозможна."
    }
    $script:ExcelPidUnknown = $false
    $lastPid = 0
    $pidUnknown = $false
    $seed = New-BvsExcelSeedFile
    $ok = $false
    try {
        for ($attempt = 1; $attempt -le 2; $attempt++) {
            $ownPid = [int](Start-IsolatedExcelProcess -Exe $exe -SeedPath $seed)
            if ($ownPid -le 0) {
                $pidUnknown = $true
                if ($attempt -lt 2) { Start-Sleep -Seconds 2; continue }
                break
            }
            $lastPid = $ownPid
            $excel = Wait-ExcelComForProcessId -ProcessId $ownPid -TimeoutSec 60
            if ($null -ne $excel) {
                $excel.Visible = $false
                $excel.DisplayAlerts = $false
                $excel.ScreenUpdating = $false
                $ok = $true
                $hwndPid = [BvsExcelRot]::GetExcelProcessId($excel)
                if ($hwndPid -gt 0) { $ownPid = $hwndPid }           # по hwnd своего COM-объекта
                return [pscustomobject]@{
                    Application = $excel
                    ProcessId   = $ownPid
                    SeedPath    = $seed
                }
            }
            if (Test-ProcessAlive $ownPid) {
                Stop-OwnExcelProcess $ownPid                           # наш процесс, COM не подключился — закрываем его
                $pidUnknown = $false
            }
            else {
                $pidUnknown = $true                                    # запущенный процесс сразу завершился: свой экземпляр мог остаться под другим PID — не угадываем
            }
            Start-Sleep -Seconds 2
        }
    }
    finally {
        # при успехе файл нужен, пока Excel не закрыт: удалит Stop-IsolatedExcelApplication; при ошибке — удалить свою seed-книгу
        if (-not $ok) {
            Start-Sleep -Milliseconds 300
            Remove-Item -LiteralPath $seed -Force -ErrorAction SilentlyContinue
        }
    }
    if ($pidUnknown) {
        $script:ExcelPidUnknown = $true
        throw "Excel не закрыт, PID неизвестен: свой экземпляр определить не удалось, другие EXCEL.EXE не тронуты. Закройте лишний Excel вручную и повторите сборку."
    }
    throw ("Excel запущен (PID {0}), но COM к нему не подключился; процесс закрыт. Один раз откройте Excel вручную (лицензия) и повторите сборку." -f $lastPid)
}

function Stop-IsolatedExcelApplication {
    param($Handle)
    if ($null -eq $Handle) { return }
    $excel = $Handle.Application
    $excelPid = [int]$Handle.ProcessId
    if ($excel) {
        try { $excel.DisplayAlerts = $false } catch { }
        try { $excel.Quit() } catch { }
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch { }
    }
    if ($excelPid -gt 0) {
        # даём Excel штатно завершиться (до 5 с), и только потом убиваем, чтобы не оставлять зависший EXCEL.EXE
        $p = $null
        for ($i = 0; $i -lt 25; $i++) {
            $p = Get-Process -Id $excelPid -ErrorAction SilentlyContinue
            if (-not $p) { break }
            Start-Sleep -Milliseconds 200
        }
        if ($p) { Stop-OwnExcelProcess $excelPid }
    }
    if ($Handle.SeedPath) {
        Remove-Item -LiteralPath $Handle.SeedPath -Force -ErrorAction SilentlyContinue
    }
}
