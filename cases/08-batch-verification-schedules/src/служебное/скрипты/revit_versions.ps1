# UTF-8 WITH BOM required (PowerShell 5.1).
# Версия Revit для запуска Revit Batch Processor. Не по имени/пути, а по самим моделям:
#   .rvt  — год из BasicFileInfo внутри файла (строка «Format: 2022», читается без Revit, из нескольких килобайт);
#   RSN://хост/…  — год из servers.cfg по хосту (строка host|название|год).
# Перед запуском проверяется: версия определена у всех, модели одной версии, общий файл не новее моделей
# (общий файл старше моделей — норма: он хранится в ранней версии и не сохраняется). Подключается из ui_schedules.ps1.

if (-not ([System.Management.Automation.PSTypeName]'BvsCompoundFile').Type) {
    Add-Type -TypeDefinition @"
using System;
using System.IO;
using System.Collections.Generic;
using System.Text;

public static class BvsCompoundFile
{
    // Читает поток по имени из файла Microsoft Compound File (так хранится BasicFileInfo в .rvt). Возвращает null, если потока нет.
    public static byte[] ReadStream(string path, string streamName)
    {
        using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
        {
            byte[] h = new byte[512];
            if (fs.Read(h, 0, 512) < 512) return null;
            if (h[0] != 0xD0 || h[1] != 0xCF || h[2] != 0x11 || h[3] != 0xE0) return null;
            int ss = 1 << BitConverter.ToUInt16(h, 30);
            int mss = 1 << BitConverter.ToUInt16(h, 32);
            uint firstDir = BitConverter.ToUInt32(h, 48);
            uint cutoff = BitConverter.ToUInt32(h, 56);
            uint firstMini = BitConverter.ToUInt32(h, 60);
            uint firstDifat = BitConverter.ToUInt32(h, 68);
            uint nDifat = BitConverter.ToUInt32(h, 72);
            List<uint> fatSectors = new List<uint>();
            for (int i = 0; i < 109; i++) { uint v = BitConverter.ToUInt32(h, 76 + 4 * i); if (v < 0xFFFFFFFA) fatSectors.Add(v); }
            uint d = firstDifat;
            for (uint n = 0; n < nDifat && d < 0xFFFFFFFA; n++)
            {
                byte[] s = Sector(fs, ss, d);
                for (int i = 0; i < ss / 4 - 1; i++) { uint v = BitConverter.ToUInt32(s, 4 * i); if (v < 0xFFFFFFFA) fatSectors.Add(v); }
                d = BitConverter.ToUInt32(s, ss - 4);
            }
            List<uint> fat = new List<uint>();
            foreach (uint fsId in fatSectors) { byte[] s = Sector(fs, ss, fsId); for (int i = 0; i < ss / 4; i++) fat.Add(BitConverter.ToUInt32(s, 4 * i)); }
            // каталог
            byte[] dir = Chain(fs, ss, fat, firstDir, -1);
            uint startSec = 0; long size = -1; uint rootStart = 0; long rootSize = 0;
            for (int off = 0; off + 128 <= dir.Length; off += 128)
            {
                int nameLen = BitConverter.ToUInt16(dir, off + 64);
                int type = dir[off + 66];
                if (nameLen < 2) continue;
                string name = Encoding.Unicode.GetString(dir, off, nameLen - 2);
                uint st = BitConverter.ToUInt32(dir, off + 116);
                long sz = (long)BitConverter.ToUInt64(dir, off + 120);
                if (type == 5) { rootStart = st; rootSize = sz; }
                if (type == 2 && string.Equals(name, streamName, StringComparison.OrdinalIgnoreCase)) { startSec = st; size = sz; }
            }
            if (size < 0) return null;
            if (size >= cutoff) return Chain(fs, ss, fat, startSec, size);
            // малый поток: через мини-FAT и контейнер корневого элемента (читаются только нужные сектора)
            byte[] miniFatBytes = Chain(fs, ss, fat, firstMini, -1);
            List<uint> containerSectors = new List<uint>();
            uint c = rootStart; int guard = 0;
            while (c < 0xFFFFFFFA && guard++ < 10000000) { containerSectors.Add(c); c = c < fat.Count ? fat[(int)c] : 0xFFFFFFFE; }
            byte[] result = new byte[size];
            long done = 0; uint m = startSec; guard = 0;
            while (m < 0xFFFFFFFA && done < size && guard++ < 10000000)
            {
                long pos = (long)m * mss;
                int secIdx = (int)(pos / ss); int inSec = (int)(pos % ss);
                if (secIdx >= containerSectors.Count) break;
                fs.Seek((long)(containerSectors[secIdx] + 1) * ss + inSec, SeekOrigin.Begin);
                int take = (int)Math.Min(mss, size - done);
                int got = fs.Read(result, (int)done, take);
                if (got <= 0) break;
                done += got;
                m = (int)m * 4 + 4 <= miniFatBytes.Length ? BitConverter.ToUInt32(miniFatBytes, (int)m * 4) : 0xFFFFFFFE;
            }
            return result;
        }
    }

    static byte[] Sector(FileStream fs, int ss, uint id)
    {
        byte[] b = new byte[ss];
        fs.Seek((long)(id + 1) * ss, SeekOrigin.Begin);
        int r = fs.Read(b, 0, ss);
        if (r < ss) Array.Resize(ref b, ss);
        return b;
    }

    static byte[] Chain(FileStream fs, int ss, List<uint> fat, uint start, long size)
    {
        List<byte> acc = new List<byte>();
        uint c = start; int guard = 0;
        while (c < 0xFFFFFFFA && guard++ < 10000000 && (size < 0 || acc.Count < size))
        {
            acc.AddRange(Sector(fs, ss, c));
            c = c < fat.Count ? fat[(int)c] : 0xFFFFFFFE;
        }
        if (size >= 0 && acc.Count > size) acc.RemoveRange((int)size, acc.Count - (int)size);
        return acc.ToArray();
    }
}
"@
}

# Год Revit из текста BasicFileInfo: «Format: 2022», запасной вариант — «Revit Build: Autodesk Revit 2022». Пусто — не нашли.
function Get-RevitYearFromInfoText([string]$Text) {
    if ($Text -match '(?i)Format:\s*(20\d{2})') { return [int]$Matches[1] }
    if ($Text -match '(?i)Autodesk Revit\s*(20\d{2})') { return [int]$Matches[1] }
    return $null
}

# Год из файла .rvt (BasicFileInfo). Возвращает @{ Year = int|$null; Why = причина }
function Get-RvtFileYear([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return @{ Year = $null; Why = "файл не найден" } }
    $bytes = $null
    try { $bytes = [BvsCompoundFile]::ReadStream($Path, "BasicFileInfo") }
    catch { return @{ Year = $null; Why = ("не удалось прочитать файл: " + $_.Exception.Message) } }
    if ($null -eq $bytes -or $bytes.Length -lt 8) { return @{ Year = $null; Why = "нет BasicFileInfo (не файл Revit?)" } }
    # текст в UTF-16; смещение 0 или 1 байт (перед ним целые числа)
    foreach ($skip in 0, 1, 2, 3, 4) {
        if ($bytes.Length - $skip -lt 4) { continue }
        $text = [System.Text.Encoding]::Unicode.GetString($bytes, $skip, (($bytes.Length - $skip) -band (-bnot 1)))
        $y = Get-RevitYearFromInfoText $text
        if ($y) { return @{ Year = $y; Why = "" } }
    }
    $y = Get-RevitYearFromInfoText ([System.Text.Encoding]::ASCII.GetString($bytes))
    if ($y) { return @{ Year = $y; Why = "" } }
    return @{ Year = $null; Why = "в BasicFileInfo нет версии Revit" }
}

# servers.cfg: host|название|год -> хэш host(нижний регистр) -> год
function Get-BvsServerYears([string]$ServersCfg) {
    $map = @{}
    if ($ServersCfg -and (Test-Path -LiteralPath $ServersCfg)) {
        $text = [System.IO.File]::ReadAllText($ServersCfg, (New-Object System.Text.UTF8Encoding $true))
        foreach ($line in @($text -split "`r?`n")) {
            $t = $line.Trim()
            if (-not $t -or $t.StartsWith("#")) { continue }
            $p = $t.Split("|")
            if ($p.Count -ge 3 -and $p[2].Trim() -match '^20\d{2}$') { $map[$p[0].Trim().ToLowerInvariant()] = [int]$p[2].Trim() }
        }
    }
    return $map
}

# Год модели: RSN://хост/… — из servers.cfg по хосту; иначе из самого файла. @{ Year; Source = server|file; Why }
function Get-ModelRevitYear([string]$Path, $ServerYears) {
    $p = ([string]$Path).Trim()
    if ($p -match '^(?i)RSN://([^/\\]+)') {
        $hostName = $Matches[1].ToLowerInvariant()
        if ($ServerYears -and $ServerYears.ContainsKey($hostName)) { return @{ Year = [int]$ServerYears[$hostName]; Source = "server"; Why = "" } }
        return @{ Year = $null; Source = "server"; Why = ("сервера {0} нет в servers.cfg" -f $hostName) }
    }
    $r = Get-RvtFileYear $p
    return @{ Year = $r.Year; Source = "file"; Why = $r.Why }
}

function Get-RevitLeaf([string]$Path) {
    $t = ([string]$Path).Trim().Replace("\", "/")
    return $t.Split("/")[-1]
}

# Проверка перед запуском RBP. Models — файлы, которые откроет RBP; Template — общий файл (или пусто, если он не открывается).
# Возвращает @{ Year = год для --revit_version | $null; Problem = короткое сообщение | "" ; Years = хэш год -> число моделей }
function Resolve-RevitRun($Models, [string]$Template, $ServerYears) {
    $years = @{}
    $unknown = New-Object System.Collections.Generic.List[string]
    foreach ($m in @($Models | Where-Object { $_ } | Sort-Object -Unique)) {
        $r = Get-ModelRevitYear ([string]$m) $ServerYears
        if ($null -eq $r.Year) { $unknown.Add(("{0} ({1})" -f (Get-RevitLeaf $m), $r.Why)); continue }
        if ($years.ContainsKey($r.Year)) { $years[$r.Year]++ } else { $years[$r.Year] = 1 }
    }
    if ($unknown.Count -gt 0) {
        $first = ($unknown | Select-Object -First 2) -join "; "
        $more = ""; if ($unknown.Count -gt 2) { $more = (" и ещё {0}" -f ($unknown.Count - 2)) }
        return @{ Year = $null; Problem = ("Не удалось определить версию Revit: {0}{1}." -f $first, $more); Years = $years }
    }
    if ($years.Count -eq 0) { return @{ Year = $null; Problem = "Нет моделей, по которым определить версию Revit."; Years = $years } }
    if ($years.Count -gt 1) {
        $parts = @($years.Keys | Sort-Object | ForEach-Object { "{0} — {1}" -f $_, $years[$_] }) -join ", "
        return @{ Year = $null; Problem = ("Модели разных версий Revit ({0}). Запускайте отдельно по каждой версии." -f $parts); Years = $years }
    }
    $modelYear = [int]@($years.Keys)[0]
    if ($Template) {
        $t = Get-ModelRevitYear $Template $ServerYears
        if ($null -eq $t.Year) { return @{ Year = $null; Problem = ("Не удалось определить версию общего файла ({0})." -f $t.Why); Years = $years } }
        if ($t.Year -gt $modelYear) {
            return @{ Year = $null; Problem = ("Общий файл ({0}) новее моделей ({1}): Revit {0} обновил бы модели. Пересохраните общий файл в версии {1} или раньше." -f $t.Year, $modelYear); Years = $years }
        }
    }
    return @{ Year = $modelYear; Problem = ""; Years = $years }
}
