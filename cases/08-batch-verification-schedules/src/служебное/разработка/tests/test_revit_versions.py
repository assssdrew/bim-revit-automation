"""revit_versions.ps1: версия Revit по моделям (BasicFileInfo внутри .rvt, RSN:// — год из servers.cfg по хосту), проверка перед запуском RBP."""
import os
import shutil
import struct
import sys
import tempfile
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from test_chain import pwsh, PWSH  # noqa: E402

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(HERE)))
TOOL = os.path.join(ROOT, "служебное", "скрипты")
LIB = os.path.join(TOOL, "revit_versions.ps1")
END, FATSECT, FREE = 0xFFFFFFFE, 0xFFFFFFFD, 0xFFFFFFFF


def info_text(year, extra=""):
    t = ("Worksharing: Not enabled\r\nUsername: user\r\nRevit Build: Autodesk Revit %s (Build: 20210224_1515(x64))\r\nFormat: %s\r\n%s" % (year, year, extra))
    return b"\x00\x00\x00\x00" + t.encode("utf-16-le")


def entry(name, typ, start=0, size=0, child=-1):
    nm = (name + "\0").encode("utf-16-le")
    e = nm.ljust(64, b"\0") + struct.pack("<HBB", len(nm) if name else 0, typ if name else 0, 1) + struct.pack("<iii", -1, -1, child if name else -1)
    e += b"\0" * 16 + struct.pack("<I", 0) + b"\0" * 16 + struct.pack("<IQ", start if name else 0, size if name else 0)
    return e.ljust(128, b"\0")


def make_rvt(path, stream, name="BasicFileInfo", big=False, tail_mb=0):
    """Минимальный Microsoft Compound File (сектор 512 байт): FAT, каталог, малый поток через мини-FAT (или обычный, если big)."""
    sectors = {}
    if not big:
        # малый поток: мини-сектора 64 байта, начиная с третьего (проверка смещений); контейнер корневого элемента — обычные сектора с 3-го
        data = b"\0" * (3 * 64) + stream
        data += b"\0" * ((-len(data)) % 64)
        n_mini = len(data) // 64
        minifat = [FREE] * 128
        for i in range(3, 3 + (len(stream) + 63) // 64):
            minifat[i] = i + 1
        minifat[3 + (len(stream) + 63) // 64 - 1] = END
        n_cont = (len(data) + 511) // 512
        container = data.ljust(n_cont * 512, b"\0")
        root_start = 3
        fat = [FREE] * 128
        fat[0], fat[1], fat[2] = FATSECT, END, END
        for i in range(n_cont):
            fat[3 + i] = 3 + i + 1 if i < n_cont - 1 else END
        dir_s = entry("Root Entry", 5, root_start, len(data), child=1) + entry(name, 2, 3, len(stream)) + entry("", 0) + entry("", 0)
        body = struct.pack("<128I", *fat) + dir_s + struct.pack("<128I", *minifat) + container
        first_mini, n_mini_fat = 2, 1
    else:
        # обычный поток (>= 4096 байт) по цепочке FAT
        n = (len(stream) + 511) // 512
        fat = [FREE] * 128
        fat[0], fat[1] = FATSECT, END
        for i in range(n):
            fat[2 + i] = 2 + i + 1 if i < n - 1 else END
        dir_s = entry("Root Entry", 5, END, 0, child=1) + entry(name, 2, 2, len(stream)) + entry("", 0) + entry("", 0)
        body = struct.pack("<128I", *fat) + dir_s + stream.ljust(n * 512, b"\0")
        first_mini, n_mini_fat = END, 0
    difat = [0] + [FREE] * 108
    header = (b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\0" * 16 + struct.pack("<HHHHH", 0x3E, 3, 0xFFFE, 9, 6) + b"\0" * 6
              + struct.pack("<IIIIIIIIII", 0, 1, 1, 0, 4096, first_mini, n_mini_fat, END, 0, 0))
    header = header[:76] + struct.pack("<109I", *difat)
    assert len(header) == 512, len(header)
    with open(path, "wb") as f:
        f.write(header + body)
        if tail_mb:
            f.write(b"\xAA" * (tail_mb * 1024 * 1024))                  # «тяжёлая» модель: читаться должны только нужные сектора


def run(tmp, body):
    script = os.path.join(tmp, "t.ps1")
    with open(script, "w", encoding="utf-8-sig") as f:
        f.write(". '%s'\n%s\n" % (LIB, body))
    r = pwsh("-File", script)
    assert r.returncode == 0, r.stdout + r.stderr
    return [l for l in r.stdout.splitlines() if l.strip()]


@unittest.skipUnless(PWSH, "нет pwsh")
class RevitVersionTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="bvs_ver_")
        self.addCleanup(shutil.rmtree, self.tmp, True)

    def rvt(self, name, year, **kw):
        p = os.path.join(self.tmp, "Папка с пробелами", name)
        os.makedirs(os.path.dirname(p), exist_ok=True)
        make_rvt(p, info_text(year), **kw)
        return p

    def servers(self):
        p = os.path.join(self.tmp, "servers.cfg")
        with open(p, "w", encoding="utf-8") as f:
            f.write("# Format: host|display name|REST year\nrevit-server.example.local|Revit Server 2022|2022\nsrv-2024.example|Revit Server 2024|2024\n")
        return p

    def test_year_is_read_from_file_not_from_name_or_path(self):
        a = self.rvt("PROJ_IOS_OV_ORG_K01_R24.rvt", 2022)             # имя говорит 2024 (_R24), файл — 2022
        b = self.rvt("Модель_Revit_2019.rvt", 2024)
        big = os.path.join(self.tmp, "big.rvt")
        make_rvt(big, info_text(2023, extra="x" * 9000), big=True)       # поток >= 4096 — по цепочке FAT
        heavy = os.path.join(self.tmp, "heavy.rvt")
        make_rvt(heavy, info_text(2021), tail_mb=30)                    # 30 МБ после служебных секторов
        out = run(self.tmp, r"""
foreach ($p in @('%s','%s','%s','%s')) { $r = Get-RvtFileYear $p; [string]$r.Year + '|' + $r.Why }
""" % (a, b, big, heavy))
        self.assertEqual(out, ["2022|", "2024|", "2023|", "2021|"])

    def test_unreadable_files_are_reported_not_guessed(self):
        bad = os.path.join(self.tmp, "not_revit_R24.rvt")
        open(bad, "wb").write(b"just text " * 100)
        nostream = os.path.join(self.tmp, "nostream_R24.rvt")
        make_rvt(nostream, info_text(2022), name="Other")
        noyear = os.path.join(self.tmp, "noyear.rvt")
        make_rvt(noyear, b"\0\0\0\0" + "Worksharing: Not enabled".encode("utf-16-le"))
        out = run(self.tmp, r"""
foreach ($p in @('%s','%s','%s','%s')) { $r = Get-RvtFileYear $p; '[' + [string]$r.Year + ']' + $r.Why }
""" % (bad, nostream, noyear, os.path.join(self.tmp, "нет_R24.rvt")))
        self.assertEqual(out[0], "[]нет BasicFileInfo (не файл Revit?)")
        self.assertEqual(out[1], "[]нет BasicFileInfo (не файл Revit?)")
        self.assertEqual(out[2], "[]в BasicFileInfo нет версии Revit")
        self.assertEqual(out[3], "[]файл не найден")

    def test_rsn_year_from_servers_cfg_by_host(self):
        cfg = self.servers()
        out = run(self.tmp, r"""
$y = Get-BvsServerYears '%s'
'count: ' + $y.Count
foreach ($p in @('RSN://revit-server.example.local/Проекты/PROJ_IOS_OV_ORG_K01_R24.rvt', 'rsn://SRV-2024.example/x/y.rvt', 'RSN://unknown.host/a.rvt')) { $r = Get-ModelRevitYear $p $y; $r.Source + '|' + [string]$r.Year + '|' + $r.Why }
""" % cfg)
        self.assertEqual(out[0], "count: 2")
        self.assertEqual(out[1], "server|2022|")
        self.assertEqual(out[2], "server|2024|")
        self.assertEqual(out[3], "server||сервера unknown.host нет в servers.cfg")

    def test_run_check_same_year_template_older_or_equal_is_fine(self):
        m1 = self.rvt("PROJ_IOS_OV_ORG_K01_R24.rvt", 2024)
        m2 = self.rvt("PROJ_IOS_VK_ORG_K01_R24.rvt", 2024)
        t22 = self.rvt("общий_2022.rvt", 2022)
        t24 = self.rvt("общий_2024.rvt", 2024)
        out = run(self.tmp, r"""
$y = @{}
foreach ($t in @('%s', '%s', '')) { $r = Resolve-RevitRun @('%s', '%s') $t $y; [string]$r.Year + '|' + $r.Problem }
""" % (t22, t24, m1, m2))
        self.assertEqual(out, ["2024|", "2024|", "2024|"])                      # общий файл старше или той же версии — норма; без общего файла — по моделям

    def test_run_check_stops_on_mixed_versions_unknown_and_newer_template(self):
        m22 = self.rvt("PROJ_IOS_OV_ORG_K01_R24.rvt", 2022)
        m24 = self.rvt("PROJ_IOS_VK_ORG_K01_R24.rvt", 2024)
        t24 = self.rvt("общий_2024.rvt", 2024)
        cfg = self.servers()
        junk = os.path.join(self.tmp, "junk.rvt")
        open(junk, "wb").write(b"x" * 600)
        out = run(self.tmp, r"""
$y = Get-BvsServerYears '%s'
$r = Resolve-RevitRun @('%s', '%s') '' $y; 'mixed: ' + [string]$r.Year + '|' + $r.Problem
$r = Resolve-RevitRun @('%s', '%s') '' $y; 'unknown: ' + [string]$r.Year + '|' + $r.Problem
$r = Resolve-RevitRun @('RSN://revit-server.example.local/p/a.rvt', 'RSN://revit-server.example.local/p/b.rvt') '%s' $y; 'newer: ' + [string]$r.Year + '|' + $r.Problem
$r = Resolve-RevitRun @('%s') '%s' $y; 'tpl unknown: ' + [string]$r.Year + '|' + $r.Problem
$r = Resolve-RevitRun @() '' $y; 'none: ' + $r.Problem
""" % (cfg, m22, m24, m22, junk, t24, m22, junk))
        self.assertEqual(out[0], "mixed: |Модели разных версий Revit (2022 — 1, 2024 — 1). Запускайте отдельно по каждой версии.")
        self.assertIn("unknown: |Не удалось определить версию Revit: junk.rvt (нет BasicFileInfo (не файл Revit?)).", out[1])
        self.assertEqual(out[2], "newer: |Общий файл (2024) новее моделей (2022): Revit 2024 обновил бы модели. Пересохраните общий файл в версии 2022 или раньше.")
        self.assertIn("Не удалось определить версию общего файла", out[3])
        self.assertEqual(out[4], "none: Нет моделей, по которым определить версию Revit.")

    def test_rsn_models_2022_with_2024_file_scenario(self):
        """Общий файл в Revit 2024, модели на RSN:// 2022: раньше открывался Revit 2024 — теперь остановка с сообщением."""
        t24 = self.rvt("общий_2024.rvt", 2024)
        cfg = self.servers()
        out = run(self.tmp, r"""
$y = Get-BvsServerYears '%s'
$r = Resolve-RevitRun @('RSN://revit-server.example.local/Проект/PROJ_IOS_OV_ORG_K01_R24.rvt') '%s' $y
[string]$r.Year + '|' + $r.Problem
""" % (cfg, t24))
        self.assertTrue(out[0].startswith("|Общий файл (2024) новее моделей (2022)"), out[0])


class WiringTests(unittest.TestCase):
    def read(self, p):
        with open(os.path.join(TOOL, p), encoding="utf-8-sig") as f:
            return f.read()

    def test_year_is_never_guessed_by_path_and_always_passed(self):
        ui, dist = self.read("ui_schedules.ps1"), self.read("project_dist_ui.ps1")
        self.assertNotIn("Guess-Year", ui)
        self.assertNotIn("Guess-Year", dist)
        self.assertIn('. (Join-Path $ToolDir "revit_versions.ps1")', ui)
        self.assertIn('"--revit_version", [string]$ver.Year', ui)                      # всегда, а не «если одна версия угадана»
        self.assertNotIn("$years.Count -eq 1", ui + dist)
        self.assertIn('"--revit_version", $Year', dist)
        self.assertIn('if (-not $Year) { throw', dist)

    def test_check_runs_before_launch_on_every_step(self):
        ui, dist = self.read("ui_schedules.ps1"), self.read("project_dist_ui.ps1")
        start = ui[ui.index("function Start-Batch"):]
        self.assertLess(start.index("Resolve-RevitRun"), start.index("Write-RunFiles"))          # Сводная, Разделы, Задание
        self.assertIn('if ($Mode -eq "summary" -or $Mode -eq "sections") { $verTemplate = $source }', start)                 # общий файл участвует в проверке, когда его открывают
        self.assertIn("return", start[start.index("if ($ver.Problem)"):start.index("$runId = Get-Date")])
        load = dist[dist.index("function Invoke-DistLoad"):dist.index("function Invoke-DistAssign")]
        self.assertLess(load.index("Resolve-RevitRun"), load.index("Invoke-RbpAndWait"))
        assign = dist[dist.index("function Invoke-DistAssign"):dist.index("function Invoke-DistCellClick")]
        self.assertLess(assign.index("Resolve-RevitRun"), assign.index("Invoke-RbpAndWait"))
        for call in ("Invoke-RbpAndWait (Join-Path $ToolDir \"read_template_specs.py\")", "Invoke-RbpAndWait (Join-Path $ToolDir \"insert_schedules.py\")"):
            self.assertIn(call, dist)

    def test_selfcheck_knows_the_module_and_servers_cfg_comes_from_root(self):
        ui = self.read("ui_schedules.ps1")
        self.assertIn('"revit_versions.ps1"', ui[:ui.index("SELFCHECK OK")])
        self.assertIn('$ServersCfgPath = Join-Path (Get-BvsRootDir $ToolDir) "servers.cfg"', ui)


if __name__ == "__main__":
    unittest.main()
