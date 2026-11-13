"""Раскладка папки «отчеты» (report_layout.ps1): отчеты\\ — xlsx, отчеты\\история\\ — постоянные файлы, отчеты\\служебные\\<время>\\ — файлы прогона
(автоочистка), отчеты\\операции\\ — отчёты служебных .bat, отчеты\\логи\\ — журналы этого инструмента. Миграция старой раскладки без потерь; история и «Динамика» продолжаются."""
import hashlib
import os
import re
import shutil
import subprocess
import sys
import tempfile
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(os.path.dirname(HERE)))
TOOL = os.path.join(ROOT, "служебное", "скрипты")      # папка с .ps1/.py окна
sys.path.insert(0, HERE)
import test_chain as chain  # noqa: E402
from test_chain import make_custom_export, pwsh, PWSH  # noqa: E402

LIB = os.path.join(TOOL, "report_layout.ps1")
BUILD = os.path.join(TOOL, "build_assignment.ps1")
MODEL = "PROJ_IOS_OV_ORG_K01_R24.rvt"


def read(p):
    with open(p, "rb") as f:
        return f.read()


def tree(root):
    out = {}
    for d, _, files in os.walk(root):
        for f in files:
            p = os.path.join(d, f)
            out[os.path.relpath(p, root)] = hashlib.sha256(read(p)).hexdigest()
    return out


def run_ps(tmp, body):
    script = os.path.join(tmp, "t.ps1")
    with open(script, "w", encoding="utf-8-sig") as f:
        f.write(". '%s'\n%s\n" % (LIB, body))
    r = pwsh("-File", script)
    assert r.returncode == 0, r.stdout + r.stderr
    return [l for l in r.stdout.splitlines() if l.strip()]


OLD_FILES = {
    "история_PROJ.txt": "hist\n", "история_PROJ_буфер.txt": "buf\n", "история_PROJ_проверки.txt": "chk\n", "история_PROJ_каталог.txt": "cat\n",
    "задание_история.txt": "old hist\n", "задание_журнал.txt": "ledger\n",
    "google_push_payload.json": "{}", "фильтры_спецификаций.txt": "filters\n",
    "исправление_фильтров_20261001_120000_apply.csv": "a;b\n", "замена_фильтра_20261001_120000_report.csv": "a;b\n",
    "приведение_к_эталону_20261016_112449_report.csv": "a;b\n", "сравнение_фильтров_20261016_120000.csv": "a;b\n",
    "2026-10-16_10-00-00_PROJ.xlsx": "xlsx", "2026-10-16_10-00-00_PROJ_динамика.xlsx": "dyn", "2026-10-15.xlsx": "step",
}


@unittest.skipUnless(PWSH, "нет pwsh")
class MigrationTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="bvs_lay_")
        self.addCleanup(shutil.rmtree, self.tmp, True)
        self.rep = os.path.join(self.tmp, "отчеты")
        os.makedirs(self.rep)
        for n, t in OLD_FILES.items():
            with open(os.path.join(self.rep, n), "w", encoding="utf-8", newline="") as f:
                f.write(t)
        os.makedirs(os.path.join(self.rep, "дамп_фильтров_20261016_092750"))
        with open(os.path.join(self.rep, "дамп_фильтров_20261016_092750", "фильтры_X.txt"), "w", encoding="utf-8", newline="") as f:
            f.write("dump\n")

    def init(self, extra=""):
        # фиксированное «сейчас»: файлы октября 2026 в этом наборе ещё не старше 30 дней
        args = extra if extra else "0 '' $now"
        return run_ps(self.tmp, "$now = [datetime]'2026-10-21 12:00'; $r = Initialize-ReportLayout '%s' %s; 'moves=' + @($r.Migration.Moves).Count; 'warn=' + @($r.Migration.Warnings).Count; 'removed=' + (@($r.Removed) -join ',')"
                      % (self.rep, args))

    def test_old_files_are_sorted_into_new_layout_without_loss(self):
        before = {n: hashlib.sha256(read(os.path.join(self.rep, n))).hexdigest() for n in OLD_FILES}
        out = self.init()
        self.assertEqual(out[0], "moves=13")                               # xlsx не переносятся
        hist = os.path.join(self.rep, "история")
        for n in ("история_PROJ.txt", "история_PROJ_буфер.txt", "история_PROJ_проверки.txt", "история_PROJ_каталог.txt", "задание_история.txt", "задание_журнал.txt"):
            self.assertEqual(hashlib.sha256(read(os.path.join(hist, n))).hexdigest(), before[n], n)      # содержимое то же
            self.assertFalse(os.path.exists(os.path.join(self.rep, n)))
        ops = os.path.join(self.rep, "операции")
        for n in ("исправление_фильтров_20261001_120000_apply.csv", "замена_фильтра_20261001_120000_report.csv",
                  "приведение_к_эталону_20261016_112449_report.csv", "сравнение_фильтров_20261016_120000.csv"):
            self.assertTrue(os.path.isfile(os.path.join(ops, n)), n)
        self.assertEqual(read(os.path.join(ops, "дамп_фильтров_20261016_092750", "фильтры_X.txt")), b"dump\n")
        old_aux = os.path.join(self.rep, "служебные", "_до_переноса")
        self.assertEqual(sorted(os.listdir(old_aux)), ["google_push_payload.json", "фильтры_спецификаций.txt"])
        self.assertEqual(sorted(n for n in os.listdir(self.rep) if n.endswith(".xlsx")),
                         ["2026-10-15.xlsx", "2026-10-16_10-00-00_PROJ.xlsx", "2026-10-16_10-00-00_PROJ_динамика.xlsx"])
        self.assertEqual(sorted(os.listdir(self.rep)), sorted(["история", "служебные", "операции", "логи", "2026-10-15.xlsx", "2026-10-16_10-00-00_PROJ.xlsx",
                                                              "2026-10-16_10-00-00_PROJ_динамика.xlsx"]))     # в корне остались только xlsx и папки
        # журнал переноса
        logs = [d for d in os.listdir(os.path.join(self.rep, "служебные")) if d.startswith("_перенос_")]
        self.assertEqual(len(logs), 1)
        self.assertEqual(len(read(os.path.join(self.rep, "служебные", logs[0], "перенос.txt")).decode("utf-8-sig").splitlines()), 13)

    def test_second_run_moves_nothing(self):
        self.init()
        snap = tree(self.rep)
        out = self.init()
        self.assertEqual(out[0], "moves=0")
        self.assertEqual(tree(self.rep), snap)

    def test_conflicts_keep_both_files_and_never_delete(self):
        os.makedirs(os.path.join(self.rep, "история"))
        with open(os.path.join(self.rep, "история", "история_PROJ.txt"), "w", encoding="utf-8", newline="") as f:
            f.write("NEW different\n")                                       # уже есть другой файл на месте
        with open(os.path.join(self.rep, "история", "задание_журнал.txt"), "w", encoding="utf-8", newline="") as f:
            f.write("ledger\n")                                              # такой же
        out = self.init()
        self.assertEqual(out[1], "warn=0")
        hist = os.path.join(self.rep, "история")
        self.assertEqual(read(os.path.join(hist, "история_PROJ.txt")), b"NEW different\n")
        conf = os.path.join(hist, "_конфликты_переноса")
        self.assertEqual(read(os.path.join(conf, "история_PROJ.txt")), b"hist\n")            # старый не потерян
        self.assertEqual(read(os.path.join(conf, "задание_журнал.txt")), b"ledger\n")
        self.assertEqual(read(os.path.join(hist, "задание_журнал.txt")), b"ledger\n")

    def test_xlsx_and_unknown_files_untouched(self):
        with open(os.path.join(self.rep, "заметка.txt"), "w", encoding="utf-8", newline="") as f:
            f.write("mine\n")
        self.init()
        self.assertEqual(read(os.path.join(self.rep, "заметка.txt")), b"mine\n")
        self.assertEqual(read(os.path.join(self.rep, "2026-10-15.xlsx")), b"step")


@unittest.skipUnless(PWSH, "нет pwsh")
class CleanupTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="bvs_clean_")
        self.addCleanup(shutil.rmtree, self.tmp, True)
        self.rep = os.path.join(self.tmp, "отчеты")
        self.aux = os.path.join(self.rep, "служебные")
        for d in ("20260901_100000", "20260920_100000", "20261010_100000", "20261020_100000", "20260801_000000_2", "ручная_папка"):
            os.makedirs(os.path.join(self.aux, d))
            with open(os.path.join(self.aux, d, "f.txt"), "w") as f:
                f.write("x")
        os.makedirs(os.path.join(self.rep, "история"))
        os.makedirs(os.path.join(self.rep, "операции", "дамп_фильтров_20250101_000000"))
        os.makedirs(os.path.join(self.rep, "операции", "дамп_фильтров_20261020_100000"))
        for p in (("история", "история_X.txt"), ("2020-01-01.xlsx",), ("операции", "дамп_фильтров_20250101_000000", "ф.txt"),
                  ("операции", "дамп_фильтров_20261020_100000", "ф.txt")):
            with open(os.path.join(self.rep, *p), "w") as f:
                f.write("keep")
        # у «ручной» папки дата изменения — давняя
        os.utime(os.path.join(self.aux, "ручная_папка"), (1, 1))

    def init(self, extra):
        return run_ps(self.tmp, "$now = [datetime]'2026-10-21 12:00'; $r = Initialize-ReportLayout '%s' %s; 'removed=' + ((@($r.Removed) | Sort-Object) -join ',') + ' keep=' + $r.KeepDays"
                      % (self.rep, extra))[0]

    def test_default_30_days_removes_only_old_aux_dirs(self):
        out = self.init("0 '' $now")
        self.assertEqual(out, "removed=20260801_000000_2,20260901_100000,20260920_100000,ручная_папка keep=30")
        self.assertEqual(sorted(os.listdir(self.aux)), ["20261010_100000", "20261020_100000"])
        self.assertTrue(os.path.isfile(os.path.join(self.rep, "история", "история_X.txt")))
        self.assertTrue(os.path.isfile(os.path.join(self.rep, "2020-01-01.xlsx")))          # итоговый xlsx не чистится, даже старый
        self.assertFalse(os.path.exists(os.path.join(self.rep, "операции", "дамп_фильтров_20250101_000000")))
        self.assertTrue(os.path.isfile(os.path.join(self.rep, "операции", "дамп_фильтров_20261020_100000", "ф.txt")))

    def test_setting_in_ui_settings_changes_retention(self):
        settings = os.path.join(self.tmp, "ui_settings.txt")
        with open(settings, "w", encoding="utf-8-sig") as f:
            f.write("source=x\nauxkeepdays=7\n")
        out = self.init("0 '%s' $now" % settings)
        self.assertEqual(out, "removed=20260801_000000_2,20260901_100000,20260920_100000,20261010_100000,ручная_папка keep=7")
        self.assertEqual(sorted(os.listdir(self.aux)), ["20261020_100000"])

    def test_bad_setting_falls_back_to_30(self):
        settings = os.path.join(self.tmp, "ui_settings.txt")
        with open(settings, "w", encoding="utf-8-sig") as f:
            f.write("auxkeepdays=0\n")
        self.assertTrue(self.init("0 '%s' $now" % settings).endswith("keep=30"))

    def test_new_run_dirs_and_bulk_cleanup(self):
        out = run_ps(self.tmp, "$d = New-RunAuxDir '%s' '20261021_120000'; $d2 = New-RunAuxDir '%s' '20261021_120000'; "
                     "$d.Substring($d.Length - 15); $d2.Substring($d2.Length - 17); "
                     "Set-Content -LiteralPath (Join-Path $d '0001.tsv') 'x'; Set-Content -LiteralPath (Join-Path $d 'manifest.txt') 'm'; "
                     "New-Item -ItemType Directory (Join-Path $d '_sections') | Out-Null; Remove-RunBulk $d; "
                     "(Get-ChildItem -LiteralPath $d | ForEach-Object Name) -join ','" % (self.aux, self.aux))
        self.assertEqual(out, ["20261021_120000", "20261021_120000_2", "manifest.txt"])


@unittest.skipUnless(PWSH, "нет pwsh")
class HistoryContinuesTests(unittest.TestCase):
    """Сборка в старой раскладке -> миграция -> сборка в новой: история, журнал, «Динамика» продолжаются; служебные файлы — только в служебной папке."""

    def build(self, hist, aux, n, items, dyn=False):
        import time
        time.sleep(1.1)
        exp = os.path.join(self.tmp, "exp%d" % n)
        make_custom_export(exp, [("П        4.1_Имя системы_Воздуховоды_Без значения", [(MODEL, "Сем", i) for i in items])], filt="ADSK_Размер")
        out_x = os.path.join(self.rep, "2026-10-2%d_PROJ.xlsx" % n)
        r = pwsh("-File", BUILD, "-ExportDir", exp, "-OutXlsx", out_x, "-HistoryPath", hist, "-AuxDir", aux)
        self.assertEqual(r.returncode, 0, r.stdout + r.stderr)
        return out_x

    def test_old_to_new_layout(self):
        self.tmp = tempfile.mkdtemp(prefix="bvs_hc_")
        self.addCleanup(shutil.rmtree, self.tmp, True)
        self.rep = os.path.join(self.tmp, "отчеты")
        os.makedirs(self.rep)
        old_hist = os.path.join(self.rep, "история_PROJ.txt")
        self.build(old_hist, os.path.join(self.rep), 1, ["Эл 1", "Эл 2", "Эл 3"])            # прежняя раскладка: всё в отчеты\
        self.assertTrue(os.path.isfile(os.path.join(self.rep, "google_push_payload.json")))
        led1 = read(os.path.join(self.rep, "задание_журнал_PROJ.txt")).decode("utf-8-sig").splitlines()
        hist1 = read(old_hist).decode("utf-8-sig").splitlines()
        out = run_ps(self.tmp, "$r = Initialize-ReportLayout '%s'; @($r.Migration.Moves).Count" % self.rep)
        self.assertGreaterEqual(int(out[0]), 3)                                               # история, журнал, payload (+фильтры)
        new_hist = os.path.join(self.rep, "история", "история_PROJ.txt")
        self.assertEqual(read(new_hist).decode("utf-8-sig").splitlines(), hist1)               # без потерь и дублей
        aux = os.path.join(self.rep, "служебные", "20261022_120000")
        os.makedirs(aux)
        x2 = self.build(new_hist, aux, 2, ["Эл 1", "Эл 2"])                                   # новая раскладка; «Эл 3» исправлен
        led2 = read(os.path.join(self.rep, "история", "задание_журнал_PROJ.txt")).decode("utf-8-sig").splitlines()
        self.assertEqual(len(led2), len(led1))
        self.assertFalse(os.path.exists(os.path.join(self.rep, "задание_журнал_PROJ.txt")))     # журнал проекта не «родился» заново в корне
        f1 = {l.split("\t")[4]: l.split("\t") for l in led1}
        f2 = {l.split("\t")[4]: l.split("\t") for l in led2}
        self.assertEqual(f2["Эл 1"][7], f1["Эл 1"][7])                                        # FirstSeen прежний
        self.assertEqual((f2["Эл 1"][9], f2["Эл 3"][9]), ("open", "fixed"))
        rows = [l.split("\t") for l in read(new_hist).decode("utf-8-sig").splitlines()]
        self.assertEqual(len({l[7] for l in rows}), 2)                                          # два среза — история продолжилась
        self.assertTrue(os.path.isfile(os.path.join(aux, "google_push_payload.json")))        # JSON для Google — в служебной папке
        self.assertFalse(os.path.exists(os.path.join(self.rep, "google_push_payload.json")))   # и не рядом с xlsx
        self.assertTrue(os.path.isfile(x2))
        # «Динамика» из истории в новом месте
        dyn_x = os.path.join(self.rep, "dyn.xlsx")
        r = pwsh("-File", os.path.join(TOOL, "build_assignment.ps1"), "-ExportDir", "", "-OutXlsx", dyn_x,
                 "-HistoryPath", new_hist, "-DynamicsFrom", "2026-01-01", "-DynamicsTo", "2030-01-01",
                 "-DynamicsOnly", "-AuxDir", os.path.join(self.rep, "служебные", "dyn"))
        self.assertEqual(r.returncode, 0, r.stdout + r.stderr)
        self.assertTrue(os.path.isfile(dyn_x))
        self.assertEqual(sorted(n for n in os.listdir(self.rep) if not n.endswith(".xlsx")), ["история", "логи", "операции", "служебные"])


def src(name):
    base = ROOT if name.startswith("служебное") else TOOL       # служебные .bat-скрипты лежат в служебное\, общие — в служебное\скрипты
    with open(os.path.join(base, name), encoding="utf-8-sig") as f:
        return f.read()


class WiringTests(unittest.TestCase):
    def test_ui_uses_new_layout(self):
        t = src("ui_schedules.ps1")
        self.assertIn('. (Join-Path $ToolDir "report_layout.ps1")', t)
        self.assertIn('Join-Path $script:Layout.History ("история_" + $code + ".txt")', t)
        self.assertEqual(t.count("Get-LedgerPath $summary"), 2)                                 # журнал — на проект: список «Выборочно» и запуск выгрузки
        self.assertIn('("задание_журнал_" + (Get-ProjectCode $Summary) + ".txt")', t)
        self.assertNotIn('Join-Path $ReportDir "задание_журнал.txt"', t)
        self.assertNotIn("verif_schedules_", t)                                                # файлы прогона больше не в %TEMP%
        self.assertIn("New-RunAuxDir $script:Layout.Aux $runId", t)
        self.assertIn("Invoke-BvsHousekeeping -ReportDir $ReportDir -KeepDays $script:AuxKeepDays -AppDir $AppDir", t)
        self.assertIn("-AuxDir $auxDir", t)                                                   # «Динамика»
        self.assertIn('-PayloadPath (Join-Path $WorkDir "google_push_payload.json")', t)
        self.assertIn("Remove-RunBulk $WorkDir", t)
        self.assertNotIn("Remove-Item -LiteralPath $WorkDir -Recurse", t)
        # настройка срока хранения читается и сохраняется
        self.assertIn("auxkeepdays", t)
        self.assertIn('("auxkeepdays={0}" -f $script:AuxKeepDays)', t)
        # миграция при запуске, после чтения настроек
        self.assertLess(t.index("$saved = Read-Settings"), t.index("Initialize-ReportLayout $ReportDir $script:AuxKeepDays"))
        # кнопка «Открыть отчёт» и путь к xlsx в статусе
        self.assertIn('$ui.OpenFolder.Text = "Открыть отчёт"', t)
        self.assertIn("Open-ReportFolder", t)
        self.assertIn("/select,", t)
        self.assertGreaterEqual(t.count("Set-OpenFolderTarget"), 4)
        self.assertIn("Файл: {0}", t)
        self.assertIn("--log_folder", t)
        self.assertIn("Quote-BvsProcessArg $logDir", t)
        self.assertIn("Quote-BvsProcessArg $task", t)
        self.assertIn("Quote-BvsProcessArg $files.List", t)
        self.assertIn("Open-BvsLogFolder", t)
        self.assertIn('$ui.OpenLogs.Text = "Логи"', t)
        self.assertIn("Initialize-ReportLayout $ReportDir $script:AuxKeepDays -AppDir $AppDir", t)

    def test_service_scripts_write_to_ops_dir(self):
        self.assertIn("layout.Ops", src(os.path.join("служебное", "разработка", "инструменты", "dump_schedule_filters.ps1")))
        self.assertIn("Layout.Ops", src(os.path.join("служебное", "разработка", "инструменты", "compare_filter_dumps.ps1")))

    def test_build_scripts_accept_aux_dir(self):
        t = src("build_assignment.ps1")
        self.assertIn("[string]$AuxDir = ''", t)
        self.assertIn('Join-Path $auxOut "google_push_payload.json"', t)
        self.assertIn('Join-Path $auxOut "фильтры_спецификаций.txt"', t)

    def test_ironpython_scripts_use_dotnet_for_paths(self):
        for name in ("export_schedules.py", "insert_schedules.py"):
            t = src(name)
            self.assertNotIn("os.path.isfile(", t.replace("return os.path.isfile(path)", ""), name)     # только запасной путь без .NET
            self.assertNotIn("os.makedirs(folder)", t, name)
        self.assertIn("lower_text(", src("insert_schedules.py"))
        self.assertNotIn(".lower()", src("insert_schedules.py").split("def lower_text")[0] + src("insert_schedules.py").split("def norm_path")[1])


if __name__ == "__main__":
    unittest.main()
