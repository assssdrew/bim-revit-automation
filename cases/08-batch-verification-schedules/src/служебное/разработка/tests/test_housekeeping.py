"""Временные файлы: что больше не создаётся, что чистится через 30 дней, что нельзя трогать."""
import os
import sys
import tempfile
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(os.path.dirname(HERE)))
TOOL = os.path.join(ROOT, "служебное", "скрипты")
sys.path.insert(0, HERE)
from test_chain import pwsh, PWSH  # noqa: E402

LAYOUT = os.path.join(TOOL, "report_layout.ps1")
UI = os.path.join(TOOL, "ui_schedules.ps1")
DIST = os.path.join(TOOL, "project_dist_ui.ps1")


def read(path):
    with open(path, encoding="utf-8-sig") as f:
        return f.read()


def func_text(src, name):
    i = src.index("function " + name)
    j = src.find("\nfunction ", i + 1)
    return src[i:] if j < 0 else src[i:j]


def run_ps(body):
    tmp = tempfile.mkdtemp(prefix="bvs_hk_")
    script = os.path.join(tmp, "t.ps1")
    with open(script, "w", encoding="utf-8-sig") as f:
        f.write(body)
    try:
        return pwsh("-File", script)
    finally:
        import shutil
        shutil.rmtree(tmp, True)


@unittest.skipUnless(PWSH, "нет pwsh")
class RetentionTests(unittest.TestCase):
    def test_old_ops_and_bad_xlsx_go_protected_files_stay(self):
        tmp = tempfile.mkdtemp(prefix="bvs_ret_")
        self.addCleanup(__import__("shutil").rmtree, tmp, True)
        rep = os.path.join(tmp, "отчеты")
        keep = []
        def put(rel, text="x"):
            path = os.path.join(rep, *rel.split("/"))
            os.makedirs(os.path.dirname(path), exist_ok=True)
            with open(path, "w", encoding="utf-8") as f:
                f.write(text)
            return path

        protected = [
            "проект/распределение_PROJ.txt",
            "проект/распределение_PROJ.txt.bak",
            "проект/выбор_PROJ.txt",
            "проект/google_PROJ.cfg",
            "проект/справочник_разделов.txt",
            "история/история_PROJ.txt",
            "история/история_PROJ_буфер.txt",
            "история/задание_журнал_PROJ.txt",
            "PROJ/2020-01-01_00-00-00_PROJ.xlsx",
            "PROJ/2020-01-01_00-00-00_PROJ_динамика.xlsx",
            "PROJ/2026-01-02_классификатор.xlsx",
            "2020-01-01.xlsx",
            "операции/дамп_фильтров_20261020_100000/ф.txt",
            "служебные/20261020_100000/google_push_payload.json",
            "PROJ/2026-10-20_00-00-00_PROJ.bad.xlsx",
        ]
        doomed = [
            "операции/дамп_фильтров_20250101_000000/ф.txt",
            "операции/сравнение_фильтров_20250101_000000.csv",
            "служебные/20250101_000000/google_push_payload.json",
            "PROJ/2020-01-01_00-00-00_PROJ.bad.xlsx",
        ]
        for rel in protected + doomed:
            put(rel, rel)
        r = run_ps("""
. '%s'
$now = [datetime]'2026-10-21 12:00'
$h = Invoke-BvsHousekeeping -ReportDir '%s' -KeepDays 30 -Now $now
'ops=' + ((@($h.Ops) | Sort-Object) -join ',')
'bad=' + ((@($h.Bad) | Sort-Object) -join ',')
""" % (LAYOUT.replace("'", "''"), rep.replace("'", "''")))
        self.assertEqual(r.returncode, 0, r.stdout + r.stderr)
        lines = [x for x in r.stdout.splitlines() if x.startswith("ops=") or x.startswith("bad=")]
        self.assertIn("дамп_фильтров_20250101_000000", lines[0])
        self.assertIn("сравнение_фильтров_20250101_000000.csv", lines[0])
        self.assertNotIn("20261020", lines[0])
        self.assertIn("2020-01-01_00-00-00_PROJ.bad.xlsx", lines[1])
        self.assertNotIn("2026-10-20", lines[1])
        for rel in protected:
            self.assertTrue(os.path.isfile(os.path.join(rep, *rel.split("/"))), rel)
        for rel in doomed:
            self.assertFalse(os.path.exists(os.path.join(rep, *rel.split("/"))), rel)

    def test_log_and_own_temp_errors_are_trimmed_other_temp_files_are_not(self):
        tmp = tempfile.mkdtemp(prefix="bvs_log_")
        self.addCleanup(__import__("shutil").rmtree, tmp, True)
        logdir = os.path.join(tmp, "логи")
        app = os.path.join(tmp, "app")
        os.makedirs(logdir)
        os.makedirs(app)
        big = b"HEAD\n" + (b"x" * 1200000) + b"\nTAIL-MARKER\n"
        for name in ("google_push.log", "partial_sets.log", "read_template_specs_error.txt", "dump_schedule_filters_error.txt"):
            with open(os.path.join(logdir, name), "wb") as f:
                f.write(big)
        leftover = os.path.join(app, "google_push.log")
        with open(leftover, "wb") as f:
            f.write(big)
        other = os.path.join(tmp, "someone_else_error.txt")
        with open(other, "wb") as f:
            f.write(big)
        r = run_ps("""
. '%s'
Trim-BvsToolLogs -LogDir '%s' -AppDir '%s'
""" % (LAYOUT.replace("'", "''"), logdir.replace("'", "''"), app.replace("'", "''")))
        self.assertEqual(r.returncode, 0, r.stdout + r.stderr)
        for name in ("google_push.log", "partial_sets.log", "read_template_specs_error.txt", "dump_schedule_filters_error.txt"):
            data = open(os.path.join(logdir, name), "rb").read()
            self.assertLess(len(data), 1048576, name)
            self.assertIn(b"TAIL-MARKER", data)
            self.assertNotIn(b"HEAD", data)
        data = open(leftover, "rb").read()
        self.assertLess(len(data), 1048576)
        self.assertIn(b"TAIL-MARKER", data)
        self.assertEqual(len(open(other, "rb").read()), len(big))


class RemovedFileTests(unittest.TestCase):
    def test_insert_csvs_only_when_the_report_is_read(self):
        """Четыре CSV читает отчёт вставки. Для «Посчитать», выгрузки и классификатора они не создаются."""
        ui = read(UI)
        fn = func_text(ui, "Write-RunFiles")
        self.assertIn("[switch]$InsertReport", fn)
        self.assertIn("if ($InsertReport)", fn)
        start = ui[ui.index("function Start-Batch"):]
        self.assertIn('-InsertReport:($Mode -eq "summary" -or $Mode -eq "sections")', start)
        if not PWSH:
            self.skipTest("нет pwsh")
        tmp = tempfile.mkdtemp(prefix="bvs_csv_")
        self.addCleanup(__import__("shutil").rmtree, tmp, True)
        app = os.path.join(tmp, "app")
        report = os.path.join(tmp, "work")
        os.makedirs(app)
        r = run_ps("""
%s
$AppDir = '%s'
$CfgPath = Join-Path $AppDir 'run.cfg'
$rep = '%s'
$a = Write-RunFiles -Models @('C:\\m\\A.rvt') -Source 'C:\\t.rvt' -ReportDir $rep -RunId 'r1' -Mode 'prefill' -ProjectPath '' -PrefixFrom '' -PrefixTo '' -InsertMode 'initial'
'prefill=' + @((Get-ChildItem -LiteralPath $rep -Filter '*.csv' -ErrorAction SilentlyContinue)).Count
$b = Write-RunFiles -Models @('C:\\m\\A.rvt') -Source 'C:\\t.rvt' -ReportDir $rep -RunId 'r2' -Mode 'summary' -ProjectPath '' -PrefixFrom '' -PrefixTo '' -InsertMode 'add_diff' -InsertReport
'insert=' + @((Get-ChildItem -LiteralPath $rep -Filter '*.csv')).Count
'names=' + ((Get-ChildItem -LiteralPath $rep -Filter '*.csv' | ForEach-Object Name | Sort-Object) -join ',')
""" % (fn, app.replace("'", "''"), report.replace("'", "''")))
        self.assertEqual(r.returncode, 0, r.stdout + r.stderr)
        lines = r.stdout.splitlines()
        self.assertIn("prefill=0", lines)
        self.assertIn("insert=4", lines)
        self.assertIn("r2_вставлено.csv", lines[-1])
        self.assertIn("r2_параметры.csv", lines[-1])
        self.assertIn("r2_пропуски.csv", lines[-1])
        self.assertIn("r2_сводка.csv", lines[-1])

    def test_count_uses_the_single_run_list(self):
        dist = read(DIST)
        count = func_text(dist, "Invoke-DistCount")
        wait = func_text(dist, "Invoke-RbpAndWait")
        self.assertIn("-FileList $files.List", count)
        self.assertNotIn("rvt_list_dist.txt", count)
        self.assertIn("[string]$FileList", wait)
        self.assertIn("if ($FileList)", wait)
        self.assertIn('Join-Path $AppDir "rvt_list_dist.txt"', wait)

    def test_template_read_shares_run_cfg_dump_keeps_its_own(self):
        dist = read(DIST)
        self.assertNotIn("template_read.cfg", dist)
        self.assertIn('Join-Path $AppDir "run.cfg"', dist)
        self.assertIn('"[template]"', dist)
        reader = read(os.path.join(TOOL, "read_template_specs.py"))
        self.assertIn('u"run.cfg"', reader)
        self.assertIn('u"template_read.cfg"', reader)   # запасной путь, окно его больше не пишет
        dump = read(os.path.join(ROOT, "служебное", "разработка", "инструменты", "dump_schedule_filters.ps1"))
        self.assertIn('Join-Path $AppDir "dump_filters.cfg"', dump)
        self.assertNotIn("run.cfg", dump)

    def test_export_writes_one_filter_file(self):
        export = read(os.path.join(TOOL, "export_schedules.py"))
        self.assertIn('фильтры_спецификаций.txt', export)
        self.assertNotIn("filters_debug.txt", export)
        build = read(os.path.join(TOOL, "build_assignment.ps1"))
        self.assertIn("GetFullPath($srcFilters)", build)
        self.assertIn('Join-Path $auxOut "фильтры_спецификаций.txt"', build)
        ui = read(UI)
        copy = func_text(ui, "Copy-RunDiagnosticsToProject")
        self.assertNotIn("Copy-Item", copy)



@unittest.skipUnless(PWSH, "нет pwsh")
class LogFolderTests(unittest.TestCase):
    """Журналы этого скрипта — в отчеты\\логи. Чужие логи RBP не трогаются."""

    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="bvs_logs_")
        self.addCleanup(__import__("shutil").rmtree, self.tmp, True)
        self.rep = os.path.join(self.tmp, "отчеты")
        self.app = os.path.join(self.tmp, "app")
        self.foreign = os.path.join(self.tmp, "RevitBatchProcessor")
        os.makedirs(self.app)
        os.makedirs(self.foreign)

    def run_body(self, body):
        r = run_ps(". '%s'\n%s" % (LAYOUT.replace("'", "''"), body))
        self.assertEqual(r.returncode, 0, r.stdout + r.stderr)
        return [l for l in r.stdout.splitlines() if l.strip()]

    def test_layout_creates_logs_folder(self):
        out = self.run_body("$r = Initialize-ReportLayout '%s'; $r.Layout.Logs" % self.rep.replace("'", "''"))
        self.assertTrue(out[0].endswith("логи") or out[0].endswith("логи"), out)
        self.assertTrue(os.path.isdir(os.path.join(self.rep, "логи")))
        self.assertTrue(os.path.isdir(os.path.join(self.rep, "история")))

    def test_quote_process_arg_only_when_spaces(self):
        out = self.run_body(r"""
'plain=' + (Quote-BvsProcessArg 'C:\BVS\script.py')
'space=' + (Quote-BvsProcessArg 'C:\My Tools\отчеты\логи')
""")
        self.assertEqual(out[0], r"plain=C:\BVS\script.py")
        self.assertEqual(out[1], r'space="C:\My Tools\отчеты\логи"')

    def test_old_appdir_logs_are_moved_content_kept(self):
        with open(os.path.join(self.app, "google_push.log"), "w", encoding="utf-8") as f:
            f.write("OLD-GOOGLE\n")
        with open(os.path.join(self.app, "partial_sets.log"), "w", encoding="utf-8") as f:
            f.write("OLD-PARTIAL\n")
        logs = os.path.join(self.rep, "логи")
        os.makedirs(logs)
        with open(os.path.join(logs, "google_push.log"), "w", encoding="utf-8") as f:
            f.write("NEW-GOOGLE\n")
        out = self.run_body("""
$h = Invoke-BvsHousekeeping -ReportDir '%s' -KeepDays 30 -AppDir '%s' -Now ([datetime]'2026-10-21')
'moved=' + ((@($h.Moved) | Sort-Object) -join ',')
""" % (self.rep.replace("'", "''"), self.app.replace("'", "''")))
        self.assertIn("google_push.log", out[0])
        self.assertIn("partial_sets.log", out[0])
        self.assertFalse(os.path.exists(os.path.join(self.app, "google_push.log")))
        self.assertFalse(os.path.exists(os.path.join(self.app, "partial_sets.log")))
        g = open(os.path.join(self.rep, "логи", "google_push.log"), encoding="utf-8-sig").read()
        self.assertIn("NEW-GOOGLE", g)
        self.assertIn("OLD-GOOGLE", g)
        self.assertIn("OLD-PARTIAL", open(os.path.join(self.rep, "логи", "partial_sets.log"), encoding="utf-8-sig").read())

    def test_old_batchrvt_logs_in_tool_folder_are_deleted_ui_logs_stay(self):
        logs = os.path.join(self.rep, "логи")
        os.makedirs(logs)
        def put(name, text="x"):
            with open(os.path.join(logs, name), "w", encoding="utf-8") as f:
                f.write(text)
        put("BatchRvt_20250101_000000_111.log", "old rbp")
        put("BatchRvt_20261020_100000_222.log", "recent rbp")
        put("google_push.log", "keep-google")
        put("partial_sets.log", "keep-partial")
        put("read_template_specs_error.txt", "keep-err")
        out = self.run_body("""
$h = Invoke-BvsHousekeeping -ReportDir '%s' -KeepDays 30 -Now ([datetime]'2026-10-21 12:00')
'logs=' + ((@($h.Logs) | Sort-Object) -join ',')
""" % self.rep.replace("'", "''"))
        self.assertIn("BatchRvt_20250101_000000_111.log", out[0])
        self.assertNotIn("20261020", out[0])
        self.assertFalse(os.path.exists(os.path.join(logs, "BatchRvt_20250101_000000_111.log")))
        self.assertTrue(os.path.isfile(os.path.join(logs, "BatchRvt_20261020_100000_222.log")))
        self.assertEqual(open(os.path.join(logs, "google_push.log"), encoding="utf-8").read(), "keep-google")
        self.assertEqual(open(os.path.join(logs, "partial_sets.log"), encoding="utf-8").read(), "keep-partial")
        self.assertTrue(os.path.isfile(os.path.join(logs, "read_template_specs_error.txt")))

    def test_foreign_rbp_logs_are_not_moved_or_deleted(self):
        """--log_folder только в запуске BVS; %LOCALAPPDATA%\\RevitBatchProcessor скрипт не читает и не чистит."""
        foreign_log = os.path.join(self.foreign, "BatchRvt_20250101_000000_999.log")
        with open(foreign_log, "w", encoding="utf-8") as f:
            f.write("FOREIGN\n")
        self.run_body("""
$h = Invoke-BvsHousekeeping -ReportDir '%s' -KeepDays 30 -AppDir '%s' -Now ([datetime]'2026-10-21 12:00')
[void]$h
# прямо указать чужую папку скрипт не должен: Remove-OldLogItems зовётся только для layout.Logs
""" % (self.rep.replace("'", "''"), self.app.replace("'", "''")))
        self.assertTrue(os.path.isfile(foreign_log))
        self.assertEqual(open(foreign_log, encoding="utf-8").read(), "FOREIGN\n")
        tool_logs = os.path.join(self.rep, "логи")
        if os.path.isdir(tool_logs):
            self.assertNotIn("BatchRvt_20250101_000000_999.log", os.listdir(tool_logs))

    def test_launchers_pass_log_folder_per_run_and_do_not_change_rbp_settings(self):
        ui = read(UI)
        dist = read(DIST)
        self.assertIn('"--log_folder", (Quote-BvsProcessArg $logDir)', ui)
        self.assertIn('"--log_folder", (Quote-BvsProcessArg $logDir)', dist)
        self.assertIn("Quote-BvsProcessArg $task", ui)
        self.assertIn("Quote-BvsProcessArg $files.List", ui)
        self.assertIn("Quote-BvsProcessArg $TaskScript", dist)
        self.assertIn("Quote-BvsProcessArg $list", dist)
        # глобальные настройки / папка RBP по умолчанию не переключаются
        for src, name in ((ui, "ui_schedules.ps1"), (dist, "project_dist_ui.ps1"), (read(LAYOUT), "report_layout.ps1")):
            for line in src.splitlines():
                s = line.strip()
                if s.startswith("#"):
                    continue
                self.assertNotIn("RevitBatchProcessor\\Logs", line, name)
                self.assertNotIn('Join-Path $env:LOCALAPPDATA "RevitBatchProcessor"', line, name + ": " + line)
        dump = read(os.path.join(ROOT, "служебное", "разработка", "инструменты", "dump_schedule_filters.ps1"))
        self.assertIn("--log_folder", dump)
        self.assertIn("log_folder=", dump)


if __name__ == "__main__":
    unittest.main()
