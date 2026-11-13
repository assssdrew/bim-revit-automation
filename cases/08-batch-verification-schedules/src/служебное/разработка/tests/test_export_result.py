"""Итог выгрузки из сводной — один файл export.txt ([result] / [manifest] / [catalog]) вместо manifest.txt + export_result.txt + catalog.txt.

Проверяется: чтение обоих форматов (bvs_export_result.ps1) даёт одно и то же; задание, payload Google, история, журнал, буфер
и <история>_каталог.txt из export.txt — те же, что из прежних трёх файлов (выгрузки прежних версий читаются как раньше);
«Выгрузка оборвалась» — по наличию итога в любом формате."""
import json
import os
import shutil
import subprocess
import sys
import tempfile
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from test_chain import pwsh, PWSH  # noqa: E402
from baseline_lib import dump_xlsx, run_pwsh  # noqa: E402
import test_partial as tp  # noqa: E402

TOOL = tp.TOOL
MOD = os.path.join(TOOL, "bvs_export_result.ps1")
OLD = ("manifest.txt", "export_result.txt", "catalog.txt")


def q(s):
    return str(s).replace("'", "''")


def read_lines(path):
    with open(path, encoding="utf-8-sig", newline="") as f:
        text = f.read()
    return [l for l in text.replace("\r\n", "\n").split("\n") if l.strip()]


def to_merged(folder):
    """Прежние три файла папки выгрузки -> export.txt в формате export_schedules.py (как он пишет сейчас)."""
    res = read_lines(os.path.join(folder, "export_result.txt")) if os.path.exists(os.path.join(folder, "export_result.txt")) else []
    out = ["[result]"] + res + ["[manifest]"] + read_lines(os.path.join(folder, "manifest.txt"))
    if os.path.exists(os.path.join(folder, "catalog.txt")):
        out += ["[catalog]"] + read_lines(os.path.join(folder, "catalog.txt"))
    for name in OLD:
        p = os.path.join(folder, name)
        if os.path.exists(p):
            os.remove(p)
    with open(os.path.join(folder, "export.txt"), "w", encoding="utf-8-sig", newline="") as f:
        f.write("\r\n".join(out) + "\r\n")


def outputs(out):
    """Всё, что видит пользователь после сборки: xlsx (значения ячеек), payload, история/журнал/буфер/проверки/каталог."""
    got = {}
    for name in sorted(os.listdir(out)):
        p = os.path.join(out, name)
        if name.endswith(".xlsx"):
            got[name] = [(s["name"], [(c.get("r"), c.get("v")) for c in s["cells"]]) for s in dump_xlsx(p)["sheets"]]
        elif name.endswith(".json"):
            with open(p, encoding="utf-8") as f:
                got[name] = json.load(f)
        elif os.path.isfile(p):
            with open(p, "rb") as f:
                data = f.read()
            # порядок строк истории/журнала зависит от порядка ключей хэш-таблицы (в pwsh 7 он меняется от запуска к запуску)
            # тестовая выгрузка пишет прежние файлы с LF, export_schedules.py — с CRLF; сравниваются строки
            bom = data.startswith(b"\xef\xbb\xbf")
            got[name] = (bom, sorted(l.rstrip(b"\r") for l in data[3 if bom else 0:].split(b"\n")))
    return got


@unittest.skipUnless(PWSH, "нет pwsh")
class ReadBothFormatsTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="bvs_expres_")
        self.addCleanup(shutil.rmtree, self.tmp, True)

    def ps(self, body):
        script = os.path.join(self.tmp, "t.ps1")
        with open(script, "w", encoding="utf-8-sig") as f:
            f.write(". '%s'\n%s" % (q(MOD), body))
        r = pwsh("-File", script)
        self.assertEqual(r.returncode, 0, r.stdout + r.stderr)
        return [l for l in r.stdout.splitlines() if l.strip()]

    def test_same_content_from_both_formats_and_same_history_catalog(self):
        legacy, merged = os.path.join(self.tmp, "legacy"), os.path.join(self.tmp, "merged")
        for d in (legacy, merged):
            tp.write_export(d, tp.ROWS, [(tp.E1, tp.S1)], tp.ALL_CAT_SPECS, tp.ALL_CAT_MODELS + [("X_UNL", "unloaded")])
            with open(os.path.join(d, "export_result.txt"), "w", encoding="utf-8-sig", newline="") as f:
                f.write("matched=3\r\nexported=1\r\nskipped=1\r\npartial=1\r\npairs=1\r\nmissing_specs=\r\nunloaded_models=X_UNL\r\n")
        to_merged(merged)
        self.assertEqual(sorted(os.listdir(merged)), ["0001.tsv", "0002.tsv", "export.txt"])
        out = self.ps(r"""
foreach ($d in @('%s', '%s')) {
    $e = Read-BvsExportResult $d
    'has=' + $e.Has + ' fmt=' + $e.Format + ' cat=' + $e.HasCatalog + ' done=' + (Test-BvsExportDone $d)
    'manifest=' + (@($e.Manifest | Where-Object { $_.Trim() }) -join '|')
    'result=' + ($e.ResultText -replace "`r`n", ';') + ' exported=' + $e.Result['exported'] + ' unloaded=' + $e.Result['unloaded_models']
    'catalog=' + (@($e.Catalog | Where-Object { $_.Trim() }) -join '|')
    [void](Save-BvsExportCatalog $e (Join-Path $d 'hist_каталог.txt'))
}
$empty = Join-Path '%s' 'empty'; New-Item -ItemType Directory -Path $empty | Out-Null
$e = Read-BvsExportResult $empty
'empty has=' + $e.Has + ' done=' + (Test-BvsExportDone $empty) + ' save=' + (Save-BvsExportCatalog $e (Join-Path $empty 'h.txt')) + ' file=' + (Test-Path (Join-Path $empty 'h.txt'))
""" % (q(legacy), q(merged), q(self.tmp)))
        self.assertEqual(out[0], "has=True fmt=legacy cat=True done=True")
        self.assertEqual(out[4], "has=True fmt=export.txt cat=True done=True")
        self.assertEqual(out[1:4], out[5:8])                       # одно и то же содержимое
        self.assertIn("*scope*\tpartial", out[1])
        self.assertIn("exported=1 unloaded=X_UNL", out[2])
        self.assertEqual(out[8], "empty has=False done=False save=False file=False")
        # постоянный файл истории <история>_каталог.txt — в прежнем формате; из export.txt — как writer прежнего catalog.txt (CRLF, BOM)
        with open(os.path.join(merged, "hist_каталог.txt"), "rb") as f:
            from_merged = f.read()
        self.assertTrue(from_merged.startswith(b"\xef\xbb\xbf"))
        self.assertEqual(from_merged.decode("utf-8-sig").replace("\r\n", "\n"), open(os.path.join(legacy, "hist_каталог.txt"), encoding="utf-8-sig").read().replace("\r\n", "\n"))

    def test_legacy_without_catalog_does_not_create_history_catalog(self):
        d = os.path.join(self.tmp, "old")
        tp.write_export(d, tp.ROWS)
        out = self.ps("$e = Read-BvsExportResult '%s'\n'cat=' + $e.HasCatalog + ' save=' + (Save-BvsExportCatalog $e (Join-Path '%s' 'h.txt'))" % (q(d), q(d)))
        self.assertEqual(out, ["cat=False save=False"])

    def test_read_once_per_build_and_new_export_is_reread(self):
        d = os.path.join(self.tmp, "c")
        tp.write_export(d, tp.ROWS)
        to_merged(d)
        out = self.ps(r"""
$a = Get-BvsExportResult '%s'; $b = Get-BvsExportResult '%s'
'same=' + [object]::ReferenceEquals($a, $b)
Start-Sleep -Milliseconds 30
[System.IO.File]::WriteAllText((Join-Path '%s' 'export.txt'), "[result]`r`nexported=7`r`n[manifest]`r`n", (New-Object System.Text.UTF8Encoding $true))
$c = Get-BvsExportResult '%s'
'reread=' + (-not [object]::ReferenceEquals($a, $c)) + ' exported=' + $c.Result['exported'] + ' cat=' + $c.HasCatalog
""" % (q(d), q(d), q(d), q(d)))
        self.assertEqual(out, ["same=True", "reread=True exported=7 cat=False"])


@unittest.skipUnless(PWSH, "нет pwsh")
class SameReportFromBothFormatsTests(unittest.TestCase):
    """Одна и та же выгрузка в двух форматах -> побайтно те же история/журнал/буфер/каталог, те же ячейки xlsx и payload."""

    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="bvs_expres2_")
        self.addCleanup(shutil.rmtree, self.tmp, True)

    def run_sequence(self, merged, steps):
        out = os.path.join(self.tmp, "merged" if merged else "legacy", "out")
        os.makedirs(out)
        results = []
        for i, (now, rows, pairs, flag) in enumerate(steps):
            exp = os.path.join(self.tmp, "merged" if merged else "legacy", "exp%d" % i)
            tp.write_export(exp, rows, pairs, tp.ALL_CAT_SPECS, tp.ALL_CAT_MODELS)
            if merged:
                to_merged(exp)
            args = ["-File", os.path.join(TOOL, "build_assignment.ps1"), "-ExportDir", exp, "-OutXlsx", os.path.join(out, "task.xlsx")]
            if flag:
                args.append("-Partial")
            r = run_pwsh(PWSH, args, {"BVS_FAKE_NOW": now})
            results.append((r.returncode, (r.stdout + r.stderr).replace(exp, "<exp>").replace(out, "<out>")))
            if r.returncode == 0:
                results.append(outputs(out))
        return results

    def test_full_partial_full_and_mode_mismatch(self):
        rows2 = {tp.S1: [("Светильник: тип А", "Светильник 1", "", tp.E1)], tp.S2: tp.ROWS[tp.S2]}
        steps = [("2026-02-01 10:00:00", tp.ROWS, None, False),
                 ("2026-02-02 10:00:00", rows2, [(tp.E1, tp.S1)], True),
                 ("2026-02-03 10:00:00", tp.ROWS, [(tp.E1, tp.S1)], False),          # выборочная выгрузка, сборка «Всё» — отказ
                 ("2026-02-03 11:00:00", tp.ROWS, None, True),                       # полная выгрузка, сборка «Выборочно» — отказ
                 ("2026-02-04 10:00:00", {tp.S1: [], tp.S2: []}, None, False)]
        legacy = self.run_sequence(False, steps)
        merged = self.run_sequence(True, steps)
        self.assertEqual([x[0] for x in legacy if isinstance(x, tuple)], [0, 0, 1, 1, 0])
        self.assertEqual(len(legacy), len(merged))
        for a, b in zip(legacy, merged):
            if isinstance(a, tuple):
                self.assertEqual(a[0], b[0])
                if a[0] != 0:
                    for phrase in ("Выгрузка выборочная (assign_scope=partial)", "Режим «Выборочно», но выгрузка полная"):
                        self.assertEqual(phrase in a[1], phrase in b[1], phrase)
            else:
                self.assertEqual(sorted(a), sorted(b))
                for k in a:
                    self.assertEqual(a[k], b[k], k)
        self.assertIn("Выгрузка выборочная (assign_scope=partial)", merged[4][1])
        self.assertIn("(в итоге выгрузки нет *scope*)", merged[5][1])

    def test_empty_export_note(self):
        """«Строк не выгружено» (exported=0) — из [result] так же, как из export_result.txt."""
        for merged in (False, True):
            base = os.path.join(self.tmp, "m" if merged else "l")
            exp, out = os.path.join(base, "exp"), os.path.join(base, "out")
            os.makedirs(out)
            subprocess.check_call([sys.executable, os.path.join(HERE, "make_sample_export.py"), exp, "--empty"])
            os.remove(os.path.join(base, "задание_история.txt"))          # без истории: пустой отчёт
            if merged:
                to_merged(exp)
            xlsx = os.path.join(out, "task.xlsx")
            r = pwsh("-File", os.path.join(TOOL, "build_assignment.ps1"), "-ExportDir", exp, "-OutXlsx", xlsx)
            self.assertEqual(r.returncode, 0, r.stdout + r.stderr)
            import openpyxl
            a1 = openpyxl.load_workbook(xlsx)["Сводка"]["A1"].value
            self.assertTrue(str(a1).startswith("Строк не выгружено"), (merged, a1))


class SourceTests(unittest.TestCase):
    def test_no_reader_opens_old_files_directly(self):
        """Прежние имена читает только bvs_export_result.ps1 (для выгрузок прежних версий)."""
        for name in os.listdir(TOOL):
            if not name.endswith((".ps1", ".py")) or name == "bvs_export_result.ps1":
                continue
            with open(os.path.join(TOOL, name), encoding="utf-8-sig") as f:
                src = f.read()
            for old in OLD:
                for pat in ('"%s"' % old, "'%s'" % old):
                    self.assertNotIn(pat, src, "%s: %s" % (name, old))

    def test_exporter_writes_one_result_file_last(self):
        with open(os.path.join(TOOL, "export_schedules.py"), encoding="utf-8") as f:
            src = f.read()
        tail = src[src.index("def main"):]
        self.assertEqual(tail.count("write_utf8("), 2)          # фильтры_спецификаций.txt и export.txt
        self.assertLess(tail.index("фильтры_спецификаций.txt"), tail.index("EXPORT_RESULT_FILE"))
        self.assertIn('EXPORT_RESULT_FILE = u"export.txt"', src)


if __name__ == "__main__":
    unittest.main()
