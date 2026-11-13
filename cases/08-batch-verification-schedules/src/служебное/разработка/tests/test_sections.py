"""Один файл разделы.txt: порядок листов, столбцы окна, Google, запасной список."""
import ast
import os
import sys
import tempfile
import unittest

sys.dont_write_bytecode = True

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(os.path.dirname(HERE)))
TOOL = os.path.join(ROOT, "служебное", "скрипты")
sys.path.insert(0, HERE)
sys.path.insert(0, TOOL)
from test_chain import pwsh, PWSH  # noqa: E402
import bvs_sections  # noqa: E402

SECTIONS_FILE = os.path.join(TOOL, "разделы.txt")
REPORT = ["EOM", "SS", "OV", "PD", "PT", "VK", "OT", "HS", "ITP", "AVT", "HC"]
COLUMNS = ["OV", "PD", "VK", "PT", "OT", "HS", "ITP", "EOM", "SS", "AVT", "HC"]
TITLES = {
    "EOM": "ЭОМ", "SS": "СС", "OV": "ОВ", "PD": "ПД", "PT": "ПТ", "VK": "ВК",
    "OT": "ОТ", "HS": "ХС", "ITP": "ИТП", "AVT": "АВТ", "HC": "ХЦ",
}


def parse_file(path):
    text = open(path, encoding="utf-8-sig").read()
    rows = []
    for raw in text.splitlines():
        line = raw.strip()
        if not line or line.startswith("#"):
            continue
        code, title, names, column = (line.split("|") + [""])[:4]
        rows.append((code.strip(), title.strip(), [n.strip() for n in names.split(",") if n.strip()], int(column)))
    return rows


def literal_section_names(path):
    src = open(path, encoding="utf-8").read()
    tree = ast.parse(src)
    for node in tree.body:
        if not isinstance(node, ast.Assign):
            continue
        for target in node.targets:
            if isinstance(target, ast.Name) and target.id == "SECTION_NAMES":
                return ast.literal_eval(ast.get_source_segment(src, node.value))
    raise AssertionError("нет SECTION_NAMES в " + path)


class SectionFileTests(unittest.TestCase):
    def test_file_keeps_report_order_column_order_and_titles(self):
        rows = parse_file(SECTIONS_FILE)
        self.assertEqual([r[0] for r in rows], REPORT)
        self.assertEqual([r[0] for r in sorted(rows, key=lambda r: r[3])], COLUMNS)
        self.assertEqual({r[0]: r[1] for r in rows}, TITLES)
        self.assertIn("ITP", [r[0] for r in rows])
        self.assertEqual(rows[REPORT.index("ITP")][1], "ИТП")

    def test_python_loader_matches_file_and_falls_back_without_crashing(self):
        rows, failed = bvs_sections.load_sections(SECTIONS_FILE)
        self.assertFalse(failed)
        self.assertEqual([r["code"] for r in rows], REPORT)
        self.assertEqual([r["code"] for r in bvs_sections.column_rows(rows)], COLUMNS)
        self.assertEqual(bvs_sections.column_rows(rows)[COLUMNS.index("ITP")]["title"], "ИТП")
        builtin_codes = [item[0] for item in bvs_sections.BUILTIN]
        self.assertEqual(builtin_codes, REPORT)
        missing = os.path.join(tempfile.mkdtemp(prefix="bvs_sec_"), "нет.txt")
        fallback, failed = bvs_sections.load_sections(missing)
        self.assertTrue(failed)
        self.assertEqual([r["code"] for r in fallback], REPORT)
        self.assertEqual([r["title"] for r in fallback], [TITLES[c] for c in REPORT])

    def test_former_python_literals_match_column_order(self):
        for name in ("insert_schedules.py", "set_assembly_code.py"):
            names = literal_section_names(os.path.join(TOOL, name))
            self.assertEqual([code for code, _variants in names], COLUMNS, name)
            self.assertIn("ITP", [code for code, _variants in names])

    def test_sibling_import_uses_script_dir(self):
        folder = bvs_sections.script_dir()
        self.assertEqual(os.path.normpath(folder), os.path.normpath(TOOL))
        self.assertTrue(os.path.isfile(bvs_sections.default_file()))


@unittest.skipUnless(PWSH, "нет pwsh")
class SectionPowerShellTests(unittest.TestCase):
    def run_ps(self, body):
        tmp = tempfile.mkdtemp(prefix="bvs_secps_")
        self.addCleanup(__import__("shutil").rmtree, tmp, True)
        script = os.path.join(tmp, "t.ps1")
        with open(script, "w", encoding="utf-8-sig") as handle:
            handle.write(body)
        return pwsh("-File", script)

    def test_report_google_stream_and_columns_match_the_file(self):
        r = self.run_ps("""
. '%s'
. '%s'
. '%s'
. '%s'
. '%s'
'report=' + ((Get-BvsReportCodes) -join ',')
'columns=' + ((Get-BvsColumnCodes) -join ',')
'google=' + ((Get-BvsGoogleSheetNames) -join ',')
'push=' + ($script:BvsGooglePushSheetNames -join ',')
'partial=' + ($script:PartialSectionOrder -join ',')
'pc=' + (($script:PcSections | ForEach-Object { $_.Code }) -join ',')
'stream=' + ($script:BvsKnownSectionCodes -join ',')
'itp-title=' + (Get-BvsTitleMap)['ITP']
'bf=' + (Get-ModelSectionFromName 'RG_ORG_BF_K1')
'gu=' + (Get-NonTargetSectionToken 'X_IOS_GU_K1')
""" % (os.path.join(TOOL, "bvs_sections.ps1").replace("'", "''"),
       os.path.join(TOOL, "project_config.ps1").replace("'", "''"),
       os.path.join(TOOL, "assignment_export_stream.ps1").replace("'", "''"),
       os.path.join(TOOL, "google_payload_lib.ps1").replace("'", "''"),
       os.path.join(TOOL, "assignment_partial_ui.ps1").replace("'", "''")))
        self.assertEqual(r.returncode, 0, r.stdout + r.stderr)
        lines = [ln for ln in r.stdout.splitlines() if "=" in ln]
        got = dict(ln.split("=", 1) for ln in lines)
        self.assertEqual(got["report"], ",".join(REPORT))
        self.assertEqual(got["columns"], ",".join(COLUMNS))
        self.assertEqual(got["google"], "Сводка," + ",".join(TITLES[c] for c in REPORT))
        self.assertEqual(got["push"], got["google"])
        self.assertEqual(got["partial"], ",".join(REPORT))
        self.assertEqual(got["pc"], ",".join(COLUMNS))
        self.assertEqual(got["stream"], ",".join(COLUMNS))
        self.assertEqual(got["itp-title"], "ИТП")
        self.assertEqual(got["bf"], "")
        self.assertEqual(got["gu"], "GU")

    def test_missing_file_falls_back_and_warns(self):
        r = self.run_ps("""
. '%s'
$w = @()
$rows = @(Read-BvsSections -Path '%s' -WarningVariable w -WarningAction Continue)
'codes=' + (($rows | ForEach-Object { $_.Code }) -join ',')
'warn=' + @($w).Count
""" % (os.path.join(TOOL, "bvs_sections.ps1").replace("'", "''"),
       os.path.join(tempfile.gettempdir(), "bvs-sections-missing.txt").replace("'", "''")))
        self.assertEqual(r.returncode, 0, r.stdout + r.stderr)
        text = r.stdout + r.stderr
        self.assertIn("codes=" + ",".join(REPORT), text)
        self.assertIn("встроенный список разделов", text)


if __name__ == "__main__":
    unittest.main()
