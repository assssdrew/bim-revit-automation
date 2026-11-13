"""SampleSite: разделы АВТ и ХЦ (по коду и по папке), справочник дописывается один раз, Google без полей адреса/секрета,
выделение узла в дереве «Выборочно», отчёты по папкам проектов; АВТ3 работает как раньше."""
import os
import re
import shutil
import sys
import tempfile
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from test_chain import pwsh, PWSH  # noqa: E402

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(HERE)))
TOOL = os.path.join(ROOT, "служебное", "скрипты")
SERVICE = os.path.join(ROOT, "служебное")
UI = os.path.join(TOOL, "ui_schedules.ps1")
DIST = os.path.join(TOOL, "project_dist_ui.ps1")
CFG = os.path.join(TOOL, "project_config.ps1")
GCFG = os.path.join(TOOL, "google_push_cfg.ps1")
LAYOUT = os.path.join(TOOL, "report_layout.ps1")
STREAM = os.path.join(TOOL, "assignment_export_stream.ps1")


def read(p):
    with open(p, encoding="utf-8-sig") as f:
        return f.read()


def q(s):
    return str(s).replace("'", "''")


def func_text(src, name):
    m = re.search(r"^function %s\b.*?^\}" % re.escape(name), src, re.M | re.S)
    assert m, name
    return m.group(0)


class PsBase(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="bvs_sg_")
        self.addCleanup(shutil.rmtree, self.tmp, True)

    def run_ps(self, body):
        script = os.path.join(self.tmp, "t.ps1")
        with open(script, "w", encoding="utf-8-sig") as f:
            f.write(body)
        r = pwsh("-File", script)
        self.assertEqual(r.returncode, 0, r.stdout + r.stderr)
        return [l for l in r.stdout.splitlines() if l.strip()]


@unittest.skipUnless(PWSH, "нет pwsh")
class SectionTests(PsBase):
    def test_samplesite_models_get_avt_and_hc_including_prk(self):
        out = self.run_ps(". '%s'\n" % q(CFG) + r"""
$names = @('SITE_ORG_AVT_K1_R22','SITE_ORG_AVT_K2_R22','SITE_ORG_AVT_K3_R22','SITE_ORG_AVT_K4_R22','SITE_ORG_AVT_K5_R22','SITE_ORG_AVT_PRK_R22',
  'SITE_ORG_HC_K1_R22','SITE_ORG_HC_K2_R22','SITE_ORG_HC_K3_R22','SITE_ORG_HC_K4_R22')
($names | ForEach-Object { $_ + '=' + (Get-ModelSectionFromName $_) }) -join ' '
$c = New-ProjectConfig 'П'
Update-ProjectModels $c $names
'sections: ' + (($c.Sections | ForEach-Object { $_.Code + ':' + $_.Title }) -join ',')
'models: ' + $c.Models.Count
""")
        self.assertEqual(out[0], " ".join(["SITE_ORG_AVT_K%d_R22=AVT" % i for i in range(1, 6)] + ["SITE_ORG_AVT_PRK_R22=AVT"] + ["SITE_ORG_HC_K%d_R22=HC" % i for i in range(1, 5)]))
        self.assertEqual(out[1], "sections: AVT:АВТ,HC:ХЦ")                       # в таблице только разделы проекта
        self.assertEqual(out[2], "models: 10")

    def test_fallback_by_folder_and_unknown_column(self):
        out = self.run_ps(". '%s'\n" % q(CFG) + r"""
'folder avt: ' + (Get-ModelSectionFromName 'C:\m\AVT\SITE_X_K1.rvt')
'folder hc ru: ' + (Get-ModelSectionFromName 'C:\m\ХЦ\SITE_X_K1.rvt')
'name beats folder: ' + (Get-ModelSectionFromName 'C:\m\AVT\SITE_ORG_HC_K1_R22.rvt')
'no code: [' + (Get-ModelSectionFromName 'C:\m\SITE_X_K1.rvt') + ']'
'task not folder: [' + (Get-ModelSectionFromName 'C:\m\AVT\SITE_TASK_K1.rvt') + ']'
$c = New-ProjectConfig 'П'
Update-ProjectModels $c @('SITE_ORG_AVT_K1_R22', 'SITE_ORG_XYZ_K1_R22', 'RG_ORG_BF_K1', 'RG_ORG_GU_K1')
'sections: ' + (($c.Sections | ForEach-Object { $_.Code }) -join ',')
'xyz in: ' + (@($c.Models | Where-Object { $_.Name -eq 'SITE_ORG_XYZ_K1_R22' }).Count)
'bf in: ' + (@($c.Models | Where-Object { $_.Name -like '*_BF_*' -or $_.Name -like '*_GU_*' }).Count)
'skipped: ' + (Measure-NonTargetModels @('SITE_ORG_XYZ_K1_R22', 'RG_ORG_BF_K1', 'RG_ORG_GU_K1', 'SITE_ORG_AVT_K1_R22'))
""")
        self.assertEqual(out[0], "folder avt: AVT")
        self.assertEqual(out[1], "folder hc ru: HC")
        self.assertEqual(out[2], "name beats folder: HC")
        self.assertEqual(out[3], "no code: []")
        self.assertEqual(out[4], "task not folder: []")                            # TASK/AR/KR… в запасной вариант по папке не идут
        self.assertEqual(out[5], "sections: AVT")                                  # чужой код в таблицу не входит
        self.assertEqual(out[6], "xyz in: 0")
        self.assertEqual(out[7], "bf in: 0")
        self.assertEqual(out[8], "skipped: 3")

    def test_non_target_model_is_not_in_the_table_or_distribution(self):
        out = self.run_ps(". '%s'\n" % q(CFG) + r"""
$c = New-ProjectConfig 'П'
Update-ProjectSpecs $c @('П        7.2_Диаметр_Трубы_Все')
$c.Categories[(Get-ProjectSpecKey 'П        7.2_Диаметр_Трубы_Все')] = 'Трубы'
$c.InProject = $null
Update-ProjectModels $c @('SITE_ORG_HC_K1_R22', 'SITE_ORG_XYZ_K1_R22')
$f = Join-Path ([IO.Path]::GetTempPath()) ('bvs_u_' + [guid]::NewGuid().ToString('N') + '.txt')
[IO.File]::WriteAllText($f, (Get-DefaultSectionCatalogText), (New-Object System.Text.UTF8Encoding $true))
$cat = Read-SectionCatalog $f; Remove-Item $f
$r = Set-ProjectDistribution $c $cat
'hc: ' + (@(Get-ProjectEnabledSpecs $c 'SITE_ORG_HC_K1_R22').Count)
'xyz in: ' + (@($c.Models | Where-Object { $_.Name -eq 'SITE_ORG_XYZ_K1_R22' }).Count)
'sections: ' + (($c.Sections | ForEach-Object { $_.Code }) -join ',')
'eng: ' + (Test-ProjectEngineeringModel 'RG_ORG_BF_K1') + ' ' + (Test-ProjectEngineeringModel 'SITE_ORG_HC_K1_R22')
""")
        self.assertEqual(out[0], "hc: 1")
        self.assertEqual(out[1], "xyz in: 0")
        self.assertEqual(out[2], "sections: HC")
        self.assertEqual(out[3], "eng: False True")
        dist = read(DIST)
        self.assertNotIn("Get-UnknownSectionNote", dist)
        self.assertIn("Format-SkippedModelsStatus", dist)
        self.assertNotIn("Неизв.", dist + read(CFG))

    def test_proj_names_unchanged(self):
        out = self.run_ps(". '%s'\n" % q(CFG) + r"""
foreach ($n in 'PROJ_IOS_OV_T_1','PROJ_IOS_VK_T_1','PROJ_IOS_PD_T_1','PROJ_IOS_EOM_T_1','PROJ_IOS_SS_T_1','PROJ_IOS_ITP_T_1') { $n + '=' + (Get-ModelSectionFromName $n) }
""")
        self.assertEqual(out, ["PROJ_IOS_OV_T_1=OV", "PROJ_IOS_VK_T_1=VK", "PROJ_IOS_PD_T_1=PD", "PROJ_IOS_EOM_T_1=EOM", "PROJ_IOS_SS_T_1=SS", "PROJ_IOS_ITP_T_1=ITP"])
        s = read(STREAM)
        self.assertIn("_IOS_([A-Z]+)_", s)                                         # прежний regex остался первым

    def test_bvs_link_section_for_new_and_old(self):
        out = self.run_ps(". '%s'\n" % q(STREAM) + r"""
foreach ($n in 'PROJ_IOS_OV_T_1','SITE_ORG_AVT_PRK_R22','SITE_ORG_HC_K2_R22') { $n + '=' + (Get-BvsLinkSection $n) }
""")
        self.assertEqual(out, ["PROJ_IOS_OV_T_1=OV", "SITE_ORG_AVT_PRK_R22=AVT", "SITE_ORG_HC_K2_R22=HC"])

    def test_chain_tables_know_new_sections(self):
        for f, needles in (("build_assignment.ps1", ('"AVT"', '"HC"', "АВТ", "ХЦ")), ("assignment_partial_ui.ps1", ("AVT", "HC")),
                           ("google_payload_lib.ps1", ('"АВТ"', '"ХЦ"')), ("project_config.ps1", ("AVT", "HC", "АВТ", "ХЦ"))):
            t = read(os.path.join(TOOL, f))
            for n in needles:
                self.assertIn(n, t, f + ": " + n)


@unittest.skipUnless(PWSH, "нет pwsh")
class CatalogTests(PsBase):
    def test_default_text_has_avt_and_hc(self):
        out = self.run_ps(". '%s'\n" % q(CFG) + r"""
$f = Join-Path ([IO.Path]::GetTempPath()) ('bvs_c_' + [guid]::NewGuid().ToString('N') + '.txt')
[IO.File]::WriteAllText($f, (Get-DefaultSectionCatalogText), (New-Object System.Text.UTF8Encoding $true))
$r = Read-SectionCatalog $f; Remove-Item $f
'avt cats: ' + ($r.Sections['AVT'].Categories -join ';')
'hc cats: ' + ($r.Sections['HC'].Categories -join ';')
'avt nums: ' + (($r.Sections['AVT'].Numbers | ForEach-Object { "$($_.From)-$($_.To)" }) -join ',')
""")
        self.assertEqual(out[0], "avt cats: Кабельные лотки;Соединительные детали кабельных лотков")
        self.assertEqual(out[1], "hc cats: Трубы;Соединительные детали трубопроводов;Арматура трубопроводов;Оборудование")

    def test_old_file_gets_blocks_once_and_user_text_is_kept(self):
        f = os.path.join(self.tmp, "справочник_разделов.txt")
        old = "[ОВ]\nномера = 0, 1\nкатегории = Воздуховоды\n# мой комментарий\n"
        with open(f, "w", encoding="utf-8-sig") as fh:
            fh.write(old)
        out = self.run_ps(". '%s'\n" % q(CFG) + "$f = '%s'\n" % q(f) + r"""
Initialize-SectionCatalogFile $f
$t1 = [IO.File]::ReadAllText($f, [Text.Encoding]::UTF8)
Initialize-SectionCatalogFile $f
$t2 = [IO.File]::ReadAllText($f, [Text.Encoding]::UTF8)
'same after 2nd: ' + ($t1 -eq $t2)
'kept: ' + $t1.Contains('мой комментарий')
'avt: ' + $t1.Contains('[АВТ]') + ' hc: ' + $t1.Contains('[ХЦ]')
$r = Read-SectionCatalog $f
'read ok: ' + $r.Ok + ' ' + $r.Sections.ContainsKey('AVT') + ' ' + $r.Sections.ContainsKey('HC')
""")
        self.assertEqual(out[0], "same after 2nd: True")
        self.assertEqual(out[1], "kept: True")
        self.assertEqual(out[2], "avt: True hc: True")
        self.assertEqual(out[3], "read ok: True True True")
        text = read(f)
        self.assertTrue(text.startswith(old))                                      # исходное начало файла не тронуто
        self.assertEqual(text.count("[АВТ]"), 1)

    def test_user_deleted_block_is_not_returned(self):
        f = os.path.join(self.tmp, "справочник_разделов.txt")
        with open(f, "w", encoding="utf-8-sig") as fh:
            fh.write("[ОВ]\nномера = 0\nкатегории = Воздуховоды\n")
        self.run_ps(". '%s'\n" % q(CFG) + "Initialize-SectionCatalogFile '%s'\n" % q(f))
        text = read(f).replace("[ХЦ]", "[хц_убран]")                                # пользователь переименовал/убрал блок
        text = re.sub(r"\[хц_убран\].*?(?=\n\[|\Z)", "", text, flags=re.S)
        with open(f, "w", encoding="utf-8-sig") as fh:
            fh.write(text)
        self.run_ps(". '%s'\n" % q(CFG) + "Initialize-SectionCatalogFile '%s'\n" % q(f))
        self.assertNotIn("[ХЦ]", read(f))


@unittest.skipUnless(PWSH, "нет pwsh")
class GoogleUiTests(PsBase):
    def test_no_fields_and_status_from_cfg_without_touching_it(self):
        ui = read(UI)
        for gone in ("$ui.GoogleUrl", "$ui.GoogleSecret", "Save-ProjectGoogle"):
            self.assertNotIn(gone, ui)
        self.assertNotIn("Save-ProjectGoogle", read(DIST))
        rep = os.path.join(self.tmp, "отчеты")
        os.makedirs(os.path.join(rep, "проект"))
        cfg = os.path.join(rep, "проект", "google_SITE.cfg")
        body = "webapp_url=https://script.google.com/macros/s/AKfycbQQQQQQQQQQQQQQQQQQQQ/exec\nsecret=TOPSECRETVALUE\n# моя строка\n"
        with open(cfg, "w", encoding="utf-8-sig") as fh:
            fh.write(body)
        before = open(cfg, "rb").read()
        out = self.run_ps(". '%s'\n. '%s'\n. '%s'\n" % (q(CFG), q(GCFG), q(LAYOUT)) + func_text(ui, "Get-GoogleStatusText") + "\n"
                          + "$script:ProjectPath = '%s'\n" % q(os.path.join(rep, "проект", "распределение_SITE.txt"))
                          + "Get-GoogleStatusText\n$script:ProjectPath = ''\nGet-GoogleStatusText\n")
        self.assertEqual(open(cfg, "rb").read(), before)                           # cfg не переписан
        self.assertTrue(out[0].startswith("Google: проект SITE, адрес …") and out[0].endswith("/exec"), out[0])
        self.assertNotIn("TOPSECRETVALUE", "\n".join(out))
        self.assertEqual(out[1], "Google: нет проекта")

    def test_save_connection_validates_and_keeps_secret(self):
        ui = read(UI)
        rep = os.path.join(self.tmp, "отчеты", "проект")
        os.makedirs(rep)
        pc = os.path.join(rep, "распределение_SITE.txt")
        url = "https://script.google.com/macros/s/AKfycb" + "A" * 40 + "/exec"
        url2 = "https://script.google.com/macros/s/AKfycb" + "B" * 40 + "/exec"
        out = self.run_ps(". '%s'\n. '%s'\n" % (q(CFG), q(GCFG)) + func_text(ui, "Save-GoogleConnection") + "\n$pc = '%s'\n" % q(pc) + r"""
'bad: ' + (@(Save-GoogleConnection $pc 'http://x' 's').Count -gt 0) + ' exists: ' + (Test-Path -LiteralPath (Get-ProjectGoogleConfigPath $pc))
'ok: ' + (@(Save-GoogleConnection $pc '%s' 'sec-1234567890').Count)
'url changed, empty secret: ' + (@(Save-GoogleConnection $pc '%s' '').Count)
$c = Read-ProjectGoogleConfig (Get-ProjectGoogleConfigPath $pc)
'now: ' + ($c.WebAppUrl -eq '%s') + ' ' + $c.Secret
""" % (url, url2, url2))
        self.assertEqual(out[0], "bad: True exists: False")
        self.assertEqual(out[1], "ok: 0")
        self.assertEqual(out[2], "url changed, empty secret: 0")
        self.assertEqual(out[3], "now: True sec-1234567890")


class TreeSelectionTests(unittest.TestCase):
    def setUp(self):
        self.ui = read(UI)

    def test_click_selects_node_and_refresh_restores_it(self):
        ui = self.ui
        self.assertIn("$ui.AsgTree.Add_NodeMouseClick", ui)
        click = ui[ui.index("$ui.AsgTree.Add_NodeMouseClick"):]
        self.assertIn("$ui.AsgTree.SelectedNode = $e.Node", click[:400])
        refresh = func_text(ui, "Refresh-PartialLists")
        self.assertIn("Get-AsgNodeKey", refresh)
        self.assertIn("$ui.AsgTree.SelectedNode = $pick", refresh)
        self.assertLess(refresh.index("$ui.AsgTree.SelectedNode = $pick"), refresh.rindex("Show-AsgSpecList"))
        spec = func_text(ui, "Show-AsgSpecList")
        self.assertIn("finally", spec)
        self.assertIn("EndUpdate()", spec[spec.index("finally"):])
        after = ui[ui.index("$ui.AsgTree.Add_AfterSelect"):]
        self.assertIn("catch", after[:1200])
        self.assertIn("Список спецификаций не показан", after[:1500])

    @unittest.skipUnless(PWSH, "нет pwsh")
    def test_node_key(self):
        r = pwsh("-Command", func_text(self.ui, "Get-AsgNodeKey") + "\n"
                 "$s = [pscustomobject]@{ Text = 'Вентиляция (12)'; Tag = [pscustomobject]@{ Kind = 'section' } }\n"
                 "Get-AsgNodeKey $s")
        self.assertEqual(r.returncode, 0, r.stdout + r.stderr)


@unittest.skipUnless(PWSH, "нет pwsh")
class ProjectReportDirTests(PsBase):
    def fns(self):
        ui = read(UI)
        return "\n".join(func_text(ui, n) for n in ("Get-ProjectReportDir", "Get-AssignmentPath", "Get-DynamicsOnlyPath", "Get-NextReportPath", "Copy-RunDiagnosticsToProject", "Get-ProjectCode"))

    def prelude(self, project=True, name="SITE"):
        rep = os.path.join(self.tmp, "отчеты")
        os.makedirs(os.path.join(rep, "проект"), exist_ok=True)
        pp = os.path.join(rep, "проект", "распределение_%s.txt" % name) if project else ""
        return rep, ("$ReportDir = '%s'\n$script:ProjectPath = '%s'\nfunction Get-CurrentProjectName { '%s' }\nfunction Get-ModelLeaf($p) { Split-Path -Leaf $p }\n" % (q(rep), q(pp), name)) + self.fns() + "\n"

    def test_reports_go_to_project_folder(self):
        rep, head = self.prelude()
        out = self.run_ps(head + r"""
$a = Get-AssignmentPath 'C:\x\SITE_SVOD.rvt'
$d = Get-DynamicsOnlyPath 'C:\x\SITE_SVOD.rvt'
$n = Get-NextReportPath (Get-ProjectReportDir 'SITE')
(Split-Path -Parent $a) + '|' + (Split-Path -Leaf $a)
(Split-Path -Parent $d) + '|' + (Split-Path -Leaf $d)
(Split-Path -Parent $n)
""")
        d = os.path.join(rep, "SITE")
        self.assertTrue(os.path.isdir(d))                                           # создана при первой записи
        a_dir, a_name = out[0].split("|")
        self.assertEqual(a_dir, d)
        self.assertTrue(a_name.endswith("_SITE.xlsx"), a_name)
        d_dir, d_name = out[1].split("|")
        self.assertEqual(d_dir, d)
        self.assertTrue(d_name.endswith("_SITE_динамика.xlsx"), d_name)
        self.assertEqual(out[2], d)

    def test_two_projects_do_not_share_folder_and_proj_keeps_its_name(self):
        rep, head = self.prelude(name="PROJ")
        out = self.run_ps(head + "(Get-ProjectReportDir 'PROJ')\n(Get-ProjectReportDir 'SITE')\n(Get-ProjectReportDir 'A/B')\n")
        self.assertEqual(out[0], os.path.join(rep, "PROJ"))
        self.assertEqual(out[1], os.path.join(rep, "SITE"))
        self.assertEqual(out[2], os.path.join(rep, "A_B"))                          # имя очищается как в имени файла настройки

    def test_no_project_uses_report_root(self):
        rep, head = self.prelude(project=False)
        out = self.run_ps(head + "(Get-ProjectReportDir 'SITE')\n(Split-Path -Parent (Get-AssignmentPath 'C:\\x\\SITE_IOS_SVOD.rvt'))\n")
        self.assertEqual(out[0], rep)
        self.assertEqual(out[1], rep)
        self.assertFalse(os.path.exists(os.path.join(rep, "SITE")))

    def test_filters_and_payload_stay_in_aux_and_are_not_copied_to_project(self):
        rep, head = self.prelude()
        work = os.path.join(rep, "служебные", "20260101_000000")
        os.makedirs(work)
        for n, t in (("фильтры_спецификаций.txt", "f"), ("google_push_payload.json", "{}")):
            with open(os.path.join(work, n), "w", encoding="utf-8") as fh:
                fh.write(t)
        proj = os.path.join(rep, "SITE")
        os.makedirs(proj)
        with open(os.path.join(proj, "2026-01-01_SITE.xlsx"), "w", encoding="utf-8") as fh:
            fh.write("x")
        self.run_ps(head + "Copy-RunDiagnosticsToProject -WorkDir '%s' -XlsxPath '%s'\n" % (q(work), q(os.path.join(proj, "2026-01-01_SITE.xlsx"))))
        self.assertEqual(os.listdir(proj), ["2026-01-01_SITE.xlsx"])                         # в папке проекта только xlsx
        self.assertEqual(open(os.path.join(work, "google_push_payload.json"), encoding="utf-8").read(), "{}")
        self.assertEqual(open(os.path.join(work, "фильтры_спецификаций.txt"), encoding="utf-8").read(), "f")
        self.assertEqual(os.listdir(rep).count("фильтры_спецификаций.txt"), 0)

    def test_wiring_auxdir_and_open_folder(self):
        ui = read(UI)
        self.assertIn("-Partial:$script:Run.Partial -AuxDir $script:Run.WorkDir", ui)       # фильтры и payload — в папку прогона, а не в корень отчетов
        self.assertIn('Join-Path $WorkDir "google_push_payload.json"', ui)
        self.assertIn("Get-NextReportPath (Get-ProjectReportDir", ui)
        opn = func_text(ui, "Open-ReportFolder")
        self.assertIn("$script:ProjectPath", opn)
        self.assertIn("Test-Path -LiteralPath $pd -PathType Container", opn)

    def test_migration_leaves_project_folders_alone(self):
        rep = os.path.join(self.tmp, "отчеты")
        for sub in ("PROJ", "SITE"):
            os.makedirs(os.path.join(rep, sub))
            for n in ("история_%s.txt" % sub, "google_push_payload.json", "фильтры_спецификаций.txt", "2026-01-01_x.xlsx"):
                with open(os.path.join(rep, sub, n), "w", encoding="utf-8") as fh:
                    fh.write("x")
        with open(os.path.join(rep, "2026-01-01_old.xlsx"), "w") as fh:
            fh.write("x")
        before = {s: sorted(os.listdir(os.path.join(rep, s))) for s in ("PROJ", "SITE")}
        self.run_ps(". '%s'\n[void](Invoke-ReportMigration '%s')\n" % (q(LAYOUT), q(rep)))
        for s in ("PROJ", "SITE"):
            self.assertEqual(sorted(os.listdir(os.path.join(rep, s))), before[s])
        self.assertTrue(os.path.exists(os.path.join(rep, "2026-01-01_old.xlsx")))            # файлы корня не переносятся
        self.assertFalse(os.path.exists(os.path.join(rep, "история")))                       # история\ не создана переносом из папок проектов


if __name__ == "__main__":
    unittest.main()


class PythonSectionTests(unittest.TestCase):
    """section_code_from_name из insert_schedules.py (без Revit: берутся только нужные определения)."""

    def load(self, path):
        import ast
        src = read(path)
        tree = ast.parse(src)
        want = {"SECTION_NAMES", "NON_SECTION_MARKERS", "lower_text", "upper_text", "project_model_name", "as_text", "section_code_from_folder", "section_code_from_name"}
        parts = []
        for node in tree.body:
            names = [t.id for t in getattr(node, "targets", []) if isinstance(t, ast.Name)] + ([node.name] if isinstance(node, (ast.FunctionDef, ast.ClassDef)) else [])
            if want & set(names):
                parts.append(ast.get_source_segment(src, node))
        ns = {"os": os, "re": re}
        exec("\n".join(parts), ns)
        return ns["section_code_from_name"]

    def test_both_scripts(self):
        for path in (os.path.join(TOOL, "insert_schedules.py"),):
            f = self.load(path)
            for name, code in (("SITE_ORG_AVT_K1_R22", "AVT"), ("SITE_ORG_AVT_PRK_R22", "AVT"), ("SITE_ORG_HC_K4_R22.rvt", "HC"),
                               ("PROJ_IOS_OV_T_1", "OV"), ("PROJ_IOS_VK_T_1.rvt", "VK"), ("C:\\m\\AVT\\SITE_X_K1.rvt", "AVT"), ("C:\\m\\ХЦ\\SITE_X_K1.rvt", "HC"),
                               ("C:\\m\\SITE_X_K1.rvt", ""), ("C:\\m\\AVT\\SITE_TASK_K1.rvt", "")):
                self.assertEqual(f(name), code, "%s: %s" % (os.path.basename(path), name))
