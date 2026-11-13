"""Исправления по АУДИТ.md: итог прогона «Сводная/Разделы», префикс проекта в «Задании», версии Revit в инструментах, замена без сохранения при неполной вставке,
RSN:// у общего файла, подсчёт при open_all."""
import json
import os
import re
import shutil
import sys
import tempfile
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from test_chain import pwsh, PWSH, make_export  # noqa: E402

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(HERE)))
SERVICE = os.path.join(ROOT, "служебное")
TOOL = os.path.join(SERVICE, "скрипты")
TOOLS = os.path.join(SERVICE, "разработка", "инструменты")


def read(p):
    with open(p, encoding="utf-8-sig") as f:
        return f.read()


def run_ps(tmp, libs, body):
    script = os.path.join(tmp, "t.ps1")
    with open(script, "w", encoding="utf-8-sig") as f:
        f.write("".join(". '%s'\n" % os.path.join(TOOL, l) for l in libs) + body + "\n")
    r = pwsh("-File", script)
    assert r.returncode == 0, r.stdout + r.stderr
    return [l for l in r.stdout.splitlines() if l.strip()]


@unittest.skipUnless(PWSH, "нет pwsh")
class RunSummaryTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="bvs_rs_")
        self.addCleanup(shutil.rmtree, self.tmp, True)

    n = 0

    def csv(self, rows):
        self.n += 1
        p = os.path.join(self.tmp, "сводка%d.csv" % self.n)
        with open(p, "w", encoding="utf-8-sig", newline="") as f:
            f.write("Модель;Раздел;Удалено;Вставлено;Пропущено;Ошибка\r\n" + "".join(r + "\r\n" for r in rows))
        return p

    def test_ready_only_when_exit_code_zero_and_no_errors(self):
        ok = self.csv(["A_OV_X;ОВ;0;5;0;", "B_VK_X;ВК;0;7;1;"])
        bad = self.csv(["A_OV_X;ОВ;0;5;0;", "B_VK_X;ВК;0;0;0;Набор спецификаций модели не определён", "C_PD_X;ПД;0;0;0;файл занят"])
        empty = self.csv([])
        out = run_ps(self.tmp, ["report_layout.ps1"], r"""
foreach ($case in @(@('%s', 0), @('%s', 0), @('%s', 0), @('%s', 5), @('%s', 5), @('нет.csv', 0))) {
    $s = Get-RunSummaryStats $case[0]
    $s.Total.ToString() + '/' + $s.Errors + ' | ' + (Get-RunFailureText ([int]$case[1]) $s)
}
""" % (ok, bad, empty, ok, bad))
        self.assertEqual(out[0], "2/0 | ")                                                    # код 0, ошибок нет — «готова»
        self.assertEqual(out[1], "3/2 | Не готово: с ошибкой 2 из 3 моделей (код RBP 0). Подробности — в отчёте, лист «Сводка».")
        self.assertEqual(out[2], "0/0 | Не готово: ни одна модель не обработана (код RBP 0) — Revit не запустился или не открыл модели.")
        self.assertEqual(out[3], "2/0 | Не готово: Revit Batch Processor завершился с кодом 5; обработано моделей: 2.")
        self.assertIn("с ошибкой 2 из 3", out[4])
        self.assertTrue(out[5].startswith("0/0 | Не готово: ни одна модель"))

    def test_ui_uses_stats_before_saying_ready(self):
        ui = read(os.path.join(TOOL, "ui_schedules.ps1"))
        tick = ui[ui.index("$runStats = Get-RunSummaryStats"):ui.index("$form.Controls.Add($ui.Timer)")]
        self.assertLess(tick.index("if ($failText)"), tick.index("Сводная готова"))
        self.assertLess(tick.index("if ($failText)"), tick.index("Модели разделов готовы"))


@unittest.skipUnless(PWSH, "нет pwsh")
class ProjectPrefixTests(unittest.TestCase):
    def test_prefix_from_config_file_then_summary(self):
        tmp = tempfile.mkdtemp(prefix="bvs_px_")
        self.addCleanup(shutil.rmtree, tmp, True)
        out = run_ps(tmp, ["project_config.ps1"], r"""
'cfg: ' + (Get-AssignProjectPrefix 'отчеты\проект\распределение_PROJ.txt' '')
'cfg unix: ' + (Get-AssignProjectPrefix 'x/проект/распределение_XYZ1.txt' 'Q_IOS_SVOD_R24')
'summary: ' + (Get-AssignProjectPrefix '' '\\srv\a\ZZ9_IOS_SVOD_ORG_R24.rvt')
'PROJ ignored: ' + (Get-AssignProjectPrefix 'распределение_PROJ.txt' 'AB2_IOS_SVOD_R24')
'none: [' + (Get-AssignProjectPrefix '' '') + ']'
""")
        self.assertEqual(out, ["cfg: PROJ", "cfg unix: XYZ1", "summary: ZZ9", "PROJ ignored: AB2", "none: []"])

    def test_link_section_and_prefixed_names(self):
        tmp = tempfile.mkdtemp(prefix="bvs_px_")
        self.addCleanup(shutil.rmtree, tmp, True)
        out = run_ps(tmp, ["assignment_export_stream.ps1"], r"""
function S($t) { $r = Get-BvsLinkSection $t; if ($null -eq $r) { '-' } else { $r } }
'generic: ' + (S 'ORG_K1_PROJ_IOS_EOM_R23.rvt') + ' ' + (S 'XYZ1_IOS_OV_ORG') + ' ' + (S 'Чужое_имя')
$script:BvsProjectPrefix = 'PROJ'
'proj: ' + (S 'ORG_K1_PROJ_IOS_EOM_R23.rvt') + ' ' + (S 'XYZ1_IOS_OV_ORG')
'prefixed: ' + (Test-BvsIosPrefixedName 'PROJ_IOS_EOM_x') + ' ' + (Test-BvsIosPrefixedName 'XYZ1_IOS_EOM_x') + ' ' + (Test-BvsIosPrefixedName 'Обычное имя')
$script:BvsProjectPrefix = 'XY.1'
'escaped: ' + (S 'XY.1_IOS_OV_T') + ' ' + (S 'XYZ1_IOS_OV_T')
""")
        self.assertEqual(out[0], "generic: EOM OV -")
        self.assertEqual(out[1], "proj: EOM -")
        self.assertEqual(out[2], "prefixed: True False False")
        self.assertEqual(out[3], "escaped: OV -")                                           # префикс — не regex

    def build(self, tmp, name, replace=None, prefix=None):
        folder = os.path.join(tmp, name)
        os.makedirs(folder)
        exp, out = make_export(folder)
        if replace:
            for d in (exp, out):                                                             # и выгрузка, и посеянная история прошлых сборок
                for f in os.listdir(d):
                    p = os.path.join(d, f)
                    if os.path.isfile(p):
                        text = read(p)
                        with open(p, "w", encoding="utf-8-sig", newline="") as fh:
                            fh.write(text.replace(*replace))
        args = ["-File", os.path.join(TOOL, "build_assignment.ps1"), "-ExportDir", exp, "-OutXlsx", os.path.join(out, "task.xlsx")]
        if prefix is not None:
            args += ["-ProjectPrefix", prefix]
        r = pwsh(*args)
        self.assertEqual(r.returncode, 0, r.stdout + r.stderr)
        with open(os.path.join(out, "google_push_payload.json"), encoding="utf-8-sig") as f:
            sheets = json.load(f)["sheets"]
        return [(s["name"], len(s["rows"])) for s in sheets]

    def test_assignment_is_the_same_for_proj_and_works_for_other_prefixes(self):
        tmp = tempfile.mkdtemp(prefix="bvs_px_")
        self.addCleanup(shutil.rmtree, tmp, True)
        base = self.build(tmp, "a")                                                          # без префикса — любой (как раньше для PROJ)
        self.assertEqual([n for n, _ in base], ["Сводка", "ЭОМ", "СС", "Динамика"])
        self.assertEqual(self.build(tmp, "b", prefix="PROJ"), base)                          # префикс из настройки: результат для PROJ не меняется
        other = ("PROJ_IOS", "XYZ1_IOS")
        self.assertEqual(self.build(tmp, "c", replace=other, prefix="XYZ1"), base)           # другой проект
        self.assertEqual(self.build(tmp, "d", replace=other), base)                          # префикс не задан
        wrong = self.build(tmp, "e", replace=other, prefix="PROJ")                           # чужой префикс — разделов нет
        self.assertNotEqual(wrong, base)


class StaticWiringTests(unittest.TestCase):
    def test_no_hardcoded_project_prefix_in_assignment(self):
        for name in ("build_assignment.ps1", "assignment_export_stream.ps1"):
            t = read(os.path.join(TOOL, name))
            code = "\n".join(l for l in t.splitlines() if not l.lstrip().startswith("#"))
            self.assertNotIn("PROJ_IOS", code, name)
        ui = read(os.path.join(TOOL, "ui_schedules.ps1"))
        self.assertIn("$params.ProjectPrefix = $prefix", ui)

    def test_prefill_counts_with_open_all_and_template_read_unchanged(self):
        d = read(os.path.join(TOOL, "project_dist_ui.ps1"))
        self.assertIn('[string]$Worksets = "close_all"', d)                                  # чтение общего файла — как раньше
        self.assertIn('"--worksets", $Worksets', d)
        assign = d[d.index("function Invoke-DistAssign"):d.index("function Invoke-DistCellClick")]
        self.assertIn('"open_all" @($template)', assign)                                      # подсчёт — по всем рабочим наборам, RSN у общего файла учитывается
        load = d[d.index("function Invoke-DistLoad"):d.index("function Invoke-DistAssign")]
        self.assertNotIn("open_all", load)

    def test_rsn_template_accepted_and_patch_enabled(self):
        ui = read(os.path.join(TOOL, "ui_schedules.ps1"))
        start = ui[ui.index("function Start-Batch"):]
        self.assertIn("-not (Test-ModelReady $source)", start)
        self.assertNotIn("Test-Path -LiteralPath $source", start)
        self.assertIn('$hasRsn = @(@($chosen) + @($(if ($Mode -eq "summary" -or $Mode -eq "sections") { $source }))', start)
        d = read(os.path.join(TOOL, "project_dist_ui.ps1"))
        self.assertIn("@(@($Files) + @($ExtraPaths)", d)

    def test_dump_filters_resolves_revit_year(self):
        """Дамп фильтров открывает модели и передаёт год Revit всегда."""
        p = os.path.join(TOOLS, "dump_schedule_filters.ps1")
        t = read(p)
        self.assertNotIn("Guess-Year", t)
        self.assertIn("revit_versions.ps1", t)
        self.assertNotIn("$years.Count -eq 1", t)
        self.assertIsNone(re.search(r"192\\?\.168", t))
        self.assertRegex(t, r"--revit_version \{0\}'")
        self.assertIn("Get-ModelRevitYear", t)
        self.assertIn("Не удалось определить версию Revit", t)


@unittest.skipUnless(PWSH, "нет pwsh")
class ToolsVersionStopTests(unittest.TestCase):
    """Остановка до запуска RBP: инструмент проверки версий вызывается как в окне (Resolve-RevitRun) — неизвестная и смешанная версии."""

    def test_unknown_and_mixed_stop(self):
        sys.path.insert(0, HERE)
        from test_revit_versions import make_rvt, info_text
        tmp = tempfile.mkdtemp(prefix="bvs_tv_")
        self.addCleanup(shutil.rmtree, tmp, True)
        a, b, junk = (os.path.join(tmp, n) for n in ("a.rvt", "b.rvt", "junk.rvt"))
        make_rvt(a, info_text(2022))
        make_rvt(b, info_text(2024))
        with open(junk, "wb") as f:
            f.write(b"x" * 300)
        out = run_ps(tmp, ["revit_versions.ps1"], r"""
$r = Resolve-RevitRun @('%s','%s') '' @{}; 'mixed: ' + [string]$r.Year + '|' + $r.Problem
$r = Resolve-RevitRun @('%s','%s') '' @{}; 'unknown: ' + [string]$r.Year + '|' + $r.Problem.Substring(0, 40)
$r = Resolve-RevitRun @('%s') '%s' @{}; 'ref newer: ' + [string]$r.Year + '|' + $r.Problem.Substring(0, 23)
$r = Resolve-RevitRun @('%s') '%s' @{}; 'ok: ' + [string]$r.Year + '|' + $r.Problem
""" % (a, b, a, junk, a, b, b, a))
        self.assertTrue(out[0].startswith("mixed: |Модели разных версий"))
        self.assertTrue(out[1].startswith("unknown: |Не удалось определить версию Revit"))
        self.assertEqual(out[2], "ref newer: |Общий файл (2024) новее")
        self.assertEqual(out[3], "ok: 2024|")


if __name__ == "__main__":
    unittest.main()
