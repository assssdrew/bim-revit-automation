"""План вставки (plan.txt от окна) против прежнего разбора файла настройки в insert_schedules.py.

Для каждого файла настройки (PROJ и SampleSite) окно (Get-InsertPlanLines в project_config.ps1) пишет plan.txt, затем insert_schedules.py
на подставном Revit API запускается дважды: с plan_path и без него (LEGACY: разбор файла). Итог должен совпасть полностью:
вставленные виды, сохранение, CSV отчёта, текст ошибки и журнал (кроме строки plan=). Проверяются «В сводную», «В разделы»,
режимы initial / add_diff / replace, автопереход на первичную загрузку, модели без кода, по папке, не размеченные, ИТП,
исключённые модели SampleSite (RG_ORG_BF_, GU_), пустые [inproject]/[summary], файл без stage=."""
import os
import shutil
import sys
import tempfile
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from test_chain import pwsh, PWSH  # noqa: E402
import test_insert_project as tip  # noqa: E402  Run — insert_schedules.py на подставном Revit API
import project_fixture  # noqa: E402

TOOL = tip.TOOL
CFG = os.path.join(TOOL, "project_config.ps1")
SITE_TSV = os.path.join(os.path.dirname(HERE), "данные", "samplesite_template_specs.tsv")
P8 = tip.P8


def q(s):
    return str(s).replace("'", "''")


def specs_of(text):
    out, sec = [], ""
    for line in text.splitlines():
        t = line.strip()
        if t.startswith("[") and t.endswith("]"):
            sec = t[1:-1]
        elif t and sec == "specs" and t not in out:
            out.append(t)
    return out


def proj_projects():
    base = tip.PROJECT
    itp = base.replace("[specs]", "PROJ_IOS_ITP_ORG_K01_R24\tITP\nPROJ_IOS_EOM_ORG_K01_R24\tEOM\n[specs]").replace(
        "[cells]\n", "[cells]\nPROJ_IOS_ITP_ORG_K01_R24\t7.2_Диаметр_Трубы_Все\t1\tm\t-1\nPROJ_IOS_ITP_ORG_K01_R24\t1_Классификатор_Все\t0\ta\t0\n")
    marked = itp.replace("[cells]", "[inproject]\n1_Классификатор_Все\n7.2_Диаметр_Трубы_Все\n5.1_Ширина_Воздуховоды_Все\n[summary]\n7.2_Диаметр_Трубы_Все\n[cells]")
    in_only = itp.replace("[cells]", "[inproject]\n1_Классификатор_Все\n7.2_Диаметр_Трубы_Без значения\n[cells]")
    empty_in = itp.replace("[cells]", "[inproject]\n[summary]\n[cells]")
    empty_sum = itp.replace("[cells]", "[inproject]\n1_Классификатор_Все\n[summary]\n[cells]")
    no_stage = itp.replace("stage=РД\n", "")
    folder = itp.replace("[specs]", "PROJ_X_K05_R24\tOV\n[specs]").replace("[cells]\n", "[cells]\nPROJ_X_K05_R24\t1_Классификатор_Все\t1\tm\t-1\n")
    fixture = project_fixture.project_text().replace("stage=П", "stage=П\ntemplate=C:\\шаблон.rvt")
    fixture_in = project_fixture.project_text(in_project=lambda s: not s.startswith("1_")).replace("stage=П", "stage=П\ntemplate=C:\\шаблон.rvt")
    return {"proj": base, "proj_itp": itp, "proj_marked": marked, "proj_in_only": in_only, "proj_empty_in": empty_in,
            "proj_empty_sum": empty_sum, "proj_no_stage": no_stage, "proj_folder": folder, "proj_fixture": fixture, "proj_fixture_in": fixture_in}


PROJ_MODELS = ["C:\\m\\PROJ_IOS_OV_ORG_K01_R24.rvt", "\\\\srv\\проект\\PROJ_IOS_VK_ORG_K01_R24.rvt", "C:\\m\\PROJ_IOS_VK_ORG_K02_R24.rvt",
               "C:\\m\\PROJ_IOS_VK_ORG_K09_R24.rvt",          # нет в распределении: объединение раздела ВК
               "C:\\m\\PROJ_IOS_ITP_ORG_K01_R24.rvt", "C:\\m\\PROJ_IOS_ITP_ORG_K07_R24.rvt",
               "C:\\m\\PROJ_IOS_EOM_ORG_K01_R24.rvt",          # в [models], ячеек нет: «не размечены»
               "C:\\m\\PROJ_IOS_SS_ORG_K01_R24.rvt",           # раздел без моделей
               "C:\\m\\PROJ_X_K01_R24.rvt",                      # кода нет
               "C:\\m\\ОВ\\PROJ_X_K05_R24.rvt",                  # раздел только по папке
               "C:\\m\\ОВ\\PROJ_TASK_K01_R24.rvt",
               "C:\\m\\PROJ_IOS_OT_ORG_K01_R24.rvt", "C:\\m\\PROJ_IOS_EOM_ORG_K02_R24.rvt"]


def site_project():
    names = []
    with open(SITE_TSV, encoding="utf-8-sig") as f:
        for line in f.read().splitlines():
            n = line.split("\t")[0].strip()
            if n and n not in names:
                names.append(n)
    specs = names[:24]
    models = [("SITE_ORG_AVT_K1_R22", "AVT"), ("SITE_ORG_AVT_PRK_R22", "AVT"), ("SITE_ORG_HC_K1_R22", "HC"), ("SITE_ORG_HC_K2_R22", "HC"),
              ("SITE_X_K7_R22", "AVT"),                     # раздел по папке «АВТ»
              ("RG_ORG_BF_K1_R22", "UNKNOWN"), ("RG_ORG_GU_K1_R22", "BF")]          # исключённые: не в таблице, отметки остаются в файле
    out = ["# SampleSite", "[project]", "stage=РД", "template=C:\\шаблон.rvt", "[models]"]
    out += ["%s\t%s" % m for m in models]
    out.append("[specs]")
    out += specs
    out += ["[inproject]"] + [s for i, s in enumerate(specs) if i % 5 != 4]
    out += ["[summary]"] + [s for i, s in enumerate(specs) if i % 3 != 2]
    out.append("[cells]")
    for mi, (m, _) in enumerate(models):
        if m == "SITE_ORG_HC_K2_R22":
            continue                                    # модель без отметок
        for si, s in enumerate(specs):
            out.append("%s\t%s\t%d\t%s\t" % (m, s, 1 if (si + mi) % 3 != 0 else 0, "a" if si % 2 else "m"))
    return "\n".join(out) + "\n"


SITE_MODELS = ["C:\\m\\АВТ\\SITE_ORG_AVT_K1_R22.rvt", "C:\\m\\SITE_ORG_AVT_PRK_R22.rvt", "C:\\m\\SITE_ORG_HC_K1_R22.rvt",
               "C:\\m\\SITE_ORG_HC_K2_R22.rvt", "C:\\m\\SITE_ORG_AVT_K9_R22.rvt", "C:\\m\\SITE_ORG_HC_K9_R22.rvt",
               "C:\\m\\АВТ\\SITE_X_K7_R22.rvt", "C:\\m\\ХЦ\\SITE_X_K8_R22.rvt",
               "C:\\m\\RG_ORG_BF_K1_R22.rvt", "C:\\m\\RG_ORG_GU_K1_R22.rvt"]


@unittest.skipUnless(PWSH, "нет pwsh")
class PlanVsLegacyTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.mkdtemp(prefix="bvs_plan_")
        cls.projects = dict(proj_projects())
        cls.projects["site"] = site_project()
        cls.plans = {}
        body = [". '%s'" % q(CFG)]
        for key, text in cls.projects.items():
            pp = os.path.join(cls.tmp, "распределение_%s.txt" % key)
            with open(pp, "w", encoding="utf-8-sig", newline="") as f:
                f.write(text)
            plan = os.path.join(cls.tmp, key, "plan.txt")
            cls.plans[key] = plan
            body.append("[void](Write-InsertPlan '%s' '%s')" % (q(pp), q(plan)))
        body.append("'done'")
        script = os.path.join(cls.tmp, "plan.ps1")
        with open(script, "w", encoding="utf-8-sig") as f:
            f.write("\n".join(body) + "\n")
        r = pwsh("-File", script)
        assert r.returncode == 0 and "done" in r.stdout, r.stdout + r.stderr

    @classmethod
    def tearDownClass(cls):
        shutil.rmtree(cls.tmp, True)

    def run_pair(self, key, model, mode="", extra="", existing=None):
        text = self.projects[key]
        names = specs_of(text)
        out = []
        for use_plan in (True, False):
            cfg = extra + ("plan_path=%s\n" % self.plans[key] if use_plan else "")
            ex = [tip.template_view(n, 1) for n in (existing or [])]
            r = tip.Run(model, mode=mode, project=text, template_names=names, existing=ex, extra_cfg=cfg).go()
            self.addCleanup(r.cleanup)
            plan_lines = [l for l in r.log if "plan=" in l]
            res = {"error": str(r.error) if r.error else "", "names": r.names(), "saved": r.doc.saved,
                   "log": [l for l in r.log if "plan=" not in l],
                   "csv": [r.read("r1_%s.csv" % c) for c in ("сводка", "параметры", "пропуски", "вставлено")]}
            out.append((res, plan_lines))
        (plan_res, plan_log), (legacy_res, legacy_log) = out
        info = "%s / %s / mode=%s %s" % (key, model, mode or "sections", extra.strip())
        self.assertTrue(any("plan=used" in l for l in plan_log), info + str(plan_log))
        self.assertTrue(any("LEGACY" in l for l in legacy_log), info + str(legacy_log))
        self.assertEqual(plan_res, legacy_res, info)
        return plan_res

    def test_sections_every_model_proj(self):
        inserted = 0
        for key in [k for k in self.projects if k.startswith("proj")]:
            for model in PROJ_MODELS:
                res = self.run_pair(key, model)
                inserted += len(res["names"])
        self.assertGreater(inserted, 50)                 # сравнение не на пустых прогонах

    def test_sections_every_model_samplesite(self):
        reasons = set()
        for model in SITE_MODELS:
            res = self.run_pair("site", model)
            summary = res["csv"][0]
            reasons.add(summary.split("не определён:")[-1].strip() if "не определён:" in summary else "ok")
        self.assertIn("ok", reasons)
        self.assertTrue(any("не размечены" in x for x in reasons), reasons)
        self.assertTrue(any("нет размеченных моделей" in x or "не определён" in x for x in reasons), reasons)

    def test_summary_all_projects(self):
        for key in self.projects:
            self.run_pair(key, "C:\\m\\%s_IOS_SVOD_ORG_R24.rvt" % ("SITE" if key == "site" else "PROJ"), mode="summary")

    def test_add_diff_replace_and_auto_initial(self):
        some = ["РД" + P8 + "1_Классификатор_Все"]
        for extra in ("insert_mode=add_diff\n", "insert_mode=replace\n"):
            for existing in (None, some):                 # нет спецификаций проекта в модели — автоматически первичная
                for model in ("C:\\m\\PROJ_IOS_VK_ORG_K01_R24.rvt", "C:\\m\\PROJ_IOS_VK_ORG_K09_R24.rvt", "C:\\m\\PROJ_IOS_ITP_ORG_K01_R24.rvt"):
                    res = self.run_pair("proj_itp", model, extra=extra, existing=existing)
                    if existing is None:
                        self.assertTrue(any("auto" in l for l in res["log"]), res["log"])
                self.run_pair("proj_marked", "C:\\m\\PROJ_IOS_SVOD_ORG_R24.rvt", mode="summary", extra=extra, existing=existing)
            site_some = [specs_of(self.projects["site"])[1]]
            self.run_pair("site", "C:\\m\\SITE_ORG_AVT_K1_R22.rvt", extra=extra, existing=site_some)
            self.run_pair("site", "C:\\m\\SITE_ORG_AVT_K1_R22.rvt", extra=extra)

    def test_broken_or_foreign_plan_falls_back_to_project_file(self):
        bad = os.path.join(self.tmp, "bad_plan.txt")
        with open(self.plans["proj"], encoding="utf-8-sig") as f:
            lines = f.read().splitlines()
        for variant in ([l for l in lines if not l.startswith("end")], [l.replace("version\t1", "version\t9") for l in lines], lines[:-3] + lines[-1:]):
            with open(bad, "w", encoding="utf-8-sig") as f:
                f.write("\n".join(variant) + "\n")
            r = tip.Run("C:\\m\\PROJ_IOS_VK_ORG_K01_R24.rvt", extra_cfg="plan_path=%s\n" % bad).go()
            self.addCleanup(r.cleanup)
            self.assertIsNone(r.error)
            self.assertTrue(any("plan=invalid" in l and "LEGACY" in l for l in r.log), r.log)
            self.assertEqual(r.names(), sorted(["РД" + P8 + n for n in ("1_Классификатор_Все", "7.2_Диаметр_Трубы_Все", "7.2_Диаметр_Трубы_Без значения")]))

    def test_plan_is_used_without_reading_project_file(self):
        """С планом файл настройки не разбирается: подменённый после записи плана файл не влияет на набор."""
        text = self.projects["proj"]
        r = tip.Run("C:\\m\\PROJ_IOS_VK_ORG_K01_R24.rvt", project=text.replace("\t1\ta\t7", "\t0\ta\t7"),
                    extra_cfg="plan_path=%s\n" % self.plans["proj"]).go()
        self.addCleanup(r.cleanup)
        self.assertIsNone(r.error)
        self.assertIn("РД" + P8 + "7.2_Диаметр_Трубы_Все", r.names())

    def test_plan_file_shape(self):
        with open(self.plans["site"], encoding="utf-8-sig") as f:
            lines = [l for l in f.read().splitlines() if l and not l.startswith("#")]
        self.assertEqual(lines[0], "version\t1")
        self.assertEqual(lines[1], "stage\tРД")
        self.assertEqual(lines[-1], "end\t%d" % (len(lines) - 1))
        models = [l.split("\t")[1] for l in lines if l.startswith("model\t")]
        self.assertNotIn("RG_ORG_BF_K1_R22", models)          # исключённые модели в план не попадают
        self.assertNotIn("RG_ORG_GU_K1_R22", models)
        self.assertNotIn("SITE_X_K7_R22", models)               # раздел по папке: как в прежнем разборе, ищется по разделу
        self.assertIn("model\tSITE_ORG_HC_K2_R22\tunmarked", lines)


class PlanWiringTests(unittest.TestCase):
    def test_ui_writes_plan_for_summary_and_sections_only(self):
        with open(os.path.join(TOOL, "ui_schedules.ps1"), encoding="utf-8-sig") as f:
            ui = f.read()
        start = ui[ui.index("function Start-Batch"):]
        start = start[:start.index("\nfunction ")]
        i = start.index("Write-InsertPlan $projectPath (Join-Path $work \"plan.txt\")")
        self.assertIn('if ($Mode -eq "summary" -or $Mode -eq "sections") {', start[i - 400:i])
        self.assertIn('"plan_path={0}"', start[i:i + 300])
        self.assertLess(i, start.index("Write-RunFiles -Models"))           # plan_path попадает в run.cfg
        self.assertIn("catch", start[i:i + 400])                           # не записался — прогон идёт по файлу настройки

    def test_prefill_keeps_project_file(self):
        with open(os.path.join(TOOL, "insert_schedules.py"), encoding="utf-8") as f:
            src = f.read()
        self.assertIn('if mode != u"prefill":\n            plan = load_plan(plan_path)', src)
        self.assertIn("LEGACY", src[src.index("def load_project"):][:300] + src[:src.index("def load_project")][-900:])


if __name__ == "__main__":
    unittest.main()
