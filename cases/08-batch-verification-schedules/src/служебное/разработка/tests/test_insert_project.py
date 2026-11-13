"""insert_schedules.py на подставном Revit API: наборы спецификаций из настройки проекта (по моделям), стадия и шифр, предзаполнение (подсчёт с откатом)."""
import builtins
import os
import shutil
import sys
import tempfile
import types
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import revit_api_stub as al  # noqa: E402  подставной Revit API (Doc, View, Definition, install)
import test_compare_filters as cmp  # noqa: E402  IronPythonDotNet

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(HERE)))
TOOL = os.path.join(ROOT, "служебное", "скрипты")      # папка с .ps1/.py окна
with open(os.path.join(TOOL, "insert_schedules.py"), encoding="utf-8") as _f:
    SOURCE = _f.read()
REAL_OPEN = builtins.open
P8 = " " * 8


class Param:
    def __init__(self):
        self.Definition = types.SimpleNamespace(Name="ADSK_Назначение вида")
        self.IsReadOnly = False
        self.StorageType = "String"
        self.value = None

    def Set(self, v):
        self.value = v


class TView(al.View):
    def __init__(self, name, definition, count=3):
        super().__init__(name, definition)
        self.Parameters = [Param()]
        self.count = count


def template_view(name, count=3, with_link=True, value=None):
    d = al.Definition()
    d.add_field("Код по классификатору")
    fid = d.add_field("ADSK_Размер")
    d.filters.append(al.Flt(fid, "NotEqual", value if value is not None else "999"))
    if with_link:
        d.filters.append(al.Flt(d.add_field(al.LINK), "NotContains", "_AR_"))      # поле связи — последнее: после его удаления номера остальных не сдвигаются
    return TView(name, d, count)


al.Definition.SetFilter = lambda self, i, f: self.filters.__setitem__(i, f)
al.Doc.SaveAs = lambda self, path, opts: self.Save()


class Group:
    def __init__(self, doc, name):
        self.doc = doc

    def Start(self):
        self.snap = [(v, v.Name, [(f.FieldId, f.FilterType, f.value) for f in v.Definition.filters]) for v in self.doc.views]

    def RollBack(self):
        self.doc.views[:] = [v for v, _, _ in self.snap]
        for v, name, filters in self.snap:
            v.Name = name


def install_insert(doc, path):
    saved, rsu = al.install(doc, path)
    db = sys.modules["Autodesk.Revit.DB"]
    base_collector = db.FilteredElementCollector

    class Collector(base_collector):
        def __init__(self, d, view_id=None):
            super().__init__(d)
            self.view_id = view_id

        def GetElementCount(self):
            for v in self.d.views:
                if v.Id == self.view_id:
                    if getattr(v, "count_fail", False):
                        raise Exception("collector failed")
                    return v.count
            raise Exception("view not found")

    db.FilteredElementCollector = Collector
    base_copy = db.ElementTransformUtils.CopyElements

    class Utils:
        @staticmethod
        def CopyElements(source, ids, target, transform, options):
            Utils.calls.append(len(list(ids)))
            if Utils.fail_batch and len(list(ids)) > 1:
                raise Exception("пачка не поддерживается")
            result = base_copy(source, ids, target, transform, options)
            for eid, sid in zip(result, ids):
                v = target.GetElement(eid)
                v.Parameters = [Param()]
                src = [x for x in source.views if x.Id == sid][0]
                v.count = getattr(src, "count", 0)
                v.count_fail = getattr(src, "count_fail", False)
            return result

        calls = []
        fail_batch = False

    db.ElementTransformUtils = Utils
    db.TransactionGroup = Group
    for n in ("SaveAsOptions", "WorksharingSaveAsOptions"):
        setattr(db, n, type(n, (), {}))
    db.SectionType = types.SimpleNamespace(Body="Body")
    db.ScheduleFieldType = types.SimpleNamespace(Count="Count", Formula="Formula")
    # View.Name: имя занято — как в Revit, исключение
    return saved, rsu


class Writers:
    """.NET StreamReader/StreamWriter поверх настоящих файлов (open() на кириллице под эмуляцией падает)."""

    class Reader:
        def __init__(self, p, detect=True):
            with REAL_OPEN(p, encoding="utf-8-sig") as f:
                self.text = f.read()

        def ReadToEnd(self):
            return self.text

        def Close(self):
            pass

    class Writer:
        def __init__(self, p, append, enc=None):
            self.f = REAL_OPEN(p, "a" if append else "w", encoding="utf-8", newline="")

        def Write(self, t):
            self.f.write(t)

        def Close(self):
            self.f.close()


PROJECT = """# test
[project]
stage=РД
template=C:\\шаблон.rvt
[models]
PROJ_IOS_OV_ORG_K01_R24\tOV
PROJ_IOS_VK_ORG_K01_R24\tVK
PROJ_IOS_VK_ORG_K02_R24\tVK
[specs]
1_Классификатор_Все
5.1_Ширина_Воздуховоды_Все
7.2_Диаметр_Трубы_Все
7.2_Диаметр_Трубы_Без значения
[cells]
PROJ_IOS_OV_ORG_K01_R24\t1_Классификатор_Все\t1\ta\t10
PROJ_IOS_OV_ORG_K01_R24\t5.1_Ширина_Воздуховоды_Все\t1\ta\t5
PROJ_IOS_OV_ORG_K01_R24\t7.2_Диаметр_Трубы_Все\t0\ta\t0
PROJ_IOS_OV_ORG_K01_R24\t7.2_Диаметр_Трубы_Без значения\t0\ta\t0
PROJ_IOS_VK_ORG_K01_R24\t1_Классификатор_Все\t1\ta\t10
PROJ_IOS_VK_ORG_K01_R24\t5.1_Ширина_Воздуховоды_Все\t0\ta\t0
PROJ_IOS_VK_ORG_K01_R24\t7.2_Диаметр_Трубы_Все\t1\ta\t7
PROJ_IOS_VK_ORG_K01_R24\t7.2_Диаметр_Трубы_Без значения\t1\tm\t-1
PROJ_IOS_VK_ORG_K02_R24\t1_Классификатор_Все\t1\ta\t10
PROJ_IOS_VK_ORG_K02_R24\t5.1_Ширина_Воздуховоды_Все\t0\ta\t0
PROJ_IOS_VK_ORG_K02_R24\t7.2_Диаметр_Трубы_Все\t1\ta\t7
PROJ_IOS_VK_ORG_K02_R24\t7.2_Диаметр_Трубы_Без значения\t0\ta\t0
"""


class Run:
    def __init__(self, model, mode="", project=PROJECT, template_names=None, existing=None, extra_cfg="", project_exists=True, counts=None,
                 empty_path=False, fail_batch=False):
        self.tmp = tempfile.mkdtemp(prefix="bvs_ins_")
        self.model = model
        self.reports = os.path.join(self.tmp, "служебные", "20261023_120000")
        os.makedirs(self.reports)
        names = template_names or ["1_Классификатор_Все", "5.1_Ширина_Воздуховоды_Все", "7.2_Диаметр_Трубы_Все", "7.2_Диаметр_Трубы_Без значения"]
        counts = counts or {}
        tviews = [template_view(n, counts.get(n, 3), value="{ШИФР}_{стадия}" if n.startswith("5.1") else None) for n in names]
        self.template = al.Doc(tviews)
        self.doc = al.Doc(list(existing or []), {"C:\\шаблон.rvt": self.template})
        self.fail_batch = fail_batch
        self.project_path = os.path.join(self.tmp, "проект", "распределение_PROJ.txt")
        os.makedirs(os.path.dirname(self.project_path))
        if project_exists:
            with REAL_OPEN(self.project_path, "w", encoding="utf-8-sig", newline="") as f:
                f.write(project)
        self.out = os.path.join(self.reports, "prefill.tsv")
        # LOCAL_CFG в скрипте собран с обратными слэшами (Windows); на Linux это одно имя файла
        cfg_path = os.path.join(self.tmp, "BatchRvt\\batch_verification_schedules\\run.cfg")
        if os.sep == "\\":                                                   # Windows: это настоящие папки, их нужно создать
            os.makedirs(os.path.dirname(cfg_path), exist_ok=True)
        with REAL_OPEN(cfg_path, "w", encoding="utf-8-sig") as f:
            f.write("source=C:\\шаблон.rvt\nreport_dir=%s\nrun_id=r1\nproject_path=%s\nmode=%s\nprefill_out=%s\ninsert_mode=initial\n%s"
                    % (self.reports, "" if empty_path else self.project_path, mode, self.out, extra_cfg))
        self.old = os.environ.get("LOCALAPPDATA")
        os.environ["LOCALAPPDATA"] = self.tmp
        self.addCleanup = shutil.rmtree

    def go(self):
        saved, rsu = install_insert(self.doc, self.model)
        sys.modules["Autodesk.Revit.DB"].ElementTransformUtils.fail_batch = self.fail_batch
        sys.modules["Autodesk.Revit.DB"].ElementTransformUtils.calls = self.calls = []
        self.error = None
        try:
            with cmp.IronPythonDotNet() as net:
                sio = sys.modules["System.IO"]
                sio.StreamReader, sio.StreamWriter = Writers.Reader, Writers.Writer
                exists = sio.File.Exists
                sio.File.Exists = lambda p: p == "C:\\шаблон.rvt" or exists(p)
                sys.modules["System.Text"].UTF8Encoding = lambda b=False: b
                try:
                    exec(compile(SOURCE, "insert_schedules.py", "exec"), {"__name__": "__main__"})
                except Exception as ex:
                    self.error = ex
        finally:
            al.restore(saved)
            if self.old is None:
                os.environ.pop("LOCALAPPDATA", None)
            else:
                os.environ["LOCALAPPDATA"] = self.old
        self.log = list(rsu.LOG)
        return self

    def read(self, name):
        p = os.path.join(self.reports, name)
        if not os.path.exists(p):
            return ""
        with REAL_OPEN(p, encoding="utf-8-sig") as f:
            return f.read()

    def names(self):
        return sorted(v.Name for v in self.doc.views)

    def cleanup(self):
        shutil.rmtree(self.tmp, True)


class InsertProjectTests(unittest.TestCase):
    def run_it(self, *a, **k):
        r = Run(*a, **k).go()
        self.addCleanup(r.cleanup)
        return r

    def test_each_model_gets_its_own_marked_specs_with_stage_prefix_and_profile(self):
        r = self.run_it("\\\\srv\\проект\\PROJ_IOS_VK_ORG_K01_R24.rvt")
        self.assertIsNone(r.error)
        self.assertEqual(r.names(), sorted(["РД" + P8 + n for n in ("1_Классификатор_Все", "7.2_Диаметр_Трубы_Все", "7.2_Диаметр_Трубы_Без значения")]))
        r2 = self.run_it("C:\\m\\PROJ_IOS_VK_ORG_K02_R24.rvt")                      # в том же разделе — другой набор (у К02 нет «Без значения»)
        self.assertEqual(r2.names(), sorted(["РД" + P8 + n for n in ("1_Классификатор_Все", "7.2_Диаметр_Трубы_Все")]))
        r3 = self.run_it("C:\\m\\PROJ_IOS_OV_ORG_K01_R24.rvt")
        self.assertEqual(r3.names(), sorted(["РД" + P8 + n for n in ("1_Классификатор_Все", "5.1_Ширина_Воздуховоды_Все")]))
        v = [x for x in r3.doc.views if "5.1_" in x.Name][0]
        self.assertIn(("ADSK_Размер", "NotEqual", "{ШИФР}_РД"), [(v.Definition.fields[f.FieldId].name, f.FilterType, f.value) for f in v.Definition.filters])   # {СТАДИЯ} подставлена, {ШИФР} больше не поддерживается
        self.assertFalse(any(f.name == al.LINK for f in v.Definition.fields))          # поле «Связанный файл» убрано
        self.assertEqual(v.Parameters[0].value, "_Проверка проекта")
        self.assertEqual(r3.doc.saved, 1)
        self.assertEqual(len(r3.read("r1_вставлено.csv").splitlines()), 2)

    def test_template_names_with_old_prefix_and_typo_are_found(self):
        r = self.run_it("C:\\m\\PROJ_IOS_VK_ORG_K01_R24.rvt",
                        template_names=["П" + P8 + "1_Классификатор_Все", "П" + P8 + "7.2_Диаметр_Трубы_Все", "П" + P8 + "7.2_Диаметр_Трубы_Без значния"])
        self.assertEqual(r.names(), sorted(["РД" + P8 + n for n in ("1_Классификатор_Все", "7.2_Диаметр_Трубы_Все", "7.2_Диаметр_Трубы_Без значения")]))

    def test_latin_c_in_model_and_cyrillic_in_template_are_one_spec_no_duplicate(self):
        lat, cyr = "\u0043антехприборы", "\u0421антехприборы"
        proj = PROJECT.replace("7.2_Диаметр_Трубы_Без значения\n[cells]", "7.2_Диаметр_Трубы_Без значения\n2.5_Клас-я системы_%s_Все\n[cells]" % lat)
        proj = proj.replace("[specs]\n", "[specs]\n2.5_Клас-я системы_%s_Все\n" % lat).replace(
            "PROJ_IOS_VK_ORG_K01_R24\t7.2_Диаметр_Трубы_Все\t1\ta\t7\n", "PROJ_IOS_VK_ORG_K01_R24\t7.2_Диаметр_Трубы_Все\t1\ta\t7\nPROJ_IOS_VK_ORG_K01_R24\t2.5_Клас-я системы_%s_Все\t1\ta\t1\n" % lat, 1)
        tmpl = ["1_Классификатор_Все", "2.5_Клас-я системы_%s_Все" % cyr, "7.2_Диаметр_Трубы_Все", "7.2_Диаметр_Трубы_Без значения", "5.1_Ширина_Воздуховоды_Все"]
        old = "РД" + P8 + "2.5_Клас-я системы_%s_Все" % lat
        for extra in ("insert_mode=add_diff\n", ""):
            r = self.run_it("C:\\m\\PROJ_IOS_VK_ORG_K01_R24.rvt", project=proj, template_names=tmpl, extra_cfg=extra,
                            existing=[template_view(old)])
            self.assertIsNone(r.error)
            self.assertEqual([n for n in r.names() if "антехприборы" in n], [old])               # дубль не создан, старое имя не тронуто

    def test_summary_mode_takes_union_and_keeps_links(self):
        r = self.run_it("C:\\m\\PROJ_IOS_SVOD_ORG_R24.rvt", mode="summary")
        self.assertIsNone(r.error)
        self.assertEqual(len(r.doc.views), 4)                                    # объединение всех отмеченных
        v = r.doc.views[0]
        self.assertTrue(any(f.name == al.LINK for f in v.Definition.fields))     # в сводной связи остаются

    def test_replace_mode_does_not_save_when_insertion_is_partial(self):
        """replace: старые спецификации удаляются в памяти до вставки; если вставилось не всё — модель не сохраняется (нет неполного комплекта)."""
        old = [template_view("РД" + P8 + "1_Классификатор_Все", 1)]
        partial = ["1_Классификатор_Все", "7.2_Диаметр_Трубы_Все"]                      # в шаблоне нет «7.2_…_Без значения», нужной модели VK K01
        r = self.run_it("C:\\m\\PROJ_IOS_VK_ORG_K01_R24.rvt", extra_cfg="insert_mode=replace\n", template_names=partial, existing=list(old))
        self.assertEqual(r.doc.saved, 0)
        self.assertIn("Модель не сохранена", r.read("r1_сводка.csv"))
        self.assertIn("нет в файле со спецификациями", r.read("r1_пропуски.csv"))
        full = self.run_it("C:\\m\\PROJ_IOS_VK_ORG_K01_R24.rvt", extra_cfg="insert_mode=replace\n")
        self.assertEqual(full.doc.saved, 1)                                            # всё вставилось — сохраняется
        self.assertNotIn("Модель не сохранена", full.read("r1_сводка.csv"))
        initial = self.run_it("C:\\m\\PROJ_IOS_VK_ORG_K01_R24.rvt", template_names=partial)
        self.assertEqual(initial.doc.saved, 1)                                         # не replace: частичная вставка сохраняется, как раньше

    def test_model_without_project_schedules_gets_initial_load_in_any_mode(self):
        """4.2: в модели нет ни одной спецификации списка — первичная загрузка независимо от режима; в отчёте «первичная (автоматически)»."""
        for mode in ("add_diff", "replace"):
            r = self.run_it("C:\\m\\PROJ_IOS_VK_ORG_K01_R24.rvt", extra_cfg="insert_mode=%s\n" % mode)
            summary = r.read("r1_сводка.csv").splitlines()[-1].split(";")
            self.assertEqual(summary[2], "0", mode)                                     # ничего не удалялось
            self.assertEqual(summary[6], "первичная (автоматически)", mode)
            self.assertEqual(r.doc.saved, 1, mode)
            self.assertEqual(len(r.names()), 3, mode)                                   # вставлен весь набор VK
            self.assertTrue(any("auto: no project schedules" in l for l in r.log), mode)

    def test_mode_label_for_models_that_already_have_schedules(self):
        old = [template_view("РД" + P8 + "1_Классификатор_Все", 1)]
        add = self.run_it("C:\\m\\PROJ_IOS_VK_ORG_K01_R24.rvt", extra_cfg="insert_mode=add_diff\n", existing=list(old))
        row = add.read("r1_сводка.csv").splitlines()[-1].split(";")
        self.assertEqual(row[6], "добавление")                                        # режим сохранён: часть спецификаций уже есть
        self.assertIn("уже в пакете", add.read("r1_пропуски.csv"))
        rep = self.run_it("C:\\m\\PROJ_IOS_VK_ORG_K01_R24.rvt", extra_cfg="insert_mode=replace\n", existing=[template_view("РД" + P8 + "1_Классификатор_Все", 1)])
        row = rep.read("r1_сводка.csv").splitlines()[-1].split(";")
        self.assertEqual(row[6], "замена")                                           # режим сохранён: спецификация списка в модели была
        first = self.run_it("C:\\m\\PROJ_IOS_VK_ORG_K01_R24.rvt")
        self.assertEqual(first.read("r1_сводка.csv").splitlines()[-1].split(";")[6], "первичная")      # режим initial выбран явно (по умолчанию в run.cfg)

    def test_unknown_model_uses_section_union_or_reports_reason(self):
        r = self.run_it("C:\\m\\PROJ_IOS_VK_ORG_K09_R24.rvt")                  # новой модели нет в настройке: объединение раздела ВК
        self.assertEqual(len(r.doc.views), 3)
        r2 = self.run_it("C:\\m\\Чужая_модель.rvt")
        self.assertEqual(len(r2.doc.views), 0)
        self.assertIn("Набор спецификаций модели не определён", r2.read("r1_сводка.csv"))
        self.assertEqual(r2.doc.saved, 0)

    def test_missing_project_file_is_a_clear_error_not_hardcoded_packs(self):
        r = self.run_it("C:\\m\\PROJ_IOS_VK_ORG_K01_R24.rvt", project_exists=False)
        self.assertEqual(len(r.doc.views), 0)
        self.assertIn("Файл настройки проекта не найден", r.read("r1_сводка.csv"))

    def test_empty_project_path_refuses_to_run_with_clear_message(self):
        for mode in ("", "prefill"):
            r = self.run_it("C:\\m\\PROJ_IOS_VK_ORG_K01_R24.rvt", mode=mode, empty_path=True)
            self.assertEqual(len(r.doc.views), 0)
            self.assertEqual(r.doc.saved, 0)
            text = r.read("r1_сводка.csv") + r.read("prefill.tsv")
            self.assertIn("Не задан путь к файлу настройки проекта", text)
            self.assertEqual(r.calls, [])                                          # ничего не копировалось

    def test_prefill_counts_with_rollback_and_timing(self):
        existing = [template_view("РД" + P8 + "1_Классификатор_Все", 1)]
        r = self.run_it("C:\\m\\PROJ_IOS_VK_ORG_K01_R24.rvt", mode="prefill", existing=existing,
                        counts={"1_Классификатор_Все": 412, "5.1_Ширина_Воздуховоды_Все": 0, "7.2_Диаметр_Трубы_Все": 7},
                        template_names=["1_Классификатор_Все", "5.1_Ширина_Воздуховоды_Все", "7.2_Диаметр_Трубы_Все"])
        self.assertIsNone(r.error)
        rows = [l.split("\t") for l in r.read("prefill.tsv").splitlines()]
        data = {x[1]: x for x in rows if not x[0].startswith("#")}
        self.assertEqual(data["1_Классификатор_Все"][2], "412")          # в модели уже была такая спека: считается копия из шаблона под временным именем
        self.assertEqual(data["5.1_Ширина_Воздуховоды_Все"][2], "0")
        self.assertEqual(data["7.2_Диаметр_Трубы_Все"][2], "7")
        self.assertEqual(data["7.2_Диаметр_Трубы_Без значения"][2], "-1")    # в шаблоне нет — «не удалось определить»
        self.assertEqual(data["7.2_Диаметр_Трубы_Без значения"][4], "нет в шаблоне")
        self.assertEqual(data["1_Классификатор_Все"][4], "collector")
        total = [x for x in rows if x[0] == "#МОДЕЛЬ"][0]
        self.assertEqual((total[1], total[3], total[4]), ("PROJ_IOS_VK_ORG_K01_R24", "4", "1"))
        self.assertGreaterEqual(float(total[2]), 0.0)                          # время прогона по модели записано
        self.assertTrue(all(float(x[3]) >= 0 for x in rows if not x[0].startswith("#")))
        # откат: модель как была, не сохранена
        self.assertEqual([v.Name for v in r.doc.views], ["РД" + P8 + "1_Классификатор_Все"])
        self.assertEqual(r.doc.saved, 0)
        self.assertEqual(r.template.closed, False)                             # шаблон закрыт без сохранения
        self.assertTrue(any("prefill" in l and "total=" in l and "mode=batch" in l for l in r.log))
        # одна пачка на модель: один CopyElements сразу со всеми найденными спецификациями (3 из шаблона), стадия применена
        self.assertEqual(r.calls, [3])
        self.assertEqual(total[6], "пачка")
        self.assertEqual(len(total), 12)                                           # + открытие, вставка, профиль, подсчёт, откат
        self.assertTrue(all(float(x) >= 0 for x in total[7:]))

    def test_prefill_falls_back_to_one_by_one_and_logs_it(self):
        r = self.run_it("C:\\m\\PROJ_IOS_VK_ORG_K01_R24.rvt", mode="prefill", fail_batch=True,
                        counts={"1_Классификатор_Все": 5, "7.2_Диаметр_Трубы_Все": 0})
        rows = [l.split("\t") for l in r.read("prefill.tsv").splitlines()]
        data = {x[1]: x for x in rows if not x[0].startswith("#")}
        self.assertEqual(data["1_Классификатор_Все"][2], "5")
        self.assertEqual(data["7.2_Диаметр_Трубы_Все"][2], "0")
        total = [x for x in rows if x[0] == "#МОДЕЛЬ"][0]
        self.assertEqual(total[6], "по одной")
        self.assertTrue(any("falling back to one-by-one" in l for l in r.log))
        self.assertEqual(r.calls[0], 4)                                            # сначала пачка, потом по одной
        self.assertTrue(all(c == 1 for c in r.calls[1:]))
        self.assertEqual(len(r.doc.views), 0)                                      # откат

    def test_prefill_takes_only_in_project_specs_and_only_engineering_models(self):
        marked = PROJECT.replace("[cells]", "[inproject]\n1_Классификатор_Все\n7.2_Диаметр_Трубы_Все\n[cells]")
        r = self.run_it("C:\\m\\PROJ_IOS_VK_ORG_K01_R24.rvt", mode="prefill", project=marked)
        names = [l.split("\t")[1] for l in r.read("prefill.tsv").splitlines() if not l.startswith("#")]
        self.assertEqual(names, ["1_Классификатор_Все", "7.2_Диаметр_Трубы_Все"])
        self.assertEqual(r.calls, [2])
        r2 = self.run_it("C:\\m\\PROJ_IOS_TASK_K01_R24.rvt", mode="prefill", project=marked)          # без кода раздела
        self.assertIn("без кода раздела", r2.read("prefill.tsv"))
        self.assertEqual(r2.calls, [])
        r3 = self.run_it("C:\\m\\PROJ_IOS_VK_ORG_K01_R24.rvt", mode="prefill", project=PROJECT.replace("[cells]", "[inproject]\n[cells]"))
        self.assertIn("нет отмеченных групп", r3.read("prefill.tsv"))

    def test_inproject_filters_insert_and_summary(self):
        marked = PROJECT.replace("[cells]", "[inproject]\n1_Классификатор_Все\n7.2_Диаметр_Трубы_Без значения\n[cells]")
        r = self.run_it("C:\\m\\PROJ_IOS_VK_ORG_K01_R24.rvt", project=marked)
        self.assertEqual(r.names(), sorted(["РД" + P8 + n for n in ("1_Классификатор_Все", "7.2_Диаметр_Трубы_Без значения")]))   # «Все» труб отмечена в ячейке, но группа не в проекте
        s = self.run_it("C:\\m\\PROJ_IOS_SVOD_ORG_R24.rvt", mode="summary", project=marked)
        self.assertEqual(len(s.doc.views), 2)                                      # сводная — все спецификации групп «В проекте»

    def test_summary_section_limits_summary_to_marked_specs(self):
        marked = PROJECT.replace("[cells]", "[inproject]\n1_Классификатор_Все\n5.1_Ширина_Воздуховоды_Все\n[summary]\n5.1_Ширина_Воздуховоды_Все\n[cells]")
        s = self.run_it("C:\\m\\PROJ_IOS_SVOD_ORG_R24.rvt", mode="summary", project=marked)
        self.assertEqual(len(s.doc.views), 1)                                      # из «В проекте» в сводную идёт только отмеченное в «Сводная»
        r = self.run_it("C:\\m\\PROJ_IOS_OV_ORG_K01_R24.rvt", project=marked)
        self.assertEqual(len(r.doc.views), 2)                                      # модели «Сводная» не ограничивает

    def test_prefill_falls_back_to_table_rows_when_collector_fails(self):
        r = Run("C:\\m\\PROJ_IOS_VK_ORG_K01_R24.rvt", mode="prefill", template_names=["1_Классификатор_Все"])
        self.addCleanup(r.cleanup)
        r.template.views[0].count_fail = True
        saved, rsu = None, None
        # у подставного вида нет таблицы: способ «table» тоже не сработает — спека «не определена», а не падение всего прогона
        r.go()
        rows = [l.split("\t") for l in r.read("prefill.tsv").splitlines()]
        data = [x for x in rows if x[1] == "1_Классификатор_Все"][0]
        self.assertEqual(data[2], "-1")
        self.assertIn("не удалось посчитать", data[4])
        self.assertIsNone(r.error)

    def test_script_has_no_hardcoded_section_tables(self):
        for bad in ("SECTION_PACKS", "SECTION_MARKS", "schedule_sets", "names_for_section", "detect_section", "load_sets", 'find(u"_IOS_")', "{шифр}"):
            self.assertNotIn(bad, SOURCE)


def install_link_collector(names):
    """Подставные RevitLinkType: коллектор отдаёт связи только для этого класса, спецификации — как раньше."""
    db = sys.modules["Autodesk.Revit.DB"]

    class RevitLinkType(object):
        def __init__(self, name):
            self.Name = name

    objects = [RevitLinkType(n) for n in names]
    db.RevitLinkType = RevitLinkType
    base = db.FilteredElementCollector

    class Collector(base):
        def OfClass(self, cls):
            if cls is RevitLinkType:
                return list(objects)
            return list(self.d.views)

    db.FilteredElementCollector = Collector


class ReadTemplateTests(unittest.TestCase):
    """read_template_specs.py: список спецификаций шаблона (только нумерованные проверки), файлы — через .NET."""

    def run_it(self, only_numbered="1", error=False, links=None, with_links_path=False, log_folder=None):
        tmp = tempfile.mkdtemp(prefix="bvs_rt_")
        self.addCleanup(shutil.rmtree, tmp, True)
        cfg_dir = os.path.join(tmp, "BatchRvt", "batch_verification_schedules")
        os.makedirs(cfg_dir)
        out = os.path.join(tmp, "служебные", "20261023_120000", "template_specs.tsv")
        os.makedirs(os.path.dirname(out))
        links_file = os.path.join(os.path.dirname(out), "template_links.txt")
        cfg_text = "out_path=%s\nonly_numbered=%s\n" % (out, only_numbered)
        if with_links_path:
            cfg_text += "links_path=%s\n" % links_file
        if log_folder:
            os.makedirs(log_folder, exist_ok=True)
            cfg_text += "log_folder=%s\n" % log_folder
        with REAL_OPEN(os.path.join(cfg_dir, "run.cfg"), "w", encoding="utf-8-sig") as f:
            f.write(cfg_text)
        d = al.Definition()
        names = ["П" + P8 + "3.1_Имя системы_Пустые", "2.4_Без_префикса_Все", "РД" + P8 + "13_Материал_Все", "Пров" + P8 + "Ошибка типоразмера",
                 "СЗ_ИОС   Отверстия_С_КР", "Раздел 1 2.1_Предварительные нагрузки"]
        views = [al.View(n, d) for n in names] + [al.View("0_Шаблон вида", d, template=True)]
        doc = al.Doc(views)
        saved, rsu = al.install(doc, "C:\\шаблон.rvt")
        if links is not None:
            install_link_collector(links)
        old = os.environ.get("LOCALAPPDATA")
        os.environ["LOCALAPPDATA"] = tmp
        src = open(os.path.join(TOOL, "read_template_specs.py"), encoding="utf-8").read()
        err = None
        try:
            with cmp.IronPythonDotNet() as net:
                if error:
                    rsu.GetScriptDocument = lambda: (_ for _ in ()).throw(IOError("boom"))
                try:
                    exec(compile(src, "read_template_specs.py", "exec"), {"__name__": "__main__"})
                except Exception as ex:
                    err = ex
                tf = os.path.join(net.temp, "read_template_specs_error.txt")
                if log_folder:
                    alt = os.path.join(log_folder, "read_template_specs_error.txt")
                    if os.path.exists(alt):
                        tf = alt
                err_file = REAL_OPEN(tf, encoding="utf-8-sig").read() if os.path.exists(tf) else None
        finally:
            self.last_log = list(getattr(rsu, "LOG", []) or [])
            self.links_file = links_file if with_links_path else None
            al.restore(saved)
            if old is None:
                os.environ.pop("LOCALAPPDATA", None)
            else:
                os.environ["LOCALAPPDATA"] = old
        text = REAL_OPEN(out, encoding="utf-8-sig").read() if os.path.exists(out) else None
        return text, err, err_file, doc

    def test_only_numbered_checks(self):
        text, err, _, doc = self.run_it()
        self.assertIsNone(err)
        names = [l.split("\t")[0] for l in text.splitlines()]
        self.assertEqual(names, ["П" + P8 + "3.1_Имя системы_Пустые", "2.4_Без_префикса_Все", "РД" + P8 + "13_Материал_Все"])
        self.assertEqual((doc.saved, doc.committed), (0, 0))                       # шаблон не меняется и не сохраняется

    def test_all_schedules_when_filter_off(self):
        text, err, _, _ = self.run_it("0")
        self.assertEqual(len(text.splitlines()), 6)                                 # без шаблона вида

    def test_error_writes_traceback_to_temp(self):
        text, err, err_file, _ = self.run_it(error=True)
        self.assertIsInstance(err, IOError)
        self.assertIn("Traceback", err_file)
        self.assertIn("step:", err_file)

    def test_error_writes_to_log_folder_from_run_cfg(self):
        folder = os.path.join(tempfile.mkdtemp(prefix="bvs_rtlog_"), "логи")
        self.addCleanup(shutil.rmtree, os.path.dirname(folder), True)
        text, err, err_file, _ = self.run_it(error=True, log_folder=folder)
        self.assertIsInstance(err, IOError)
        self.assertTrue(os.path.isfile(os.path.join(folder, "read_template_specs_error.txt")))
        self.assertIn("Traceback", err_file)
        self.assertIn("step:", err_file)

    def test_links_written_beside_specs_without_rvt(self):
        names = [
            "PROJ_IOS_OV_ORG_K01.rvt",
            "PROJ_IOS_VK_ORG_K01.RVT : 1",
            "RSN://srv/Проект/PROJ_IOS_ВК_ORG_K02.rvt",
            "PROJ_IOS_OV_ORG_K01.rvt",
            "Сводная.rvt",
        ]
        text, err, _, doc = self.run_it(links=names, with_links_path=True)
        self.assertIsNone(err)
        spec_names = [l.split("\t")[0] for l in text.splitlines() if l]
        self.assertEqual(spec_names, ["П" + P8 + "3.1_Имя системы_Пустые", "2.4_Без_префикса_Все", "РД" + P8 + "13_Материал_Все"])
        self.assertNotIn("PROJ_IOS", text)
        with REAL_OPEN(self.links_file, encoding="utf-8-sig") as f:
            links = [l for l in f.read().splitlines() if l]
        self.assertEqual(links, [
            "PROJ_IOS_OV_ORG_K01",
            "PROJ_IOS_VK_ORG_K01",
            "PROJ_IOS_ВК_ORG_K02",
            "PROJ_IOS_OV_ORG_K01",
            "Сводная",
        ])
        self.assertTrue(self.last_log)
        for line in self.last_log:
            self.assertTrue(all(ord(c) < 128 for c in line), line)
        self.assertTrue(any("links=5" in line for line in self.last_log))
        self.assertEqual((doc.saved, doc.committed), (0, 0))

    def test_missing_link_types_do_not_drop_specs(self):
        text, err, _, _ = self.run_it(with_links_path=True)
        self.assertIsNone(err)
        self.assertEqual(len([l for l in text.splitlines() if l]), 3)
        with REAL_OPEN(self.links_file, encoding="utf-8-sig") as f:
            self.assertEqual(f.read().strip(), "")
        self.assertTrue(any("links=0" in line for line in self.last_log))


if __name__ == "__main__":
    unittest.main()
