"""Выбор «проверяемого параметра» из фильтров спецификации (export_schedules.py) без Revit."""
import os
import sys
import types
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(HERE))), "служебное", "скрипты", "export_schedules.py")


def load_module():
    text = open(SRC, encoding="utf-8").read()
    assert text.rstrip().endswith("main()")
    text = text.rstrip()[:-len("main()")]
    db = types.ModuleType("Autodesk.Revit.DB")
    db.FilteredElementCollector = object
    db.ViewSchedule = object
    db.SectionType = object
    mods = {"clr": types.SimpleNamespace(AddReference=lambda n: None), "Autodesk": types.ModuleType("Autodesk"),
            "Autodesk.Revit": types.ModuleType("Autodesk.Revit"), "Autodesk.Revit.DB": db,
            "revit_script_util": types.SimpleNamespace(Output=lambda t: None)}
    saved = {k: sys.modules.get(k) for k in mods}
    sys.modules.update(mods)
    try:
        ns = {"__name__": "export_schedules_under_test", "__file__": SRC}
        exec(compile(text, SRC, "exec"), ns)
    finally:
        for k, v in saved.items():
            if v is None:
                sys.modules.pop(k, None)
            else:
                sys.modules[k] = v
    return ns


class FakeFilter:
    def __init__(self, fid, ftype, value):
        self.FieldId, self.FilterType, self.value = fid, ftype, value

    def GetStringValue(self):
        if self.value is None:
            raise TypeError("no string value")
        return self.value


class FakeView:
    def __init__(self, filters):
        self.names = {i: n for i, (n, t, v) in enumerate(filters)}
        flts = [FakeFilter(i, t, v) for i, (n, t, v) in enumerate(filters)]
        names = self.names

        class Field:
            def __init__(self, n):
                self.n = n

            def GetName(self):
                return self.n

        class Definition:
            def GetFilterCount(self):
                return len(flts)

            def GetFilter(self, i):
                return flts[i]

            def GetField(self, fid):
                return Field(names[fid])

        self.Definition = Definition()


class ExportFilterTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.ns = load_module()

    def names(self, filters):
        self.ns["FILTER_DEBUG"][:] = []
        return self.ns["filter_param_names"](FakeView(filters), u"Спец")

    def test_checked_parameter_is_the_has_no_value_filter(self):
        got = self.names([
            ("Код по классификатору", "HasValue", None),
            ("ADSK_Размер_Длина", "HasNoValue", None),
            ("ADSK_Материал наименование", "NotContains", "_IOS_"),
            ("Связанный файл: Имя файла", "Contains", "PROJ"),
        ])
        self.assertEqual(got, ["ADSK_Размер_Длина"])

    def test_empty_spec_equal_to_empty_string(self):
        got = self.names([("Код по классификатору", "NotEqual", ""), ("ADSK_Номер здания", "Equal", "")])
        self.assertEqual(got, ["ADSK_Номер здания"])

    def test_selection_equal_with_value_is_not_the_checked_parameter(self):
        got = self.names([("Код по классификатору", "Equal", "99.9"), ("Имя системы", "HasNoValue", None), ("ADSK_Марка изделия", "HasNoValue", None)])
        self.assertEqual(got, ["Имя системы"])      # «ADSK_Марка изделия» HasNoValue — отбор, не проверка

    def test_fallback_skips_selection_conditions(self):
        got = self.names([("Код по классификатору", "HasValue", None), ("Параметр Р", "Unknown", None)])
        self.assertEqual(got, ["Параметр Р"])

    def test_classifier_never_chosen_in_fallbacks(self):
        got = self.names([("Код по классификатору", "HasValue", None), ("Связанный файл: Имя файла", "Contains", "PROJ")])
        self.assertEqual(got, [])
        got = self.names([("Код по классификатору", "NotEqual", "999"), ("Связанный файл: Имя файла", "Equal", "")])
        self.assertEqual(got, [])

    def test_classifier_spec_checks_classifier_when_equal_empty(self):
        got = self.ns["choose_check_params"]([("Код по классификатору", "Equal", ""), ("Связанный файл: Имя файла", "HasValue", None)],
                                             u"П        1_Классификатор_Пустые")[0]
        self.assertEqual(got, ["Код по классификатору"])

    def test_two_real_checks_pick_the_one_named_in_spec(self):
        pick = lambda name, fields: self.ns["choose_check_params"]([(f, "HasNoValue", None) for f in fields], name)[0]
        self.assertEqual(pick(u"П        5.1_Ширина_Детали возд-в_Без значения", ["ADSK_Размер_Ширина", "ADSK_Размер_Диаметр"]), ["ADSK_Размер_Ширина"])
        self.assertEqual(pick(u"П        6.1_Высота_Детали возд-в_Без значения", ["ADSK_Размер_Высота", "ADSK_Размер_Диаметр"]), ["ADSK_Размер_Высота"])
        self.assertEqual(pick(u"П        7.1_Диаметр_Детали возд-в_Без значения", ["ADSK_Размер_Диаметр", "ADSK_Размер_Высота"]), ["ADSK_Размер_Диаметр"])
        # тема не нашлась среди полей — показываем все
        self.assertEqual(pick(u"П        9.9_Масса_Детали_Без значения", ["ADSK_Размер_Ширина", "ADSK_Размер_Диаметр"]), ["ADSK_Размер_Ширина", "ADSK_Размер_Диаметр"])
        # имя без номера/темы — все
        self.assertEqual(pick(u"Произвольное имя", ["А", "Б"]), ["А", "Б"])

    def test_system_type_topic_maps_to_system_name(self):
        pick = lambda name: self.ns["choose_check_params"]([("Имя системы", "HasNoValue", None), ("ADSK_Марка изделия", "HasNoValue", None)], name)[0]
        self.assertEqual(pick(u"П        3.1_Тип системы_Детали возд-в_Без значния"), ["Имя системы"])

    # Спецификации 7.8 и 13_Материал: в новых выгрузках фильтры уже исправленные
    # фильтры, в старом фильтры_спецификаций.txt — прежние. Проверяемый параметр от исправления не должен «ломаться».
    def test_fixed_13_fittings_empty_checks_material(self):
        got = self.names([("ADSK_Материал наименование", "Equal", ""), ("Связанный файл: Имя файла", "Contains", "PROJ")])
        self.assertEqual(got, ["ADSK_Материал наименование"])

    def test_before_fix_13_fittings_empty_had_no_parameter(self):
        # до исправления «равно (пусто)» стояло по «Имени файла» (служебное поле): параметр был пустым, а не материалом
        got = self.names([("Связанный файл: Имя файла", "Equal", ""), ("Связанный файл: Имя файла", "Contains", "PROJ")])
        self.assertEqual(got, [])

    def test_fixed_13_pipes_no_value_checks_material(self):
        got = self.names([("ADSK_Материал наименование", "HasNoValue", None), ("Связанный файл: Имя файла", "Contains", "PROJ")])
        self.assertEqual(got, ["ADSK_Материал наименование"])
        # до исправления было «равно (пусто)» по тому же полю — параметр тот же, ключи истории не меняются
        before = self.names([("ADSK_Материал наименование", "Equal", ""), ("Связанный файл: Имя файла", "Contains", "PROJ")])
        self.assertEqual(before, got)

    def test_fixed_7_8_not_contains_ar_is_selection_not_parameter(self):
        for kind in ("Contains", "NotContains"):
            got = self.names([("ADSK_Размер_Длина", "HasNoValue", None), ("Связанный файл: Имя файла", kind, "_AR_")])
            self.assertEqual(got, ["ADSK_Размер_Длина"], kind)

    def test_debug_lines_mark_chosen_filter(self):
        self.names([("Код по классификатору", "HasValue", None), ("ADSK_Размер_Длина", "HasNoValue", None)])
        dbg = self.ns["FILTER_DEBUG"]
        self.assertEqual(len(dbg), 2)
        self.assertEqual(dbg[1].split("\t")[4], "ПРОВЕРЯЕМЫЙ")
        self.assertEqual(dbg[0].split("\t")[4], "")
        # диагностика: распознанная тема спеки и время выгрузки
        self.assertEqual(len(dbg[0].split("\t")), 9)
        self.assertRegex(dbg[0].split("\t")[6], r"^\d{4}-\d\d-\d\d \d\d:\d\d:\d\d$")

    def test_broken_view_gives_empty_list(self):
        self.ns["FILTER_DEBUG"][:] = []
        self.assertEqual(self.ns["filter_param_names"](object(), u"x"), [])

    def test_suffixless_checks_are_an_allowlist(self):
        """Четыре проверки без суффикса — замечания с параметром из фильтра. «Все» и «Уникальные» — нет."""
        issue = self.ns["is_issue_schedule"]
        param = self.ns["literal_issue_param"]
        self.assertTrue(issue(u"П        1_Классификатор_Пустые"))
        self.assertTrue(issue(u"П        5.7_Ширина_Лотки_Без значения"))
        want = {
            u"П        0_Ошибка типоразмера": u"ADSK_Наименование",
            u"РД       0_Ошибка типоразмера": u"ADSK_Наименование",
            u"ПП       0_Ошибка типоразмера": u"ADSK_Наименование",
            u"0_Ошибка типоразмера": u"ADSK_Наименование",
            u"П0_Ошибка типоразмера": u"ADSK_Наименование",
            u"п\u00a0\u00a00_Ошибка  типоразмера": u"ADSK_Наименование",
            u"П        1_Классификатор_999": u"Код по классификатору",
            u"1 _ Классификатор _ 999": u"Код по классификатору",
            u"П        16.0_Корпус_Кириллица": u"ADSK_Номер здания",
            u"РД   16.0 _ Корпус _ Кирилица": u"ADSK_Номер здания",
            u"П        17.0_Секция_Кириллица": u"ADSK_Номер секции",
            u"17.0_Секция_Кириллица": u"ADSK_Номер секции",
        }
        for name, expect in want.items():
            self.assertTrue(issue(name), name)
            self.assertEqual(param(name), expect, name)
        for name in (
            u"Пров     Ошибка типоразмера",
            u"П        1_Классификатор_Все",
            u"П        1_Классификатор_Уникальные_По разделам",
            u"П        16.1_Корпус_Оборудование_Все",
            u"П        16.1_Корпус_Кириллица",
            u"Обычная спецификация",
        ):
            self.assertFalse(issue(name), name)
            self.assertEqual(param(name), u"", name)
        filters = [
            ("ADSK_Наименование", "Equal", u"Ошибка типоразмера"),
            ("Код по классификатору", "NotEqual", "999"),
            ("Связанный файл: Имя файла", "HasValue", None),
            ("Связанный файл: Имя файла", "NotContains", "_AR_"),
        ]
        got, marks = self.ns["choose_check_params"](filters, u"П        0_Ошибка типоразмера")
        self.assertEqual(got, ["ADSK_Наименование"])
        self.assertEqual([filters[i][0] + ":" + filters[i][1] for i, m in enumerate(marks) if m], ["ADSK_Наименование:Equal"])
        got, marks = self.ns["choose_check_params"]([("Код по классификатору", "NotEqual", "999")], u"РД       0_Ошибка типоразмера")
        self.assertEqual(got, ["ADSK_Наименование"])
        self.assertEqual(marks, [False])
        self.assertFalse(issue(u"Пров     Ошибка типоразмера"))

    def test_svod_dump_suffixless_checks_use_real_filters(self):
        """Дамп фильтров сводной: без суффикса выгружаются ровно четыре проверки, параметр — из их фильтров."""
        path = os.path.join(HERE, "fixtures", "align_dumps", "фильтры_PROJ_IOS_SVOD_ORG_R24.txt")
        if not os.path.isfile(path):
            self.skipTest("portfolio copy omits the client schedule-filter export")
        by = {}
        order = []
        with open(path, encoding="utf-8-sig") as f:
            for line in f.read().splitlines():
                p = line.split("\t")
                if len(p) < 2:
                    continue
                if p[0] not in by:
                    order.append(p[0])
                    by[p[0]] = []
                by[p[0]].append(p[1:])
        extra = [n for n in order if self.ns["is_issue_schedule"](n) and ("Пустые" not in n) and ("Без знач" not in n)]
        self.assertEqual(sorted(self.ns["issue_name_key"](n) for n in extra), [
            u"0ошибкатипоразмера", u"16.0корпускириллица", u"17.0секциякириллица", u"1классификатор999"])
        for n in order:
            if "_Все" in n or "Уникальные" in n or n.strip().startswith("Пров"):
                self.assertFalse(self.ns["is_issue_schedule"](n), n)
        expect = {
            u"0ошибкатипоразмера": ("ADSK_Наименование", "Equal"),
            u"1классификатор999": ("Код по классификатору", "BeginsWith"),
            u"16.0корпускириллица": ("ADSK_Номер здания", "HasValue"),
            u"17.0секциякириллица": ("ADSK_Номер секции", "HasValue"),
        }
        for spec in extra:
            infos = []
            for row in by[spec]:
                field = row[0]
                ftype = row[1] if len(row) > 1 else ""
                val = row[2] if len(row) > 2 else ""
                if field.startswith("("):
                    continue
                infos.append((field, ftype, val if val != "" else None))
            got, marks = self.ns["choose_check_params"](
                [(n, t, v if v is not None else (u"" if t == "Equal" else None)) for n, t, v in infos], spec)
            field, ftype = expect[self.ns["issue_name_key"](spec)]
            self.assertEqual(got, [field], spec)
            self.assertEqual([(infos[i][0], infos[i][1]) for i, m in enumerate(marks) if m], [(field, ftype)], spec)


FIXTURE = os.path.join(HERE, "fixtures", "filters_before_fix.txt")


def read_fixture():
    specs = {}
    with open(FIXTURE, encoding="utf-8-sig") as f:
        for line in f.read().splitlines()[1:]:
            p = line.split("\t")
            if len(p) >= 4:
                specs.setdefault(p[0].strip(), []).append((p[1], p[2], p[3] if p[3] != "" else None))
    return specs


class RealFiltersTests(unittest.TestCase):
    """Needs the unpublished schedule-filter export. Skipped in this portfolio copy."""

    @classmethod
    def setUpClass(cls):
        if not os.path.isfile(FIXTURE):
            raise unittest.SkipTest("portfolio copy omits the client schedule-filter export")
        cls.ns = load_module()
        cls.specs = read_fixture()

    def params(self, specs=None):
        out = {}
        for name, infos in (specs or self.specs).items():
            out[name] = self.ns["choose_check_params"]([(n, t, v if v is not None else (u"" if t == "Equal" else None)) for n, t, v in infos], name)[0]
        return out

    def test_57_specs(self):
        self.assertEqual(len(self.specs), 57)

    def test_never_classifier_or_mark(self):
        for name, got in self.params().items():
            self.assertNotIn("ADSK_Марка изделия", got, name)
            if "1_Классификатор" not in name:
                self.assertNotIn("Код по классификатору", got, name)

    def test_expected_parameters(self):
        got = self.params()
        def one(prefix_part):
            hits = [v for n, v in got.items() if prefix_part in n]
            self.assertTrue(hits, prefix_part)
            return hits
        for v in one("1_Классификатор_Пустые"): self.assertEqual(v, ["Код по классификатору"])
        for part in ("2.1_Клас-я", "2.2_Клас-я", "2.4_Клас-я"):
            for v in one(part): self.assertEqual(v, ["Классификация систем"], part)
        for part in ("3.1_Тип системы", "3.2_Тип системы", "4.1_Имя системы", "4.2_Имя системы", "4.4_Имя системы"):
            for v in one(part): self.assertEqual(v, ["Имя системы"], part)
        self.assertEqual(one("5.1_Ширина_Детали возд-в"), [["ADSK_Размер_Ширина"]])
        self.assertEqual(one("6.1_Высота_Детали возд-в"), [["ADSK_Размер_Высота"]])
        self.assertEqual(one("7.1_Диаметр_Детали возд-в"), [["ADSK_Размер_Диаметр"]])
        self.assertEqual(one("16.1_Корпус"), [["ADSK_Номер здания"]] * 2)
        self.assertEqual(one("17.1_Секция"), [["ADSK_Номер секции"]] * 2)
        self.assertEqual(one("18.1_Этаж"), [["ADSK_Этаж"]] * 2)
        self.assertEqual(one("99_Толщина стенки"), [["Толщина стенки"]])

    def test_air_fittings_show_one_parameter_by_topic(self):
        """5.1/6.1/7.1 «Детали возд-в»: два HasNoValue (размер и диаметр/высота) — показывается один, по теме из имени."""
        want = {
            "5.1_Ширина_Детали возд-в": ["ADSK_Размер_Ширина"],
            "6.1_Высота_Детали возд-в": ["ADSK_Размер_Высота"],
            "7.1_Диаметр_Детали возд-в": ["ADSK_Размер_Диаметр"],
        }
        # имена из Revit бывают с неразрывными пробелами, другим «П» и т. п. — результат не должен зависеть от префикса
        prefixes = ["П        ", "П\u00a0\u00a0\u00a0\u00a0\u00a0\u00a0\u00a0\u00a0", "П\u00a0 \u00a0 ", "п        ", "Π        ", "П", ""]
        for topic, expected in want.items():
            full = [n for n in self.specs if topic in n]
            self.assertEqual(len(full), 1, topic)
            infos = self.specs[full[0]]
            self.assertEqual(len([i for i in infos if i[1] == "HasNoValue"]), 2, topic)      # в модели действительно два условия
            for prefix in prefixes:
                name = prefix + full[0].strip()[len("П"):].strip()
                got = self.ns["choose_check_params"]([(n, tt, v if v is not None else (u"" if tt == "Equal" else None)) for n, tt, v in infos], name)[0]
                self.assertEqual(got, expected, repr(name))
        # в сводке Параметр без ADSK_Марка изделия и без второго условия
        self.assertEqual(self.params()["П        5.1_Ширина_Детали возд-в_Без значения"], ["ADSK_Размер_Ширина"])

    def test_lower_text_handles_cyrillic_without_str_lower(self):
        """В IronPython 2.7 str.lower() не меняет кириллицу («Ширина» оставалась «Ширина»): свой нижний регистр."""
        low = self.ns["lower_text"]
        self.assertEqual(low(u"Ширина"), u"ширина")
        self.assertEqual(low(u"ADSK_Размер_Диаметр"), u"adsk_размер_диаметр")
        self.assertEqual(low(u"ЁЖ Клас-я Системы"), u"ёж клас-я системы")
        self.assertEqual(low(u"уже нижний 123_x"), u"уже нижний 123_x")
        self.assertEqual(low(None), u"")

    def test_no_native_lower_on_names_and_topics(self):
        """Имена полей и тем сравниваются только через lower_text: родной .lower() в этих функциях запрещён."""
        import ast
        tree = ast.parse(open(os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(HERE))), "служебное", "скрипты", "export_schedules.py"), encoding="utf-8").read())
        allowed = {"read_cfg", "is_check_filter", "choose_check_params", "main", "_lower_diag"}   # там .lower() только у латиницы/служебных слов
        bad = []
        for fn in [n for n in ast.walk(tree) if isinstance(n, ast.FunctionDef)]:
            for c in ast.walk(fn):
                if isinstance(c, ast.Call) and isinstance(c.func, ast.Attribute) and c.func.attr == "lower" and fn.name not in allowed:
                    bad.append("%s:%d" % (fn.name, c.lineno))
        self.assertEqual(bad, [])

    def test_debug_file_marks_topic_match_per_filter(self):
        self.ns["FILTER_DEBUG"][:] = []
        name = u"П        5.1_Ширина_Детали возд-в_Без значения"
        view = FakeView([("ADSK_Размер_Ширина", "HasNoValue", None), ("ADSK_Размер_Диаметр", "HasNoValue", None),
                         ("Код по классификатору", "NotEqual", "999")])
        names = self.ns["filter_param_names"](view, name)
        self.assertEqual(names, ["ADSK_Размер_Ширина"])
        rows = [l.split("\t") for l in self.ns["FILTER_DEBUG"]]
        self.assertEqual([(r[1], r[4], r[7]) for r in rows], [
            ("ADSK_Размер_Ширина", "ПРОВЕРЯЕМЫЙ", "да"), ("ADSK_Размер_Диаметр", "", "нет"), ("Код по классификатору", "", "")])
        self.assertIn("lower_text: ok", rows[0][8])

    def test_spec_parts_plain_string_parsing(self):
        parts = self.ns["spec_parts"]
        cases = {
            u"П        5.1_Ширина_Детали возд-в_Без значения": (u"5.1", u"ширина"),
            u"П        2.1_Клас-я системы_Воздуховоды_Без значния": (u"2.1", u"клас-я системы"),
            u"П        1_Классификатор_Пустые": (u"1", u"классификатор"),
            u"П        99_Толщина стенки труб_Без значения": (u"99", u"толщина стенки труб"),
            u"П\u00a0\u00a05.1_Ширина_Х": (u"5.1", u"ширина"),
            u"П13_Материал_Х": (u"13", u"материал"),
            u"5.1_Ширина": (u"5.1", u"ширина"),
            u"Произвольное имя": (u"", u""),
            u"П        Тема_Х": (u"", u""),
            u"П        5._Х_Y": (u"", u""),
        }
        for name, want in cases.items():
            self.assertEqual(parts(name), want, repr(name))

    def test_debug_file_shows_topic_for_real_specs(self):
        # то, что попадёт в фильтры_спецификаций.txt: тема распознана у всех 57 спецификаций
        for name in self.specs:
            self.assertTrue(self.ns["spec_topic"](name), name)

    def test_topic_ignores_prefix_and_whitespace_kind(self):
        topic = self.ns["spec_topic"]
        for name in ("П        5.1_Ширина_Детали возд-в_Без значения", u"П\u00a0\u00a05.1_Ширина_Х", u"Π 5.1_Ширина_Х", "5.1_Ширина_Х"):
            self.assertEqual(topic(name), u"ширина", repr(name))
        self.assertEqual(topic(u"П        4.4_Имя\u00a0системы_Оборудование_Пустые"), u"имя системы")
        self.assertEqual(topic(u"Произвольное имя"), u"")

    def fixed(self):
        """Исправленные фильтры: 7.8 «содержит _AR_» -> «не содержит»; 13_Фитинги_Пустые — по материалу; 13_Трубы_Без значения — HasNoValue."""
        specs = {}
        for name, infos in self.specs.items():
            infos = list(infos)
            if "7.8_Диаметр_Фитингов коробов" in name:
                infos = [(n, "NotContains" if t == "Contains" and v == "_AR_" else t, v) for n, t, v in infos]
            if "13_Материал_Фитинги труб_Пустые" in name:
                infos = [(n, t, v) for n, t, v in infos if not (t == "Equal" and n.startswith(u"Связанный файл"))] + [("ADSK_Материал наименование", "Equal", u"")]
            if "13_Материал_Трубы_Без значения" in name:
                infos = [(n, "HasNoValue" if (t == "Equal" and n == "ADSK_Материал наименование") else t, None if (t == "Equal" and n == "ADSK_Материал наименование") else v) for n, t, v in infos]
            specs[name] = infos
        return specs

    def test_after_fix_13_material_checks_material(self):
        got = self.params(self.fixed())
        for name, v in got.items():
            if "13_Материал" in name:
                self.assertEqual(v, ["ADSK_Материал наименование"], name)
        self.assertEqual(self.params(self.fixed())["П        7.8_Диаметр_Фитингов коробов_Без значения"], ["ADSK_Размер_Диаметр"])

    def test_before_fix_13_pipes_and_fittings_empty_had_no_parameter(self):
        got = self.params()
        self.assertEqual(got["П        13_Материал_Фитинги труб_Пустые"], [])     # «равно пусто» стояло по «Имени файла»
        self.assertEqual(got["П        13_Материал_Трубы_Без значения"], ["ADSK_Материал наименование"])


if __name__ == "__main__":
    unittest.main()
