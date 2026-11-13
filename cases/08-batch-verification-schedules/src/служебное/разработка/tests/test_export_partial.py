"""export_schedules.main() на подставном Revit API: режим «Всё» (как раньше) и «Выборочно» (пары, фильтр строк)."""
import io
import os
import sys
import tempfile
import types
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(HERE))), "служебное", "скрипты", "export_schedules.py")

HDR = [u"Семейство и типоразмер", u"Наименование", u"ADSK_Размер", u"Связанный файл"]
M1 = u"PROJ_IOS_EOM_K1_R22.rvt"
M2 = u"PROJ_IOS_OV_K1_R22.rvt"
M3 = u"PROJ_IOS_VK_K1_R22.rvt"


class Writer:
    def __init__(self, path, append, enc=None):
        self.f = io.open(path, "a" if append else "w", encoding="utf-8", newline="")

    def Write(self, t):
        self.f.write(t)

    def Close(self):
        self.f.close()


class Reader:
    def __init__(self, path, detect):
        self.path = path

    def ReadToEnd(self):
        with io.open(self.path, encoding="utf-8-sig") as f:
            return f.read()

    def Close(self):
        pass


class Section:
    def __init__(self, rows, cols=None):
        self.rows = rows
        self.NumberOfRows = len(rows)
        self.NumberOfColumns = cols if cols is not None else (len(rows[0]) if rows else 0)


class Field:
    def __init__(self, n):
        self.n = n
        self.IsHidden = False

    def GetName(self):
        return self.n


class Flt:
    def __init__(self, fid, ftype, value):
        self.FieldId, self.FilterType, self._value = fid, ftype, value

    def GetStringValue(self):
        if self._value is None:
            raise TypeError("no string value")
        return self._value


class Definition:
    def __init__(self, names, filters=None):
        self.fields = [Field(n) for n in names]
        self._filters = list(filters or [])

    def GetFieldCount(self):
        return len(self.fields)

    def GetField(self, i):
        return self.fields[i]

    def GetFilterCount(self):
        return len(self._filters)

    def GetFilter(self, i):
        return self._filters[i]


class View:
    IsTemplate = False

    def __init__(self, name, rows, headers=None, filters=None):
        self.Name = name
        self.rows = rows            # строки тела (без шапки)
        self.headers = headers or HDR
        self.Definition = Definition(self.headers, filters)

    def GetTableData(self):
        v = self

        class TD:
            def GetSectionData(self, sec):
                if sec == "Header":
                    return Section([v.headers])
                return Section(v.rows, len(v.headers)) if v.rows else Section([], 0)
        return TD()

    def GetCellText(self, sec, r, c):
        table = self.headers if sec == "Header" else self.rows
        return table[r][c]


class LinkType:
    def __init__(self, name, loaded):
        self.Name, self.loaded = name, loaded

    def GetLinkedFileStatus(self):
        return "Loaded" if self.loaded else "Unloaded"


def load(cfg_lines, views, links):
    text = open(SRC, encoding="utf-8").read().rstrip()
    assert text.endswith("main()")
    text = text[:-len("main()")]

    class Collector:
        def __init__(self, doc):
            pass

        def OfClass(self, cls):
            return list(views) if cls is ViewSchedule else list(links)

    class ViewSchedule: pass
    class RevitLinkType: pass
    db = types.ModuleType("Autodesk.Revit.DB")
    db.FilteredElementCollector = Collector
    db.ViewSchedule = ViewSchedule
    db.SectionType = types.SimpleNamespace(Body="Body", Header="Header")
    db.RevitLinkType = RevitLinkType
    db.LinkedFileStatus = types.SimpleNamespace(Loaded="Loaded")
    sysio = types.ModuleType("System.IO")
    sysio.StreamWriter, sysio.StreamReader = Writer, Reader
    systext = types.ModuleType("System.Text")
    systext.UTF8Encoding = lambda b: None
    rsu = types.SimpleNamespace(Output=lambda t: None, GetScriptDocument=lambda: object())
    mods = {"clr": types.SimpleNamespace(AddReference=lambda n: None), "Autodesk": types.ModuleType("Autodesk"),
            "Autodesk.Revit": types.ModuleType("Autodesk.Revit"), "Autodesk.Revit.DB": db,
            "System": types.ModuleType("System"), "System.IO": sysio, "System.Text": systext,
            "revit_script_util": rsu}
    saved = {k: sys.modules.get(k) for k in mods}
    sys.modules.update(mods)
    ns = {"__name__": "export_schedules_under_test", "__file__": SRC}
    exec(compile(text, SRC, "exec"), ns)
    # классы, по которым ищет коллектор, берутся в момент вызова из sys.modules (внутри collect_links)
    return ns, mods, saved


class ExportPartialTests(unittest.TestCase):
    def run_main(self, cfg, views=None, links=None):
        d = tempfile.mkdtemp()
        out = os.path.join(d, "out")
        lines = [u"report_dir=" + out] + cfg
        cfgp = os.path.join(d, "run.cfg")
        with io.open(cfgp, "w", encoding="utf-8") as f:
            f.write(u"\n".join(lines))
        os.environ["LOCALAPPDATA"] = d
        views = views if views is not None else self.views()
        links = links if links is not None else [LinkType(M1[:-4], True), LinkType(M2[:-4], True), LinkType(M3[:-4], False)]
        ns, mods, saved = load(lines, views, links)
        ns["LOCAL_CFG"] = cfgp
        # Collector должен вернуть ссылки для RevitLinkType
        db = mods["Autodesk.Revit.DB"]
        try:
            db.FilteredElementCollector = lambda doc: types.SimpleNamespace(
                OfClass=lambda cls: list(views) if cls is db.ViewSchedule else list(links))
            ns["FilteredElementCollector"] = db.FilteredElementCollector
            ns["ViewSchedule"] = db.ViewSchedule
            ns["main"]()
        finally:
            for k, v in saved.items():
                if v is None:
                    sys.modules.pop(k, None)
                else:
                    sys.modules[k] = v
        return out

    def views(self):
        return [
            View(u"1.1 Пустые ADSK_Размер", [[u"Ф1", u"Труба", u"", M1], [u"Ф2", u"Кран", u"", M2], [u"Ф3", u"Лоток", u"", M2]]),
            View(u"1.2 Без знач ADSK_Размер", [[u"Ф4", u"Муфта", u"", M3]]),
            View(u"2.1 Пустые ADSK_Масса", []),
            View(u"Обычная спецификация", [[u"x", u"y", u"z", M1]]),
        ]

    # прежние три файла теперь — разделы одного export.txt; текст раздела — как был текст прежнего файла
    SECTIONS = {"manifest.txt": "manifest", "export_result.txt": "result", "catalog.txt": "catalog"}

    def read(self, out, name):
        if name in self.SECTIONS:
            self.assertFalse(os.path.exists(os.path.join(out, name)), name + " больше не пишется")
            with io.open(os.path.join(out, "export.txt"), encoding="utf-8-sig", newline="") as f:
                text = f.read()
            self.assertTrue(text.endswith("\r\n"))
            parts, cur = {}, None
            for line in text.split("\r\n")[:-1]:
                if line in ("[result]", "[manifest]", "[catalog]"):
                    cur = line[1:-1]
                    parts[cur] = []
                else:
                    parts[cur].append(line)
            self.assertEqual(list(parts)[:2], ["result", "manifest"])
            lines = parts.get(self.SECTIONS[name], [])
            return "".join(l + "\n" for l in lines)
        with io.open(os.path.join(out, name), encoding="utf-8-sig") as f:
            return f.read().replace("\r\n", "\n")

    def test_all_mode_unchanged(self):
        out = self.run_main([])
        man = self.read(out, "manifest.txt").splitlines()
        self.assertEqual(man[0].split("\t")[:2], ["0001.tsv", u"1.1 Пустые ADSK_Размер"])
        self.assertEqual(man[1].split("\t")[:2], ["0002.tsv", u"1.2 Без знач ADSK_Размер"])
        self.assertEqual(man[2], u"*empty*\t2.1 Пустые ADSK_Масса")
        self.assertEqual(len(man), 3)          # никаких *scope*/*pair*
        self.assertEqual(self.read(out, "export_result.txt"), "matched=3\nexported=2\nskipped=1\n")
        tsv = self.read(out, "0001.tsv").splitlines()
        self.assertEqual(len(tsv), 4)          # шапка + 3 строки всех моделей
        cat = self.read(out, "catalog.txt").splitlines()
        self.assertIn(u"spec\t2.1 Пустые ADSK_Масса", cat)
        self.assertIn(u"model\t" + M3[:-4] + u"\tunloaded", cat)
        self.assertIn(u"model\t" + M1[:-4] + u"\tloaded", cat)

    def test_partial_filters_specs_and_models(self):
        out = self.run_main([u"assign_scope=partial", u"assign_specs=1.1 Пустые ADSK_Размер|2.1 Пустые ADSK_Масса",
                             u"assign_models=" + M2[:-4] + u"|" + M3[:-4]])
        man = self.read(out, "manifest.txt").splitlines()
        self.assertEqual(man[0].split("\t")[:2], ["0001.tsv", u"1.1 Пустые ADSK_Размер"])
        self.assertEqual(man[1], u"*empty*\t2.1 Пустые ADSK_Масса")
        self.assertIn(u"*scope*\tpartial", man)
        pairs = [l for l in man if l.startswith(u"*pair*")]
        # M3 — связь не загружена и строк нет: пары нет; M2 — две выбранные спецификации (одна пустая — тоже проверена)
        self.assertEqual(sorted(pairs), sorted([u"*pair*\t" + M2[:-4] + u"\t1.1 Пустые ADSK_Размер",
                                                u"*pair*\t" + M2[:-4] + u"\t2.1 Пустые ADSK_Масса"]))
        tsv = self.read(out, "0001.tsv").splitlines()
        self.assertEqual(len(tsv), 3)          # шапка + 2 строки модели M2, без M1
        self.assertTrue(all(M1 not in l for l in tsv))
        res = self.read(out, "export_result.txt")
        self.assertIn("matched=2\n", res)
        self.assertIn("partial=1\n", res)
        self.assertIn("pairs=2\n", res)
        self.assertIn(u"unloaded_models=" + M3[:-4], res)

    def test_partial_model_without_rows_is_still_checked(self):
        out = self.run_main([u"assign_scope=partial", u"assign_specs=1.2 Без знач ADSK_Размер",
                             u"assign_models=" + M1[:-4]])
        man = self.read(out, "manifest.txt").splitlines()
        self.assertEqual(man[0], u"*empty*\t1.2 Без знач ADSK_Размер")     # у M1 строк нет — спецификация пуста для выбора
        self.assertIn(u"*pair*\t" + M1[:-4] + u"\t1.2 Без знач ADSK_Размер", man)

    def test_pair_model_name_is_report_name_without_rvt(self):
        out = self.run_main([u"assign_scope=partial", u"assign_specs=1.1 Пустые ADSK_Размер", u"assign_models=" + M2])
        man = self.read(out, "manifest.txt").splitlines()
        self.assertIn(u"*pair*\t" + M2[:-4] + u"\t1.1 Пустые ADSK_Размер", man)

    def test_partial_missing_spec_reported(self):
        out = self.run_main([u"assign_scope=partial", u"assign_specs=9.9 Пустые Нет", u"assign_models=" + M1[:-4]])
        man = self.read(out, "manifest.txt").splitlines()
        self.assertEqual([l for l in man if l.startswith(u"*pair*")], [])
        self.assertIn(u"missing_specs=9.9 Пустые Нет", self.read(out, "export_result.txt"))

    def pairs_file(self, pairs):
        d = tempfile.mkdtemp()
        path = os.path.join(d, u"assign_pairs.txt")
        with io.open(path, "w", encoding="utf-8-sig", newline="") as f:
            f.write(u"".join(m + u"\t" + sp + u"\r\n" for m, sp in pairs))
        return path

    def test_exact_pairs_file_is_not_a_cartesian_product(self):
        """Окно «Выборочно»: у каждой модели свои спецификации. Файл пар: (M2, 1.1) и (M1, 2.1) — и больше ничего."""
        path = self.pairs_file([(M2[:-4], u"1.1 Пустые ADSK_Размер"), (M1[:-4], u"2.1 Пустые ADSK_Масса")])
        out = self.run_main([u"assign_scope=partial", u"assign_specs=1.1 Пустые ADSK_Размер|2.1 Пустые ADSK_Масса",
                             u"assign_models=" + M2[:-4] + u"|" + M1[:-4], u"assign_pairs_file=" + path])
        man = self.read(out, "manifest.txt").splitlines()
        pairs = sorted(l for l in man if l.startswith(u"*pair*"))
        self.assertEqual(pairs, sorted([u"*pair*\t" + M2[:-4] + u"\t1.1 Пустые ADSK_Размер", u"*pair*\t" + M1[:-4] + u"\t2.1 Пустые ADSK_Масса"]))
        tsv = self.read(out, "0001.tsv").splitlines()
        self.assertEqual(len(tsv), 3)                                   # шапка + строки только M2 (M1 для этой спеки не выбрана)
        self.assertTrue(all(M1 not in l for l in tsv))
        self.assertIn("pairs=2\n", self.read(out, "export_result.txt"))

    def test_exact_pairs_file_overrides_lists_and_normalizes_names(self):
        """Имена из файла пар сравниваются как в режиме списков: без .rvt и пути, опечатка «значния» не мешает."""
        path = self.pairs_file([(u"\\\\srv\\" + M1, u"1.2 Без значния ADSK_Размер")])
        out = self.run_main([u"assign_scope=partial", u"assign_specs=ДРУГАЯ", u"assign_models=Другая", u"assign_pairs_file=" + path],
                            views=[View(u"1.2 Без значения ADSK_Размер", [[u"Ф4", u"Муфта", u"", M1], [u"Ф5", u"Муфта2", u"", M2]])])
        man = self.read(out, "manifest.txt").splitlines()
        self.assertTrue(man[0].startswith("0001.tsv\t1.2 Без значения ADSK_Размер"))
        self.assertEqual(len(self.read(out, "0001.tsv").splitlines()), 2)       # шапка + одна строка M1

    def test_missing_pairs_file_is_an_error(self):
        with self.assertRaises(Exception):
            self.run_main([u"assign_scope=partial", u"assign_pairs_file=" + os.path.join(tempfile.mkdtemp(), u"нет.txt")])

    def type_size_views(self, rows):
        spec = u"П        0_Ошибка типоразмера"
        headers = [u"Семейство и типоразмер", u"Наименование", u"ADSK_Наименование", u"Связанный файл"]
        code = [u"Семейство и типоразмер", u"Наименование", u"Код по классификатору", u"Связанный файл"]
        corp = [u"Семейство и типоразмер", u"Наименование", u"ADSK_Номер здания", u"Связанный файл"]
        sect = [u"Семейство и типоразмер", u"Наименование", u"ADSK_Номер секции", u"Связанный файл"]
        return spec, [
            View(spec, rows, headers, [Flt(2, u"Equal", u"Ошибка типоразмера")]),
            View(u"П        1_Классификатор_999", [[u"Щит", u"Щит 1", u"999.1", M1]], code, [Flt(2, u"BeginsWith", u"999")]),
            View(u"П        16.0_Корпус_Кириллица", [[u"Панель", u"П 1", u"Корпус-А", M2]], corp, [Flt(2, u"HasValue", None)]),
            View(u"П        17.0_Секция_Кириллица", [[u"Панель", u"П 2", u"Секция-1", M1]], sect, [Flt(2, u"HasValue", None)]),
            View(u"Пров     Ошибка типоразмера", [[u"a", u"Ошибка типоразмера", u"x", M1]], headers, [Flt(2, u"Equal", u"Ошибка типоразмера")]),
            View(u"П        1_Классификатор_Все", [[u"a", u"b", u"c", M1]]),
            View(u"П        1_Классификатор_Уникальные_По разделам", [[u"a", u"b", u"c", M1]]),
            View(u"П        1_Классификатор_Пустые", [[u"Ф", u"Имя", u"", M1]]),
        ]

    def test_suffixless_checks_exported_overviews_not(self):
        spec, views = self.type_size_views([
            [u"Отвод: 90", u"Отвод", u"Ошибка типоразмера", M1],
            [u"Труба: 100", u"Труба", u"Ошибка типоразмера", M2],
        ])
        out = self.run_main([], views=views)
        man = self.read(out, "manifest.txt").splitlines()
        by_name = {l.split("\t")[1]: l for l in man}
        self.assertEqual(by_name[spec].split("\t")[3], u"ADSK_Наименование")
        self.assertEqual(by_name[u"П        1_Классификатор_999"].split("\t")[3], u"Код по классификатору")
        self.assertEqual(by_name[u"П        16.0_Корпус_Кириллица"].split("\t")[3], u"ADSK_Номер здания")
        self.assertEqual(by_name[u"П        17.0_Секция_Кириллица"].split("\t")[3], u"ADSK_Номер секции")
        self.assertIn(u"П        1_Классификатор_Пустые", by_name)
        self.assertFalse(any("Классификатор_Все" in l or "Уникальные" in l or "Пров" in l for l in man))
        body = self.read(out, by_name[spec].split("\t")[0])
        self.assertIn(u"Ошибка типоразмера", body)
        self.assertIn(u"999.1", self.read(out, by_name[u"П        1_Классификатор_999"].split("\t")[0]))
        cat = self.read(out, "catalog.txt")
        for name in (spec, u"П        1_Классификатор_999", u"П        16.0_Корпус_Кириллица", u"П        17.0_Секция_Кириллица", u"П        1_Классификатор_Пустые"):
            self.assertIn(u"spec\t" + name, cat)
        self.assertNotIn(u"Пров", cat)
        self.assertNotIn(u"Классификатор_Все", cat)
        self.assertNotIn(u"Уникальные", cat)

    def test_suffixless_empty_manifest(self):
        for spec in (u"РД       0_Ошибка типоразмера", u"П        1_Классификатор_999", u"П        16.0_Корпус_Кирилица", u"17.0_Секция_Кириллица"):
            out = self.run_main([], views=[View(spec, [])])
            man = self.read(out, "manifest.txt").splitlines()
            self.assertEqual(man, [u"*empty*\t" + spec], spec)
            self.assertIn(u"spec\t" + spec, self.read(out, "catalog.txt"))
            res = self.read(out, "export_result.txt")
            self.assertIn("matched=1\n", res)
            self.assertIn("exported=0\n", res)
            self.assertIn("skipped=1\n", res)

    def test_type_size_error_partial_pair(self):
        spec, views = self.type_size_views([
            [u"Отвод: 90", u"Отвод", u"Ошибка типоразмера", M1],
            [u"Труба: 100", u"Труба", u"Ошибка типоразмера", M2],
        ])
        path = self.pairs_file([(M2[:-4], spec)])
        out = self.run_main([u"assign_scope=partial", u"assign_pairs_file=" + path], views=views)
        man = self.read(out, "manifest.txt").splitlines()
        data = [l for l in man if l.startswith("0001.tsv")]
        self.assertEqual(len(data), 1)
        self.assertTrue(data[0].startswith("0001.tsv\t" + spec + "\t"))
        self.assertTrue(data[0].endswith(u"\tADSK_Наименование"))
        self.assertEqual([l for l in man if l.startswith(u"*pair*")], [u"*pair*\t" + M2[:-4] + u"\t" + spec])
        tsv = self.read(out, "0001.tsv").splitlines()
        self.assertEqual(len(tsv), 2)
        self.assertIn(M2, tsv[1])
        self.assertNotIn(M1, tsv[1])
        self.assertIn("pairs=1\n", self.read(out, "export_result.txt"))


if __name__ == "__main__":
    unittest.main()
