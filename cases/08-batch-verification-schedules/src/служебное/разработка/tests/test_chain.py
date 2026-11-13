"""Тесты цепочки «xlsx -> payload -> Google» без Windows/Revit/Excel/Google.

    python -m unittest discover -s tests -v        (из папки batch_verification_schedules)

Что нужно: python3 + openpyxl; node (мок Code.gs); PowerShell Core `pwsh` (сборка xlsx настоящим кодом
build_assignment.ps1 + DocumentFormat.OpenXml.dll и сквозной тест push_google.ps1 против локального мок-сервера).
Нет pwsh/node — соответствующие тесты пропускаются.
"""
import http.server
import json
import os
import shutil
import subprocess
import sys
import tempfile
import threading
import unittest
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(os.path.dirname(HERE)))
TOOL = os.path.join(ROOT, "служебное", "скрипты")      # папка с .ps1/.py окна
sys.path.insert(0, HERE)
import check_xlsx  # noqa: E402

PWSH = os.environ.get("PWSH") or shutil.which("pwsh")
NODE = shutil.which("node")
ENV = dict(os.environ, DOTNET_SYSTEM_GLOBALIZATION_INVARIANT="1")


def _decode_out(b):
    """Вывод pwsh: UTF-8 или кодировка консоли Windows (cp866) — иначе текст ошибки теряется."""
    if b is None:
        return ""
    for enc in ("utf-8", "cp866"):
        try:
            return b.decode(enc)
        except UnicodeDecodeError:
            continue
    return b.decode("utf-8", "replace")


def pwsh(*args, timeout=300):
    r = subprocess.run([PWSH, "-NoProfile", *args], capture_output=True, timeout=timeout, env=ENV, cwd=TOOL)
    r.stdout = _decode_out(r.stdout)
    r.stderr = _decode_out(r.stderr)
    return r


def make_export(folder, empty=False, history=True):
    exp = os.path.join(folder, "exp")
    out = os.path.join(folder, "out")
    os.makedirs(out, exist_ok=True)
    subprocess.check_call([sys.executable, os.path.join(HERE, "make_sample_export.py"), exp] + (["--empty"] if empty else []))
    hist = os.path.join(folder, "задание_история.txt")
    if history:
        shutil.move(hist, os.path.join(out, "задание_история.txt"))
    else:
        os.remove(hist)
    return exp, out


@unittest.skipUnless(PWSH, "нет pwsh")
class BuildTests(unittest.TestCase):
    def build(self, **kw):
        tmp = tempfile.mkdtemp(prefix="bvs_test_")
        self.addCleanup(shutil.rmtree, tmp, True)
        exp, out = make_export(tmp, **kw)
        xlsx = os.path.join(out, "task.xlsx")
        r = pwsh("-File", os.path.join(TOOL, "build_assignment.ps1"), "-ExportDir", exp, "-OutXlsx", xlsx)
        self.assertEqual(r.returncode, 0, r.stdout + r.stderr)
        self.assertIn("BUILDER=OPENXML", r.stdout)
        return xlsx, os.path.join(out, "google_push_payload.json")

    def test_full_report_valid_and_consistent_with_payload(self):
        xlsx, payload = self.build()
        problems = []
        check_xlsx.check_zip_xml(xlsx, problems)
        wb = check_xlsx.check_workbook(xlsx, problems)
        check_xlsx.check_payload(wb, payload, problems)
        self.assertEqual(problems, [])
        ws = wb["ЭОМ"]
        # кириллица, ведущие нули и «формула» дошли как текст
        texts = {str(c.value) for row in ws.iter_rows() for c in row if c.value is not None}
        self.assertIn("ЭОМ Проверка заполнения", texts)
        self.assertIn("007", texts)
        self.assertIn("=1+1", texts)
        # колонка «Параметр» — имя параметра из manifest, а не заголовок колонки Revit
        self.assertIn("ADSK_Размер_Длина", texts)
        self.assertFalse(any("Проверить, чтобы" in t for t in texts))
        # «Динамика»: параметр показывается именем поля, а ключ истории (заголовок колонки) не менялся — отработанное найдено
        dyn = [str(c.value) for row in wb["Динамика"].iter_rows(min_row=3) for c in row if c.value is not None]
        self.assertIn("ADSK_Размер_Длина", dyn)
        # «Динамика» теперь считается с текущей сборкой, у которой имя параметра есть
        self.assertFalse(any("Проверить, чтобы" in t for t in dyn))
        hist = open(os.path.join(os.path.dirname(xlsx), "задание_история.txt"), encoding="utf-8-sig").read().splitlines()
        today = [l.split("\t") for l in hist if l.split("\t")[2].startswith("ORG_K1_PROJ_IOS_EOM")]
        self.assertTrue(any(len(p) == 8 and p[6] == "ADSK_Размер_Длина" and "Проверить" in p[4] and p[7] > p[0] + " 00:00:00" for p in today), hist)
        # вид по умолчанию «2»: спецификации (уровень 1) видны, элементы (уровень 2) свёрнуты
        levels = {(rd.outlineLevel or 0, bool(rd.hidden)) for rd in ws.row_dimensions.values()}
        self.assertIn((1, False), levels)
        self.assertIn((2, True), levels)
        self.assertNotIn((1, True), levels)
        # колонки «Замечаний» нет; строки модели и спецификации — заголовки групп (название жирным)
        self.assertEqual([c.value for c in ws[2]][:6], ["Модель", "Спецификация", "Параметр", "Семейство и типоразмер", "Наименование", "Значение"])
        self.assertFalse(any(c.value == "Замечаний" for c in ws[2]))
        for r in range(3, ws.max_row + 1):
            if ws.cell(r, 1).value:
                self.assertTrue(ws.cell(r, 1).font.b, "A%d не жирная" % r)
            elif ws.cell(r, 2).value:
                self.assertTrue(ws.cell(r, 2).font.b, "B%d не жирная" % r)
        # ширина по содержимому: колонки разные, а не фиксированные 48/12/46...
        wd = {k: round(v.width) for k, v in ws.column_dimensions.items()}
        self.assertGreater(wd["A"], 25)
        self.assertNotEqual(wd["A"], wd["C"])
        sv_w = {k: round(v.width) for k, v in wb["Сводка"].column_dimensions.items()}
        self.assertGreater(sv_w["A"], 25)
        # колонки: Параметр = параметр из фильтра, Семейство/Наименование/Значение — каждое в своей колонке
        item_rows = [r for r in range(3, ws.max_row + 1) if ws.cell(r, 4).value]
        self.assertTrue(item_rows)
        fam = {str(ws.cell(r, 4).value) for r in item_rows}
        self.assertEqual(fam, {"Светильник: тип А", "Светильник: тип Б", "Розетка"})
        self.assertTrue(all(ws.cell(r, 5).value for r in item_rows))
        self.assertIn("007", {str(ws.cell(r, 6).value) for r in item_rows})
        spec_rows = [r for r in range(3, ws.max_row + 1) if ws.cell(r, 2).value]
        self.assertTrue(all(ws.cell(r, 3).value == "ADSK_Размер_Длина" for r in spec_rows))
        self.assertFalse(any("Семейство и типоразмер" == ws.cell(r, 3).value for r in spec_rows))
        # границы внутри таблицы (Сводка и разделы), заголовки спецификаций в Сводке — вертикально
        sv0 = wb["Сводка"]
        for sheet in (sv0, ws):
            for coord in ("A2", "B3", "A3"):
                self.assertEqual(sheet[coord].border.left.style, "thin", sheet.title + coord)
        import re
        hdr_checks = [c for c in sv0[2] if c.column > 1 and c.value]
        # колонок «Замечания / дата» (итог по модели) нет: сразу за моделью идут спецификации
        self.assertFalse(any(str(c.value).startswith("Замечания") for c in sv0[2]))
        self.assertIn(sv0["B2"].value, ("ЭОМ Проверка заполнения", "СС Проверка марки"))
        self.assertIn("строк спецификации с незаполненным параметром", str(sv0["A1"].value))
        self.assertNotIn("элементов", str(sv0["A1"].value))
        self.assertTrue(sv0['B2'].alignment.wrap_text)
        self.assertTrue(hdr_checks and all(c.alignment.textRotation == 90 for c in hdr_checks))
        self.assertGreater(sv0.row_dimensions[2].height, 60)
        self.assertLess(sv0.row_dimensions[2].height, 160)  # три строки переноса, а не километровая шапка
        self.assertTrue(all(c.alignment.wrap_text for c in hdr_checks))
        # желтая шапка (Excel 0x00FFFF -> FFFF00), а не голубая
        self.assertEqual(ws["A1"].fill.fgColor.rgb[-6:], "FFFF00")
        # оранжевая ячейка «есть замечания» (0x80C0FF -> FFC080)
        sv = wb["Сводка"]
        fills = {c.fill.fgColor.rgb[-6:] for row in sv.iter_rows(min_row=3) for c in row if c.fill and c.fill.fgColor.rgb}
        self.assertIn("FFC080", fills)

    def test_empty_report_single_row_summary(self):
        # Раньше: merge A1:A1 (Excel просит «восстановить»), в payload rows=["текст"] и A1='С' (первая буква)
        xlsx, payload = self.build(empty=True, history=False)
        problems = []
        check_xlsx.check_zip_xml(xlsx, problems)
        wb = check_xlsx.check_workbook(xlsx, problems)
        check_xlsx.check_payload(wb, payload, problems)
        self.assertEqual(problems, [])
        self.assertTrue(str(wb["Сводка"]["A1"].value).startswith("Строк не выгружено"))

    def test_ps_unit(self):
        r = pwsh("-File", os.path.join(HERE, "ps_unit.ps1"))
        self.assertEqual(r.returncode, 0, r.stdout + r.stderr)

    def test_excel_seed_file_is_valid_xlsx(self):
        r = pwsh("-Command", ". ./excel_isolated.ps1; New-BvsExcelSeedFile")
        self.assertEqual(r.returncode, 0, r.stdout + r.stderr)
        path = r.stdout.strip().splitlines()[-1]
        try:
            problems = []
            check_xlsx.check_zip_xml(path, problems)
            self.assertEqual(problems, [])
            self.assertTrue(zipfile.is_zipfile(path))
        finally:
            os.remove(path)


class _HistoryBase(unittest.TestCase):
    """Общее для тестов истории/журнала: последовательные сборки в одну папку."""

    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="bvs_hist_")
        self.addCleanup(shutil.rmtree, self.tmp, True)
        self.out = os.path.join(self.tmp, "out")
        os.makedirs(self.out)
        self.xlsx = os.path.join(self.out, "task.xlsx")
        self.hist = os.path.join(self.out, "задание_история.txt")
        self.n = 0

    def build(self, *flags, collapse=False):
        import time
        self.n += 1
        exp = os.path.join(self.tmp, "exp%d" % self.n)
        subprocess.check_call([sys.executable, os.path.join(HERE, "make_sample_export.py"), exp] + list(flags))
        seed = os.path.join(self.tmp, "задание_история.txt")
        if os.path.exists(seed):
            os.remove(seed)  # без посеянной истории
        time.sleep(1.1)  # идентификатор сборки — с точностью до секунды
        args = ["-File", os.path.join(TOOL, "build_assignment.ps1"), "-ExportDir", exp, "-OutXlsx", self.xlsx]
        if collapse:
            args.append("-CollapseDay")
        r = pwsh(*args)
        self.assertEqual(r.returncode, 0, r.stdout + r.stderr)
        return r.stdout

    def history_rows(self):
        with open(self.hist, encoding="utf-8-sig") as f:
            return [l.rstrip("\n").split("\t") for l in f if l.strip()]

    def summary(self):
        import openpyxl
        import warnings
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            wb = openpyxl.load_workbook(self.xlsx)
        return wb

    @staticmethod
    def cell_for(ws, model, check):
        hdr = [c.value for c in ws[2]]
        col = hdr.index(check) + 1
        for r in range(3, ws.max_row + 1):
            if ws.cell(r, 1).value == model:
                return ws.cell(r, col)
        return None

@unittest.skipUnless(PWSH, "нет pwsh")
class HistorySliceTests(_HistoryBase):
    """История по срезам: каждая сборка — срез; исправленное записывается нулём; «Динамика» видит сегодняшнюю сборку."""

    def test_fix_between_builds_same_day_is_visible(self):
        self.build()                      # сборка 1: у s1 есть замечания
        rows1 = self.history_rows()
        self.assertTrue(all(len(r) == 8 for r in rows1))
        self.build("--fixed")             # сборка 2 в тот же день: s1 пуста
        rows2 = self.history_rows()
        builds = sorted({r[7] for r in rows2})
        self.assertEqual(len(builds), 2, builds)
        s1_check = "ЭОМ Проверка заполнения"
        zeros = [r for r in rows2 if r[3] == s1_check and r[5] == "0" and r[7] == builds[1]]
        self.assertTrue(zeros, "нет нулевых строк по исправленной спецификации")
        wb = self.summary()
        cell = self.cell_for(wb["Сводка"], "ORG_K1_PROJ_IOS_EOM_R23", s1_check)
        self.assertEqual(cell.value, "0 (2) \u21932")  # «сейчас 0, было 2, исправлено 2»
        self.assertEqual(cell.fill.fgColor.rgb[-6:], "B4D2F0")  # синий ноль
        # срезы в Сводке больше не показываются колонками (история остаётся в файле и в «Динамике»)
        hdr = [str(c.value) for c in wb["Сводка"][2]]
        self.assertEqual(len([h for h in hdr if h.startswith("Замечания")]), 0, hdr)
        # «Динамика» видит исправление сегодняшней сборки
        dyn = [str(c.value) for row in wb["Динамика"].iter_rows(min_row=3) for c in row if c.value is not None]
        self.assertIn(s1_check, dyn)
        self.assertFalse(any("не найдено" in t for t in dyn), dyn)

    def test_spec_not_exported_is_not_marked_fixed(self):
        self.build()
        self.build("--drop-s1")           # s1 убрана из списка: «не проверялось», а не «исправлено»
        rows = self.history_rows()
        last = max(r[7] for r in rows)
        self.assertFalse([r for r in rows if r[7] == last and r[3] == "ЭОМ Проверка заполнения"], "нулей быть не должно")
        wb = self.summary()
        hdr = [str(c.value) for c in wb["Сводка"][2]]
        self.assertNotIn("ЭОМ Проверка заполнения", hdr)

    def test_collapse_day_keeps_only_last_build_of_day(self):
        self.build()
        self.build("--fixed")
        self.build("--fixed", collapse=True)
        builds = sorted({r[7] for r in self.history_rows()})
        self.assertEqual(len(builds), 1, builds)

    def test_old_history_format_still_readable(self):
        # старый формат: 6 колонок, по одной записи в день
        d1 = "2026-01-01"
        with open(self.hist, "w", encoding="utf-8-sig") as f:
            f.write("%s\tEOM\tORG_K1_PROJ_IOS_EOM_R23\tЭОМ Проверка заполнения\tkey\t5\n" % d1)
        self.build()
        rows = self.history_rows()
        old = [r for r in rows if r[0] == d1]
        self.assertEqual(len(old), 1)
        self.assertEqual(old[0][7], d1 + " 00:00:00")


@unittest.skipUnless(PWSH, "нет pwsh")
class LedgerTests(_HistoryBase):
    """Журнал замечаний: статусы, исправленное серым на листах разделов, «Динамика» по каждому замечанию."""

    ledger_path = property(lambda self: os.path.join(self.out, "задание_журнал.txt"))

    def ledger_rows(self):
        with open(self.ledger_path, encoding="utf-8-sig") as f:
            return [l.rstrip("\n").split("\t") for l in f if l.strip()]

    def test_fixed_items_stay_on_section_sheet_grey_with_dates(self):
        self.build()
        led = self.ledger_rows()
        self.assertEqual(len(led), 4)
        self.assertTrue(all(r[9] == "open" for r in led))
        self.build("--fixed")
        led = self.ledger_rows()
        fixed = [r for r in led if r[9] == "fixed"]
        self.assertEqual(len(fixed), 3, led)
        self.assertTrue(all(r[10] for r in fixed))               # дата исправления проставлена
        self.assertEqual([r for r in led if r[9] == "open"][0][2], "СС Проверка марки")
        wb = self.summary()
        ws = wb["ЭОМ"]                                           # раздел остался, хотя открытых замечаний нет
        self.assertEqual([ws.cell(2, c).value for c in (7, 8, 9, 10)], ["Статус", "Найдено", "Исправлено", "Дней"])
        fixed_rows = [r for r in range(3, ws.max_row + 1) if ws.cell(r, 7).value == "Исправлено"]
        self.assertEqual(len(fixed_rows), 3)
        for r in fixed_rows:
            self.assertEqual(ws.cell(r, 4).font.color.rgb[-6:], "808080")   # серым
            self.assertTrue(ws.cell(r, 8).value and ws.cell(r, 9).value)    # найдено / исправлено
        k1 = [r for r in range(3, ws.max_row + 1) if ws.cell(r, 1).value == "ORG_K1_PROJ_IOS_EOM_R23"][0]
        self.assertFalse(ws.cell(k1, 2).value)                   # счётчика замечаний на строке модели нет
        self.assertEqual(ws.cell(k1, 9).value, 2)                # исправлено 2
        # СС: открытое осталось с обычным статусом
        ss = wb["СС"]
        statuses = {ss.cell(r, 7).value for r in range(3, ss.max_row + 1)} - {None, ""}
        self.assertEqual(statuses, {"Открыто"})
        # Google получает только открытое: листа ЭОМ в payload нет, СС — 6 колонок
        import json
        payload = json.load(open(os.path.join(self.out, "google_push_payload.json"), encoding="utf-8"))
        names = [x["name"] for x in payload["sheets"]]
        self.assertNotIn("ЭОМ", names)
        self.assertIn("СС", names)
        self.assertEqual({len(r) for r in [x for x in payload["sheets"] if x["name"] == "СС"][0]["rows"]}, {6})
        problems = []
        check_xlsx.check_zip_xml(self.xlsx, problems)
        wb2 = check_xlsx.check_workbook(self.xlsx, problems)
        check_xlsx.check_payload(wb2, os.path.join(self.out, "google_push_payload.json"), problems)
        self.assertEqual(problems, [])

    def test_dynamics_lists_each_fixed_issue(self):
        self.build()
        self.build("--fixed")
        ws = self.summary()["Динамика"]
        texts = [str(c.value) for row in ws.iter_rows(min_row=3) for c in row if c.value is not None]
        self.assertIn("Светильник: тип А", texts)
        self.assertIn("Розетка", texts)
        self.assertFalse(any("без детализации" in t for t in texts), texts)
        hdr = [ws.cell(2, c).value for c in (8, 9, 10)]
        self.assertEqual(hdr, ["Найдено", "Исправлено", "Дней"])
        items = [r for r in range(3, ws.max_row + 1) if ws.cell(r, 5).value in ("Светильник: тип А", "Светильник: тип Б", "Розетка")]   # «Динамика»: раскладка прежняя
        self.assertEqual(len(items), 3)
        self.assertTrue(all(ws.cell(r, 9).value for r in items))     # дата исправления

    def test_reopened_issue_is_marked_new(self):
        self.build()
        self.build("--fixed")
        self.build()                                                 # ошибки вернулись
        led = self.ledger_rows()
        reopened = [r for r in led if r[9] == "open" and r[11] == "1"]
        self.assertEqual(len(reopened), 3, led)
        ws = self.summary()["ЭОМ"]
        marked = [r for r in range(3, ws.max_row + 1) if ws.cell(r, 7).value == "Новое"]
        self.assertEqual(len(marked), 3)
        cell = self.cell_for(self.summary()["Сводка"], "ORG_K1_PROJ_IOS_EOM_R23", "ЭОМ Проверка заполнения")
        self.assertEqual(cell.value, "2 (0) \u21912")

    def test_fixed_older_than_retention_disappear_from_sections_but_stay_in_ledger(self):
        import datetime
        old = (datetime.datetime.now() - datetime.timedelta(days=40)).strftime("%Y-%m-%d 10:00:00")
        older = (datetime.datetime.now() - datetime.timedelta(days=60)).strftime("%Y-%m-%d 10:00:00")
        row = ["ORG_K1_PROJ_IOS_EOM_R23", "EOM", "ЭОМ Проверка заполнения", "Старое семейство", "Старое наименование", "P", "",
               older, older, "fixed", old, "0"]
        with open(self.ledger_path, "w", encoding="utf-8-sig") as f:
            f.write("\t".join(row) + "\n")
        self.build()
        ws = self.summary()["ЭОМ"]
        texts = {str(c.value) for r in ws.iter_rows() for c in r if c.value is not None}
        self.assertNotIn("Старое семейство", texts)
        self.assertTrue(any(r[3] == "Старое семейство" for r in self.ledger_rows()))

    def test_history_remainder_without_ledger_detail(self):
        # исправлено по старой истории (счётчики), журнала ещё не было: строка «без детализации»
        import datetime
        d1 = (datetime.datetime.now() - datetime.timedelta(days=1)).strftime("%Y-%m-%d")
        key = 'Проверить, чтобы была заполнена "Длина", ADSK_Размер_Длина'
        with open(self.hist, "w", encoding="utf-8-sig") as f:
            f.write("%s\tEOM\tORG_K1_PROJ_IOS_EOM_R23\tЭОМ Проверка заполнения\t%s\t5\n" % (d1, key))
        self.build("--fixed")
        ws = self.summary()["Динамика"]
        texts = [str(c.value) for row in ws.iter_rows(min_row=3) for c in row if c.value is not None]
        self.assertTrue(any("без детализации" in t for t in texts), texts)

    def test_spec_not_exported_keeps_issues_open(self):
        self.build()
        self.build("--drop-s1")
        led = self.ledger_rows()
        s1 = [r for r in led if r[2] == "ЭОМ Проверка заполнения"]
        self.assertTrue(s1 and all(r[9] == "open" for r in s1), "не проверялось ≠ исправлено")


def make_custom_export(folder, specs, filt=None):
    """specs: [(имя спецификации, [(модель.rvt, семейство, наименование), ...])] — manifest в заданном (перемешанном) порядке."""
    os.makedirs(folder, exist_ok=True)
    head = ["Семейство и типоразмер", "Наименование", "Связанный файл"]
    fields = "|".join(head)
    manifest = []
    for i, (name, rows) in enumerate(specs, 1):
        fn = "%04d.tsv" % i
        manifest.append("%s\t%s\t%s\t%s" % (fn, name, fields, filt if filt is not None else "Парам_%d" % i))
        with open(os.path.join(folder, fn), "w", encoding="utf-8-sig", newline="") as f:
            f.write("\t".join(head) + "\n")
            for mdl, fam, item in rows:
                f.write("\t".join([fam, item, mdl]) + "\n")
    with open(os.path.join(folder, "manifest.txt"), "w", encoding="utf-8-sig", newline="") as f:
        f.write("\n".join(manifest) + "\n")
    with open(os.path.join(folder, "export_result.txt"), "w") as f:
        f.write("exported=%d\n" % len(specs))


@unittest.skipUnless(PWSH, "нет pwsh")
class ParamColumnTests(unittest.TestCase):
    """«Параметр» берётся из manifest (4-й столбец), но ключи истории и журнала от него не зависят."""

    def run_build(self, tmp, filt, n):
        exp = os.path.join(tmp, "exp%d" % n)
        make_custom_export(exp, [("П        14.7_Длина_Детали лотков_Без значения",
                                  [("PROJ_IOS_EOM_ORG_K01_R24.rvt", "Колено", "Угол 1"), ("PROJ_IOS_EOM_ORG_K01_R24.rvt", "Тройник", "Т 1")])], filt=filt)
        import time
        time.sleep(1.1)
        r = pwsh("-File", os.path.join(TOOL, "build_assignment.ps1"), "-ExportDir", exp, "-OutXlsx", os.path.join(tmp, "out", "task.xlsx"))
        self.assertEqual(r.returncode, 0, r.stdout + r.stderr)

    def test_param_display_changes_but_keys_do_not(self):
        import openpyxl
        import warnings
        tmp = tempfile.mkdtemp(prefix="bvs_param_")
        self.addCleanup(shutil.rmtree, tmp, True)
        os.makedirs(os.path.join(tmp, "out"))
        self.run_build(tmp, "Код по классификатору|ADSK_Размер_Длина", 1)   # как было: лишний фильтр в «Параметре»
        hist1 = open(os.path.join(tmp, "out", "задание_история.txt"), encoding="utf-8-sig").read().splitlines()
        led1 = open(os.path.join(tmp, "out", "задание_журнал.txt"), encoding="utf-8-sig").read().splitlines()
        self.run_build(tmp, "ADSK_Размер_Длина", 2)                          # исправлено: только проверяемый параметр
        hist2 = open(os.path.join(tmp, "out", "задание_история.txt"), encoding="utf-8-sig").read().splitlines()
        led2 = open(os.path.join(tmp, "out", "задание_журнал.txt"), encoding="utf-8-sig").read().splitlines()
        # ключ истории (колонка Param) не изменился
        self.assertEqual({l.split("\t")[4] for l in hist1}, {l.split("\t")[4] for l in hist2})
        # те же замечания остались открытыми: ни «нового», ни «возврата», ни закрытых
        keys1 = sorted(l.split("\t")[:5] for l in led1)
        keys2 = sorted(l.split("\t")[:5] for l in led2)
        self.assertEqual([k[0:5] for k in keys1], [k[0:5] for k in keys2])
        for l in led2:
            f = l.split("\t")
            self.assertEqual(f[9], "open")
            self.assertEqual(f[11], "0")
            self.assertEqual(f[7], l.split("\t")[7])
        self.assertEqual(sorted(l.split("\t")[7] for l in led1), sorted(l.split("\t")[7] for l in led2))  # FirstSeen прежний
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            wb = openpyxl.load_workbook(os.path.join(tmp, "out", "task.xlsx"))
        ws = wb["ЭОМ"]
        params = {ws.cell(r, 3).value for r in range(3, ws.max_row + 1) if ws.cell(r, 2).value}
        self.assertEqual(params, {"ADSK_Размер_Длина"})


ORDER_SPECS = [
    ("П        99_Толщина стенки_Без значения", "Толщина"),
    ("П        14.7_Длина_Детали лотков_Без значения", "Длина"),
    ("П        4.4_Имя системы_Оборудование_Без значения", "Имя"),
    ("П        13_Материал_Фитинги_Пустые", "Материал"),
    ("П        4.4_Имя системы_Оборудование_Пустые", "Имя"),
    ("П        5.7_Ширина_Детали лотков_Без значения", "Ширина"),
    ("П        13_Материал_Фитинги_Без значния", "Материал"),
]
ORDER_EXPECTED = [
    "П        4.4_Имя системы_Оборудование_Пустые",
    "П        4.4_Имя системы_Оборудование_Без значения",
    "П        5.7_Ширина_Детали лотков_Без значения",
    "П        13_Материал_Фитинги_Пустые",
    "П        13_Материал_Фитинги_Без значения",
    "П        14.7_Длина_Детали лотков_Без значения",
    "П        99_Толщина стенки_Без значения",
]


@unittest.skipUnless(PWSH, "нет pwsh")
class OrderAndPairsTests(unittest.TestCase):
    def build(self):
        tmp = tempfile.mkdtemp(prefix="bvs_order_")
        self.addCleanup(shutil.rmtree, tmp, True)
        exp = os.path.join(tmp, "exp")
        out = os.path.join(tmp, "out")
        os.makedirs(out)
        specs = []
        for name, tag in ORDER_SPECS:
            rows = [("PROJ_IOS_EOM_ORG_K01_R24.rvt", "Сем %s" % tag, "Эл %s 1" % tag), ("PROJ_IOS_EOM_ORG_K01_R24.rvt", "Сем %s" % tag, "Эл %s 2" % tag)]
            specs.append((name, rows))
        make_custom_export(exp, specs)
        self.xlsx = os.path.join(out, "task.xlsx")
        r = pwsh("-File", os.path.join(TOOL, "build_assignment.ps1"), "-ExportDir", exp, "-OutXlsx", self.xlsx)
        self.assertEqual(r.returncode, 0, r.stdout + r.stderr)
        self.payload = os.path.join(out, "google_push_payload.json")
        import openpyxl
        import warnings
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            return openpyxl.load_workbook(self.xlsx)

    def test_summary_and_sections_follow_numeric_order(self):
        import json
        wb = self.build()
        hdr = [c.value for c in wb["Сводка"][2] if c.value and not str(c.value).startswith(("Модель", "Замечания"))]
        self.assertEqual(hdr, ORDER_EXPECTED)
        ws = wb["ЭОМ"]
        specs = [ws.cell(r, 2).value for r in range(3, ws.max_row + 1) if ws.cell(r, 2).value]
        self.assertEqual(specs, ORDER_EXPECTED)
        # пары «Пустые»/«Без значения»: в шапке (строка 2) между ними нет границы, в строках с числами — есть
        sv = wb["Сводка"]
        names = {c.value: c.column for c in sv[2] if c.value}

        def side(cell, which):
            b = getattr(cell.border, which)
            return b.style if b is not None else None

        for left_name, right_name in (
            ("П        4.4_Имя системы_Оборудование_Пустые", "П        4.4_Имя системы_Оборудование_Без значения"),
            ("П        13_Материал_Фитинги_Пустые", "П        13_Материал_Фитинги_Без значения"),
        ):
            cl, cr = names[left_name], names[right_name]
            self.assertEqual(cr, cl + 1)
            self.assertIsNone(side(sv.cell(2, cl), "right"))
            self.assertIsNone(side(sv.cell(2, cr), "left"))
            self.assertEqual(side(sv.cell(3, cl), "right"), "thin")
            self.assertEqual(side(sv.cell(3, cr), "left"), "thin")
        single = names["П        5.7_Ширина_Детали лотков_Без значения"]
        self.assertEqual(side(sv.cell(2, single), "left"), "thin")
        self.assertEqual(side(sv.cell(2, single), "right"), "thin")
        # Google получает тот же порядок
        payload = json.load(open(self.payload, encoding="utf-8"))
        self.assertEqual(payload["sheets"][0]["meta"]["pairCols"], [names["П        4.4_Имя системы_Оборудование_Пустые"], names["П        13_Материал_Фитинги_Пустые"]])
        sheets = {x["name"]: x for x in payload["sheets"]}
        self.assertEqual([c for c in sheets["Сводка"]["rows"][1][1:] if c], ORDER_EXPECTED)
        self.assertEqual([r[1] for r in sheets["ЭОМ"]["rows"][2:] if r[1]], ORDER_EXPECTED)


@unittest.skipUnless(PWSH, "нет pwsh")
class SpecRenameKeepsHistoryTests(unittest.TestCase):
    """Спеки «…_Без значния» переименованы в «…_Без значения»: ключи истории, журнала и буфера не меняются, история не теряется."""
    OLD = "П        4.1_Имя системы_Воздуховоды_Без значния"
    NEW = "П        4.1_Имя системы_Воздуховоды_Без значения"
    MODEL = "PROJ_IOS_OV_ORG_K01_R24.rvt"

    def build(self, tmp, n, name, items):
        import time
        exp = os.path.join(tmp, "exp%d" % n)
        make_custom_export(exp, [(name, [(self.MODEL, "Сем", i) for i in items])], filt="ADSK_Размер")
        time.sleep(1.1)
        r = pwsh("-File", os.path.join(TOOL, "build_assignment.ps1"), "-ExportDir", exp, "-OutXlsx", os.path.join(tmp, "out", "task.xlsx"))
        self.assertEqual(r.returncode, 0, r.stdout + r.stderr)

    @staticmethod
    def lines(path):
        return open(path, encoding="utf-8-sig").read().splitlines()

    def test_old_history_and_ledger_with_typo_continue_under_new_name(self):
        import openpyxl
        import warnings
        tmp = tempfile.mkdtemp(prefix="bvs_ren_")
        self.addCleanup(shutil.rmtree, tmp, True)
        out = os.path.join(tmp, "out")
        os.makedirs(out)
        hist, led = os.path.join(out, "задание_история.txt"), os.path.join(out, "задание_журнал.txt")
        self.build(tmp, 1, self.OLD, ["Эл 1", "Эл 2", "Эл 3"])
        # файлы прошлых сборок, как их писал инструмент до переименования: имя с опечаткой
        for path in (hist, led):
            text = open(path, encoding="utf-8-sig").read()
            open(path, "w", encoding="utf-8-sig", newline="").write(text.replace("Без значения", "Без значния"))
        led1 = self.lines(led)
        self.assertTrue(all("Без значния" in l for l in led1) and led1)
        self.build(tmp, 2, self.NEW, ["Эл 1", "Эл 2"])        # в Revit спеку переименовали, один элемент исправлен
        led2 = self.lines(led)
        self.assertEqual(len(led2), 3)                         # те же три записи, без дублей под двумя именами
        self.assertFalse(any("значния" in l for l in led2))
        by_item = {l.split("\t")[4]: l.split("\t") for l in led2}
        first = {l.split("\t")[4]: l.split("\t") for l in led1}
        for item in ("Эл 1", "Эл 2"):
            self.assertEqual(by_item[item][9], "open")
            self.assertEqual(by_item[item][7], first[item][7])          # FirstSeen прежний: замечание не «новое»
            self.assertEqual(by_item[item][11], "0")                    # и не «вернулось»
        self.assertEqual(by_item["Эл 3"][9], "fixed")                  # исправленное распознано именно по старой записи
        rows = [l.split("\t") for l in self.lines(hist)]
        self.assertFalse(any("значния" in l[3] for l in rows))
        slices = sorted({l[7] for l in rows if len(l) > 7})
        self.assertEqual(len(slices), 2)                               # обе сборки — срезы одной спецификации
        last = [l for l in rows if l[7] == slices[-1]]
        prev = [l for l in rows if l[7] == slices[0]]
        self.assertEqual((last[0][5], prev[0][5]), ("2", "3"))        # было 3, стало 2: история не потеряна
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            wb = openpyxl.load_workbook(os.path.join(out, "task.xlsx"))
        hdr = [c.value for c in wb["Сводка"][2] if c.value and "Имя системы" in str(c.value)]
        self.assertEqual(hdr, [self.NEW])                              # один столбец, а не два (старое и новое имя)

    def test_latin_c_in_history_and_ledger_continues_under_cyrillic_name(self):
        lat, cyr = "\u0043антехприборы", "\u0421антехприборы"
        old, new = "П        2.5_Клас-я системы_%s_Все" % lat, "П        2.5_Клас-я системы_%s_Все" % cyr
        tmp = tempfile.mkdtemp(prefix="bvs_lc_")
        self.addCleanup(shutil.rmtree, tmp, True)
        out = os.path.join(tmp, "out")
        os.makedirs(out)
        hist, led = os.path.join(out, "задание_история.txt"), os.path.join(out, "задание_журнал.txt")
        self.build(tmp, 1, old, ["Эл 1", "Эл 2", "Эл 3"])
        for path in (hist, led):
            text = open(path, encoding="utf-8-sig").read()
            open(path, "w", encoding="utf-8-sig", newline="").write(text.replace(cyr, lat))
        led1 = self.lines(led)
        self.assertTrue(led1 and all(lat in l for l in led1))
        self.build(tmp, 2, new, ["Эл 1", "Эл 2"])
        led2 = self.lines(led)
        self.assertEqual(len(led2), 3)                         # те же записи, без дублей под двумя написаниями
        self.assertFalse(any(lat in l for l in led2))
        by_item = {l.split("\t")[4]: l.split("\t") for l in led2}
        first = {l.split("\t")[4]: l.split("\t") for l in led1}
        for item in ("Эл 1", "Эл 2"):
            self.assertEqual((by_item[item][9], by_item[item][7]), ("open", first[item][7]))
        self.assertEqual(by_item["Эл 3"][9], "fixed")
        rows = [l.split("\t") for l in self.lines(hist)]
        self.assertFalse(any(lat in l[3] for l in rows))

    def test_partial_buffer_checks_and_catalog_with_typo_are_read_under_new_name(self):
        ps = os.path.join(tmp_dir := tempfile.mkdtemp(prefix="bvs_ren2_"), "t.ps1")
        self.addCleanup(shutil.rmtree, tmp_dir, True)
        buf = os.path.join(tmp_dir, "b.txt")
        chk = os.path.join(tmp_dir, "c.txt")
        cat = os.path.join(tmp_dir, "k.txt")
        old = self.OLD
        open(buf, "w", encoding="utf-8-sig").write("OV\tМодель\t%s\tПар\tПар\tСем\tЭл\tЗнач\n" % old)
        open(chk, "w", encoding="utf-8-sig").write("Модель\t%s\t2026-10-01 10:00:00\n" % old)
        open(cat, "w", encoding="utf-8-sig").write("spec\t%s\nmodel\tМодель\tloaded\n" % old)
        with open(ps, "w", encoding="utf-8-sig") as f:
            f.write(". '%s'\n. '%s'\n$b = @(Read-PartialBuffer '%s'); $c = Read-PartialChecks '%s'; $k = Read-ExportCatalog '%s'\n"
                    "Write-Output $b[0].Check; Write-Output (@($c.Keys) -join '|'); Write-Output (@($k.Specs.Keys) -join '|')\n"
                    % (os.path.join(TOOL, "assignment_export_stream.ps1"), os.path.join(TOOL, "assignment_partial.ps1"), buf, chk, cat))
        r = pwsh("-File", ps)
        self.assertEqual(r.returncode, 0, r.stdout + r.stderr)
        out = [l for l in r.stdout.splitlines() if l.strip()]
        self.assertEqual(out[0], self.NEW)
        self.assertTrue(out[1].endswith(self.NEW), out[1])
        self.assertEqual(out[2], self.NEW)


class _Handler(http.server.BaseHTTPRequestHandler):
    mode = "ok"
    seen = []

    def log_message(self, *a):
        pass

    def _send(self, code, body, ctype="application/json"):
        data = body.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", ctype + "; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def do_POST(self):  # как Apps Script: POST -> 302 -> GET с результатом
        n = int(self.headers.get("Content-Length", 0))
        _Handler.seen.append(json.loads(self.rfile.read(n).decode("utf-8")))
        if _Handler.mode == "404":
            return self._send(404, "nope", "text/plain")
        self.send_response(302)
        self.send_header("Location", "/echo")
        self.send_header("Content-Length", "0")
        self.end_headers()

    def do_GET(self):
        m = _Handler.mode
        if m == "ok":
            self._send(200, json.dumps({"ok": True, "version": 3}))
        elif m == "legacy":
            self._send(200, "ok", "text/plain")
        elif m == "html":
            self._send(200, "<!DOCTYPE html><html><body>Нужна авторизация</body></html>", "text/html")
        elif m == "error":
            self._send(200, "error: Exception: boom", "text/plain")
        elif m == "false":
            self._send(200, json.dumps({"ok": False, "error": "row group depth cannot exceed 8"}))
        else:
            self._send(200, "POST JSON from push_google.ps1", "text/plain")


@unittest.skipUnless(PWSH, "нет pwsh")
class PushTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.srv = http.server.ThreadingHTTPServer(("127.0.0.1", 0), _Handler)
        cls.port = cls.srv.server_address[1]
        threading.Thread(target=cls.srv.serve_forever, daemon=True).start()
        cls.tool = tempfile.mkdtemp(prefix="bvs_tool_")
        with open(os.path.join(cls.tool, "google_push.cfg"), "w", encoding="utf-8") as f:
            f.write("webapp_url=http://127.0.0.1:%d/macros/s/AKfycbTEST/exec\nsecret=test-secret-123\n" % cls.port)

    @classmethod
    def tearDownClass(cls):
        cls.srv.shutdown()
        shutil.rmtree(cls.tool, True)

    def push(self, mode, payload="payload_multi.json"):
        _Handler.mode = mode
        _Handler.seen = []
        return pwsh("-File", os.path.join(TOOL, "push_google.ps1"), "-PayloadOnly",
                    "-PayloadPath", os.path.join(HERE, "fixtures", payload), "-ToolDir", self.tool)

    def test_success_json(self):
        r = self.push("ok")
        self.assertEqual(r.returncode, 0, r.stdout + r.stderr)
        self.assertRegex(r.stdout.strip().splitlines()[-1], r'^OK \{"ok": ?true')
        body = _Handler.seen[0]
        self.assertEqual(body["secret"], "test-secret-123")
        names = [s["name"] for s in body["sheets"]]
        self.assertEqual(names, ["Сводка", "ЭОМ"])  # кириллица дошла без искажений
        for s in body["sheets"]:
            self.assertTrue(all(isinstance(r, list) for r in s["rows"]))
        # разделы, которых нет в выгрузке, отправлены на очистку; «Сводка» и присутствующие — нет
        self.assertEqual(sorted(body["clear"]), sorted(["СС", "ОВ", "ПД", "ПТ", "ВК", "ОТ", "ХС", "ИТП", "АВТ", "ХЦ"]))
        self.assertEqual(body["hide"], ["Динамика"])

    def test_partial_payload_clears_only_its_own_list(self):
        r = self.push("ok", "payload_partial.json")
        self.assertEqual(r.returncode, 0, r.stdout + r.stderr)
        body = _Handler.seen[0]
        self.assertEqual([s["name"] for s in body["sheets"]], ["Сводка", "ЭОМ"])
        self.assertEqual(body["clear"], ["ВК"])      # остальные разделы (СС, ОВ…) в выборочной отправке не очищаются

    def test_one_row_payload_stays_nested(self):
        r = self.push("ok", "payload_empty.json")
        self.assertEqual(r.returncode, 0, r.stdout + r.stderr)
        rows = _Handler.seen[0]["sheets"][0]["rows"]
        self.assertTrue(isinstance(rows, list) and isinstance(rows[0], list), rows)

    def test_legacy_ok_accepted_with_note(self):
        r = self.push("legacy")
        self.assertEqual(r.returncode, 0, r.stdout + r.stderr)
        self.assertIn("старая версия Code.gs", r.stdout)

    def test_failures_are_not_reported_as_ok(self):
        for mode, needle in (("html", "HTML"), ("error", "boom"), ("false", "depth"), ("other", "POST JSON"), ("404", "404")):
            with self.subTest(mode=mode):
                r = self.push(mode)
                self.assertNotEqual(r.returncode, 0, "режим %s ошибочно принят: %s" % (mode, r.stdout))
                self.assertNotRegex(r.stdout, r"(?m)^OK ")
                self.assertIn(needle, r.stdout + r.stderr)


@unittest.skipUnless(NODE, "нет node")
class CodeGsMockTests(unittest.TestCase):
    def test_code_gs_mock(self):
        r = subprocess.run([NODE, os.path.join(HERE, "code_gs_mock_test.js")], capture_output=True, text=True, encoding="utf-8")
        self.assertEqual(r.returncode, 0, r.stdout + r.stderr)


class CheckerSelfTests(unittest.TestCase):
    """Проверяльщик xlsx сам должен ловить прежние дефекты (на файлах, собранных openpyxl)."""

    def make(self, mutate):
        import openpyxl
        from openpyxl.styles import PatternFill
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Сводка"
        ws["A1"] = "x"
        ws["A2"] = "Модель"
        mutate(ws, PatternFill)
        d = tempfile.mkdtemp(prefix="bvs_chk_")
        self.addCleanup(shutil.rmtree, d, True)
        p = os.path.join(d, "t.xlsx")
        wb.save(p)
        return p, d

    def test_detects_single_cell_merge_and_swapped_color(self):
        p, d = self.make(lambda ws, PF: (ws.merge_cells("A1:A1") if False else None, setattr(ws["A1"], "fill", PF("solid", fgColor="FF00FFFF"))))
        payload = os.path.join(d, "p.json")
        with open(payload, "w", encoding="utf-8") as f:
            json.dump({"sheets": [{"name": "Сводка", "rows": [["x"], ["Модель"]], "formats": [{"r": 1, "c": 1, "bg": 65535}]}]}, f)
        problems = []
        wb = check_xlsx.check_workbook(p, problems)
        check_xlsx.check_payload(wb, payload, problems)
        self.assertTrue(any("цвет" in x for x in problems), problems)  # жёлтый 0x00FFFF записан как голубой

    def test_detects_flat_rows(self):
        p, d = self.make(lambda ws, PF: None)
        payload = os.path.join(d, "p.json")
        with open(payload, "w", encoding="utf-8") as f:
            json.dump({"sheets": [{"name": "Сводка", "rows": ["x", "Модель"]}]}, f)
        problems = []
        wb = check_xlsx.check_workbook(p, problems)
        check_xlsx.check_payload(wb, payload, problems)
        self.assertTrue(any("массивом массивов" in x for x in problems), problems)


@unittest.skipUnless(PWSH, "нет pwsh")
class TypeSizeErrorReportTests(unittest.TestCase):
    """«0_Ошибка типоразмера»: колонка Сводки первой, на листах разделов «Параметр»/«Значение», пустой manifest закрывает замечание, «Динамика» помнит его."""

    SPEC0 = "П        0_Ошибка типоразмера"
    SPEC1 = "П        1_Классификатор_Пустые"
    OV = "PROJ_IOS_OV_ORG_K01_R24.rvt"
    VK = "PROJ_IOS_VK_ORG_K01_R24.rvt"

    def write_export(self, folder, type_rows):
        os.makedirs(folder, exist_ok=True)
        head0 = ["Семейство и типоразмер", "Наименование", "ADSK_Наименование", "Связанный файл"]
        head1 = ["Семейство и типоразмер", "Наименование", "Код по классификатору", "Связанный файл"]
        man = []
        # классификатор в файле раньше ошибки типоразмера: порядок отчёта — не порядок manifest
        with open(os.path.join(folder, "0001.tsv"), "w", encoding="utf-8-sig", newline="") as f:
            f.write("\t".join(head1) + "\n")
            f.write("\t".join(["Колено", "Колено 1", "", self.OV]) + "\n")
        man.append("0001.tsv\t%s\t%s\tКод по классификатору" % (self.SPEC1, "|".join(head1)))
        if type_rows is None:
            man.append("*empty*\t" + self.SPEC0)
        else:
            with open(os.path.join(folder, "0002.tsv"), "w", encoding="utf-8-sig", newline="") as f:
                f.write("\t".join(head0) + "\n")
                for fam, item, model in type_rows:
                    f.write("\t".join([fam, item, "Ошибка типоразмера", model]) + "\n")
            man.append("0002.tsv\t%s\t%s\tADSK_Наименование" % (self.SPEC0, "|".join(head0)))
        with open(os.path.join(folder, "manifest.txt"), "w", encoding="utf-8-sig", newline="") as f:
            f.write("\n".join(man) + "\n")
        with open(os.path.join(folder, "export_result.txt"), "w", encoding="utf-8") as f:
            f.write("exported=%d\n" % (1 if type_rows is None else 2))

    def build(self, exp, xlsx, when):
        env = dict(ENV, BVS_FAKE_NOW=when)
        r = subprocess.run(
            [PWSH, "-NoProfile", "-File", os.path.join(TOOL, "build_assignment.ps1"), "-ExportDir", exp, "-OutXlsx", xlsx],
            capture_output=True, timeout=300, env=env, cwd=TOOL)
        r.stdout = _decode_out(r.stdout)
        r.stderr = _decode_out(r.stderr)
        self.assertEqual(r.returncode, 0, r.stdout + r.stderr)
        return r

    def test_parser_keeps_literal_in_value_and_leaves_other_specs(self):
        tmp = tempfile.mkdtemp(prefix="bvs_tsz_")
        self.addCleanup(shutil.rmtree, tmp, True)
        script = os.path.join(tmp, "t.ps1")
        stream = os.path.join(TOOL, "assignment_export_stream.ps1")
        with open(script, "w", encoding="utf-8-sig") as f:
            f.write(r"""
. '%s'
function To-Text($Value) { if ($null -eq $Value) { return "" }; return ([string]$Value).Trim() }
function Get-ModelFileName([string]$Link) {
    $text = $Link.Trim()
    if ($text -match '[^\\/]+$') { $text = $Matches[0] }
    return ($text -replace '(?i)\.rvt$', '')
}
$link = 'PROJ_IOS_OV_ORG_K01_R24.rvt'
$h = @('Семейство и типоразмер','Наименование','Связанный файл')
$fld = @('Семейство и типоразмер','ADSK_Наименование','Связанный файл')
$a = Parse-ExportTsvIssue $h @('Отвод: 90','Ошибка типоразмера',$link) 'П        0_Ошибка типоразмера' $fld @('ADSK_Наименование')
'swallow: ' + $a.ParamName + ' | [' + $a.Item + '] | ' + $a.Value + ' | ' + $a.Section
$h2 = @('Семейство и типоразмер','Наименование','ADSK_Наименование','Связанный файл')
$b = Parse-ExportTsvIssue $h2 @('Отвод: 90','Отвод 90','Ошибка типоразмера',$link) 'П        0_Ошибка типоразмера' $h2 @('ADSK_Наименование')
'split: ' + $b.ParamName + ' | ' + $b.Item + ' | ' + $b.Value
$h3 = @('Семейство и типоразмер','ADSK_Материал наименование','Наименование','Связанный файл')
$d = Parse-ExportTsvIssue $h3 @('Труба','медь','Сталь',$link) 'П        13_Материал_Трубы_Пустые' $h3 @('ADSK_Материал наименование')
'material: ' + $d.Item + ' | [' + $d.Value + '] | ' + $d.ParamName
$h4 = @('Проверить длину','Наименование','ADSK_Размер_Длина','Количество','Связанный файл')
$f4 = @('Семейство и типоразмер','Наименование','ADSK_Размер_Длина','Количество','Связанный файл')
$e = Parse-ExportTsvIssue $h4 @('Отвод','Отвод 1','007','1',$link) 'П        14.2_Длина_Трубы_Пустые' $f4 @('ADSK_Размер_Длина')
'length: ' + $e.Family + ' | ' + $e.Item + ' | ' + $e.Value + ' | ' + $e.ParamName
$h5 = @('Семейство и типоразмер','Наименование','Код по классификатору','Связанный файл')
$g = Parse-ExportTsvIssue $h5 @('Щит','Щит 1','999.12',$link) 'П        1_Классификатор_999' $h5 @('Код по классификатору')
'code999: ' + $g.ParamName + ' | ' + $g.Item + ' | ' + $g.Value
$h6 = @('Семейство и типоразмер','Наименование','ADSK_Номер здания','Связанный файл')
$k = Parse-ExportTsvIssue $h6 @('Панель','Панель 1','Корпус-А',$link) 'П        16.0_Корпус_Кириллица' $h6 @('ADSK_Номер здания')
'cyr: ' + $k.ParamName + ' | ' + $k.Item + ' | ' + $k.Value
""" % stream)
        r = pwsh("-File", script)
        self.assertEqual(r.returncode, 0, r.stdout + r.stderr)
        lines = [l for l in r.stdout.splitlines() if l.strip()]
        self.assertEqual(lines[0], "swallow: ADSK_Наименование | [] | Ошибка типоразмера | OV")
        self.assertEqual(lines[1], "split: ADSK_Наименование | Отвод 90 | Ошибка типоразмера")
        self.assertEqual(lines[2], "material: Сталь | [] | ADSK_Материал наименование")
        self.assertEqual(lines[3], "length: Отвод | Отвод 1 | 007 | ADSK_Размер_Длина")
        self.assertEqual(lines[4], "code999: Код по классификатору | Щит 1 | 999.12")
        self.assertEqual(lines[5], "cyr: ADSK_Номер здания | Панель 1 | Корпус-А")

    def test_summary_sections_and_dynamics(self):
        import openpyxl
        import warnings
        tmp = tempfile.mkdtemp(prefix="bvs_tsz_")
        self.addCleanup(shutil.rmtree, tmp, True)
        out = os.path.join(tmp, "out")
        os.makedirs(out)
        xlsx = os.path.join(out, "task.xlsx")
        exp1 = os.path.join(tmp, "exp1")
        self.write_export(exp1, [("Отвод: 90", "Отвод 90", self.OV), ("Труба: 100", "Труба 100", self.VK)])
        self.build(exp1, xlsx, "2026-11-01 10:00:00")
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            wb = openpyxl.load_workbook(xlsx)
        self.assertEqual([c.value for c in wb["Сводка"][2]][:3], ["Модель", self.SPEC0, self.SPEC1])
        ov = wb["ОВ"]
        spec_rows = [r for r in range(3, ov.max_row + 1) if ov.cell(r, 2).value == self.SPEC0]
        self.assertEqual(len(spec_rows), 1)
        self.assertEqual(ov.cell(spec_rows[0], 3).value, "ADSK_Наименование")
        later = [r for r in range(3, ov.max_row + 1) if ov.cell(r, 2).value == self.SPEC1]
        self.assertTrue(later and spec_rows[0] < later[0])
        self.assertTrue(any(ov.cell(r, 5).value == "Отвод 90" and ov.cell(r, 6).value == "Ошибка типоразмера" for r in range(3, ov.max_row + 1)))
        vk = wb["ВК"]
        self.assertTrue(any(vk.cell(r, 2).value == self.SPEC0 and vk.cell(r, 3).value == "ADSK_Наименование" for r in range(3, vk.max_row + 1)))
        self.assertTrue(any(vk.cell(r, 6).value == "Ошибка типоразмера" for r in range(3, vk.max_row + 1)))
        self.assertNotIn("ЭОМ", wb.sheetnames)
        hist = open(os.path.join(out, "задание_история.txt"), encoding="utf-8-sig").read()
        self.assertIn(self.SPEC0, hist)
        self.assertIn("ADSK_Наименование", hist)
        exp2 = os.path.join(tmp, "exp2")
        self.write_export(exp2, None)
        self.build(exp2, xlsx, "2026-11-01 11:00:00")
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            wb = openpyxl.load_workbook(xlsx)
        self.assertEqual(wb["Сводка"]["B2"].value, self.SPEC0)
        hdr = [c.value for c in wb["Сводка"][2]]
        col = hdr.index(self.SPEC0) + 1
        model = self.OV[:-4]
        cell = None
        for r in range(3, wb["Сводка"].max_row + 1):
            if wb["Сводка"].cell(r, 1).value == model:
                cell = wb["Сводка"].cell(r, col).value
        self.assertEqual(cell, "0 (1) \u21931")
        dyn = [str(c.value) for row in wb["Динамика"].iter_rows(min_row=3) for c in row if c.value is not None]
        self.assertIn(self.SPEC0, dyn)
        self.assertIn("ADSK_Наименование", dyn)
        self.assertIn("Ошибка типоразмера", dyn)
        ov = wb["ОВ"]
        self.assertTrue(any(ov.cell(r, 2).value == self.SPEC0 and ov.cell(r, 3).value == "ADSK_Наименование" for r in range(3, ov.max_row + 1)))
        self.assertTrue(any(ov.cell(r, 6).value == "Ошибка типоразмера" for r in range(3, ov.max_row + 1)))


if __name__ == "__main__":
    unittest.main()
