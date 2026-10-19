# -*- coding: utf-8 -*-
# Dynamo (Revit 2024, CPython3). Выгрузка СЫРЫХ строк помещений из связи АР.
# Ничего не считает: одна строка = одно помещение. Порядок столбцов = лист «Выгрузка».
# IN[0] - часть имени связи АР (необязательно), IN[1] - папка отчётов (там лежит шаблон).
# OUT = [путь к файлу корпуса, строки для листа «Выгрузка»]

# ======================= НАСТРОЙКИ =======================
# Example parameter names from one office's shared parameters (KLWT_*).
# They are not a fixed standard: change this block when the architectural
# team uses different names. Примеры имён — на другом объекте их меняют здесь.
P_FLOOR = "KLWT_Этаж"
P_TYPE  = "KLWT_Тип квартиры"
P_NUM   = "KLWT_Номер квартиры на этаже"
P_SOBSH = "KLWT_Площадь квартиры общая"
P_SZH   = "KLWT_Площадь квартиры жилая"
# Имена помещений для выгрузки (регистр, пробелы, ё/е не важны).
# Те же варианты должны быть на листе Excel «Настройки», иначе строка выгрузится, но не попадёт в столбец.
ROOM_NAMES = ["Совмещенный санузел", "Санузел", "Ванная", "Уборная", "Туалет", "Гардеробная"]
EXPORT_ALL_APT_ROOMS = False   # True = выгрузить все помещения, у которых заполнен номер квартиры
AREA_DECIMALS = 1
TEMPLATE_NAME = "Типовые_этажи_ШАБЛОН.xlsx"
REPORT_PREFIX = "Типовые_этажи_"
HEADER = ["Этаж", "Тип", "Номер", "Sобщ", "Sж", "Имя", "Число"]
# =========================================================

import clr
clr.AddReference("RevitAPI")
clr.AddReference("RevitServices")
from Autodesk.Revit.DB import (FilteredElementCollector, RevitLinkInstance, BuiltInCategory,
                               BuiltInParameter, StorageType, UnitUtils, UnitTypeId, SpecTypeId)
from RevitServices.Persistence import DocumentManager

doc = DocumentManager.Instance.CurrentDBDocument

def norm(s):
    return " ".join((s or "").replace("ё", "е").replace("Ё", "Е").split()).casefold()

NAMES = set(norm(n) for n in ROOM_NAMES)

def pval(el, name):
    p = el.LookupParameter(name)
    if p is None:
        return None
    if not p.HasValue:
        return ""
    st = p.StorageType
    if st == StorageType.String:
        return (p.AsString() or "").strip()
    if st == StorageType.Integer:
        return str(p.AsInteger())
    if st == StorageType.Double:
        try:
            if p.Definition.GetDataType() == SpecTypeId.Area:
                v = UnitUtils.ConvertFromInternalUnits(p.AsDouble(), UnitTypeId.SquareMeters)
                return round(v, AREA_DECIMALS)
        except Exception:
            pass
        return (p.AsValueString() or "").strip()
    return (p.AsValueString() or "").strip()

def nkey(s):
    # кириллица-двойники -> латиница, без регистра, пробелов, _ - .
    tr = str.maketrans("АВЕКМНОРСТХаеорсхук", "ABEKMHOPCTXaeopcxyk")
    s = (s or "").translate(tr).casefold()
    return "".join(ch for ch in s if ch not in " _-.\t")

import re
def toks(s):
    tr = str.maketrans("АВЕКМНОРСТХаеорсхук", "ABEKMHOPCTXaeopcxyk")
    return [t for t in re.split(r"[ _\-.:\t]+", (s or "").translate(tr).casefold()) if t]

key = "" if IN[0] is None else str(IN[0]).strip()
k = nkey(key)
links = FilteredElementCollector(doc).OfClass(RevitLinkInstance).ToElements()
found, found_name, names, ar = None, "", [], {}
for li in links:
    ld = li.GetLinkDocument()
    names.append(li.Name + ("" if ld else "  [НЕ ЗАГРУЖЕНА]"))
    if ld is None:
        continue
    if k:
        if found is None and (k in nkey(li.Name) or k in nkey(ld.Title)):
            found, found_name = ld, ld.Title
    elif "ar" in toks(ld.Title) or "ar" in toks(li.Name):
        ar[ld.Title] = ld          # несколько экземпляров одной связи = одна связь

if not k:
    if len(ar) == 1:
        found_name, found = list(ar.items())[0]
    elif len(ar) > 1:
        OUT = [["ОШИБКА: загружено несколько связей АР. Впишите часть имени нужной связи в поле «Связь АР»:"]] + [[n] for n in sorted(ar)]
if found is None and not (not k and len(ar) > 1):
    what = ("с '%s' в имени" % key) if k else "АР (с '_AR_' в имени)"
    OUT = [["ОШИБКА: загруженная связь %s не найдена. Загрузите связь или впишите часть имени. Связи в модели:" % what]] + [[n] for n in names]
if found is None:
    pass
else:
    rows, skipped, missing = [HEADER], 0, set()
    rooms = FilteredElementCollector(found).OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType()
    for r in rooms:
        if r.Area <= 0:          # не размещённые / неокружённые
            continue
        nm_p = r.get_Parameter(BuiltInParameter.ROOM_NAME)
        nm = (nm_p.AsString() if nm_p else "") or ""
        num = pval(r, P_NUM)
        if EXPORT_ALL_APT_ROOMS:
            if not num:
                continue
        elif norm(nm) not in NAMES:
            continue
        vals = [pval(r, P_FLOOR), pval(r, P_TYPE), num, pval(r, P_SOBSH), pval(r, P_SZH)]
        for pn, v in zip([P_FLOOR, P_TYPE, P_NUM, P_SOBSH, P_SZH], vals):
            if v is None:
                missing.add(pn)
        if not vals[0] or not vals[2]:
            skipped += 1
            continue
        rows.append([("" if v is None else v) for v in vals] + [nm.strip(), 1])
    def sk(x):
        def f(v):
            try: return float(str(v).replace(",", "."))
            except: return 9e9
        return (f(x[0]), f(x[2]), x[5])
    OUT = [rows[0]] + sorted(rows[1:], key=sk)
    if missing:
        OUT.append(["ВНИМАНИЕ: нет параметров в связи: " + ", ".join(sorted(missing))])
    if skipped:
        OUT.append(["Пропущено помещений без этажа/номера: %d" % skipped])
    OUT.append(["Связь АР:", found_name])

# ---------- файл корпуса ----------
import os, shutil, datetime

def building_code(title):
    tr = str.maketrans("АВЕКМНОРСТХаеорсхук", "ABEKMHOPCTXaeopcxyk")
    t = (title or "").translate(tr).upper()
    m = re.search(r"(?:^|[ _\-.])K(\d{2})(?=$|[ _\-.])", t)
    if m:
        return "K" + m.group(1)
    t = re.sub(r"\.rvt$", "", (title or "").strip(), flags=re.I)
    return re.sub(r"[^\w\-]+", "_", t).strip("_") or "БЕЗ_ИМЕНИ"

def prepare_file(folder, code):
    folder = (folder or "").strip().strip('"')
    if not folder or not os.path.isdir(folder):
        raise Exception("Папка отчётов не найдена: '%s'. Укажите её в поле «Папка отчётов»." % folder)
    tpl = os.path.join(folder, TEMPLATE_NAME)
    path = os.path.join(folder, REPORT_PREFIX + code + ".xlsx")
    if not os.path.exists(path):
        if not os.path.exists(tpl):
            raise Exception("Нет шаблона %s в папке %s" % (TEMPLATE_NAME, folder))
        shutil.copyfile(tpl, path)
    try:
        with open(path, "ab"):
            pass
    except (PermissionError, IOError, OSError):
        raise Exception("Закройте файл %s и запустите снова." % path)
    return path

is_error = found is None
code = "ОШИБКА" if is_error else building_code(found_name)
OUT.append(["Выгружено:", datetime.datetime.now().strftime("%d.%m.%Y %H:%M")])
OUT = [prepare_file(IN[1], code), OUT]
