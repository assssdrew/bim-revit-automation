# -*- coding: utf-8 -*-
# «Выгрузка квартирографии» v7. Dynamo (Revit 2024, CPython3).
# Выгружает помещения квартир (СУ/уборная/гардеробная) и помещения МОП из ОДНОЙ связи АР
# в файл корпуса Квартирография_Kxx.xlsx. Excel НЕ запускается: файл пишется напрямую
# (только стандартная библиотека Python), остальные листы и формулы шаблона не трогаются.
# IN[0] - часть имени связи АР (необязательно). OUT - текст результата для Проигрывателя.

# ======================= НАСТРОЙКИ =======================
# Example shared-parameter names (KLWT_*). Change this block for another architectural model.
# Примеры имён — на другом объекте их меняют здесь.
P_FLOOR = "KLWT_Этаж"
P_TYPE  = "KLWT_Тип квартиры"
P_NUM   = "KLWT_Номер квартиры на этаже"
P_SOBSH = "KLWT_Площадь квартиры общая"
P_SZH   = "KLWT_Площадь квартиры жилая"
# Имена помещений квартир для выгрузки (регистр, пробелы, ё/е не важны).
ROOM_NAMES = ["Совмещенный санузел", "Санузел", "Ванная", "Уборная", "Туалет", "Гардеробная"]
EXPORT_ALL_APT_ROOMS = False   # True = выгрузить все помещения, у которых заполнен номер квартиры
# Помещения МОП (лист «МОП»). Те же имена должны стоять на листе «Настройки», столбец F.
MOP_NAMES = ["Лестничная клетка", "Тамбур-шлюз", "Тамбур-шлюз (ПБЗ)", "Коридор", "Лифтовой холл"]
AREA_DECIMALS = 2
# Папка скрипта (здесь лежит шаблон):
SCRIPT_DIR = r"\\server\share\Dynamo\Выгрузка квартирографии"
TEMPLATE_NAME = "Квартирография_ШАБЛОН.xlsx"
REPORT_DIR_NAME = "Квартирография"     # папка отчётов рядом с моделью ОВ
REPORT_PREFIX = "Квартирография_"
HEADER = ["Этаж", "Тип", "Номер", "Sобщ", "Sж", "Имя", "Число"]
MOP_HEADER = ["Этаж", "Имя", "Площадь"]
SHEET_DATA = "Выгрузка"
SHEET_MOP = "Выгрузка МОП"
# =========================================================

import os, re, shutil, datetime, zipfile, tempfile
from xml.sax.saxutils import escape

class UserError(Exception):
    pass

def norm(s):
    return " ".join((s or "").replace("ё", "е").replace("Ё", "Е").split()).casefold()

TR = str.maketrans("АВЕКМНОРСТХаеорсхук", "ABEKMHOPCTXaeopcxyk")

def nkey(s):
    return "".join(ch for ch in (s or "").translate(TR).casefold() if ch not in " _-.\t")

def toks(s):
    return [t for t in re.split(r"[ _\-.:\t]+", (s or "").translate(TR).casefold()) if t]

def to_num(v):
    try:
        return float(str(v).replace(",", ".").strip())
    except Exception:
        return None

def building_code(title):
    t = (title or "").translate(TR).upper()
    m = re.search(r"(?:^|[ _\-.])K(\d{2})(?=$|[ _\-.])", t)
    if m:
        return "K" + m.group(1)
    t = re.sub(r"\.rvt$", "", (title or "").strip(), flags=re.I)
    return re.sub(r"[^\w\-]+", "_", t).strip("_") or "БЕЗ_ИМЕНИ"

# ---------------- запись xlsx без Excel (stdlib) ----------------
def col_letter(i):
    s = ""
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s

_bad = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f]")

def sheet_xml(rows):
    out = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
           '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
           'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
           '<sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/>'
           '<sheetData>']
    for ri, row in enumerate(rows, 1):
        cells = []
        for ci, v in enumerate(row, 1):
            if v is None or v == "":
                continue
            ref = col_letter(ci) + str(ri)
            if isinstance(v, bool):
                v = str(v)
            if isinstance(v, (int, float)):
                cells.append('<c r="%s"><v>%s</v></c>' % (ref, repr(float(v)) if isinstance(v, float) else v))
            else:
                t = escape(_bad.sub("", str(v)))
                cells.append('<c r="%s" t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>' % (ref, t))
        out.append('<row r="%d">%s</row>' % (ri, "".join(cells)))
    out.append('</sheetData><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>')
    return "".join(out).encode("utf-8")

def _sheet_paths(z):
    wb = z.read("xl/workbook.xml").decode("utf-8")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    rid2t = {}
    for m in re.finditer(r"<Relationship\b[^>]*>", rels):
        tag = m.group(0)
        i = re.search(r'\bId="([^"]+)"', tag); t = re.search(r'\bTarget="([^"]+)"', tag)
        if i and t:
            tg = t.group(1)
            tg = tg[1:] if tg.startswith("/") else "xl/" + tg
            rid2t[i.group(1)] = tg
    res = {}
    for m in re.finditer(r"<sheet\b[^>]*>", wb):
        tag = m.group(0)
        n = re.search(r'\bname="([^"]*)"', tag); r = re.search(r'\br:id="([^"]+)"', tag)
        if n and r:
            name = n.group(1).replace("&amp;", "&").replace("&quot;", '"').replace("&lt;", "<").replace("&gt;", ">").replace("&apos;", "'")
            res[name] = rid2t.get(r.group(1))
    return res

def sheet_names(path):
    with zipfile.ZipFile(path) as z:
        return list(_sheet_paths(z).keys())

def _full_calc(wb):
    if "<calcPr" in wb:
        def fix(m):
            t = re.sub(r'\s(fullCalcOnLoad|forceFullCalc)="[^"]*"', "", m.group(0))
            return t.replace("<calcPr", '<calcPr fullCalcOnLoad="1"', 1)
        return re.sub(r"<calcPr\b[^>]*?/?>", fix, wb, count=1)
    return wb.replace("</workbook>", '<calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>')

def write_sheets(path, data):
    """data = {имя листа: строки}. Заменяет листы целиком (старое удаляется), остальное не трогает."""
    with zipfile.ZipFile(path) as z:
        paths = _sheet_paths(z)
        miss = [n for n in data if n not in paths]
        if miss:
            raise UserError("В файле %s нет листа «%s»." % (path, "», «".join(miss)))
        repl = dict((paths[n], sheet_xml(rows)) for n, rows in data.items())
        fd, tmp = tempfile.mkstemp(suffix=".xlsx", dir=os.path.dirname(path))
        os.close(fd)
        try:
            with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as o:
                for it in z.infolist():
                    n = it.filename
                    if n == "xl/calcChain.xml":
                        continue
                    b = z.read(n)
                    if n in repl:
                        b = repl[n]
                    elif n == "xl/workbook.xml":
                        b = _full_calc(b.decode("utf-8")).encode("utf-8")
                    elif n == "xl/_rels/workbook.xml.rels":
                        b = re.sub(r'<Relationship\b[^>]*calcChain[^>]*/>', "", b.decode("utf-8")).encode("utf-8")
                    elif n == "[Content_Types].xml":
                        b = re.sub(r'<Override\b[^>]*calcChain[^>]*/>', "", b.decode("utf-8")).encode("utf-8")
                    o.writestr(it, b)
        except Exception:
            try: os.remove(tmp)
            except OSError: pass
            raise
    try:
        os.replace(tmp, path)
    except OSError:
        try: os.remove(tmp)
        except OSError: pass
        raise UserError("Файл занят: %s, закройте его и запустите снова." % path)

def check_free(path):
    if not os.path.exists(path):
        return
    try:
        with open(path, "r+b"):
            pass
    except OSError:
        raise UserError("Файл занят: %s, закройте его и запустите снова." % path)

def prepare_file(folder, code, template):
    """Создаёт файл корпуса из шаблона; файл старого формата (без листа МОП) сохраняет копией."""
    path = os.path.join(folder, REPORT_PREFIX + code + ".xlsx")
    note = ""
    check_free(path)
    if os.path.exists(path):
        try:
            ok = SHEET_MOP in sheet_names(path)
        except Exception:
            ok = False
        if not ok:
            bak = os.path.join(folder, REPORT_PREFIX + code + "_старый_" + datetime.datetime.now().strftime("%Y%m%d_%H%M") + ".xlsx")
            try:
                os.replace(path, bak)
            except OSError:
                raise UserError("Файл занят: %s, закройте его и запустите снова." % path)
            note = "Файл был старого формата (без листа МОП) - сохранён копией: %s" % os.path.basename(bak)
    if not os.path.exists(path):
        if not os.path.exists(template):
            raise UserError("Нет шаблона %s в папке скрипта %s" % (TEMPLATE_NAME, os.path.dirname(template)))
        shutil.copyfile(template, path)
    return path, note

def write_report(folder, code, template, rows, mop_rows, link_name):
    path, note = prepare_file(folder, code, template)
    stamp = datetime.datetime.now().strftime("%d.%m.%Y %H:%M")
    data = [HEADER] + rows + [["Связь АР:", link_name], ["Выгружено:", stamp]]
    mop = [MOP_HEADER] + mop_rows
    write_sheets(path, {SHEET_DATA: data, SHEET_MOP: mop})
    return path, note

# ======================= Revit =======================
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
        return p.AsInteger()
    if st == StorageType.Double:
        try:
            if p.Definition.GetDataType() == SpecTypeId.Area:
                v = UnitUtils.ConvertFromInternalUnits(p.AsDouble(), UnitTypeId.SquareMeters)
                return round(v, AREA_DECIMALS)
        except Exception:
            pass
        return (p.AsValueString() or "").strip()
    return (p.AsValueString() or "").strip()

def room_area(r):
    return round(UnitUtils.ConvertFromInternalUnits(r.Area, UnitTypeId.SquareMeters), AREA_DECIMALS)

def numify(v):
    if isinstance(v, str):
        n = to_num(v)
        if n is not None:
            return int(n) if n == int(n) and "." not in v and "," not in v else n
    return v

def find_link(doc, key):
    k = nkey(key)
    links = FilteredElementCollector(doc).OfClass(RevitLinkInstance).ToElements()
    names, ar = [], {}
    for li in links:
        ld = li.GetLinkDocument()
        names.append(li.Name + ("" if ld else "  [НЕ ЗАГРУЖЕНА]"))
        if ld is None:
            continue
        if k:
            if k in nkey(li.Name) or k in nkey(ld.Title):
                return ld
        elif "ar" in toks(ld.Title) or "ar" in toks(li.Name):
            ar[ld.Title] = ld
    if not k and len(ar) == 1:
        return list(ar.values())[0]
    lst = "\n  - " + "\n  - ".join(names or ["(связей нет)"])
    if not k and len(ar) > 1:
        raise UserError("Загружено несколько связей АР. Впишите часть имени нужной связи (например K04) "
                        "в поле «Связь АР» или выгрузите лишние связи в «Управление связями»:\n  - " + "\n  - ".join(sorted(ar)))
    what = ("с «%s» в имени" % key) if k else "АР (с «_AR_» в имени)"
    raise UserError("Загруженная связь %s не найдена. Загрузите связь или впишите часть имени в поле «Связь АР». Связи в модели:%s" % (what, lst))

def report_folder(doc):
    path = ""
    if doc.IsWorkshared:
        mp = doc.GetWorksharingCentralModelPath()
        if mp is not None:
            if mp.CloudPath:
                raise UserError("Модель ОВ хранится в облаке (BIM 360/ACC). Папку отчётов рядом с ней создать нельзя.")
            path = ModelPathUtils.ConvertModelPathToUserVisiblePath(mp) or ""
    if not path:
        path = doc.PathName or ""
    low = path.strip().lower()
    if not low:
        raise UserError("Модель ОВ не сохранена. Сохраните модель и запустите снова.")
    if low.startswith(("bim 360://", "autodesk docs://", "cloud://", "rsn://", "http://", "https://")):
        raise UserError("Модель ОВ хранится в облаке/на Revit Server (%s). Папку отчётов рядом с ней создать нельзя." % path)
    folder = os.path.join(os.path.dirname(path), REPORT_DIR_NAME)
    if not os.path.isdir(folder):
        try:
            os.makedirs(folder)
        except OSError:
            raise UserError("Не удалось создать папку отчётов %s (нет доступа?)." % folder)
    return folder

def collect(found):
    NAMES = set(norm(n) for n in ROOM_NAMES)
    MOP = dict((norm(n), n) for n in MOP_NAMES)
    rows, mop, skipped, mop_skipped, missing = [], [], 0, 0, set()
    rooms = FilteredElementCollector(found).OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType()
    for r in rooms:
        if r.Area <= 0:
            continue
        nm_p = r.get_Parameter(BuiltInParameter.ROOM_NAME)
        nm = ((nm_p.AsString() if nm_p else "") or "").strip()
        n = norm(nm)
        if n in MOP:
            fl = pval(r, P_FLOOR)
            if fl is None:
                missing.add(P_FLOOR)
            if fl in (None, ""):
                mop_skipped += 1
            else:
                mop.append([numify(fl), MOP[n], room_area(r)])
            continue
        num = pval(r, P_NUM)
        if EXPORT_ALL_APT_ROOMS:
            if num in (None, ""):
                continue
        elif n not in NAMES:
            continue
        vals = [pval(r, p) for p in (P_FLOOR, P_TYPE, P_NUM, P_SOBSH, P_SZH)]
        for pn, v in zip([P_FLOOR, P_TYPE, P_NUM, P_SOBSH, P_SZH], vals):
            if v is None:
                missing.add(pn)
        if vals[0] in (None, "") or vals[2] in (None, ""):
            skipped += 1
            continue
        rows.append([numify("" if v is None else v) for v in vals] + [nm, 1])
    def f(v):
        n = to_num(v)
        return 9e9 if n is None else n
    rows.sort(key=lambda x: (f(x[0]), f(x[2]), x[5]))
    mop.sort(key=lambda x: (f(x[0]), x[1]))
    return rows, mop, skipped, mop_skipped, missing

def main(key):
    doc = DocumentManager.Instance.CurrentDBDocument
    found = find_link(doc, key)
    folder = report_folder(doc)
    rows, mop, skipped, mop_skipped, missing = collect(found)
    code = building_code(found.Title)
    path, note = write_report(folder, code, os.path.join(SCRIPT_DIR, TEMPLATE_NAME), rows, mop, found.Title)
    msg = ["ГОТОВО: корпус %s" % code, "Файл: %s" % path, "Связь АР: %s" % found.Title,
           "Помещений квартир: %d, помещений МОП: %d" % (len(rows), len(mop))]
    if note: msg.append(note)
    if missing: msg.append("ВНИМАНИЕ: нет параметров в связи: " + ", ".join(sorted(missing)))
    if skipped: msg.append("ВНИМАНИЕ: пропущено помещений квартир без этажа/номера: %d" % skipped)
    if mop_skipped: msg.append("ВНИМАНИЕ: пропущено помещений МОП без этажа: %d" % mop_skipped)
    if not rows: msg.append("ВНИМАНИЕ: помещения квартир не найдены - проверьте ROOM_NAMES.")
    return "\n".join(msg)

if "IN" in globals():
    try:
        import clr
        clr.AddReference("RevitAPI")
        clr.AddReference("RevitServices")
        from Autodesk.Revit.DB import (ModelPathUtils, FilteredElementCollector, RevitLinkInstance, BuiltInCategory,
                                       BuiltInParameter, StorageType, UnitUtils, UnitTypeId, SpecTypeId)
        from RevitServices.Persistence import DocumentManager
        OUT = main("" if IN[0] is None else str(IN[0]).strip())
    except UserError as e:
        OUT = "ОШИБКА: " + str(e)
    except Exception as e:
        import traceback
        OUT = "ОШИБКА (непредвиденная): %s\n%s" % (e, traceback.format_exc())
