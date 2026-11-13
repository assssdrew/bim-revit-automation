"""Проверка задания xlsx без Excel (openpyxl + разбор XML) и его соответствия google_push_payload.json.

Запуск на любом ПК (в том числе на Windows после сборки):
    python check_xlsx.py <задание.xlsx> [google_push_payload.json]
Код выхода 0 — файл годен, 1 — есть проблемы (список печатается).
"""
import json
import re
import sys
import zipfile
import warnings
import xml.etree.ElementTree as ET

import openpyxl

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
# порядок дочерних элементов <worksheet> по схеме ECMA-376 (нужные нам)
WS_ORDER = ["sheetPr", "dimension", "sheetViews", "sheetFormatPr", "cols", "sheetData", "mergeCells"]
PUSH_SHEETS = ["Сводка", "ЭОМ", "СС", "ОВ", "ПД", "ПТ", "ВК", "ОТ", "ХС"]


def bgr_to_rgb(n):
    """Цвет Excel (.Interior.Color = R + G*256 + B*65536) -> 'RRGGBB'."""
    return "%02X%02X%02X" % (n & 0xFF, (n >> 8) & 0xFF, (n >> 16) & 0xFF)


def check_zip_xml(path, problems):
    with zipfile.ZipFile(path) as z:
        bad = z.testzip()
        if bad:
            problems.append("zip: битая запись " + bad)
        names = set(z.namelist())
        for need in ("[Content_Types].xml", "xl/workbook.xml", "xl/styles.xml"):
            if need not in names:
                problems.append("нет части " + need)
        for n in sorted(names):
            if not re.match(r"xl/worksheets/sheet\d+\.xml$", n):
                continue
            raw = z.read(n)
            try:
                root = ET.fromstring(raw)
            except ET.ParseError as e:
                problems.append("%s: XML не разбирается: %s" % (n, e))
                continue
            order = [c.tag.replace(NS, "") for c in root]
            idx = [WS_ORDER.index(t) for t in order if t in WS_ORDER]
            if idx != sorted(idx):
                problems.append("%s: неверный порядок элементов %s" % (n, order))
            if b't="str"' in raw:
                problems.append(n + ': есть t="str" (Excel просит восстановить файл)')
            for mc in root.iter(NS + "mergeCell"):
                a, _, b = mc.get("ref").partition(":")
                if not b or a == b:
                    problems.append("%s: объединение одной ячейки %s" % (n, mc.get("ref")))
            # недопустимые управляющие символы
            if re.search(rb"[\x00-\x08\x0b\x0c\x0e-\x1f]", raw):
                problems.append(n + ": управляющие символы в XML")


def cell_text(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v)


def check_workbook(path, problems):
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        wb = openpyxl.load_workbook(path)
    if "Сводка" not in wb.sheetnames:
        problems.append("нет листа «Сводка»")
    for ws in wb.worksheets:
        depth = 0
        for rd in ws.row_dimensions.values():
            depth = max(depth, rd.outlineLevel or 0)
        if depth > 3:
            problems.append("%s: глубина группировки %d > 3" % (ws.title, depth))
    return wb


def check_payload(wb, payload_path, problems):
    with open(payload_path, encoding="utf-8") as f:
        payload = json.load(f)
    sheets = payload.get("sheets")
    if not isinstance(sheets, list) or not sheets:
        problems.append("payload: нет sheets")
        return
    for sh in sheets:
        name = sh.get("name")
        rows = sh.get("rows")
        if not isinstance(rows, list) or not all(isinstance(r, list) for r in rows):
            problems.append("payload[%s]: rows должен быть массивом массивов" % name)
            continue
        widths = {len(r) for r in rows}
        if len(widths) > 1:
            problems.append("payload[%s]: строки разной длины %s" % (name, sorted(widths)))
        if name not in wb.sheetnames:
            problems.append("payload[%s]: такого листа нет в xlsx" % name)
            continue
        ws = wb[name]
        if ws.cell(2, 7).value == "Статус":
            # лист раздела с журналом: в xlsx есть колонки статуса и исправленные строки (в Google их нет);
            # строки модели и спецификации Google получает, только если под ними есть открытые строки
            def kind(r):
                if ws.cell(r, 1).value:
                    return "model"
                if ws.cell(r, 2).value:
                    return "spec"
                return "item"
            xl = []
            last = {"model": None, "spec": None}
            open_under = {"model": False, "spec": False}
            pending = []                     # [(вид, индекс в xl)]
            for r in range(3, ws.max_row + 1):
                k = kind(r)
                row = [cell_text(ws.cell(r, c).value) for c in range(1, 7)]
                if k == "item":
                    if ws.cell(r, 7).value == "Исправлено":
                        continue
                    for pk, pi in pending:
                        xl[pi][1] = True
                    xl.append([row, True, "item"])
                else:
                    if k == "model":
                        pending = []
                    else:
                        pending = [x for x in pending if x[0] == "model"]
                    pending.append((k, len(xl)))
                    xl.append([row, False, k])
            xl = [x[0] for x in xl if x[1]]
            pay = [[str(v) for v in row[:6]] for row in rows[2:]]
            if xl != pay:
                problems.append("расхождение %s: открытые строки xlsx и payload не совпадают (xlsx %d, payload %d)" % (name, len(xl), len(pay)))
            continue
        for ri, row in enumerate(rows, 1):
            for ci, val in enumerate(row, 1):
                got = cell_text(ws.cell(ri, ci).value)
                if got != str(val):
                    problems.append("расхождение %s!R%dC%d: xlsx=%r payload=%r" % (name, ri, ci, got, val))
                    break
            else:
                continue
            break
        for f in sh.get("formats", []):
            if "bg" not in f:
                continue
            want = bgr_to_rgb(int(f["bg"]))
            got = ws.cell(f["r"], f["c"]).fill.fgColor.rgb
            if not isinstance(got, str) or got[-6:].upper() != want:
                problems.append("цвет %s!R%dC%d: xlsx=%s ожидалось %s (путаница R/B?)" % (name, f["r"], f["c"], got, want))
                break


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        return 2
    problems = []
    check_zip_xml(sys.argv[1], problems)
    try:
        wb = check_workbook(sys.argv[1], problems)
    except Exception as e:  # noqa: BLE001
        print("openpyxl не открыл файл:", e)
        return 1
    if len(sys.argv) > 2:
        check_payload(wb, sys.argv[2], problems)
    for p in problems:
        print("ПРОБЛЕМА:", p)
    print("OK" if not problems else "ПРОБЛЕМ: %d" % len(problems))
    return 1 if problems else 0


if __name__ == "__main__":
    sys.exit(main())
