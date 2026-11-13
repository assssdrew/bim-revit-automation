# -*- coding: utf-8 -*-
"""Список разделов для IronPython 2.7 и CPython.

Файл разделы.txt лежит рядом с этим модулем. Нет файла или он не читается —
возвращается встроенный список (тот же, что в файле) и признак fallback.
Исключение наружу не бросает.
"""

BUILTIN = (
    (u"EOM", u"ЭОМ", (u"ЭОМ", u"EOM"), 8),
    (u"SS", u"СС", (u"СС", u"SS"), 9),
    (u"OV", u"ОВ", (u"ОВ", u"OV"), 1),
    (u"PD", u"ПД", (u"ПД", u"PD"), 2),
    (u"PT", u"ПТ", (u"ПТ", u"PT"), 4),
    (u"VK", u"ВК", (u"ВК", u"VK"), 3),
    (u"OT", u"ОТ", (u"ОТ", u"OT"), 5),
    (u"HS", u"ХС", (u"ХС", u"HS"), 6),
    (u"ITP", u"ИТП", (u"ИТП", u"ITP"), 7),
    (u"AVT", u"АВТ", (u"АВТ", u"AVT"), 10),
    (u"HC", u"ХЦ", (u"ХЦ", u"HC"), 11),
)


def _as_text(value):
    if value is None:
        return u""
    return u"{}".format(value)


def builtin_rows():
    rows = []
    index = 0
    for code, title, names, column in BUILTIN:
        index += 1
        rows.append({
            "code": code,
            "title": title,
            "names": [n for n in names],
            "column": column,
            "report": index,
        })
    return rows


def _read_text(path):
    try:
        from System.IO import File
        from System.Text import UTF8Encoding
        return _as_text(File.ReadAllText(path, UTF8Encoding(True)))
    except Exception:
        # Нет .NET (обычный CPython) или в тестах подставлен чужой модуль System.
        pass
    handle = open(path, "rb")
    try:
        raw = handle.read()
    finally:
        handle.close()
    bom = "\xef\xbb\xbf"
    if isinstance(raw, bytes):
        if raw[:3] == bom.encode("latin1"):
            raw = raw[3:]
        return raw.decode("utf-8")
    if raw[:3] == bom:
        raw = raw[3:]
    if isinstance(raw, bytes):
        return raw.decode("utf-8")
    return raw.decode("utf-8")


def _parse(text):
    rows = []
    index = 0
    for raw in _as_text(text).splitlines():
        line = raw.strip()
        if not line or line.startswith(u"#"):
            continue
        parts = line.split(u"|")
        if len(parts) < 2:
            raise Exception("section line needs code and title")
        code = parts[0].strip()
        title = parts[1].strip()
        if not code or not title:
            raise Exception("empty code or title")
        names = []
        if len(parts) >= 3 and parts[2].strip():
            for item in parts[2].split(u","):
                item = item.strip()
                if item:
                    names.append(item)
        if not names:
            names = [title, code]
        column = 0
        if len(parts) >= 4 and parts[3].strip():
            try:
                column = int(parts[3].strip())
            except Exception:
                column = 0
        index += 1
        if column < 1:
            column = 1000 + index
        rows.append({
            "code": code,
            "title": title,
            "names": names,
            "column": column,
            "report": index,
        })
    if not rows:
        raise Exception("no sections")
    return rows


def load_sections(path):
    """(rows, used_fallback). rows — в порядке листов задания."""
    try:
        return _parse(_read_text(path)), False
    except Exception:
        return builtin_rows(), True


def column_rows(rows):
    return sorted(rows, key=lambda row: (row["column"], row["report"]))


def script_dir():
    try:
        return os_path_dirname(__file__)
    except Exception:
        return u""


def os_path_dirname(path):
    import os
    return os.path.dirname(os.path.abspath(path))


def default_file():
    import os
    folder = script_dir()
    if not folder:
        return u""
    return os.path.join(folder, u"разделы.txt")
