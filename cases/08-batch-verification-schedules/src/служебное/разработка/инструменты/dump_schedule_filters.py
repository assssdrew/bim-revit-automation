# -*- coding: utf-8 -*-
"""
Revit Batch Processor — выгрузка фильтров всех спецификаций модели в текстовый файл (только чтение).

Модель не меняется и не сохраняется, в Google ничего не пишется. Для каждой спецификации (кроме шаблонов, ведомостей
изменений и ключей) пишутся все её фильтры в том же виде, что и фильтры_спецификаций.txt:
  Спецификация <TAB> Поле фильтра <TAB> Условие <TAB> Значение
Спецификация без фильтров — одна строка с полем «(нет фильтров)». Значение: строка фильтра; числовое/идентификатор — как
есть; у условий без значения («имеет значение», «не имеет значения») пусто.

Файл: <report_dir>\\фильтры_<имя модели>.txt (UTF-8 с BOM). Сравнение двух таких файлов — compare_filter_dumps.ps1.
Запуск: служебное\\разработка\\инструменты\\Выгрузить фильтры спек.bat -> dump_schedule_filters.ps1 (через RBP).
Настройки: %LOCALAPPDATA%\\BatchRvt\\batch_verification_schedules\\dump_filters.cfg
  report_dir=папка для файлов
  name_contains=часть имени|часть имени   (необязательно: только спецификации, в имени которых есть одна из частей)
Совместимо с IronPython 2.7: без f-строк, flags= и регулярных выражений с lookbehind.
Файлы — только System.IO; при ошибке traceback пишется в отчеты\\логи\\dump_schedule_filters_error.txt (если в dump_filters.cfg есть log_folder=), иначе в %TEMP%, и в лог RBP.
"""

import clr
import sys
import traceback

clr.AddReference("RevitAPI")

from Autodesk.Revit.DB import FilteredElementCollector, ViewSchedule

import revit_script_util
from revit_script_util import Output

# Все файловые операции — только через .NET (System.IO), без встроенных файловых функций и os: в IronPython 2.7 под RBP они на путях с
# кириллицей (служебное, отчеты) дают [IOError]/[WindowsError] без стека. Путь самого скрипта не нужен.
HEADER = u"Спецификация\tПоле фильтра\tУсловие\tЗначение"
NO_FILTERS = u"(нет фильтров)"
ERROR_FILE = u"dump_schedule_filters_error.txt"
STEP = [u"start"]          # где остановились: в RBP у IOError нет стека, поэтому шаг пишется в файл ошибки


def as_text(value):
    if value is None:
        return u""
    try:
        return unicode(value)  # IronPython 2
    except NameError:
        return u"{}".format(value)
    except Exception:
        return u"{}".format(value)


def log_msg(text):
    safe = u"".join((c if ord(c) < 128 else u"?") for c in as_text(text))
    try:
        Output(safe)
    except Exception:
        pass


def utf8():
    from System.Text import UTF8Encoding
    return UTF8Encoding(True)


def env_var(name, default=u""):
    from System import Environment
    value = Environment.GetEnvironmentVariable(name)
    return as_text(value) if value else default


def local_app_data():
    """%LOCALAPPDATA% текущего пользователя; нет переменной — <профиль>\AppData\Local (без зашитых имён пользователей)."""
    return env_var(u"LOCALAPPDATA") or path_join(env_var(u"USERPROFILE"), u"AppData", u"Local")


def path_join(*parts):
    from System.IO import Path
    result = parts[0]
    for part in parts[1:]:
        result = Path.Combine(result, part)
    return as_text(result)


def read_text(path):
    """UTF-8 (с BOM или без) через .NET."""
    from System.IO import File
    return as_text(File.ReadAllText(path, utf8()))


def write_utf8(path, text):
    from System.IO import File
    File.WriteAllText(path, text, utf8())


def append_utf8(path, text):
    from System.IO import File
    File.AppendAllText(path, text, utf8())


def file_exists(path):
    from System.IO import File
    return bool(File.Exists(path))


def ensure_dir(path):
    """Папку (со всеми родителями) создаёт .NET; если уже есть — ничего не делает."""
    from System.IO import Directory
    Directory.CreateDirectory(path)


def cfg_path():
    return path_join(local_app_data(),
                     u"BatchRvt", u"batch_verification_schedules", u"dump_filters.cfg")


def read_cfg():
    data = {}
    path = cfg_path()
    STEP[0] = u"read cfg " + path
    if not file_exists(path):
        raise Exception(u"Нет файла настроек: " + path)
    for line in read_text(path).splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        data[key.strip().lower()] = value.strip()
    return data


def cfg_log_folder():
    """Папка журналов из dump_filters.cfg (log_folder=). Нет ключа или файла — пустая строка."""
    try:
        path = cfg_path()
        if not file_exists(path):
            return u""
        for line in read_text(path).splitlines():
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, value = line.split("=", 1)
            if key.strip().lower() == u"log_folder" and value.strip():
                return value.strip()
    except Exception:
        pass
    return u""


def error_path():
    from System.IO import Path, Directory
    folder = cfg_log_folder()
    if folder:
        try:
            Directory.CreateDirectory(folder)
            return path_join(folder, ERROR_FILE)
        except Exception:
            pass
    return path_join(as_text(Path.GetTempPath()), ERROR_FILE)


def report_error(model_path):
    """Полный traceback (+ шаг и CLR-исключение) — в отчеты\\логи, если задан log_folder=; иначе в %TEMP%. И в лог RBP."""
    parts = [u"model: " + as_text(model_path), u"step: " + as_text(STEP[0])]
    try:
        parts.append(as_text(traceback.format_exc()))
    except Exception as ex:
        parts.append(u"traceback недоступен: " + as_text(ex))
    try:
        exc = sys.exc_info()[1]
        cls = getattr(exc, "clsException", None)
        if cls is not None:
            parts.append(u"CLR: " + as_text(cls.ToString()))
        else:
            parts.append(u"exception: " + as_text(repr(exc)))
    except Exception:
        pass
    text = u"\r\n".join(parts) + u"\r\n" + (u"-" * 60) + u"\r\n"
    where = u""
    try:
        where = error_path()
        append_utf8(where, text)
    except Exception:
        where = u"(не удалось записать файл)"
    log_msg(u"DUMP ERROR, details: " + where)
    log_msg(text)


def clean(value):
    return as_text(value).replace(u"\t", u" ").replace(u"\r", u" ").replace(u"\n", u" ").strip()


def model_leaf(path):
    text = as_text(path).strip().replace(u"\\", u"/")
    leaf = text.split(u"/")[-1]
    if leaf.lower().endswith(u".rvt"):
        leaf = leaf[:-4]
    bad = u'<>:"/\\|?*'
    return u"".join((u"_" if c in bad else c) for c in leaf) or u"model"


def usable_schedules(document):
    result = []
    for view in FilteredElementCollector(document).OfClass(ViewSchedule):
        try:
            if view.IsTemplate:
                continue
        except Exception:
            pass
        skip = False
        for flag in ("IsTitleblockRevisionSchedule", "IsInternalKeynoteSchedule", "IsKeynoteSchedule"):
            try:
                if getattr(view, flag):
                    skip = True
                    break
            except Exception:
                pass
        if not skip:
            result.append(view)
    return result


def filter_value_text(flt):
    """Строка фильтра; иначе число; иначе идентификатор элемента; иначе пусто (условия без значения)."""
    try:
        return clean(flt.GetStringValue())
    except Exception:
        pass
    for getter in ("GetDoubleValue", "GetIntegerValue"):
        try:
            return clean(getattr(flt, getter)())
        except Exception:
            pass
    try:
        return u"id:" + clean(flt.GetElementIdValue().IntegerValue)
    except Exception:
        pass
    return u""


def schedule_rows(view):
    name = clean(view.Name)
    rows = []
    definition = view.Definition
    count = definition.GetFilterCount()
    for index in range(count):
        flt = definition.GetFilter(index)
        try:
            field = clean(definition.GetField(flt.FieldId).GetName())
        except Exception:
            field = u"(поле недоступно)"
        try:
            kind = clean(flt.FilterType)
        except Exception:
            kind = u""
        rows.append(u"\t".join([name, field, kind, filter_value_text(flt)]))
    if count == 0:
        rows.append(u"\t".join([name, NO_FILTERS, u"", u""]))
    return rows


def main(model_path):
    document = revit_script_util.GetScriptDocument()
    cfg = read_cfg()
    out_dir = cfg.get("report_dir", u"")
    if not out_dir:
        raise Exception(u"В dump_filters.cfg нет report_dir")
    STEP[0] = u"create dir " + out_dir
    if file_exists(out_dir) or out_dir.lower().endswith((u".rvt", u".rte", u".rfa", u".rft")):
        raise Exception(u"report_dir указывает на файл, а не на папку: " + out_dir)
    ensure_dir(out_dir)
    parts = [p.strip() for p in cfg.get("name_contains", u"").split(u"|") if p.strip()]
    lines = [HEADER]
    total = 0
    errors = 0
    STEP[0] = u"read schedules"
    for view in usable_schedules(document):
        name = as_text(view.Name)
        if parts and not any(p in name for p in parts):
            continue
        total += 1
        STEP[0] = u"read schedule " + name
        try:
            lines.extend(schedule_rows(view))
        except Exception as ex:
            errors += 1
            lines.append(u"\t".join([clean(name), u"[ОШИБКА чтения: " + clean(ex) + u"]", u"", u""]))
    target = path_join(out_dir, u"фильтры_" + model_leaf(model_path) + u".txt")
    STEP[0] = u"write " + target
    write_utf8(target, u"\r\n".join(lines) + u"\r\n")
    log_msg(u"filters dump specs={0} errors={1}".format(total, errors))


def run():
    model_path = u""
    try:
        model_path = revit_script_util.GetRevitFilePath()
        main(model_path)
    except Exception:
        report_error(model_path)
        raise


run()
