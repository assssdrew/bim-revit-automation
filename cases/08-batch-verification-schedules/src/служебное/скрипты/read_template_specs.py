# -*- coding: utf-8 -*-
"""
Revit Batch Processor — список спецификаций ШАБЛОНА для вкладки «Распределение» (только чтение, модель не меняется и не сохраняется).

Шаблон открывается как обычная модель задания RBP; в файл out_path пишется по строке на спецификацию: имя TAB категория.
Берутся спецификации (не шаблоны видов, не ведомости изменений и не ключи), у которых название без префикса стадии («П», «РД», «ПП»)
начинается с цифры («3.1_Имя системы_Воздуховоды_Пустые»); only_numbered=0 в настройках — брать все.
Файлы — только через .NET (пути с кириллицей); при ошибке traceback в отчеты\\логи\\read_template_specs_error.txt (если в run.cfg есть log_folder=), иначе в %TEMP%\\read_template_specs_error.txt, и в лог RBP.

Настройки: %LOCALAPPDATA%\\BatchRvt\\batch_verification_schedules\\run.cfg (секция [template]; старый template_read.cfg читается, только если в run.cfg нет out_path).
  out_path=файл результата, only_numbered=1|0, links_path=файл имён связей (необязательно)
Если задан links_path, в тот же прогон пишутся имена RevitLinkType без .rvt (по строке на связь).
Совместимо с IronPython 2.7: без f-строк, flags=, lookbehind, str.lower() для кириллицы.
"""

import clr
import sys
import traceback

clr.AddReference("RevitAPI")

from Autodesk.Revit.DB import FilteredElementCollector, ViewSchedule

import revit_script_util
from revit_script_util import Output

STEP = [u"start"]


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


def path_join(*parts):
    from System.IO import Path
    result = parts[0]
    for part in parts[1:]:
        result = Path.Combine(result, part)
    return as_text(result)


def env_var(name, default=u""):
    from System import Environment
    value = Environment.GetEnvironmentVariable(name)
    return as_text(value) if value else default


def local_app_data():
    """%LOCALAPPDATA% текущего пользователя; нет переменной — <профиль>\AppData\Local (без зашитых имён пользователей)."""
    return env_var(u"LOCALAPPDATA") or path_join(env_var(u"USERPROFILE"), u"AppData", u"Local")


def parse_cfg_file(path):
    from System.IO import File
    data = {}
    for line in as_text(File.ReadAllText(path, utf8())).splitlines():
        line = line.strip()
        if not line or line.startswith(u"#") or line.startswith(u"[") or u"=" not in line:
            continue
        key, value = line.split(u"=", 1)
        data[key.strip().lower()] = value.strip()
    return data


def read_cfg():
    """Настройки чтения шаблона пишутся в общий run.cfg (секция [template]).
    template_read.cfg остаётся запасным, если run.cfg ещё без out_path (старая копия на диске)."""
    from System.IO import File
    base = path_join(local_app_data(), u"BatchRvt", u"batch_verification_schedules")
    primary = path_join(base, u"run.cfg")
    fallback = path_join(base, u"template_read.cfg")
    path = primary
    data = parse_cfg_file(path) if File.Exists(path) else {}
    if not data.get(u"out_path") and File.Exists(fallback):
        path = fallback
        data = parse_cfg_file(path)
    STEP[0] = u"read cfg " + path
    if not data.get(u"out_path"):
        raise Exception(u"Нет файла настроек: " + primary)
    return data


def strip_stage_prefix(name):
    text = as_text(name).strip()
    for prefix in (u"РД", u"ПП", u"П"):
        if text.startswith(prefix) and len(text) > len(prefix):
            nxt = text[len(prefix)]
            if nxt.isspace() or nxt.isdigit():
                return text[len(prefix):].strip()
    return text


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


def strip_rvt_name(raw):
    """Имя связи без пути и без .rvt (в том числе «Имя.rvt : 1»). Расширение латиницей, без str.lower() по кириллице."""
    text = as_text(raw).strip()
    for sep in (u"\\", u"/"):
        if sep in text:
            text = text.split(sep)[-1]
            break
    i = 0
    n = len(text)
    while i < n:
        if text[i] == u"." and i + 3 < n:
            a = text[i + 1]
            b = text[i + 2]
            c = text[i + 3]
            if a in (u"r", u"R") and b in (u"v", u"V") and c in (u"t", u"T"):
                return text[:i].strip()
        i += 1
    return text.strip()


def is_revit_link(element, link_cls):
    try:
        if isinstance(element, link_cls):
            return True
    except Exception:
        pass
    try:
        return as_text(element.GetType().Name) == u"RevitLinkType"
    except Exception:
        return False


def collect_links(document):
    """Имена RevitLinkType без .rvt. Сбой чтения связей не должен ломать список спецификаций."""
    out = []
    try:
        from Autodesk.Revit.DB import RevitLinkType
    except Exception:
        return out
    try:
        for lt in FilteredElementCollector(document).OfClass(RevitLinkType):
            try:
                if not is_revit_link(lt, RevitLinkType):
                    continue
                name = strip_rvt_name(getattr(lt, "Name", u""))
                if name:
                    out.append(name)
            except Exception:
                continue
    except Exception:
        return out
    return out


def category_name(view):
    try:
        cat = view.Definition.CategoryId
        element = view.Document.Settings.Categories
        for c in element:
            if c.Id == cat:
                return as_text(c.Name)
    except Exception:
        pass
    return u""


def cfg_log_folder():
    """Папка журналов из run.cfg (log_folder=). Нет ключа или файла — пустая строка (ошибка уйдёт в %TEMP%)."""
    try:
        from System.IO import File
        base = path_join(local_app_data(), u"BatchRvt", u"batch_verification_schedules")
        for name in (u"run.cfg", u"template_read.cfg"):
            path = path_join(base, name)
            if File.Exists(path):
                folder = parse_cfg_file(path).get(u"log_folder", u"")
                if folder:
                    return folder
    except Exception:
        pass
    return u""


def error_file_path(name=u"read_template_specs_error.txt"):
    from System.IO import Path, Directory
    folder = cfg_log_folder()
    if folder:
        try:
            Directory.CreateDirectory(folder)
            return path_join(folder, name)
        except Exception:
            pass
    return path_join(as_text(Path.GetTempPath()), name)


def report_error(model_path):
    from System.IO import File
    parts = [u"model: " + as_text(model_path), u"step: " + as_text(STEP[0])]
    try:
        parts.append(as_text(traceback.format_exc()))
    except Exception as ex:
        parts.append(u"traceback недоступен: " + as_text(ex))
    text = u"\r\n".join(parts) + u"\r\n" + (u"-" * 60) + u"\r\n"
    where = u""
    try:
        where = error_file_path()
        File.AppendAllText(where, text, utf8())
    except Exception:
        where = u"(не удалось записать файл)"
    log_msg(u"READ TEMPLATE ERROR, details: " + where)
    log_msg(text)


def main():
    from System.IO import File
    document = revit_script_util.GetScriptDocument()
    cfg = read_cfg()
    out_path = cfg.get(u"out_path", u"")
    if not out_path:
        raise Exception(u"Не задан out_path")
    only_numbered = cfg.get(u"only_numbered", u"1") != u"0"
    STEP[0] = u"read schedules"
    lines = []
    total = 0
    for view in usable_schedules(document):
        total += 1
        name = as_text(view.Name).replace(u"\t", u" ").strip()
        if only_numbered:
            base = strip_stage_prefix(name)
            if not base or not base[0].isdigit():
                continue
        lines.append(name + u"\t" + category_name(view))
    STEP[0] = u"write " + out_path
    File.WriteAllText(out_path, u"\r\n".join(lines) + (u"\r\n" if lines else u""), utf8())
    link_count = 0
    links_path = cfg.get(u"links_path", u"")
    if links_path:
        STEP[0] = u"write links"
        try:
            link_names = collect_links(document)
            link_count = len(link_names)
            File.WriteAllText(links_path, u"\r\n".join(link_names) + (u"\r\n" if link_names else u""), utf8())
        except Exception:
            log_msg(u"links write failed")
    log_msg(u"template specs={0} of schedules={1} links={2}".format(len(lines), total, link_count))


def run():
    model_path = u""
    try:
        model_path = revit_script_util.GetRevitFilePath()
        main()
    except Exception:
        report_error(model_path)
        raise


run()
