# -*- coding: utf-8 -*-
"""
Revit Batch Processor — вставка проверочных спецификаций в модели ИОС.

Для каждой модели:
  - набор спецификаций — из настройки проекта (вкладка «Распределение», cfg: project_path): отметки именно этой модели;
    модели нет в настройке — объединение её раздела (раздел — из имени файла _OV_, _VK_ ... или из настройки);
  - имя вида = стадия проекта + 8 пробелов + название из шаблона (шаблон — без префиксов стадии);
  - в значениях фильтров {СТАДИЯ}/{STAGE} и {ШИФР}/{CODE} подставляются стадия и шифр проекта;
  - mode=prefill — пробная вставка спецификаций шаблона в модель и подсчёт элементов с откатом (модель не меняется и не сохраняется);
  - режим insert_mode в run.cfg (окно «Файлы» / «Разделы»):
    initial — первичная загрузка: ничего не удалять, вставить весь список;
    replace — заменить пакет: удалить в модели спецификации из списка (по хвосту имени), вставить список заново;
    add_diff — добавить только новые: ничего не удалять, вставить строки, которых ещё нет (по хвосту имени);
  - копия вставляется с тем именем, которое уже стоит в файле-источнике;
  - снимается «Включить элементы из связей»;
  - убирается поле «Связанный файл»;
  - в отчёт пишутся параметры, которых нет в модели.

run.cfg — в %LOCALAPPDATA%\\BatchRvt\\batch_verification_schedules\\run.cfg
  source=полный путь к rvt со спецификациями
  report_dir=папка отчёта
  run_id=метка прогона
"""

import clr
import os

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    ViewSchedule,
    Transaction,
    ElementTransformUtils,
    CopyPasteOptions,
    Transform,
    IDuplicateTypeNamesHandler,
    DuplicateTypeAction,
    ElementId,
    ScheduleFieldType,
    BuiltInParameter,
    ScheduleFilter,
    TransactionGroup,
    SectionType,
)
from System.Collections.Generic import List

import revit_script_util
from revit_script_util import Output


def file_exists(path):
    """.NET вместо os.path.isfile: в IronPython 2.7 на путях с кириллицей (отчеты, сервер) os.path.* отвечает неверно."""
    try:
        from System.IO import File
        return bool(File.Exists(path))
    except ImportError:
        return os.path.isfile(path)


def ensure_dir(path):
    """Папка через .NET (не падает, если уже есть, и понимает кириллицу в пути)."""
    try:
        from System.IO import Directory
        Directory.CreateDirectory(path)
    except ImportError:
        if not os.path.isdir(path):
            os.makedirs(path)


def lower_text(value):
    """Нижний регистр БЕЗ str.lower(): в IronPython 2.7 оно меняет только ASCII, кириллица остаётся как была."""
    try:
        make_char = unichr  # IronPython 2
    except NameError:
        make_char = chr
    out = []
    for ch in as_text(value):
        code = ord(ch)
        if 65 <= code <= 90 or 1040 <= code <= 1071:
            ch = make_char(code + 32)
        elif code == 1025:
            ch = make_char(1105)
        out.append(ch)
    return u"".join(out)


def norm_path(path):
    return lower_text(as_text(path).replace(u"/", u"\\"))


def log_msg(text):
    try:
        t = unicode(text)
    except Exception:
        t = u"{}".format(text)
    safe = u"".join((c if ord(c) < 128 else u"?") for c in t)
    try:
        Output(safe)
    except Exception:
        pass

# Пути — только от расположения скриптов и профиля пользователя: без зашитых имён пользователей и серверов (папку можно копировать куда угодно, в т. ч. на UNC).
LOCAL_APPDATA = os.environ.get("LOCALAPPDATA") or os.path.join(os.path.expanduser("~"), "AppData", "Local")
LOCAL_CFG = os.path.join(LOCAL_APPDATA, r"BatchRvt\batch_verification_schedules\run.cfg")

class UseDestinationTypes(IDuplicateTypeNamesHandler):
    def OnDuplicateTypeNamesFound(self, args):
        return DuplicateTypeAction.UseDestinationTypes


def get_script_dir():
    """RBP запускает файл через exec, поэтому __file__ часто нет; тогда — текущая папка."""
    try:
        folder = os.path.dirname(os.path.abspath(__file__))
        if folder and file_exists(os.path.join(folder, "insert_schedules.py")):
            return folder
    except Exception:
        pass
    return os.getcwd()



def read_text(path):
    """Чтение UTF-8. codecs utf-8-sig в IronPython падает на кириллице в пути."""
    from System.IO import StreamReader
    reader = StreamReader(path, True)
    try:
        text = reader.ReadToEnd()
    finally:
        reader.Close()
    if text is None:
        return u""
    return u"{}".format(text)


def read_cfg():
    data = {}
    # настройки прогона — только у текущего пользователя (LOCALAPPDATA): на общем компьютере файл рядом со скриптами был бы общим для всех
    candidates = [LOCAL_CFG]
    for path in candidates:
        if not path or not file_exists(path):
            continue
        for line in read_text(path).splitlines():
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, value = line.split("=", 1)
            data[lower_text(key.strip())] = value.strip()
        if data.get("source"):
            return data
    return data


def csv_cell(value):
    if value is None:
        return u""
    text = u"{}".format(value)
    return text.replace(u";", u",").replace(u"\r", u" ").replace(u"\n", u" ")


def append_csv(path, cells):
    if not path:
        return
    from System.IO import StreamWriter
    from System.Text import UTF8Encoding
    line = u";".join([csv_cell(c) for c in cells]) + u"\r\n"
    writer = StreamWriter(path, True, UTF8Encoding(False))
    try:
        writer.Write(line)
    finally:
        writer.Close()


# ---------- настройка проекта (отчеты\\проект\\распределение_<префикс моделей>.txt) ----------
# ЕДИНСТВЕННЫЙ источник «какие спецификации в какие модели/разделы» (вкладка «Распределение»). Парсер повторяет project_config.ps1.
def upper_text(value):
    """Верхний регистр БЕЗ str.upper(): в IronPython 2.7 оно меняет только ASCII."""
    try:
        make_char = unichr  # IronPython 2
    except NameError:
        make_char = chr
    out = []
    for ch in as_text(value):
        code = ord(ch)
        if 97 <= code <= 122 or 1072 <= code <= 1103:
            ch = make_char(code - 32)
        elif code == 1105:
            ch = make_char(1025)
        out.append(ch)
    return u"".join(out)


def project_model_name(path_or_name):
    text = as_text(path_or_name).strip().replace(u"\\", u"/")
    leaf = text.split(u"/")[-1]
    if lower_text(leaf).endswith(u".rvt"):
        leaf = leaf[:-4]
    return leaf


def strip_stage_prefix(name):
    """Префикс стадии («РД», «ПП», «П») перед пробелом или цифрой убирается; «Пров …» префиксом не является."""
    text = as_text(name).strip()
    for prefix in (u"РД", u"ПП", u"П"):
        if text.startswith(prefix) and len(text) > len(prefix):
            nxt = text[len(prefix)]
            if nxt.isspace() or nxt.isdigit():
                return text[len(prefix):].strip()
    return text


# Опечатки в именах спецификаций: (было, стало). «Cантехприборы» с латинской C -> «Сантехприборы» с кириллической С (регистр первой буквы сохраняется)
SPEC_NAME_FIXES = (
    (u"значния", u"значения"),
    (u"\u0043антехприборы", u"\u0421антехприборы"),
    (u"\u0063антехприборы", u"\u0441антехприборы"),
)


def canon_text(name):
    """Каноническое имя: исправляются известные опечатки (SPEC_NAME_FIXES)."""
    text = as_text(name)
    for old, new in SPEC_NAME_FIXES:
        text = text.replace(old, new)
    return text


def project_spec_key(name):
    text = canon_text(strip_stage_prefix(name))
    return u" ".join(lower_text(text).split())


# ЕДИНСТВЕННОЕ правило «раздел по имени файла»: код ищется только в виде «_КОД_» (регистр не важен, кириллица и латиница).
# Признаки IOS/ИОС/ВИС не используются (они есть и у сводной, и у TASK). Код раздела — латиница, как в файле настройки.
SECTION_NAMES = (
    (u"OV", (u"ОВ", u"OV")), (u"PD", (u"ПД", u"PD")), (u"VK", (u"ВК", u"VK")), (u"PT", (u"ПТ", u"PT")), (u"OT", (u"ОТ", u"OT")),
    (u"HS", (u"ХС", u"HS")), (u"ITP", (u"ИТП", u"ITP")), (u"EOM", (u"ЭОМ", u"EOM")), (u"SS", (u"СС", u"SS")),
    (u"AVT", (u"АВТ", u"AVT")), (u"HC", (u"ХЦ", u"HC")),
)

# Модели, по папке которых раздел не определяется (задание, базовые и т. п.)
NON_SECTION_MARKERS = (u"_TASK_", u"_AR_", u"_KR_", u"_TH_", u"_BASE_")


def section_code_from_folder(path):
    """Запасной вариант: раздел по имени папки, где лежит модель («…/АВТ/RVT/модель.rvt» -> AVT; папка RVT пропускается). Нет папки/кода — пусто."""
    parts = [p for p in as_text(path).strip().replace(u"\\", u"/").split(u"/") if p]
    if len(parts) < 2:
        return u""
    folders = parts[:-1]
    i = len(folders) - 1
    while i >= 0 and upper_text(folders[i]) == u"RVT":
        i -= 1
    if i < 0:
        return u""
    name = upper_text(folders[i])
    for sep in (u"_", u"-", u".", u"(", u")"):
        name = name.replace(sep, u" ")
    tokens = name.split()
    for code, variants in SECTION_NAMES:
        for variant in variants:
            if variant in tokens:
                return code
    return u""



def section_code_from_name(path_or_name):
    """Код раздела по имени файла («_ОВ_», «_vk_» …); сводная и модели без кода — пусто."""
    up = upper_text(project_model_name(path_or_name))
    if u"_SVOD_" in up or u"_СВОД_" in up or u"СВОДН" in up:
        return u""
    best, best_pos = u"", -1
    for code, variants in SECTION_NAMES:
        for variant in variants:
            pos = up.find(u"_" + variant + u"_")
            if pos >= 0 and (best_pos < 0 or pos < best_pos):
                best, best_pos = code, pos
    if not best and not any(marker in up for marker in NON_SECTION_MARKERS):
        return section_code_from_folder(path_or_name)
    return best


SECTION_TITLES = {u"OV": u"ОВ", u"PD": u"ПД", u"VK": u"ВК", u"PT": u"ПТ", u"OT": u"ОТ", u"HS": u"ХС", u"ITP": u"ИТП", u"EOM": u"ЭОМ", u"SS": u"СС", u"AVT": u"АВТ", u"HC": u"ХЦ"}


def _refresh_sections_from_file():
    """Подменяет SECTION_NAMES/SECTION_TITLES списком из разделы.txt. Нет файла или нет модуля — остаётся встроенный кортеж."""
    global SECTION_NAMES, SECTION_TITLES
    import sys
    try:
        here = os.path.dirname(os.path.abspath(__file__))
    except Exception:
        return
    if here not in sys.path:
        sys.path.insert(0, here)
    try:
        import bvs_sections
    except Exception:
        return
    rows, failed = bvs_sections.load_sections(os.path.join(here, u"разделы.txt"))
    if failed or not rows:
        try:
            log_msg("sections file unreadable, built-in list")
        except Exception:
            try:
                sys.stderr.write("sections file unreadable, built-in list\n")
            except Exception:
                pass
        return
    ordered = sorted(rows, key=lambda row: (row["column"], row["report"]))
    names = []
    titles = {}
    for row in ordered:
        names.append((row["code"], tuple(row["names"])))
        titles[row["code"]] = row["title"]
    SECTION_NAMES = tuple(names)
    SECTION_TITLES = titles


_refresh_sections_from_file()


def project_path_blocker(project_path):
    """Пустой или несуществующий путь настройки — запускать нельзя; возвращает понятную причину или пустую строку."""
    if not as_text(project_path).strip():
        return u"Не задан путь к файлу настройки проекта (project_path). Откройте вкладку «Распределение» и нажмите «Загрузить спецификации и модели»."
    if not file_exists(project_path):
        return u"Файл настройки проекта не найден: {}. Нажмите «Загрузить спецификации и модели» на вкладке «Распределение».".format(project_path)
    return u""


# ---------- LEGACY: разбор файла настройки проекта ----------
# Запасной путь. Обычно окно передаёт готовый план (plan_path в run.cfg, см. load_plan); эти функции нужны, когда плана нет
# (прогон из старого окна, план не записался или повреждён) и для «Посчитать элементы» (prefill). Логику не менять отдельно от
# Get-InsertPlanLines в project_config.ps1: тест test_insert_plan.py сравнивает оба пути.
REASON_UNMARKED = u"модель в распределении есть, но спецификации для неё не размечены"
REASON_NOSECTION = u"раздел не определён по имени (ожидается «_ОВ_», «_VK_» и т. п.)"
REASON_EMPTYSECTION = u"модели нет в распределении, а у раздела {} нет размеченных моделей"


def load_project(path):
    """LEGACY (запасной путь, см. выше). Файл настройки проекта -> словарь: stage, template, models [(имя, раздел по имени)], specs [имена], inproject (None — все | {ключ спеки}),
    cells {(модель, ключ спеки): (0|1, a|m|x, число)}."""
    proj = {"stage": u"", "template": u"", "models": [], "specs": [], "inproject": None, "summary": None, "cells": {}}
    section = u""
    for raw in read_text(path).splitlines():
        line = raw.rstrip()
        stripped = line.strip()
        if not stripped or stripped.startswith(u"#"):
            continue
        if stripped.startswith(u"[") and stripped.endswith(u"]"):
            section = stripped[1:-1]
            if section == u"inproject" and proj["inproject"] is None:
                proj["inproject"] = set()
            if section == u"summary" and proj["summary"] is None:
                proj["summary"] = set()
            continue
        if section == u"project":
            if u"=" in line:
                key, value = line.split(u"=", 1)
                key = lower_text(key.strip())
                if key in (u"stage", u"template"):
                    proj[key] = value.strip()
        elif section == u"models":
            name = line.split(u"\t")[0].strip()
            code = section_code_from_name(name)
            if code:
                proj["models"].append((name, code))
        elif section == u"specs":
            proj["specs"].append(stripped)
        elif section == u"inproject":
            proj["inproject"].add(project_spec_key(stripped))
        elif section == u"summary":
            proj["summary"].add(project_spec_key(stripped))
        elif section == u"cells":
            parts = line.split(u"\t")
            if len(parts) >= 4:
                try:
                    count = int(parts[4].strip()) if len(parts) > 4 and parts[4].strip() else -1
                    state = int(parts[2].strip())
                except ValueError:
                    continue
                proj["cells"][(lower_text(project_model_name(parts[0])), project_spec_key(parts[1]))] = (state, parts[3].strip(), count)
    return proj


def project_find_model(proj, path_or_name):
    key = lower_text(project_model_name(path_or_name))
    for name, sec in proj["models"]:
        if lower_text(name) == key:
            return name, sec
    return None


def project_spec_in(proj, spec):
    """Спецификация участвует в проекте (отмечена группа «В проекте»); в файле без раздела [inproject] — все."""
    return proj["inproject"] is None or project_spec_key(spec) in proj["inproject"]


def project_enabled_specs(proj, model_name):
    key = lower_text(project_model_name(model_name))
    return [s for s in proj["specs"]
            if project_spec_in(proj, s) and proj["cells"].get((key, project_spec_key(s)), (0, u"", -1))[0] == 1]


def project_section_specs(proj, code):
    models = [lower_text(n) for n, sec in proj["models"] if sec == code]
    return [s for s in proj["specs"]
            if project_spec_in(proj, s)
            and any(proj["cells"].get((m, project_spec_key(s)), (0, u"", -1))[0] == 1 for m in models)]


def project_summary_specs(proj):
    """Спецификации сводной: спецификации проекта, отмеченные в столбце «Сводная» (нет раздела [summary] — все спецификации проекта;
    нет и [inproject] — объединение отмеченных ячеек, как в файлах прежних версий)."""
    if proj["inproject"] is not None:
        return [s for s in proj["specs"]
                if project_spec_in(proj, s) and (proj["summary"] is None or project_spec_key(s) in proj["summary"])]
    models = [lower_text(n) for n, sec in proj["models"]]
    return [s for s in proj["specs"]
            if any(proj["cells"].get((m, project_spec_key(s)), (0, u"", -1))[0] == 1 for m in models)]


def project_model_specs(proj, path_or_name):
    """(список спецификаций | None, причина). Модель есть в настройке — её отметки; иначе объединение раздела по имени файла; иначе None."""
    found = project_find_model(proj, path_or_name)
    if found is not None:
        key = lower_text(found[0])
        if any((key, project_spec_key(s)) in proj["cells"] for s in proj["specs"]):
            return project_enabled_specs(proj, found[0]), u""
        return None, REASON_UNMARKED
    code = section_code_from_name(path_or_name)
    if not code:
        return None, REASON_NOSECTION
    specs = project_section_specs(proj, code)
    if not specs:
        return None, REASON_EMPTYSECTION.format(code)
    return specs, u""


# ---------- план вставки от окна (plan.txt в служебной папке прогона) ----------
PLAN_VERSION = u"1"


def load_plan(path):
    """plan.txt (пишет Get-InsertPlanLines в project_config.ps1) -> словарь или None (нет файла, другая версия, нет строки end с верным числом строк)."""
    if not path or not file_exists(path):
        return None
    try:
        text = read_text(path)
    except Exception:
        return None
    plan = {"stage": u"", "summary": [], "models": {}, "unmarked": {}, "sections": {}}
    version = None
    count = 0
    ended = False
    for raw in text.splitlines():
        line = raw.rstrip(u"\r\n")
        if not line.strip() or line.startswith(u"#"):
            continue
        parts = line.split(u"\t")
        kind = parts[0].strip()
        if kind == u"end":
            ended = len(parts) > 1 and parts[1].strip() == u"{}".format(count)
            break
        count += 1
        if kind == u"version" and len(parts) > 1:
            version = parts[1].strip()
        elif kind == u"stage":
            plan["stage"] = parts[1].strip() if len(parts) > 1 else u""
        elif kind == u"summary" and len(parts) > 1:
            plan["summary"].append(parts[1])
        elif kind == u"model" and len(parts) > 2:
            key = lower_text(project_model_name(parts[1]))
            if parts[2].strip() == u"unmarked":
                plan["unmarked"][key] = True
            else:
                plan["models"].setdefault(key, [])
        elif kind == u"mspec" and len(parts) > 2:
            plan["models"].setdefault(lower_text(project_model_name(parts[1])), []).append(parts[2])
        elif kind == u"sspec" and len(parts) > 2:
            plan["sections"].setdefault(parts[1].strip(), []).append(parts[2])
    if not ended or version != PLAN_VERSION:
        return None
    return plan


def plan_model_specs(plan, path_or_name):
    """То же, что project_model_specs, но по готовому плану: (список | None, причина)."""
    key = lower_text(project_model_name(path_or_name))
    if key in plan["unmarked"]:
        return None, REASON_UNMARKED
    if key in plan["models"]:
        return list(plan["models"][key]), u""
    code = section_code_from_name(path_or_name)
    if not code:
        return None, REASON_NOSECTION
    specs = plan["sections"].get(code) or []
    if not specs:
        return None, REASON_EMPTYSECTION.format(code)
    return list(specs), u""


def project_view_name(proj, spec_name):
    """Имя вида в Revit: <стадия> + 8 пробелов + название (папка «П» в диспетчере проекта задаётся пробелами)."""
    stage = proj.get("stage") or u""
    base = canon_text(strip_stage_prefix(spec_name))
    return (stage + u" " * 8 + base) if stage else base


def expand_profile(value, stage):
    """{СТАДИЯ}/{STAGE} в значении фильтра -> стадия проекта (регистр не важен)."""
    text = as_text(value)
    for token in (u"{стадия}", u"{stage}"):
        while True:
            pos = lower_text(text).find(token)
            if pos < 0:
                break
            text = text[:pos] + as_text(stage) + text[pos + len(token):]
    return text


def open_source(app, path):
    """Источник открывается отдельно и закрывается без сохранения."""
    try:
        from Autodesk.Revit.DB import OpenOptions, DetachFromCentralOption, ModelPathUtils
        model_path = ModelPathUtils.ConvertUserVisiblePathToModelPath(path)
        opts = OpenOptions()
        opts.DetachFromCentralOption = DetachFromCentralOption.DetachAndPreserveWorksets
        try:
            opts.Audit = False
        except Exception:
            pass
        return app.OpenDocumentFile(model_path, opts)
    except Exception as ex:
        Output(u"Open source detach failed, plain open: {}".format(ex))
        return app.OpenDocumentFile(path)


def index_schedules(document):
    found = {}
    collector = FilteredElementCollector(document).OfClass(ViewSchedule)
    for view in collector:
        try:
            if view.IsTemplate:
                continue
        except Exception:
            pass
        try:
            found[view.Name] = view
        except Exception:
            pass
    return found


def is_link_field(field):
    try:
        name = field.GetName() or u""
    except Exception:
        name = u""
    if u"Связанный файл" in name or u"RVT Link" in name or u"Linked file" in name:
        return True
    try:
        pid = field.ParameterId
        if pid is not None and pid.IntegerValue == int(BuiltInParameter.RVT_LINK_FILE_NAME_PARAM):
            return True
    except Exception:
        pass
    return False


def skip_field_type(field):
    try:
        kind = field.FieldType
    except Exception:
        return False
    return kind in (
        ScheduleFieldType.Count,
        ScheduleFieldType.Formula,
    )


def remove_link_setup(view):
    definition = view.Definition
    try:
        definition.IncludeLinkedFiles = False
    except Exception as ex:
        Output(u"WARN IncludeLinkedFiles: {}".format(ex))

    remove_ids = []
    count = definition.GetFieldCount()
    for index in range(count):
        field = definition.GetField(index)
        if is_link_field(field):
            remove_ids.append(field.FieldId)

    for field_id in remove_ids:
        try:
            index = definition.GetFilterCount() - 1
            while index >= 0:
                schedule_filter = definition.GetFilter(index)
                same = False
                try:
                    same = schedule_filter.FieldId == field_id
                except Exception:
                    same = False
                if same:
                    definition.RemoveFilter(index)
                index -= 1
        except Exception as ex:
            Output(u"WARN filter: {}".format(ex))
        try:
            definition.RemoveField(field_id)
        except Exception as ex:
            Output(u"WARN RemoveField: {}".format(ex))


def param_bound_to_category(document, parameter_name, category_id):
    if category_id is None or category_id.IntegerValue < 0:
        return True
    iterator = document.ParameterBindings.ForwardIterator()
    iterator.Reset()
    while iterator.MoveNext():
        definition = iterator.Key
        binding = iterator.Current
        try:
            def_name = definition.Name
        except Exception:
            continue
        if def_name != parameter_name:
            continue
        try:
            if binding.Categories.Contains(document.Settings.Categories.get_Item(category_id)):
                return True
        except Exception:
            try:
                for category in binding.Categories:
                    if category.Id == category_id:
                        return True
            except Exception:
                return True
        return False
    return False


def audit_schedule(document, view, model_path, params_csv):
    definition = view.Definition
    try:
        category_id = definition.CategoryId
    except Exception:
        category_id = None
    holes = 0
    count = definition.GetFieldCount()
    for index in range(count):
        field = definition.GetField(index)
        if skip_field_type(field) or is_link_field(field):
            continue
        try:
            pid = field.ParameterId
        except Exception:
            continue
        if pid is None or pid.IntegerValue < 0:
            continue
        try:
            field_name = field.GetName()
        except Exception:
            field_name = u""
        element = document.GetElement(pid)
        if element is None:
            holes += 1
            append_csv(
                params_csv,
                [model_path, view.Name, field_name, u"нет параметра в модели"],
            )
            continue
        try:
            single = category_id is not None and category_id.IntegerValue > 0
        except Exception:
            single = False
        if single and not param_bound_to_category(document, field_name, category_id):
            holes += 1
            append_csv(
                params_csv,
                [
                    model_path,
                    view.Name,
                    field_name,
                    u"параметр не привязан к категории спецификации",
                ],
            )
    return holes


def copy_options():
    options = CopyPasteOptions()
    try:
        options.SetDuplicateTypeNamesHandler(UseDestinationTypes())
    except Exception as ex:
        Output(u"WARN duplicate handler: {}".format(ex))
    return options


def as_text(value):
    if value is None:
        return u""
    try:
        return unicode(value)
    except Exception:
        return u"{}".format(value)


def set_text_param(view, param_name, value):
    """Записывает текстовый параметр вида. Папки диспетчера — это значения таких параметров."""
    from Autodesk.Revit.DB import StorageType
    for param in view.Parameters:
        try:
            pname = as_text(param.Definition.Name)
        except Exception:
            continue
        if pname != param_name:
            continue
        if param.IsReadOnly:
            raise Exception(u"параметр «{}» только для чтения".format(pname))
        if param.StorageType != StorageType.String:
            raise Exception(u"параметр «{}» не текстовый".format(pname))
        param.Set(value)
        return
    raise Exception(u"нет параметра «{}»".format(param_name))


def set_view_purpose(view):
    """Верхняя папка «_Проверка проекта». Папка «П» задаётся пробелами в имени вида."""
    set_text_param(view, u"ADSK_Назначение вида", u"_Проверка проекта")


def live_status_path():
    base = os.environ.get("LOCALAPPDATA", "")
    if not base:
        return u""
    return os.path.join(base, "BatchRvt", "batch_verification_schedules", "live_status.txt")


def model_position(target_path):
    list_path = os.path.join(os.path.dirname(LOCAL_CFG), "rvt_list_run.txt")
    names = []
    try:
        for line in read_text(list_path).splitlines():
            text = line.strip()
            if text and not text.startswith(u"#"):
                names.append(text)
    except Exception:
        names = []
    total = len(names) or 1
    current = norm_path(target_path)
    index = 1
    for i, item in enumerate(names):
        if norm_path(item) == current:
            index = i + 1
            break
    return index, total


def write_live_status(target_path, inserted, skipped, holes):
    path = live_status_path()
    if not path:
        return
    folder = os.path.dirname(path)
    try:
        ensure_dir(folder)
    except Exception:
        return
    index, total = model_position(target_path)
    from System.IO import StreamWriter
    from System.Text import UTF8Encoding
    lines = [
        u"index={}".format(index),
        u"total={}".format(total),
        u"name={}".format(os.path.basename(as_text(target_path))),
        u"inserted={}".format(inserted),
        u"skipped={}".format(skipped),
        u"params={}".format(holes),
    ]
    try:
        writer = StreamWriter(path, False, UTF8Encoding(True))
        try:
            writer.Write(u"\r\n".join(lines) + u"\r\n")
        finally:
            writer.Close()
    except Exception as ex:
        Output(u"WARN live status: {}".format(ex))


MODE_LABELS = {u"initial": u"первичная", u"add_diff": u"добавление", u"replace": u"замена"}


def normalize_insert_mode(value):
    mode = lower_text(as_text(value).strip())
    if mode in (u"first", u"first_load"):
        mode = u"initial"
    elif mode in (u"add_missing", u"add_new_tails", u"add_all"):
        mode = u"add_diff"
    if mode in (u"initial", u"replace", u"add_diff"):
        return mode
    return u"initial"


def model_has_tail(doc_index, tail, extra_prefixes):
    if not tail:
        return False
    for existing_name in doc_index.keys():
        if schedule_tail(existing_name, extra_prefixes) == tail:
            return True
    return False


def model_has_variant(doc_index, placed_name, extra_prefixes):
    """В модели уже есть та же спецификация под другим написанием («Cантехприборы» с латинской C при вставке «Сантехприборы») — дубль не нужен."""
    tail = schedule_tail(placed_name, extra_prefixes)
    raw = schedule_tail(placed_name, extra_prefixes, False)
    if not tail:
        return None
    for existing_name in doc_index.keys():
        if schedule_tail(existing_name, extra_prefixes) == tail and schedule_tail(existing_name, extra_prefixes, False) != raw:
            return existing_name
    return None


def canon_spec(name):
    """Опечатки «Без значния», «Cантехприборы» (латинская C) = верное написание (в списке и в моделях имена могут отличаться)."""
    return canon_text(name)


def schedule_tail(name, extra=None, canon=True):
    """Хвост имени после заголовка стадии и пробелов: «0_Ошибка типоразмера». canon=False — без исправления опечаток."""
    text = (canon_spec(name) if canon else as_text(name)).strip()
    prefixes = [u"РД", u"ПП", u"П"]
    if extra:
        for item in extra:
            item = as_text(item).strip()
            if item and item not in prefixes:
                prefixes.append(item)
    prefixes.sort(key=lambda item: len(item), reverse=True)
    for prefix in prefixes:
        if text.startswith(prefix):
            text = text[len(prefix):]
            break
    return text.lstrip(u" ")


def apply_prefix(name, prefix_from, prefix_to):
    text = as_text(name)
    src = as_text(prefix_from).strip()
    dst = as_text(prefix_to).strip()
    if not src or not dst or src == dst:
        return text
    if text.startswith(src):
        return dst + text[len(src):]
    return text


def delete_schedules(document, names):
    """Удаление отдельной операцией, чтобы имя освободилось до копирования."""
    pending = []
    index = index_schedules(document)
    for name in names:
        view = index.get(name)
        if view is not None:
            pending.append((name, view.Id))
    if not pending:
        return 0, []
    deleted = 0
    failed = []
    transaction = Transaction(document, "Delete verification schedules")
    transaction.Start()
    try:
        for name, element_id in pending:
            try:
                document.Delete(element_id)
                deleted += 1
            except Exception as ex:
                failed.append((name, u"{}".format(ex)))
        transaction.Commit()
    except Exception:
        try:
            transaction.RollBack()
        except Exception:
            pass
        raise
    still = index_schedules(document)
    confirmed = 0
    confirmed_failed = []
    for name, element_id in pending:
        if name in still:
            reason = u"не удалилась"
            for failed_name, failed_reason in failed:
                if failed_name == name:
                    reason = failed_reason
                    break
            confirmed_failed.append((name, reason))
        else:
            confirmed += 1
    return confirmed, confirmed_failed


def insert_one(document, source, source_view, target_name=None):
    target_name = as_text(target_name) if target_name else as_text(source_view.Name)
    options = copy_options()
    transaction = Transaction(document, "Insert verification schedule")
    transaction.Start()
    try:
        ids = List[ElementId]()
        ids.Add(source_view.Id)
        copied = ElementTransformUtils.CopyElements(
            source, ids, document, Transform.Identity, options
        )
        if copied is None or copied.Count == 0:
            raise Exception(u"CopyElements вернул пустой список")
        new_id = None
        for element_id in copied:
            new_id = element_id
            break
        view = document.GetElement(new_id)
        if view is None:
            raise Exception(u"скопированный вид не найден")
        set_view_purpose(view)
        if as_text(view.Name) != target_name:
            view.Name = target_name
        if as_text(view.Name) != target_name:
            raise Exception(u"имя стало «{}»".format(as_text(view.Name)))
        transaction.Commit()
        return view
    except Exception:
        try:
            transaction.RollBack()
        except Exception:
            pass
        raise


def is_rsn_path(path):
    if not path:
        return False
    return path.upper().startswith("RSN://")


def clear_readonly(path):
    if is_rsn_path(path):
        return
    try:
        import System
        attrs = System.IO.File.GetAttributes(path)
        readonly = System.IO.FileAttributes.ReadOnly
        if (attrs & readonly) == readonly:
            System.IO.File.SetAttributes(path, attrs & ~readonly)
    except Exception as ex:
        Output(u"WARN ReadOnly: {}".format(ex))


def make_relinquish_options():
    from Autodesk.Revit.DB import RelinquishOptions
    opts = RelinquishOptions(True)
    try:
        opts.StandardWorksets = True
        opts.ViewWorksets = True
        opts.FamilyWorksets = True
        opts.UserWorksets = True
        opts.CheckedOutElements = True
    except Exception:
        pass
    return opts


def save_or_sync(document, path):
    from Autodesk.Revit.DB import (
        SynchronizeWithCentralOptions,
        TransactWithCentralOptions,
        SaveAsOptions,
        WorksharingSaveAsOptions,
        WorksharingUtils,
        ModelPathUtils,
    )

    is_detached = False
    try:
        is_detached = document.IsDetached
    except Exception:
        is_detached = False

    workshared_live = (
        document.IsWorkshared
        and (not is_detached)
        and document.GetWorksharingCentralModelPath() is not None
    )
    if is_rsn_path(path) or workshared_live:
        if is_detached:
            raise Exception(u"Файл отсоединён от хранилища. В RBP: Create New Local, Detach = OFF.")
        sync_opts = SynchronizeWithCentralOptions()
        sync_opts.SetRelinquishOptions(make_relinquish_options())
        try:
            sync_opts.Comment = u"Проверочные спецификации"
            sync_opts.SaveLocalFile = True
        except Exception:
            pass
        try:
            central = ModelPathUtils.ConvertModelPathToUserVisiblePath(
                document.GetWorksharingCentralModelPath()
            )
            Output(u"Sync: {}".format(central))
        except Exception:
            Output(u"Sync")
        document.SynchronizeWithCentral(TransactWithCentralOptions(), sync_opts)
        Output(u"SynchronizedWithCentral")
        return

    clear_readonly(path)
    save_opts = SaveAsOptions()
    save_opts.OverwriteExistingFile = True
    if document.IsWorkshared:
        workset_opts = WorksharingSaveAsOptions()
        workset_opts.SaveAsCentral = True
        save_opts.SetWorksharingOptions(workset_opts)
    document.SaveAs(path, save_opts)
    try:
        if document.IsWorkshared:
            WorksharingUtils.RelinquishOwnership(
                document, make_relinquish_options(), TransactWithCentralOptions()
            )
    except Exception as ex:
        Output(u"WARN relinquish: {}".format(ex))
    Output(u"SavedAs")


def apply_profile(view, stage):
    """Подстановка {СТАДИЯ} в строковые значения фильтров вставленной спецификации. Возвращает число изменённых фильтров."""
    definition = view.Definition
    changed = 0
    for index in range(definition.GetFilterCount()):
        flt = definition.GetFilter(index)
        try:
            value = flt.GetStringValue()
        except Exception:
            continue
        if value is None:
            continue
        new_value = expand_profile(value, stage)
        if new_value != as_text(value):
            definition.SetFilter(index, ScheduleFilter(flt.FieldId, flt.FilterType, new_value))
            changed += 1
    return changed


def source_index_by_key(source_index):
    """Спецификации шаблона по ключу без префикса стадии: имена шаблона могут быть без префикса или со старым «П        »."""
    result = {}
    for name, view in source_index.items():
        result.setdefault(project_spec_key(name), view)
    return result


def count_elements(document, view):
    """Число элементов спецификации: FilteredElementCollector по виду спецификации; если он не сработал — строки таблицы. (число, способ)."""
    try:
        return int(FilteredElementCollector(document, view.Id).GetElementCount()), u"collector"
    except Exception:
        pass
    try:
        body = view.GetTableData().GetSectionData(SectionType.Body)
        return max(0, int(body.NumberOfRows) - 1), u"table"
    except Exception as ex:
        raise Exception(u"не удалось посчитать элементы: {}".format(ex))


def append_tsv(path, cells):
    if not path:
        return
    from System.IO import StreamWriter
    from System.Text import UTF8Encoding
    line = u"\t".join([as_text(c).replace(u"\t", u" ").replace(u"\r", u" ").replace(u"\n", u" ") for c in cells]) + u"\r\n"
    writer = StreamWriter(path, True, UTF8Encoding(False))
    try:
        writer.Write(line)
    finally:
        writer.Close()


def insert_batch(document, source, source_views):
    """Все спецификации одной пачкой (один CopyElements в одной транзакции). Возвращает новые виды в порядке исходных.
    Любое несоответствие (число, имена) — исключение: транзакция откатывается, вызывающий переходит на вставку по одной."""
    options = copy_options()
    transaction = Transaction(document, "Insert verification schedules (batch)")
    transaction.Start()
    try:
        ids = List[ElementId]()
        for view in source_views:
            ids.Add(view.Id)
        copied = ElementTransformUtils.CopyElements(source, ids, document, Transform.Identity, options)
        if copied is None or copied.Count != len(source_views):
            raise Exception(u"CopyElements вернул {} элементов из {}".format(0 if copied is None else copied.Count, len(source_views)))
        views = []
        for number, element_id in enumerate(copied):
            view = document.GetElement(element_id)
            if view is None:
                raise Exception(u"скопированный вид не найден")
            if not as_text(view.Name).startswith(as_text(source_views[number].Name)):
                raise Exception(u"порядок скопированных видов не совпал с порядком шаблона («{}» вместо «{}»)".format(as_text(view.Name), as_text(source_views[number].Name)))
            set_view_purpose(view)
            view.Name = u"__bvs_prefill_{}".format(number)
            views.append(view)
        transaction.Commit()
        return views
    except Exception:
        try:
            transaction.RollBack()
        except Exception:
            pass
        raise


def adjust_prefill_view(document, view, stage):
    """Снятие связей и подстановка стадии в только что вставленной пробной спецификации (отдельная транзакция)."""
    fix = Transaction(document, "Prefill adjust")
    fix.Start()
    try:
        remove_link_setup(view)
        apply_profile(view, stage)
        fix.Commit()
    except Exception:
        try:
            fix.RollBack()
        except Exception:
            pass
        raise


def prefill_model(document, source, spec_names, proj, out_path, target_path, open_seconds=0.0):
    """Пробная вставка спецификаций шаблона в модель и подсчёт элементов; всё откатывается (одна TransactionGroup), модель не сохраняется.
    Все спецификации копируются одной пачкой; не получилось — вставка по одной (причина в логе). В out_path (TSV, дописывается): модель,
    спецификация, число (-1 — не удалось), секунды, способ/ошибка; в конце строка «#МОДЕЛЬ»: модель, секунды, спеки, ошибки, способы подсчёта,
    «пачка»/«по одной», секунды: открытие источника, вставка, профиль, подсчёт, откат."""
    import time
    model_name = project_model_name(target_path)
    keys = source_index_by_key(index_schedules(source))
    stage = proj.get("stage") or u""
    started = time.time()
    errors = 0
    done = 0
    methods = set()
    t_paste = t_profile = t_count = t_rollback = 0.0
    found = []
    for number, name in enumerate(spec_names):
        view = keys.get(project_spec_key(name))
        if view is None:
            errors += 1
            done += 1
            append_tsv(out_path, [model_name, name, -1, u"0.00", u"нет в шаблоне"])
        else:
            found.append((name, view))
    group = TransactionGroup(document, u"Распределение: пробная вставка (откат)")
    group.Start()
    mode_text = u"пачка"
    try:
        views = None
        if found:
            t0 = time.time()
            try:
                views = insert_batch(document, source, [v for _, v in found])
            except Exception as ex:
                mode_text = u"по одной"
                log_msg(u"WARN prefill {}: batch insert failed, falling back to one-by-one: {}".format(model_name, ex))
            t_paste += time.time() - t0
        for number, (name, source_view) in enumerate(found):
            count, note = -1, u""
            t0 = time.time()
            try:
                if views is not None:
                    fresh = views[number]
                else:
                    fresh = insert_one(document, source, source_view, u"__bvs_prefill_{}".format(number))
                    t_paste += time.time() - t0
                t1 = time.time()
                adjust_prefill_view(document, fresh, stage)
                t_profile += time.time() - t1
                t2 = time.time()
                count, method = count_elements(document, fresh)
                t_count += time.time() - t2
                methods.add(method)
                note = method
            except Exception as ex:
                count = -1
                note = u"{}".format(ex)
            if count < 0:
                errors += 1
            done += 1
            append_tsv(out_path, [model_name, name, count, u"{:.2f}".format(time.time() - t0), note])
            write_live_status(target_path, done, errors, 0)
    finally:
        t0 = time.time()
        try:
            group.RollBack()
        except Exception as ex:
            log_msg(u"WARN rollback prefill: {}".format(ex))
        t_rollback = time.time() - t0
    seconds = time.time() - started
    append_tsv(out_path, [u"#МОДЕЛЬ", model_name, u"{:.1f}".format(seconds), done, errors, u"/".join(sorted(methods)), mode_text,
                          u"{:.1f}".format(open_seconds), u"{:.1f}".format(t_paste), u"{:.1f}".format(t_profile), u"{:.1f}".format(t_count),
                          u"{:.1f}".format(t_rollback)])
    log_msg(u"prefill {} mode={} specs={} errors={} open_source={:.1f}s paste={:.1f}s profile={:.1f}s count={:.1f}s rollback={:.1f}s total={:.1f}s".format(
        model_name, u"batch" if mode_text == u"пачка" else u"one-by-one", done, errors, open_seconds, t_paste, t_profile, t_count, t_rollback, seconds))
    return seconds, done, errors


def main():
    document = revit_script_util.GetScriptDocument()
    target_path = revit_script_util.GetRevitFilePath()
    cfg = read_cfg()
    source_path = cfg.get("source", u"")
    report_dir = cfg.get("report_dir", u"")
    run_id = cfg.get("run_id", u"run")
    summary_csv = os.path.join(report_dir, u"{}_сводка.csv".format(run_id)) if report_dir else u""
    params_csv = os.path.join(report_dir, u"{}_параметры.csv".format(run_id)) if report_dir else u""
    skips_csv = os.path.join(report_dir, u"{}_пропуски.csv".format(run_id)) if report_dir else u""
    inserted_csv = os.path.join(report_dir, u"{}_вставлено.csv".format(run_id)) if report_dir else u""

    mode = cfg.get("mode", u"")
    project_path = cfg.get("project_path", u"")
    proj = None
    code, label = u"", u"не распознан"
    log_msg(u"insert_schedules start")
    log_msg(u"section_code={}".format(code or u"?"))
    deleted = 0
    inserted = 0
    skipped = 0
    failed = 0          # вставки/удаления, которые не удались (в «уже в пакете» не входят): в режиме replace при failed > 0 модель не сохраняется
    holes = 0
    error_text = u""
    mode_label = u""
    source = None

    try:
        write_live_status(target_path, 0, 0, 0)
        if not source_path or not file_exists(source_path):
            raise Exception(u"Не найден файл со спецификациями: {}".format(source_path))
        if norm_path(source_path) == norm_path(target_path):
            raise Exception(u"Файл со спецификациями совпал с моделью, копирование пропущено")

        blocker = project_path_blocker(project_path)
        if blocker:
            raise Exception(blocker)
        plan = None
        plan_path = cfg.get("plan_path", u"")
        if mode != u"prefill":
            plan = load_plan(plan_path)
        if plan is not None:
            proj = {"stage": plan["stage"]}          # стадия для имён видов и профиля; остальное — в плане
            log_msg(u"plan=used {}".format(plan_path))
        else:
            # LEGACY fallback: плана нет (старое окно, «Посчитать элементы») или он повреждён — разбор файла настройки, как раньше
            if plan_path:
                log_msg(u"plan=invalid {}; LEGACY fallback: файл настройки".format(plan_path))
            elif mode != u"prefill":
                log_msg(u"plan=none; LEGACY fallback: файл настройки")
            proj = load_project(project_path)
        code = section_code_from_name(target_path)
        label = SECTION_TITLES.get(code, code) if code else u"не распознан"
        log_msg(u"section_code={}".format(code or u"?"))
        prefix_from = cfg.get("prefix_from", u"")
        prefix_to = cfg.get("prefix_to", u"")
        keep_links = False
        if mode == u"prefill":
            # пробная вставка всех спецификаций шаблона и подсчёт элементов; модель не сохраняется
            out_path = cfg.get("prefill_out", u"")
            if not out_path:
                raise Exception(u"Не задан prefill_out")
            import time
            opened = time.time()
            source = open_source(document.Application, source_path)
            names = [n for n in proj["specs"] if project_spec_in(proj, n)]      # только спецификации групп «В проекте»
            if not names:
                raise Exception(u"В распределении нет отмеченных групп «В проекте»")
            if not section_code_from_name(target_path):
                raise Exception(u"Модель без кода раздела в имени в распределении не участвует")
            prefill_model(document, source, names, proj, out_path, target_path, time.time() - opened)
            return
        if mode == u"summary":
            code, label = u"SVOD", u"Сводная"
            raw_names = plan["summary"] if plan is not None else project_summary_specs(proj)
            keep_links = True
            if not raw_names:
                raise Exception(u"В распределении нет отмеченных спецификаций")
        else:
            if plan is not None:
                raw_names, why = plan_model_specs(plan, target_path)
            else:
                raw_names, why = project_model_specs(proj, target_path)
            if raw_names is None:
                raise Exception(u"Набор спецификаций модели не определён: {}".format(why))

        pairs = []
        for raw_name in raw_names:
            if proj.get("stage"):
                pairs.append((raw_name, project_view_name(proj, raw_name)))
            else:
                pairs.append((raw_name, apply_prefix(raw_name, prefix_from, prefix_to)))
        log_msg(u"schedules_in_list={}".format(len(pairs)))

        insert_mode = normalize_insert_mode(cfg.get("insert_mode", u"initial"))
        log_msg(u"insert_mode={}".format(insert_mode))

        extra_prefixes = [prefix_from, prefix_to]
        tails = []
        for source_name, placed_name in pairs:
            for variant in (source_name, placed_name):
                tail = schedule_tail(variant, extra_prefixes)
                if tail and tail not in tails:
                    tails.append(tail)
        # модель, в которой спецификаций списка ещё нет, получает первичную загрузку независимо от выбранного режима (удалять и сравнивать нечего)
        primary = False
        if insert_mode in (u"add_diff", u"replace"):
            existing_now = index_schedules(document)
            if not any(model_has_tail(existing_now, t, extra_prefixes) for t in tails):
                insert_mode = u"initial"
                primary = True
                log_msg(u"insert_mode=initial (auto: no project schedules in model yet)")
        mode_label = MODE_LABELS.get(insert_mode, u"")
        if primary:
            mode_label = u"первичная (автоматически)"
        deleted = 0
        if insert_mode == u"replace":
            tail_set = {}
            for item in tails:
                tail_set[item] = True
            old_names = []
            for existing_name in index_schedules(document).keys():
                if schedule_tail(existing_name, extra_prefixes) in tail_set:
                    old_names.append(existing_name)
            deleted, delete_failed = delete_schedules(document, old_names)
            for failed_name, failed_reason in delete_failed:
                skipped += 1
                failed += 1
                append_csv(skips_csv, [target_path, failed_name, u"не удалена: {}".format(failed_reason)])
            log_msg(u"deleted_old={}".format(deleted))
        else:
            log_msg(u"deleted_old=0")

        app = document.Application
        source = open_source(app, source_path)
        source_index = index_schedules(source)
        source_keys = source_index_by_key(source_index)

        for source_name, placed_name in pairs:
            doc_index = index_schedules(document)
            if insert_mode == u"add_diff":
                placed_tail = schedule_tail(placed_name, extra_prefixes)
                if model_has_tail(doc_index, placed_tail, extra_prefixes):
                    skipped += 1
                    append_csv(
                        skips_csv,
                        [target_path, placed_name, u"уже в пакете модели ({})".format(placed_tail)],
                    )
                    write_live_status(target_path, inserted, skipped, holes)
                    continue
            else:
                variant = model_has_variant(doc_index, placed_name, extra_prefixes)
                if variant is not None:
                    skipped += 1
                    append_csv(skips_csv, [target_path, placed_name, u"уже есть под другим написанием ({})".format(variant)])
                    write_live_status(target_path, inserted, skipped, holes)
                    continue
            source_view = source_index.get(source_name)
            if source_view is None:
                # шаблон без префикса стадии или со старым «П        », со старой опечаткой «Без значния» — ищем по ключу без префикса
                source_view = source_keys.get(project_spec_key(source_name))
            if source_view is None:
                skipped += 1
                failed += 1
                append_csv(skips_csv, [target_path, placed_name, u"нет в файле со спецификациями"])
                write_live_status(target_path, inserted, skipped, holes)
                continue
            try:
                fresh = insert_one(document, source, source_view)
            except Exception as ex:
                skipped += 1
                failed += 1
                append_csv(skips_csv, [target_path, placed_name, u"{}".format(ex)])
                write_live_status(target_path, inserted, skipped, holes)
                continue
            if fresh is None:
                skipped += 1
                failed += 1
                append_csv(skips_csv, [target_path, placed_name, u"после копирования вид пустой"])
                write_live_status(target_path, inserted, skipped, holes)
                continue
            fix = Transaction(document, "Adjust verification schedule")
            fix.Start()
            try:
                if as_text(fresh.Name) != placed_name:
                    fresh.Name = placed_name
                if not keep_links:
                    remove_link_setup(fresh)
                apply_profile(fresh, proj.get("stage") or u"")
                fix.Commit()
            except Exception as ex:
                try:
                    fix.RollBack()
                except Exception:
                    pass
                failed += 1
                append_csv(
                    skips_csv,
                    [target_path, placed_name, u"вставлена, но не настроена: {}".format(ex)],
                )
            if as_text(fresh.Name) != placed_name:
                skipped += 1
                failed += 1
                append_csv(
                    skips_csv,
                    [target_path, placed_name, u"после копирования имя «{}»".format(as_text(fresh.Name))],
                )
                write_live_status(target_path, inserted, skipped, holes)
                continue
            holes += audit_schedule(document, fresh, target_path, params_csv)
            append_csv(
                inserted_csv,
                [target_path, label if code else u"не распознан", as_text(fresh.Name)],
            )
            inserted += 1
            write_live_status(target_path, inserted, skipped, holes)

        if insert_mode == u"replace" and failed:
            # старые спецификации уже удалены в памяти, а новые вставились не все: сохранение оставило бы неполный комплект — модель не сохраняется
            error_text = u"Модель не сохранена: вставилось не всё (не удалось {}), прежние спецификации остались".format(failed)
            log_msg(u"replace: not saved, failed={}".format(failed))
        elif deleted or inserted:
            save_or_sync(document, target_path)
        else:
            log_msg(u"No changes, skip save")
    except Exception as ex:
        error_text = u"{}".format(ex)
        log_msg(u"ERROR (see summary csv)")
    finally:
        if source is not None:
            try:
                source.Close(False)
            except Exception as ex:
                Output(u"WARN close source: {}".format(ex))

    if mode == u"prefill":
        if error_text:
            append_tsv(cfg.get("prefill_out", u""), [u"#ОШИБКА", project_model_name(target_path), error_text])
        return
    append_csv(
        summary_csv,
        [target_path, label if code else u"не распознан", deleted, inserted, skipped, error_text, mode_label],
    )
    log_msg(
        u"Done deleted={} inserted={} skipped={}".format(
            deleted, inserted, skipped
        )
    )


main()
