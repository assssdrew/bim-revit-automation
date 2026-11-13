# -*- coding: utf-8 -*-
"""
Revit Batch Processor — файл «Код по классификатору» (Assembly Code) в моделях разделов.

Окно Revit: Управление → Дополнительные параметры → Код по классификатору.
Путь записывается как абсолютный (в диалоге — «Абсолютные»), один UNC на все модели.
Относительная копия рядом с моделью (UniformatClassifications.txt) не используется.

Как запустить для PROJ
  1. Скопировать папку служебное\\скрипты\\ на рабочий ПК (этот файл рядом с insert_schedules.py).
  2. Открыть «Проверочные спецификации» (Проверочные спецификации.vbs).
  3. В списке проектов выбрать PROJ.
  4. Нажать «Загрузить спецификации и модели» — тот же список, что открывает «В разделы».
  5. Нажать «Классификатор».
     Перед запуском Revit окно проверяет, что файл классификатора есть и читается.
     Если файла нет, модели не открываются.
     BatchRvt: Create New Local, рабочие наборы Close All.
     Сводная и служебные модели в список не входят.
     Если в модели уже стоит этот абсолютный UNC — файл не сохраняется и с централом не синхронизируется.
  6. Отчёт: отчеты\\<проект>\\<дата>_классификатор.xlsx
     Колонки: Модель, Прежний путь, Новый путь, Результат, Ошибка.
     Результат: updated / already_set / skipped_not_section / error.

Путь классификатора и проект задаёт окно: из настройки проекта (classifier= в распределение_<проект>.txt,
для PROJ по умолчанию — официальный путь) в run.cfg строками classifier_path= и classifier_project=.
Нет этих строк (старое окно) — берутся CLASSIFIER_PATH и CLASSIFIER_PROJECT ниже.
Путь — абсолютный UNC (начинается с \\\\) . Рядом с моделями файл не копировать.

API Revit 2024 (проверено по справке, не Load(string)):
  AssemblyCodeTable.GetAssemblyCodeTable(doc)
  ExternalResourceReference.CreateLocalResource(..., PathType.Absolute)
  KeyBasedTreeEntryTable.LoadFrom(reference, KeyBasedTreeEntriesLoadResults) внутри Transaction
  Проверка: ExternalFileReference.PathType == Absolute и GetPath() совпадает с UNC.
  У KeyBasedTreeEntryTable в 2024 нет метода Load(string) — есть только LoadFrom.
  Спецификации, фильтры и прочие настройки не меняются.

run.cfg (пишет окно, %LOCALAPPDATA%\\BatchRvt\\batch_verification_schedules\\run.cfg):
  report_dir=папка прогона
  run_id=метка прогона
  classifier_path=UNC-путь классификатора проекта (нет — CLASSIFIER_PATH)
  classifier_project=префикс моделей проекта (нет — CLASSIFIER_PROJECT)
"""

import clr
import os

clr.AddReference("RevitAPI")


# Запасные значения, если окно не передало classifier_path / classifier_project в run.cfg.
# Путь в этой копии — заглушка, не сетевой адрес.
CLASSIFIER_PROJECT = u"PROJ"
CLASSIFIER_PATH = u"\\\\server\\share\\BIM\\classifier.txt"

REPORT_HEADER = u"Модель;Прежний путь;Новый путь;Результат;Ошибка"

LOCAL_APPDATA = os.environ.get("LOCALAPPDATA") or os.path.join(os.path.expanduser("~"), "AppData", "Local")
LOCAL_CFG = os.path.join(LOCAL_APPDATA, r"BatchRvt\batch_verification_schedules\run.cfg")

# Код раздела — как в insert_schedules.py / project_config.ps1. IOS и ORG в имени раздела не отменяют:
# они есть и у моделей разделов (PROJ_IOS_OV_ORG_…).
SECTION_NAMES = (
    (u"OV", (u"ОВ", u"OV")), (u"PD", (u"ПД", u"PD")), (u"VK", (u"ВК", u"VK")), (u"PT", (u"ПТ", u"PT")), (u"OT", (u"ОТ", u"OT")),
    (u"HS", (u"ХС", u"HS")), (u"ITP", (u"ИТП", u"ITP")), (u"EOM", (u"ЭОМ", u"EOM")), (u"SS", (u"СС", u"SS")),
    (u"AVT", (u"АВТ", u"AVT")), (u"HC", (u"ХЦ", u"HC")),
)
NON_SECTION_MARKERS = (u"_TASK_", u"_AR_", u"_KR_", u"_TH_", u"_BASE_")


def _refresh_sections_from_file():
    """SECTION_NAMES из разделы.txt. Импорт соседнего модуля не удался или файла нет — остаётся кортеж выше."""
    global SECTION_NAMES
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
            sys.stderr.write("sections file unreadable, built-in list\n")
        except Exception:
            pass
        return
    ordered = sorted(rows, key=lambda row: (row["column"], row["report"]))
    SECTION_NAMES = tuple((row["code"], tuple(row["names"])) for row in ordered)


_refresh_sections_from_file()


def as_text(value):
    if value is None:
        return u""
    return u"{}".format(value)


def lower_text(value):
    """Нижний регистр без str.lower(): в IronPython 2.7 оно не трогает кириллицу."""
    try:
        make_char = unichr
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


def upper_text(value):
    """Верхний регистр без str.upper()."""
    try:
        make_char = unichr
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


def log_msg(text):
    try:
        import revit_script_util
        output = revit_script_util.Output
    except Exception:
        return
    raw = as_text(text)
    safe = u"".join((c if ord(c) < 128 else u"?") for c in raw)
    if len(safe) > 180:
        safe = safe[:180]
    try:
        output(safe)
    except Exception:
        pass


def file_exists(path):
    """.NET: os.path на IronPython 2.7 врёт на кириллице в пути."""
    try:
        from System.IO import File
        return bool(File.Exists(path))
    except ImportError:
        return os.path.isfile(path)


def ensure_dir(path):
    if not path:
        return
    try:
        from System.IO import Directory
        Directory.CreateDirectory(path)
        return
    except ImportError:
        pass
    except Exception:
        return
    if not os.path.isdir(path):
        os.makedirs(path)


def norm_path(path):
    """Сравнение путей Windows: слэши, регистр (включая кириллицу), лишние \\\\ в середине, префикс \\\\?\\UNC\\."""
    text = as_text(path).strip().replace(u"/", u"\\")
    folded = lower_text(text)
    if folded.startswith(u"\\\\?\\unc\\"):
        text = u"\\\\" + text[8:]
    elif text.startswith(u"\\\\?\\"):
        text = text[4:]
    if text.startswith(u"\\\\"):
        rest = text[2:]
        while u"\\\\" in rest:
            rest = rest.replace(u"\\\\", u"\\")
        text = u"\\\\" + rest
    else:
        while u"\\\\" in text:
            text = text.replace(u"\\\\", u"\\")
    if len(text) > 3 and text.endswith(u"\\"):
        text = text.rstrip(u"\\")
    return lower_text(text)


def is_unc(path):
    text = as_text(path).strip()
    return text.startswith(u"\\\\") and not text.startswith(u"\\\\?\\")


def project_model_name(path_or_name):
    text = as_text(path_or_name).strip().replace(u"\\", u"/")
    leaf = text.split(u"/")[-1]
    if lower_text(leaf).endswith(u".rvt"):
        leaf = leaf[:-4]
    return leaf


def path_type_name(value):
    text = as_text(value).strip()
    if text == u"Absolute" or text.endswith(u".Absolute"):
        return u"Absolute"
    if text == u"Relative" or text.endswith(u".Relative"):
        return u"Relative"
    if text == u"Content" or text.endswith(u".Content"):
        return u"Content"
    if text == u"Server" or text.endswith(u".Server"):
        return u"Server"
    return text


def same_absolute(stored, path_type, target):
    """Уже стоит нужный файл только если диалог показал бы «Абсолютные» и путь совпадает. Относительный путь — не то."""
    if path_type_name(path_type) != u"Absolute":
        return False
    if not as_text(stored).strip() or not as_text(target).strip():
        return False
    return norm_path(stored) == norm_path(target)


def _digits(text):
    if not text:
        return False
    for ch in text:
        if ch < u"0" or ch > u"9":
            return False
    return True


def starts_with_date(name):
    """Папка-дата: 2026-10-19 и 2026-10-19_что-то. «Старт» сюда не попадает."""
    if len(name) < 10:
        return False
    return (
        name[4] == u"-" and name[7] == u"-"
        and _digits(name[0:4]) and _digits(name[5:7]) and _digits(name[8:10])
    )


def skipped_folder_name(name):
    """Те же папки, что Test-SkippedArchiveFolderName / Test-SkipWalkFolderName / сегмент Резерв."""
    n = as_text(name).strip()
    if not n:
        return False
    low = lower_text(n)
    if low in (u"backup", u"revit_temp", u"dwg", u"pdf", u"excel", u"word", u"xlsx", u"nwc", u"nwd", u"ifc", u"хлам", u"резерв", u"семейства"):
        return True
    if low.startswith(u"архив") or low.startswith(u"archive"):
        return True
    if low == u"old" or low.startswith(u"old_"):
        return True
    if low in (u"стар", u"старые", u"старое", u"старый"):
        return True
    if starts_with_date(n):
        return True
    return False


def backup_revit_copy(low_name):
    """Резервная копия Revit: имя.0001.rvt."""
    if not low_name.endswith(u".rvt") or len(low_name) < 10:
        return False
    stem = low_name[:-4]
    dot = stem.rfind(u".")
    if dot < 0 or dot == len(stem) - 1:
        return False
    return _digits(stem[dot + 1:]) and len(stem[dot + 1:]) == 4


def is_excluded_path(path):
    text = as_text(path).strip().replace(u"/", u"\\")
    if not text:
        return True
    parts = [p for p in text.split(u"\\") if p]
    if not parts:
        return True
    for seg in parts:
        if skipped_folder_name(seg):
            return True
    low_name = lower_text(parts[-1])
    if u"_backup" in low_name:
        return True
    if backup_revit_copy(low_name):
        return True
    return False


def section_code_from_folder(path):
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
    """Пустая строка — не модель раздела (сводная, TASK, BASE, нет кода)."""
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


_ACTIVE = {u"project": CLASSIFIER_PROJECT}


def classifier_settings(cfg):
    """(путь, проект) из run.cfg; пустые или нет строк — встроенные CLASSIFIER_PATH / CLASSIFIER_PROJECT."""
    path = as_text(cfg.get(u"classifier_path", u"")).strip() or CLASSIFIER_PATH
    project = as_text(cfg.get(u"classifier_project", u"")).strip() or CLASSIFIER_PROJECT
    return path, project


def name_has_project_prefix(name):
    prefix = upper_text(_ACTIVE[u"project"])
    return upper_text(name).startswith(prefix + u"_")


def model_disposition(path):
    """(result, error). result = section | skipped_not_section | error. section значит «можно ставить путь»."""
    if is_excluded_path(path):
        return u"skipped_not_section", u""
    if not name_has_project_prefix(project_model_name(path)):
        return u"error", u"not_target_project"
    if not section_code_from_name(path):
        return u"skipped_not_section", u""
    return u"section", u""


def read_text(path):
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
    if not file_exists(LOCAL_CFG):
        return data
    for line in read_text(LOCAL_CFG).splitlines():
        line = line.strip()
        if not line or line.startswith(u"#") or u"=" not in line:
            continue
        key, value = line.split(u"=", 1)
        data[lower_text(key.strip())] = value.strip()
    return data


def csv_cell(value):
    text = as_text(value)
    return text.replace(u";", u",").replace(u"\r", u" ").replace(u"\n", u" ")


def classifier_report_path(cfg):
    report_dir = cfg.get(u"report_dir", u"")
    run_id = cfg.get(u"run_id", u"run") or u"run"
    if not report_dir:
        return u""
    return os.path.join(report_dir, u"{}_классификатор.csv".format(run_id))


def append_report(path, cells):
    if not path:
        return
    from System.IO import StreamWriter
    from System.Text import UTF8Encoding
    line = u";".join([csv_cell(c) for c in cells]) + u"\r\n"
    if not file_exists(path):
        folder = os.path.dirname(path)
        if folder:
            ensure_dir(folder)
        writer = StreamWriter(path, False, UTF8Encoding(True))
        try:
            writer.Write(REPORT_HEADER + u"\r\n" + line)
        finally:
            writer.Close()
        return
    writer = StreamWriter(path, True, UTF8Encoding(False))
    try:
        writer.Write(line)
    finally:
        writer.Close()


def classifier_readable(path):
    """Файл есть и открывается на чтение. Кириллица, пробелы и скобки — через .NET, не через open()."""
    if not as_text(path).strip():
        return False
    try:
        from System.IO import File, FileStream, FileMode, FileAccess, FileShare
    except ImportError:
        return False
    if not File.Exists(path):
        return False
    stream = None
    try:
        stream = FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
        return True
    except Exception:
        return False
    finally:
        if stream is not None:
            try:
                stream.Close()
            except Exception:
                pass


def read_reference(table):
    """(stored user-visible path, PathType name). Пусто, если файла классификатора ещё нет."""
    from Autodesk.Revit.DB import ModelPathUtils
    try:
        if not table.IsExternalFileReference():
            return u"", u""
    except Exception:
        return u"", u""
    try:
        ext = table.GetExternalFileReference()
    except Exception:
        return u"", u""
    kind = u""
    stored = u""
    try:
        kind = path_type_name(ext.PathType)
    except Exception:
        kind = u""
    try:
        stored = as_text(ModelPathUtils.ConvertModelPathToUserVisiblePath(ext.GetPath()))
    except Exception:
        stored = u""
    return stored, kind


def status_success(status):
    text = as_text(status).strip()
    return text == u"Success" or text.endswith(u".Success")


def load_absolute(document, target):
    """Ставит абсолютный UNC. (True, "") либо (False, код ошибки). При неуспехе транзакция откатывается."""
    from Autodesk.Revit.DB import (
        AssemblyCodeTable,
        ExternalResourceReference,
        ExternalResourceTypes,
        ModelPathUtils,
        PathType,
        Transaction,
        KeyBasedTreeEntriesLoadResults,
    )
    table = AssemblyCodeTable.GetAssemblyCodeTable(document)
    if table is None:
        return False, u"no_assembly_code_table"
    model_path = ModelPathUtils.ConvertUserVisiblePathToModelPath(target)
    reference = ExternalResourceReference.CreateLocalResource(
        document,
        ExternalResourceTypes.BuiltInExternalResourceTypes.AssemblyCodeTable,
        model_path,
        PathType.Absolute,
    )
    if reference is None:
        return False, u"create_reference_failed"
    results = KeyBasedTreeEntriesLoadResults()
    transaction = Transaction(document, "Load assembly code file")
    transaction.Start()
    done = False
    try:
        status = table.LoadFrom(reference, results)
        stored, kind = read_reference(table)
        if (not status_success(status)) or (not same_absolute(stored, kind, target)):
            try:
                transaction.RollBack()
            except Exception:
                pass
            done = True
            if status_success(status):
                return False, u"stored_not_absolute"
            return False, u"load_failed"
        transaction.Commit()
        done = True
        return True, u""
    except Exception as ex:
        if not done:
            try:
                transaction.RollBack()
            except Exception:
                pass
        return False, u"{}".format(ex)


def is_rsn_path(path):
    return lower_text(as_text(path).strip()).startswith(u"rsn://")


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
    )
    is_detached = False
    try:
        is_detached = document.IsDetached
    except Exception:
        is_detached = False
    central = None
    try:
        central = document.GetWorksharingCentralModelPath()
    except Exception:
        central = None
    workshared_live = bool(document.IsWorkshared) and (not is_detached) and central is not None
    if is_rsn_path(path) or workshared_live:
        if is_detached:
            raise Exception("detached_from_central")
        sync_opts = SynchronizeWithCentralOptions()
        sync_opts.SetRelinquishOptions(make_relinquish_options())
        try:
            sync_opts.Comment = u"Assembly code file"
            sync_opts.SaveLocalFile = True
        except Exception:
            pass
        log_msg(u"Sync")
        document.SynchronizeWithCentral(TransactWithCentralOptions(), sync_opts)
        log_msg(u"SynchronizedWithCentral")
        return
    try:
        from System.IO import File, FileAttributes
        attrs = File.GetAttributes(path)
        readonly = FileAttributes.ReadOnly
        if (attrs & readonly) == readonly:
            File.SetAttributes(path, attrs & ~readonly)
    except Exception as ex:
        log_msg(u"WARN readonly {}".format(ex))
    save_opts = SaveAsOptions()
    save_opts.OverwriteExistingFile = True
    if document.IsWorkshared:
        workset_opts = WorksharingSaveAsOptions()
        workset_opts.SaveAsCentral = True
        save_opts.SetWorksharingOptions(workset_opts)
    document.SaveAs(path, save_opts)
    try:
        if document.IsWorkshared:
            WorksharingUtils.RelinquishOwnership(document, make_relinquish_options(), TransactWithCentralOptions())
    except Exception as ex:
        log_msg(u"WARN relinquish {}".format(ex))
    log_msg(u"SavedAs")


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


def write_live(target_path, result):
    path = live_status_path()
    if not path:
        return
    try:
        ensure_dir(os.path.dirname(path))
        index, total = model_position(target_path)
        from System.IO import StreamWriter
        from System.Text import UTF8Encoding
        leaf = project_model_name(target_path)
        lines = [
            u"index={}".format(index),
            u"total={}".format(total),
            u"name={}".format(leaf),
            u"inserted=0",
            u"skipped=0",
            u"params=0",
            u"result={}".format(as_text(result)),
        ]
        writer = StreamWriter(path, False, UTF8Encoding(True))
        try:
            writer.Write(u"\r\n".join(lines) + u"\r\n")
        finally:
            writer.Close()
    except Exception as ex:
        log_msg(u"WARN live status {}".format(ex))


def process_model(document, model_path, classifier_path):
    """Не сохраняет модель. save=True только у result=updated — сохраняет вызывающий."""
    row = {u"previous": u"", u"new": u"", u"result": u"error", u"error": u"", u"save": False}
    if not is_unc(classifier_path):
        row[u"error"] = u"classifier_path_not_unc"
        row[u"new"] = as_text(classifier_path)
        return row
    if not classifier_readable(classifier_path):
        row[u"error"] = u"classifier_file_missing"
        row[u"new"] = as_text(classifier_path)
        return row
    kind, why = model_disposition(model_path)
    if kind != u"section":
        row[u"result"] = kind
        row[u"error"] = why
        try:
            from Autodesk.Revit.DB import AssemblyCodeTable
            table = AssemblyCodeTable.GetAssemblyCodeTable(document)
            if table is not None:
                stored, _kind = read_reference(table)
                row[u"previous"] = stored
        except Exception:
            pass
        return row
    from Autodesk.Revit.DB import AssemblyCodeTable
    table = AssemblyCodeTable.GetAssemblyCodeTable(document)
    if table is None:
        row[u"error"] = u"no_assembly_code_table"
        row[u"new"] = as_text(classifier_path)
        return row
    stored, path_type = read_reference(table)
    row[u"previous"] = stored
    row[u"new"] = as_text(classifier_path)
    if same_absolute(stored, path_type, classifier_path):
        row[u"result"] = u"already_set"
        return row
    ok, detail = load_absolute(document, classifier_path)
    if not ok:
        row[u"error"] = detail or u"load_failed"
        return row
    row[u"result"] = u"updated"
    row[u"save"] = True
    return row


def main():
    log_msg(u"set_assembly_code start")
    import revit_script_util
    document = revit_script_util.GetScriptDocument()
    target_path = revit_script_util.GetRevitFilePath()
    cfg = read_cfg()
    report = classifier_report_path(cfg)
    classifier_path, project = classifier_settings(cfg)
    _ACTIVE[u"project"] = project
    log_msg(u"classifier path from {}".format(u"run.cfg" if as_text(cfg.get(u"classifier_path", u"")).strip() else u"built-in CLASSIFIER_PATH"))
    row = {
        u"model": project_model_name(target_path),
        u"previous": u"",
        u"new": u"",
        u"result": u"error",
        u"error": u"",
    }
    try:
        if document is None:
            row[u"error"] = u"no_document"
            log_msg(u"ERROR no_document")
        else:
            outcome = process_model(document, target_path, classifier_path)
            row[u"previous"] = outcome.get(u"previous", u"")
            row[u"new"] = outcome.get(u"new", u"")
            row[u"result"] = outcome.get(u"result", u"error")
            row[u"error"] = outcome.get(u"error", u"")
            write_live(target_path, row[u"result"])
            if row[u"result"] == u"updated" and outcome.get(u"save"):
                save_or_sync(document, target_path)
                log_msg(u"updated")
            elif row[u"result"] == u"already_set":
                log_msg(u"No changes, skip save")
            elif row[u"result"] == u"skipped_not_section":
                log_msg(u"skip not_section")
            else:
                log_msg(u"ERROR {}".format(row[u"error"]))
    except Exception as ex:
        row[u"result"] = u"error"
        if not row[u"error"]:
            row[u"error"] = u"{}".format(ex)
        log_msg(u"ERROR (see report)")
    append_report(report, [row[u"model"], row[u"previous"], row[u"new"], row[u"result"], row[u"error"]])


def _batch_document():
    try:
        import revit_script_util
        return revit_script_util.GetScriptDocument()
    except Exception:
        return None


if _batch_document() is not None:
    main()
