# -*- coding: utf-8 -*-
# Name: levels_grids_audit.py
# Version: 1.21.0
# What it does: Read-only RBP compare of host levels/grids against an exemplar RVT link.
# Inputs: rvt_list.txt, exemplar link cfg, tolerances.cfg, xlsx_out_path.cfg.
# Outputs: CSV per model; one Excel report (Summary + Details) at end of batch.
# How to run: RBP task after choose_exemplar_link and reset_report_session.
# Notes: Report-only; no model edits.

"""
Revit Batch Processor — Levels & Grids Audit (только отчёт).

Сравнивает host с эталонной RVT-связью:
  - уровни: отметка по имени, лишние и недостающие имена;
  - оси: перпендикуляр и угол по имени, лишние и недостающие имена;
  - базовая точка проекта;
  - плюс предупреждения Copy/Monitor из Document.GetWarnings().
Окно «Просмотр координаций» не равно GetWarnings: сдвиг уровня/оси
пишется в отчёт по геометрии, даже если предупреждений нет.
CSV на шаре + цветной XLSX один раз в конце батча.

Важно:
  - Save / Sync / Relinquish НЕ вызываются
  - в RBP: Create New Local, Detach = OFF
  - эталон по умолчанию AUTO: из связей хоста
    (БФ → АР того же корпуса; раздел → БФ того же корпуса)
  - явный путь в exemplar_model.cfg или project_map — по желанию
"""

import clr
import os
import sys
import System

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
from Autodesk.Revit.DB import *

import revit_script_util
from revit_script_util import Output

SCRIPT_VERSION = "1.21.0"
RBP_ROOT = os.path.join(
    os.environ.get("USERPROFILE", r"C:\Users\Public"),
    r"Documents\doc\script\rbp",
)
TOOL_DIR_FALLBACK = os.path.join(RBP_ROOT, "batch_levels_grids")
TOOL_DIR_SHARE = r""


def get_script_dir():
    try:
        folder = os.path.dirname(os.path.abspath(__file__))
        if folder and os.path.isfile(os.path.join(folder, "tolerances.cfg")):
            return folder
    except Exception:
        pass
    candidates = [TOOL_DIR_FALLBACK, os.getcwd(), TOOL_DIR_SHARE]
    for folder in candidates:
        if folder and os.path.isfile(os.path.join(folder, "tolerances.cfg")):
            return folder
    return TOOL_DIR_FALLBACK


SCRIPT_DIR = get_script_dir()
if SCRIPT_DIR not in sys.path:
    sys.path.insert(0, SCRIPT_DIR)

import lg_lib

CONFIG_PATH = os.path.join(SCRIPT_DIR, "tolerances.cfg")
MAP_PATH = os.path.join(SCRIPT_DIR, "project_map.csv")
MAP_EXAMPLE = os.path.join(SCRIPT_DIR, "project_map.example.csv")
REPORTS_DIR = lg_lib.shared_reports_dir()
CSV_DIR = REPORTS_DIR
ACTIVE_MARKER = os.path.join(CSV_DIR, "_active_levels_grids_audit.txt")
ACTIVE_DETAILS = os.path.join(CSV_DIR, "_active_levels_grids_audit_details.txt")
XLSX_OUT_CFG = os.path.join(SCRIPT_DIR, "xlsx_out_path.cfg")
RVT_LIST_PATH = os.path.join(SCRIPT_DIR, "rvt_list.txt")

SUMMARY_FIELDS = [
    "timestamp",
    "model_path",
    "model_name",
    "exemplar_link",
    "levels_host",
    "levels_exemplar",
    "grids_host",
    "grids_exemplar",
    "issues_total",
    "issues_red",
    "issues_yellow",
    "status",
    "short_problems",
    "script_version",
]

SUMMARY_HEADERS_RU = [
    u"Дата/время",
    u"Путь к модели",
    u"Имя файла",
    u"Связь-эталон",
    u"Уровней в модели",
    u"Уровней в эталоне",
    u"Осей в модели",
    u"Осей в эталоне",
    u"Расхождений",
    u"Критических",
    u"Предупреждений",
    u"Статус",
    u"Коротко о проблемах",
    u"Версия скрипта",
]

DETAIL_FIELDS = [
    "model_name",
    "exemplar_link",
    "category",
    "problem",
    "element",
    "value_host",
    "value_exemplar",
    "delta",
    "tolerance",
    "status",
    "source",
    "problem_code",
    "model_path",
    "timestamp",
]

DETAIL_HEADERS_RU = [
    u"Имя модели",
    u"Связь-эталон",
    u"Категория",
    u"Проблема",
    u"Элемент",
    u"Значение в модели",
    u"Значение в эталоне",
    u"Разница",
    u"Допуск",
    u"Статус",
    u"Источник",
    u"Код",
    u"Путь модели",
    u"Дата/время",
]

# Revit Coordination Review messages (levels/grids) — exact UI wording
MSG_NAME_CHANGED = u"Имя изменено"
MSG_LEVEL_MOVED = u"Уровень перемещен"
MSG_GRID_MOVED = u"Сетка перемещена"
MSG_ELEMENT_DELETED = u"Элемент удален"
MSG_LEFTOVER_EXEMPLAR = u"Есть в эталоне, нет в модели"
MSG_PBP_MOVED = u"Базовая точка смещена"
MSG_SURVEY_MOVED = u"Точка съемки смещена"

# Flat order for "Коротко о проблемах"
COORD_MSG_ORDER = (
    MSG_NAME_CHANGED,
    MSG_LEVEL_MOVED,
    MSG_GRID_MOVED,
    MSG_PBP_MOVED,
    MSG_SURVEY_MOVED,
    MSG_ELEMENT_DELETED,
    MSG_LEFTOVER_EXEMPLAR,
)

# Internal codes -> Coordination Review message (levels/grids)
COORD_MSG_BY_CODE = {
    "name_mismatch": MSG_NAME_CHANGED,
    "elevation_delta": MSG_LEVEL_MOVED,
    "position_delta": MSG_GRID_MOVED,
    "angle_delta": MSG_GRID_MOVED,
    "pbp_delta": MSG_PBP_MOVED,
    "pbp_missing": MSG_PBP_MOVED,
    "survey_delta": MSG_SURVEY_MOVED,
    "survey_missing": MSG_SURVEY_MOVED,
    "extra_in_host": MSG_ELEMENT_DELETED,
    "leftover_exemplar": MSG_LEFTOVER_EXEMPLAR,
    "missing_in_host": MSG_LEFTOVER_EXEMPLAR,
    "whitelist_missing": MSG_ELEMENT_DELETED,
    "monitor_broken": MSG_ELEMENT_DELETED,
    "cr_grid_moved": MSG_GRID_MOVED,
    "cr_level_moved": MSG_LEVEL_MOVED,
    "cr_name_changed": MSG_NAME_CHANGED,
    "cr_element_deleted": MSG_ELEMENT_DELETED,
    "cr_link_needs_review": u"Нужна проверка координации",
    "cr_watched_link_moved": u"Наблюдаемая связь перемещена",
    "cr_alert": u"Предупреждение Copy/Monitor",
    "grid_relative_moved": MSG_GRID_MOVED,
}

# English internal codes -> Russian for Excel (all sheets)
PROBLEM_RU = {
    "missing_in_host": MSG_LEFTOVER_EXEMPLAR,
    "leftover_exemplar": MSG_LEFTOVER_EXEMPLAR,
    "extra_in_host": MSG_ELEMENT_DELETED,
    "elevation_delta": MSG_LEVEL_MOVED,
    "position_delta": MSG_GRID_MOVED,
    "angle_delta": MSG_GRID_MOVED,
    "pbp_delta": MSG_PBP_MOVED,
    "pbp_missing": MSG_PBP_MOVED,
    "survey_delta": MSG_SURVEY_MOVED,
    "survey_missing": MSG_SURVEY_MOVED,
    "name_mismatch": MSG_NAME_CHANGED,
    "whitelist_missing": MSG_ELEMENT_DELETED,
    "no_monitoring": u"Нет мониторинга",
    "duplicate_monitor": u"Дубль мониторинга",
    "monitor_broken": u"Мониторинг разорван",
    "duplicate_host": u"Дубли имени в модели",
    "duplicate_exemplar": u"Дубли имени в эталоне",
    "exemplar_not_found": u"Эталонная связь не найдена/недоступна",
    "script_error": u"Ошибка скрипта",
    "cr_grid_moved": MSG_GRID_MOVED,
    "cr_level_moved": MSG_LEVEL_MOVED,
    "cr_name_changed": MSG_NAME_CHANGED,
    "cr_element_deleted": MSG_ELEMENT_DELETED,
    "cr_link_needs_review": u"Нужна проверка координации",
    "cr_watched_link_moved": u"Наблюдаемая связь перемещена",
    "cr_alert": u"Предупреждение Copy/Monitor",
    "grid_relative_moved": MSG_GRID_MOVED,
}

PROBLEM_LINE = dict(PROBLEM_RU)

CATEGORY_RU = {
    "Level": u"Уровни",
    "Grid": u"Оси",
    "Coords": u"Координаты",
    "System": u"Система",
}

SOURCE_RU = {
    "revit_cr": u"Revit CR",
    "pbp": u"Базовая точка",
    "geometry": u"Геометрия",
}


def ensure_reports_dir():
    """First action for any report: create the shared folder if it does not exist."""
    return lg_lib.shared_reports_dir()


def clear_active_markers():
    for path in (ACTIVE_MARKER, ACTIVE_DETAILS):
        try:
            if os.path.isfile(path):
                os.remove(path)
        except Exception:
            pass


def is_first_model_in_batch():
    try:
        get_num = getattr(revit_script_util, "GetProgressNumber", None)
        if callable(get_num):
            return int(get_num()) <= 1
    except Exception:
        pass
    return False


def get_report_paths():
    """
    One CSV pair per RBP batch.
    First model in batch always starts a NEW report (old marker cleared).
    """
    ensure_reports_dir()
    if is_first_model_in_batch():
        clear_active_markers()
        Output("New report session (first model in RBP batch)")

    if os.path.isfile(ACTIVE_MARKER):
        try:
            with open(ACTIVE_MARKER, "r") as f:
                name = f.read().strip()
            if name:
                summary = os.path.join(CSV_DIR, name)
                details_name = name.replace(".csv", "_details.csv")
                if os.path.isfile(ACTIVE_DETAILS):
                    with open(ACTIVE_DETAILS, "r") as f2:
                        dn = f2.read().strip()
                        if dn:
                            details_name = dn
                return summary, os.path.join(CSV_DIR, details_name)
        except Exception:
            pass

    name = "levels_grids_audit_{}.csv".format(lg_lib.now_file_stamp())
    details_name = name.replace(".csv", "_details.csv")
    with open(ACTIVE_MARKER, "w") as f:
        f.write(name)
    with open(ACTIVE_DETAILS, "w") as f:
        f.write(details_name)
    Output("Report files: {} / {}".format(name, details_name))
    return os.path.join(CSV_DIR, name), os.path.join(CSV_DIR, details_name)


def read_xlsx_output_dir():
    """Always the shared rbp\\reports folder (created if missing)."""
    return lg_lib.shared_reports_dir()


def read_rvt_list_models(list_path=None):
    path = list_path or RVT_LIST_PATH
    if not os.path.isfile(path):
        return []
    models = []
    try:
        with open(path, "r") as f:
            for raw in f:
                line = raw.strip().strip('"')
                if not line or line.startswith("#"):
                    continue
                if lg_lib.is_dummy_rsn_path(line):
                    continue
                low = line.lower()
                if low.endswith(".rvt") or low.startswith("rsn://"):
                    models.append(line)
    except Exception as ex:
        Output("WARN read model list: {}".format(ex))
    return models


def count_csv_data_rows(report_path):
    """Count data rows; ignore newlines inside quoted CSV fields."""
    if not report_path or not os.path.isfile(report_path):
        return 0
    n = 0
    try:
        with open(report_path, "r") as f:
            text = f.read()
        if text.startswith("\xef\xbb\xbf"):
            text = text[3:]
        row = []
        field = []
        in_quotes = False
        i = 0
        first_row = True
        while i < len(text):
            ch = text[i]
            if in_quotes:
                if ch == '"':
                    if i + 1 < len(text) and text[i + 1] == '"':
                        field.append('"')
                        i += 2
                        continue
                    in_quotes = False
                    i += 1
                    continue
                field.append(ch)
                i += 1
                continue
            if ch == '"':
                in_quotes = True
                i += 1
                continue
            if ch == ";":
                row.append("".join(field))
                field = []
                i += 1
                continue
            if ch == "\r":
                i += 1
                continue
            if ch == "\n":
                row.append("".join(field))
                field = []
                if not first_row:
                    # non-empty data row
                    if any(x.strip() for x in row):
                        n += 1
                first_row = False
                row = []
                i += 1
                continue
            field.append(ch)
            i += 1
        if field or row:
            row.append("".join(field))
            if not first_row and any(x.strip() for x in row):
                n += 1
    except Exception:
        return 0
    return n


def is_last_model_in_batch(current_path, report_path):
    models = read_rvt_list_models()
    if models:
        cur = (current_path or "").strip().lower()
        cur_base = os.path.basename(current_path or "").lower()
        idx = None
        for i, m in enumerate(models):
            ml = m.strip().lower()
            if ml == cur or os.path.basename(m).lower() == cur_base:
                idx = i
                break
        if idx is not None:
            Output("RBP list (no dummy RSN): {}/{}".format(idx + 1, len(models)))
            return idx >= len(models) - 1

    try:
        get_num = getattr(revit_script_util, "GetProgressNumber", None)
        get_max = getattr(revit_script_util, "GetProgressMax", None)
        if callable(get_num) and callable(get_max):
            num = int(get_num())
            mx = int(get_max())
            Output("RBP progress: {}/{}".format(num, mx))
            if mx > 0:
                return num >= mx
    except Exception as ex:
        Output("WARN RBP progress API: {}".format(ex))

    rows = count_csv_data_rows(report_path)
    if models:
        Output(
            "Fallback progress: csv_rows={}/{} (rvt_list)".format(rows, len(models))
        )
        return rows >= len(models)

    Output("WARN cannot detect batch end — building XLSX now")
    return True


def export_xlsx_from_csv(summary_path, details_path, wait=True):
    try:
        script_path = os.path.join(SCRIPT_DIR, "csv_to_colored_xlsx.ps1")
        if not os.path.isfile(script_path):
            Output("WARN XLSX export skipped: csv_to_colored_xlsx.ps1 not found")
            return False

        out_dir = read_xlsx_output_dir()
        if not out_dir:
            Output("WARN XLSX export skipped: cannot resolve output folder")
            return False
        try:
            if not os.path.isdir(out_dir):
                os.makedirs(out_dir)
        except Exception:
            pass

        log_path = os.path.join(out_dir, "levels_grids_audit_xlsx_export.log")
        args = (
            u'-NoProfile -ExecutionPolicy Bypass -File "{0}" '
            u'-CsvPath "{1}" -DetailsCsvPath "{2}" -OutputDir "{3}" -LogPath "{4}"'
        ).format(script_path, summary_path, details_path, out_dir, log_path)

        psi = System.Diagnostics.ProcessStartInfo()
        psi.FileName = "powershell"
        psi.Arguments = args
        psi.CreateNoWindow = True
        psi.UseShellExecute = False
        psi.RedirectStandardOutput = True
        psi.RedirectStandardError = True

        Output("XLSX export: building -> {}".format(out_dir))
        proc = System.Diagnostics.Process.Start(psi)
        if wait:
            if not proc.WaitForExit(90000):
                try:
                    proc.Kill()
                except Exception:
                    pass
                Output("WARN XLSX export timeout (>90s)")
                return False
            stdout = proc.StandardOutput.ReadToEnd()
            stderr = proc.StandardError.ReadToEnd()
            if proc.ExitCode == 0:
                if stdout:
                    Output(stdout.strip())
                Output("XLSX export: OK")
                return True
            Output(
                "WARN XLSX export failed (exit {}): {}".format(
                    proc.ExitCode, stderr or stdout
                )
            )
            return False
        else:
            Output("XLSX export: started (async)")
            return True
    except Exception as ex:
        Output("WARN XLSX export exception: {}".format(ex))
        return False


def resolve_map_path():
    if os.path.isfile(MAP_PATH):
        return MAP_PATH
    return None


def empty_problem_counts():
    return {}


def count_problem_nums(details):
    return empty_problem_counts(), u""


def issues_summary(details, limit=8):
    return short_problems_from_details(details)


def short_problems_from_details(details):
    """
    Flat unique list for Excel "Коротко о проблемах".
    Coordination Review messages plus PBP / survey shift.
    """
    found = set()
    for d in details or []:
        code = d.get("problem") or d.get("problem_code") or ""
        msg = COORD_MSG_BY_CODE.get(code)
        if msg:
            found.add(msg)
    return u"; ".join([m for m in COORD_MSG_ORDER if m in found])


def localize_detail_row(d):
    row = dict(d)
    code = row.get("problem", "") or row.get("problem_code", "")
    row["problem_code"] = code
    row["problem"] = PROBLEM_RU.get(code, code)
    cat = row.get("category", "")
    row["category"] = CATEGORY_RU.get(cat, cat)
    src = row.get("source") or "revit_cr"
    row["source"] = SOURCE_RU.get(src, src)
    if not row.get("element"):
        row["element"] = row.get("name_host") or row.get("name_exemplar") or ""
    return row


def analyze(document, model_path, tolerances, project_rows):
    row_map = lg_lib.find_project_row(project_rows, model_path)
    whitelist = row_map.get("level_whitelist") if row_map else None
    link_hint = lg_lib.resolve_exemplar_hint(SCRIPT_DIR, tolerances, row_map)

    inst, link_doc, label = lg_lib.find_exemplar_link(
        document, link_hint, host_model_path=model_path
    )
    if link_doc is None:
        det = [
            {
                "timestamp": lg_lib.now_stamp(),
                "model_path": model_path or "",
                "model_name": os.path.basename(model_path) if model_path else "",
                "exemplar_link": "",
                "category": "System",
                "name_host": "",
                "name_exemplar": "",
                "element": "",
                "problem": "exemplar_not_found",
                "problem_code": "exemplar_not_found",
                "value_host": "",
                "value_exemplar": link_hint or "",
                "delta": "",
                "tolerance": "",
                "status": "ERROR",
            }
        ]
        row = {
            "timestamp": lg_lib.now_stamp(),
            "model_path": model_path or "",
            "model_name": os.path.basename(model_path) if model_path else "",
            "exemplar_link": "",
            "levels_host": "",
            "levels_exemplar": "",
            "grids_host": "",
            "grids_exemplar": "",
            "issues_total": 1,
            "issues_red": 1,
            "issues_yellow": 0,
            "status": "ERROR",
            "short_problems": short_problems_from_details(det),
            "script_version": SCRIPT_VERSION,
            "_details": det,
        }
        return row

    details, counts = lg_lib.compare_levels_grids(
        document,
        link_doc,
        tolerances,
        whitelist,
        exemplar_link_instance=inst,
        host_model_path=model_path,
    )

    link_inst_id = None
    try:
        if inst is not None:
            link_inst_id = inst.Id
    except Exception:
        link_inst_id = None

    cr_pack = lg_lib.collect_coordination_review_warnings(
        document, link_instance_id=link_inst_id, link_label=label
    )
    if isinstance(cr_pack, tuple):
        cr_details, warn_total = cr_pack
    else:
        cr_details, warn_total = cr_pack, -1
    details = lg_lib.merge_geo_with_coord_review(details, cr_details)
    counts.update(lg_lib.recount_issues(details))

    ts = lg_lib.now_stamp()
    model_name = os.path.basename(model_path) if model_path else ""
    for d in details:
        d["timestamp"] = ts
        d["model_path"] = model_path or ""
        d["model_name"] = model_name
        d["exemplar_link"] = label
        d["element"] = d.get("name_host") or d.get("name_exemplar") or ""
        d["problem_code"] = d.get("problem", "")
        if not d.get("source"):
            d["source"] = "revit_cr"

    row = {
        "timestamp": ts,
        "model_path": model_path or "",
        "model_name": model_name,
        "exemplar_link": label,
        "levels_host": counts["levels_host"],
        "levels_exemplar": counts["levels_exemplar"],
        "grids_host": counts["grids_host"],
        "grids_exemplar": counts["grids_exemplar"],
        "issues_total": counts["issues_total"],
        "issues_red": counts["issues_red"],
        "issues_yellow": counts["issues_yellow"],
        "status": counts["status"],
        "short_problems": short_problems_from_details(details),
        "script_version": SCRIPT_VERSION,
        "_details": details,
        "_cr_count": len(cr_details),
        "_warn_total": warn_total,
    }
    return row


def write_error_row(path, message, summary_path, details_path):
    ts = lg_lib.now_stamp()
    name = os.path.basename(path) if path else ""
    row = dict((h, "") for h in SUMMARY_FIELDS)
    row["timestamp"] = ts
    row["model_path"] = path or ""
    row["model_name"] = name
    row["status"] = "ERROR"
    row["short_problems"] = u""
    row["issues_total"] = 1
    row["issues_red"] = 1
    row["script_version"] = SCRIPT_VERSION
    lg_lib.append_csv(summary_path, SUMMARY_FIELDS, SUMMARY_HEADERS_RU, row)
    lg_lib.append_csv(
        details_path,
        DETAIL_FIELDS,
        DETAIL_HEADERS_RU,
        localize_detail_row(
            {
                "timestamp": ts,
                "model_path": path or "",
                "model_name": name,
                "exemplar_link": "",
                "category": "System",
                "name_host": "",
                "name_exemplar": "",
                "element": "",
                "problem": "script_error",
                "problem_code": "script_error",
                "value_host": message,
                "value_exemplar": "",
                "delta": "",
                "tolerance": "",
                "status": "ERROR",
            }
        ),
    )


# --- main ---
TOL = lg_lib.load_tolerances(CONFIG_PATH)
map_file = resolve_map_path()
PROJECT_ROWS = lg_lib.load_project_map(map_file) if map_file else []
summary_path, details_path = get_report_paths()
doc = None
revitFilePath = ""

Output()
Output("=== Levels & Grids Audit v{} ===".format(SCRIPT_VERSION))
Output("Script dir: {}".format(SCRIPT_DIR))
Output("Config: {}".format(CONFIG_PATH))
Output("Project map: {}".format(map_file or "(optional — not required)"))
hint = lg_lib.resolve_exemplar_hint(SCRIPT_DIR, TOL, None)
if not hint:
    Output("Exemplar mode: AUTO (link from each host by building code)")
else:
    Output("Exemplar hint: {!r}".format(hint))
Output("Report: {}".format(summary_path))
Output("Details: {}".format(details_path))
Output(
    "Compare: level elevation, grid perp+angle, missing/extra, PBP, GetWarnings"
)
Output(
    "Tolerances mm: level={} grid={} angle_deg={} pbp_xy={} pbp_z={}".format(
        TOL["level_delta_yellow_mm"],
        TOL["grid_delta_yellow_mm"],
        TOL["grid_angle_yellow_deg"],
        TOL["pbp_delta_yellow_mm"],
        TOL["pbp_z_yellow_mm"],
    )
)

try:
    doc = revit_script_util.GetScriptDocument()
    revitFilePath = revit_script_util.GetRevitFilePath()
    Output("File: {}".format(revitFilePath))
    Output(
        "Workshared={}, Detached={}".format(
            doc.IsWorkshared, getattr(doc, "IsDetached", False)
        )
    )

    result = analyze(doc, revitFilePath, TOL, PROJECT_ROWS)
    cr_count = result.pop("_cr_count", 0)
    warn_total = result.pop("_warn_total", -1)
    details = result.pop("_details", [])
    lg_lib.append_csv(summary_path, SUMMARY_FIELDS, SUMMARY_HEADERS_RU, result)
    if details:
        lg_lib.append_csv_rows(
            details_path,
            DETAIL_FIELDS,
            DETAIL_HEADERS_RU,
            [localize_detail_row(d) for d in details],
        )

    Output(
        "levels host/ex={}/{} grids host/ex={}/{}".format(
            result["levels_host"],
            result["levels_exemplar"],
            result["grids_host"],
            result["grids_exemplar"],
        )
    )
    Output(
        "issues total={} RED={} YELLOW={}".format(
            result["issues_total"],
            result["issues_red"],
            result["issues_yellow"],
        )
    )
    Output("Document.GetWarnings() total: {}".format(warn_total))
    Output("Coordination Review (GetWarnings): {} items".format(cr_count))
    Output("status={}".format(result["status"]))
    if result.get("short_problems"):
        Output("short: {}".format(result["short_problems"]))
    Output("exemplar={}".format(result["exemplar_link"]))

    if is_last_model_in_batch(revitFilePath, summary_path):
        Output("Last model in RBP batch — building XLSX once")
        ok = export_xlsx_from_csv(summary_path, details_path, wait=True)
        # Close session so the next RBP run cannot append into this CSV/XLSX.
        clear_active_markers()
        if ok:
            Output("Report session closed (next run = new files)")
    else:
        Output("XLSX deferred until last model in RBP batch")

    Output("NO SAVE — report only")
    Output("OK")
except Exception as ex:
    msg = lg_lib.u(ex)
    Output("ERROR: {}".format(msg))
    try:
        write_error_row(revitFilePath, msg, summary_path, details_path)
        if is_last_model_in_batch(revitFilePath, summary_path):
            export_xlsx_from_csv(summary_path, details_path, wait=True)
            clear_active_markers()
    except Exception as ex2:
        Output("WARN could not write error row: {}".format(ex2))
    raise
