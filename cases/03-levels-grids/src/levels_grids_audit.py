# -*- coding: utf-8 -*-
# Name: levels_grids_audit.py
# Version: 1.0.2
# What it does: Read-only RBP compare of host levels/grids against an exemplar RVT link.
# Inputs: rvt_list.txt, exemplar link cfg, tolerances.cfg, xlsx_out_path.cfg.
# Outputs: CSV per model; one Excel report (Summary + Details) at end of batch.
# How to run: RBP task after choose_exemplar_link and reset_report_session.
# Notes: Report-only; no model edits.

import clr
import os
import sys
import System

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
from Autodesk.Revit.DB import *

import revit_script_util
from revit_script_util import Output

SCRIPT_VERSION = "1.0.2"
TOOL_DIR_FALLBACK = r""  # optional: path to deployed copy on share


def get_script_dir():
    try:
        return os.path.dirname(os.path.abspath(__file__))
    except Exception:
        pass
    candidates = [TOOL_DIR_FALLBACK, os.getcwd()]
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
REPORTS_DIR = os.path.join(SCRIPT_DIR, "reports")
CSV_DIR = os.path.join(REPORTS_DIR, "cvc")
ACTIVE_MARKER = os.path.join(CSV_DIR, "_active_report.txt")
ACTIVE_DETAILS = os.path.join(CSV_DIR, "_active_details.txt")
XLSX_OUT_CFG = os.path.join(SCRIPT_DIR, "xlsx_out_path.cfg")
RVT_LIST_PATH = os.path.join(SCRIPT_DIR, "rvt_list.txt")

SUMMARY_FIELDS = [
    "timestamp",
    "model_path",
    "model_name",
    "building",
    "role",
    "exemplar_link",
    "levels_host",
    "levels_exemplar",
    "grids_host",
    "grids_exemplar",
    "issues_total",
    "issues_red",
    "issues_yellow",
    "status",
    "issues",
    "script_version",
]

SUMMARY_HEADERS_RU = [
    u"Дата/время",
    u"Путь к модели",
    u"Имя файла",
    u"Корпус",
    u"Роль",
    u"Связь-эталон",
    u"Уровни host",
    u"Уровни эталон",
    u"Оси host",
    u"Оси эталон",
    u"Расхождений, шт",
    u"RED, шт",
    u"YELLOW, шт",
    u"Статус",
    u"Проблемы",
    u"Версия скрипта",
]

DETAIL_FIELDS = [
    "timestamp",
    "model_path",
    "model_name",
    "exemplar_link",
    "category",
    "name_host",
    "name_exemplar",
    "problem",
    "value_host",
    "value_exemplar",
    "delta",
    "tolerance",
    "status",
]

DETAIL_HEADERS_RU = [
    u"Дата/время",
    u"Host",
    u"Имя файла",
    u"Эталон",
    u"Категория",
    u"Имя host",
    u"Имя эталон",
    u"Тип проблемы",
    u"Значение host",
    u"Значение эталон",
    u"Delta",
    u"Допуск",
    u"Статус",
]


def ensure_reports_dir():
    if not os.path.isdir(REPORTS_DIR):
        os.makedirs(REPORTS_DIR)
    if not os.path.isdir(CSV_DIR):
        os.makedirs(CSV_DIR)


def get_report_paths():
    ensure_reports_dir()
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

    name = "lg_{}.csv".format(lg_lib.now_file_stamp())
    details_name = name.replace(".csv", "_details.csv")
    with open(ACTIVE_MARKER, "w") as f:
        f.write(name)
    with open(ACTIVE_DETAILS, "w") as f:
        f.write(details_name)
    return os.path.join(CSV_DIR, name), os.path.join(CSV_DIR, details_name)


def read_xlsx_output_dir():
    if not os.path.isfile(XLSX_OUT_CFG):
        return ""
    try:
        with open(XLSX_OUT_CFG, "r") as f:
            for raw in f:
                line = raw.strip()
                if not line or line.startswith("#"):
                    continue
                return line
    except Exception:
        pass
    return ""


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
                low = line.lower()
                if low.endswith(".rvt") or low.startswith("rsn://"):
                    models.append(line)
    except Exception as ex:
        Output("WARN read model list: {}".format(ex))
    return models


def count_csv_data_rows(report_path):
    if not report_path or not os.path.isfile(report_path):
        return 0
    n = 0
    try:
        with open(report_path, "r") as f:
            first = True
            for raw in f:
                if first:
                    first = False
                    continue
                if raw.strip():
                    n += 1
    except Exception:
        return 0
    return n


def is_last_model_in_batch(current_path, report_path):
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

    models = read_rvt_list_models()
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
            Output(
                "WARN XLSX export skipped: run reset_report_session.cmd "
                "and choose local XLSX folder"
            )
            return False

        log_path = os.path.join(out_dir, "xlsx_export.log")
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


def issues_preview(details, limit=8):
    parts = []
    for d in details[:limit]:
        parts.append(
            u"{0}:{1}/{2}".format(
                d.get("problem", ""),
                d.get("name_host") or d.get("name_exemplar") or "",
                d.get("status", ""),
            )
        )
    if len(details) > limit:
        parts.append(u"...(+{})".format(len(details) - limit))
    return u" | ".join(parts)


def analyze(document, model_path, tolerances, project_rows):
    row_map = lg_lib.find_project_row(project_rows, model_path)
    building = row_map["building"] if row_map else ""
    role = row_map["role"] if row_map else ""
    whitelist = row_map.get("level_whitelist") if row_map else None
    link_hint = lg_lib.resolve_exemplar_hint(SCRIPT_DIR, tolerances, row_map)

    inst, link_doc, label = lg_lib.find_exemplar_link(document, link_hint)
    if link_doc is None:
        return {
            "timestamp": lg_lib.now_stamp(),
            "model_path": model_path or "",
            "model_name": os.path.basename(model_path) if model_path else "",
            "building": building,
            "role": role,
            "exemplar_link": label or "",
            "levels_host": "",
            "levels_exemplar": "",
            "grids_host": "",
            "grids_exemplar": "",
            "issues_total": 1,
            "issues_red": 1,
            "issues_yellow": 0,
            "status": "ERROR",
            "issues": label or u"эталон не найден",
            "script_version": SCRIPT_VERSION,
            "_details": [
                {
                    "timestamp": lg_lib.now_stamp(),
                    "model_path": model_path or "",
                    "model_name": os.path.basename(model_path) if model_path else "",
                    "exemplar_link": label or "",
                    "category": "System",
                    "name_host": "",
                    "name_exemplar": "",
                    "problem": "exemplar_not_found",
                    "value_host": "",
                    "value_exemplar": link_hint or "",
                    "delta": "",
                    "tolerance": "",
                    "status": "ERROR",
                }
            ],
        }

    details, counts = lg_lib.compare_levels_grids(
        document, link_doc, tolerances, whitelist
    )
    ts = lg_lib.now_stamp()
    model_name = os.path.basename(model_path) if model_path else ""
    for d in details:
        d["timestamp"] = ts
        d["model_path"] = model_path or ""
        d["model_name"] = model_name
        d["exemplar_link"] = label

    return {
        "timestamp": ts,
        "model_path": model_path or "",
        "model_name": model_name,
        "building": building,
        "role": role,
        "exemplar_link": label,
        "levels_host": counts["levels_host"],
        "levels_exemplar": counts["levels_exemplar"],
        "grids_host": counts["grids_host"],
        "grids_exemplar": counts["grids_exemplar"],
        "issues_total": counts["issues_total"],
        "issues_red": counts["issues_red"],
        "issues_yellow": counts["issues_yellow"],
        "status": counts["status"],
        "issues": issues_preview(details),
        "script_version": SCRIPT_VERSION,
        "_details": details,
    }


def write_error_row(path, message, summary_path, details_path):
    ts = lg_lib.now_stamp()
    name = os.path.basename(path) if path else ""
    row = dict((h, "") for h in SUMMARY_FIELDS)
    row["timestamp"] = ts
    row["model_path"] = path or ""
    row["model_name"] = name
    row["status"] = "ERROR"
    row["issues"] = message
    row["issues_total"] = 1
    row["issues_red"] = 1
    row["script_version"] = SCRIPT_VERSION
    lg_lib.append_csv(summary_path, SUMMARY_FIELDS, SUMMARY_HEADERS_RU, row)
    lg_lib.append_csv(
        details_path,
        DETAIL_FIELDS,
        DETAIL_HEADERS_RU,
        {
            "timestamp": ts,
            "model_path": path or "",
            "model_name": name,
            "exemplar_link": "",
            "category": "System",
            "name_host": "",
            "name_exemplar": "",
            "problem": "script_error",
            "value_host": message,
            "value_exemplar": "",
            "delta": "",
            "tolerance": "",
            "status": "ERROR",
        },
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
Output("Exemplar hint: {!r}".format(hint))
Output("Report: {}".format(summary_path))
Output("Details: {}".format(details_path))
Output(
    "Tolerances: level={}mm grid={}mm angle={}deg".format(
        TOL["level_mm"], TOL["grid_mm"], TOL["grid_angle_deg"]
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
    details = result.pop("_details", [])
    lg_lib.append_csv(summary_path, SUMMARY_FIELDS, SUMMARY_HEADERS_RU, result)
    if details:
        lg_lib.append_csv_rows(
            details_path, DETAIL_FIELDS, DETAIL_HEADERS_RU, details
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
    Output("status={}".format(result["status"]))
    if result["issues"]:
        Output("issues: {}".format(result["issues"]))
    Output("exemplar={}".format(result["exemplar_link"]))

    if is_last_model_in_batch(revitFilePath, summary_path):
        Output("Last model in RBP batch — building XLSX once")
        export_xlsx_from_csv(summary_path, details_path, wait=True)
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
    except Exception as ex2:
        Output("WARN could not write error row: {}".format(ex2))
    raise
