# -*- coding: utf-8 -*-
# Name: saveas_central.py
# Version: 3.1.0
# What it does: RBP Save As new central per job_paths.csv (rename and/or new folder).
# Inputs: rvt_list.txt or rvt_list_staging.txt, job_paths.csv.
# Outputs: New central paths; CSV log rows.
# How to run: start_ops.ps1 pass from presets UI.
# Notes: Create New Local, Detach off.

"""
RBP — SaveAs новой централи по job_paths.csv (новое имя и/или новая папка).

Список = rvt_list.txt (старые) ИЛИ rvt_list_staging.txt (буфер после Detach).
RBP: Create New Local, Detach OFF.
"""

from __future__ import print_function

import clr
import os
import sys
import time

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
from Autodesk.Revit.DB import *

import revit_script_util
from revit_script_util import Output

SCRIPT_VERSION = "3.1.0"
RBP_ROOT = os.path.join(
    os.environ.get("USERPROFILE", r"C:\Users\Public"),
    r"Documents\doc\script\rbp",
)
TOOL_DIR_FALLBACK = os.path.join(RBP_ROOT, "batch_model_ops")
TOOL_DIR_SHARE = r""


def get_script_dir():
    try:
        code = os.path.dirname(os.path.abspath(__file__))
        # code sits in <tool>\sluzhebnoe; job files are in the tool root one level up
        for folder in (os.path.dirname(code), code):
            if folder and os.path.isfile(os.path.join(folder, "job_paths.csv")):
                return folder
    except Exception:
        pass
    for folder in (TOOL_DIR_FALLBACK, os.getcwd(), TOOL_DIR_SHARE):
        if folder and os.path.isfile(os.path.join(folder, "job_paths.csv")):
            return folder
    return TOOL_DIR_FALLBACK


SCRIPT_DIR = get_script_dir()  # tool root: job_paths.csv, mapping.csv, job.json
try:
    CODE_DIR = os.path.dirname(os.path.abspath(__file__))  # ops_lib.py, relink_prefix.py
except Exception:
    CODE_DIR = SCRIPT_DIR
if not os.path.isfile(os.path.join(CODE_DIR, "ops_lib.py")):
    # no __file__ under RBP: the code folder is a subfolder of the tool root
    try:
        for _d in os.listdir(SCRIPT_DIR):
            if os.path.isfile(os.path.join(SCRIPT_DIR, _d, "ops_lib.py")):
                CODE_DIR = os.path.join(SCRIPT_DIR, _d)
                break
    except Exception:
        pass
for _p in (CODE_DIR, SCRIPT_DIR):
    if _p and _p not in sys.path:
        sys.path.append(_p)

import ops_lib as lib

JOB_PATHS = os.path.join(SCRIPT_DIR, "job_paths.csv")
MAPPING_PATH = os.path.join(SCRIPT_DIR, "mapping.csv")
JOB_JSON = os.path.join(SCRIPT_DIR, "job.json")
REPORTS_DIR = None
ACTIVE_SAVEAS = None


def ensure_saveas_csv():
    global REPORTS_DIR, ACTIVE_SAVEAS
    REPORTS_DIR = lib.shared_reports_dir(SCRIPT_DIR)
    ACTIVE_SAVEAS = os.path.join(REPORTS_DIR, "_active_saveas_central.txt")
    Output(u"Reports dir: {0}".format(lib.ascii_safe(REPORTS_DIR)))

    stamp = time.strftime("%Y-%m-%d_%H%M")
    path = os.path.join(REPORTS_DIR, "saveas_central_{}.csv".format(stamp))
    if lib.file_exists(ACTIVE_SAVEAS):
        try:
            p = lib.read_all_text(ACTIVE_SAVEAS).strip()
            if p and lib.file_exists(p):
                return p
        except Exception:
            pass
    lib.write_all_text(ACTIVE_SAVEAS, path)
    lib.write_all_text(
        path, "old_path;new_path;status;message;script_version\n"
    )
    return path


def make_relinquish_all_options():
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


def sync_with_relinquish(document, comment):
    is_detached = False
    try:
        is_detached = document.IsDetached
    except Exception:
        is_detached = False
    if is_detached:
        raise Exception("DETACHED — use Create New Local (Detach OFF).")
    if not document.IsWorkshared:
        raise Exception("Expected workshared for Sync.")
    sync_opts = SynchronizeWithCentralOptions()
    sync_opts.SetRelinquishOptions(make_relinquish_all_options())
    try:
        sync_opts.Comment = comment
    except Exception:
        pass
    try:
        sync_opts.SaveLocalFile = True
    except Exception:
        pass
    document.SynchronizeWithCentral(TransactWithCentralOptions(), sync_opts)
    Output("SynchronizedWithCentral + relinquish OK")


def saveas_new(document, new_path, allow_overwrite):
    """SaveAs to new path (same folder, new name). Workshared -> new central.
    An existing file is replaced only when job.json says allow_overwrite."""
    model_path = ModelPathUtils.ConvertUserVisiblePathToModelPath(new_path)
    opts = SaveAsOptions()
    opts.OverwriteExistingFile = bool(allow_overwrite)
    try:
        opts.Compact = False
    except Exception:
        pass

    if document.IsWorkshared:
        wopts = WorksharingSaveAsOptions()
        wopts.SaveAsCentral = True
        try:
            wopts.ClearTransmitted = False
        except Exception:
            pass
        opts.SetWorksharingOptions(wopts)

    document.SaveAs(model_path, opts)
    Output("SaveAs OK -> {}".format(new_path))


doc = revit_script_util.GetScriptDocument()
revitFilePath = revit_script_util.GetRevitFilePath()

Output()
Output("=== saveas_central v{} ===".format(SCRIPT_VERSION))
Output("Script dir: {}".format(SCRIPT_DIR))
Output("Old path: {}".format(revitFilePath))
Output(
    "Workshared={}, Detached={}".format(
        doc.IsWorkshared, getattr(doc, "IsDetached", False)
    )
)

csv_path = ensure_saveas_csv()

try:
    by_name, by_old, by_stg = lib.load_job_paths(JOB_PATHS)
    row = lib.find_job_row(revitFilePath, by_old, by_stg, by_name)
    if row:
        new_path = row["new_path"]
        old_base = row["old_name"] or lib.basename_no_ext(revitFilePath)
        new_base = row["new_name"] or lib.basename_no_ext(new_path)
    else:
        mapping = lib.load_mapping(MAPPING_PATH)
        new_path, old_base, new_base = lib.map_path(revitFilePath, mapping)
    Output("New path: {}".format(new_path))
    Output("Rename: {} -> {}".format(old_base, new_base))

    if revitFilePath.replace("\\", "/").lower() == new_path.replace("\\", "/").lower():
        raise Exception("Old and new path are identical — check mapping/list")

    allow_overwrite = lib.job_allow_overwrite(JOB_JSON)
    Output("allow_overwrite: {}".format(allow_overwrite))
    if lib.target_blocks_saveas(new_path, allow_overwrite):
        # Never replace an existing central silently; report and move on.
        msg = "target exists, not overwritten (enable 'allow overwrite' in the job)"
        Output("SKIP_EXISTS: {}".format(new_path))
        lib.append_csv(
            csv_path,
            [revitFilePath, new_path, "SKIP_EXISTS", msg, SCRIPT_VERSION],
        )
        Output("Done Pass1 SKIP_EXISTS")
        Output("OK")
    else:
        if lib.is_rsn_path(new_path) and not allow_overwrite:
            Output(
                "WARN: RSN target existence is not pre-checked; "
                "SaveAs runs with OverwriteExistingFile=False"
            )
        saveas_new(doc, new_path, allow_overwrite)

        # After SaveAs as new central, sync once
        if doc.IsWorkshared:
            try:
                is_detached = doc.IsDetached
            except Exception:
                is_detached = False
            if not is_detached:
                sync_with_relinquish(
                    doc,
                    "Pass1 SaveAs {} -> {}".format(old_base, new_base),
                )

        lib.append_csv(
            csv_path,
            [revitFilePath, new_path, "OK", "", SCRIPT_VERSION],
        )
        Output("Done Pass1 OK")
        Output("OK")
except Exception as ex:
    Output("ERROR: {}".format(ex))
    try:
        lib.append_csv(
            csv_path,
            [revitFilePath, "", "ERROR", str(ex)[:500], SCRIPT_VERSION],
        )
    except Exception:
        pass
    raise
