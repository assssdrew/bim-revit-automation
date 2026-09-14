# -*- coding: utf-8 -*-
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

SCRIPT_VERSION = "3.0.0"
RBP_ROOT = os.path.join(
    os.environ.get("USERPROFILE", r"C:\Users\Public"),
    r"Documents\doc\script\rbp",
)
TOOL_DIR_FALLBACK = os.path.join(RBP_ROOT, "batch_model_ops")
TOOL_DIR_SHARE = r""  # optional deployed copy on share


def get_script_dir():
    try:
        folder = os.path.dirname(os.path.abspath(__file__))
        if folder and os.path.isfile(os.path.join(folder, "job_paths.csv")):
            return folder
    except Exception:
        pass
    for folder in (TOOL_DIR_FALLBACK, os.getcwd(), TOOL_DIR_SHARE):
        if folder and os.path.isfile(os.path.join(folder, "job_paths.csv")):
            return folder
    return TOOL_DIR_FALLBACK


SCRIPT_DIR = get_script_dir()
if SCRIPT_DIR not in sys.path:
    sys.path.append(SCRIPT_DIR)

import ops_lib as lib

JOB_PATHS = os.path.join(SCRIPT_DIR, "job_paths.csv")
MAPPING_PATH = os.path.join(SCRIPT_DIR, "mapping.csv")
REPORTS_DIR = None
ACTIVE_SAVEAS = None


def ensure_saveas_csv():
    global REPORTS_DIR, ACTIVE_SAVEAS
    REPORTS_DIR = lib.shared_reports_dir()
    ACTIVE_SAVEAS = os.path.join(REPORTS_DIR, "_active_saveas_central.txt")
    Output("Reports dir: {}".format(REPORTS_DIR))

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


def saveas_new(document, new_path):
    """SaveAs to new path (same folder, new name). Workshared -> new central."""
    model_path = ModelPathUtils.ConvertUserVisiblePathToModelPath(new_path)
    opts = SaveAsOptions()
    opts.OverwriteExistingFile = True
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

    saveas_new(doc, new_path)

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
