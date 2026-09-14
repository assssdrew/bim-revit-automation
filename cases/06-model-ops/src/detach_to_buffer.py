# -*- coding: utf-8 -*-
"""
RBP — выгрузка с RSN на буфер ПК: Detach уже включён в RBP.

Список = rvt_list.txt (старые RSN).
RBP: Detach from Central = ON. Create New Local = OFF.
SaveAs в staging_path из job_paths.csv (обычный файл, не боевая централь).
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
ACTIVE_MARKER = None


def ensure_csv():
    global ACTIVE_MARKER
    csv_dir = lib.shared_reports_dir()
    ACTIVE_MARKER = os.path.join(csv_dir, "_active_detach_to_buffer.txt")
    stamp = time.strftime("%Y-%m-%d_%H%M")
    path = os.path.join(csv_dir, "detach_to_buffer_{}.csv".format(stamp))
    if lib.file_exists(ACTIVE_MARKER):
        try:
            p = lib.read_all_text(ACTIVE_MARKER).strip()
            if p and lib.file_exists(p):
                return p
        except Exception:
            pass
    lib.write_all_text(ACTIVE_MARKER, path)
    lib.write_all_text(
        path, "old_path;staging_path;status;message;script_version\n"
    )
    return path


def save_detached(document, dest_path):
    folder = os.path.dirname(dest_path)
    lib.ensure_dir(folder)
    model_path = ModelPathUtils.ConvertUserVisiblePathToModelPath(dest_path)
    opts = SaveAsOptions()
    opts.OverwriteExistingFile = True
    try:
        opts.Compact = False
    except Exception:
        pass
    detached = False
    try:
        detached = bool(document.IsDetached)
    except Exception:
        detached = False
    # Buffer must not become a live central on the RBP PC.
    if document.IsWorkshared and not detached:
        wopts = WorksharingSaveAsOptions()
        wopts.SaveAsCentral = False
        opts.SetWorksharingOptions(wopts)
    document.SaveAs(model_path, opts)
    Output("Saved detached file -> {}".format(dest_path))


doc = revit_script_util.GetScriptDocument()
revitFilePath = revit_script_util.GetRevitFilePath()

Output()
Output("=== detach_to_buffer v{} ===".format(SCRIPT_VERSION))
Output("Script dir: {}".format(SCRIPT_DIR))
Output("Old path: {}".format(revitFilePath))
try:
    Output(
        "Workshared={}, Detached={}".format(
            doc.IsWorkshared, getattr(doc, "IsDetached", False)
        )
    )
except Exception:
    pass

csv_path = ensure_csv()

try:
    by_name, by_old, by_stg = lib.load_job_paths(JOB_PATHS)
    row = lib.find_job_row(revitFilePath, by_old, by_stg, by_name)
    if not row or not row.get("staging_path"):
        raise Exception("No staging_path in job_paths.csv for this model")
    dest = row["staging_path"]
    Output("Buffer: {}".format(dest))
    try:
        detached = bool(getattr(doc, "IsDetached", False))
    except Exception:
        detached = False
    if not detached:
        Output("WARN: document IsDetached=False. Set RBP Detach = ON.")
    save_detached(doc, dest)
    lib.append_csv(
        csv_path,
        [revitFilePath, dest, "OK", "", SCRIPT_VERSION],
    )
    Output("Done buffer OK")
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
