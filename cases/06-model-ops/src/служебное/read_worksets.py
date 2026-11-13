# -*- coding: utf-8 -*-
# Name: read_worksets.py
# Version: 1.0.1
# What it does: See the case README.
# Notes: Paths and project names in this portfolio copy are placeholders.

"""
RBP read-only pass: user worksets and Revit links of one model.

Writes reports/workset_read_<stamp>.csv (UTF-8 with BOM) and the pointer
reports/_active_workset_read.txt. Does not open a transaction, save, sync,
or write element parameters. Revit 2021-2025, IronPython 2.7.

The window's fast read uses --worksets close_all. Link elements that sit in a
closed workset are not loaded, and this script does not count elements.
Full read uses --worksets open_all. Manual full run:

  BatchRvt.exe --task_script "<tool>\\служебное\\read_worksets.py" ^
    --file_list "<one path per line>.txt" ^
    --create_new_local --worksets open_all --revit_version 2024

Every UserWorkset is written, including Workset1 and the default
"30_ИОС_Элементы". Each link row carries the type workset and, when an
instance exists, the instance workset. A model that is not workshared is
marked with row_kind not_workshared and is not given invented worksets.
"""

from __future__ import print_function

import clr
import os
import sys
import time

clr.AddReference("RevitAPI")
clr.AddReference("System")
import System
from System.IO import FileAccess, FileMode, FileShare, FileStream, StreamWriter
from System.Text import UTF8Encoding
from Autodesk.Revit.DB import (
    Element,
    ExternalFileUtils,
    FilteredElementCollector,
    FilteredWorksetCollector,
    ModelPathUtils,
    RevitLinkInstance,
    RevitLinkType,
    WorksetKind,
)

import revit_script_util
from revit_script_util import Output

SCRIPT_VERSION = "1.0.1"
RBP_ROOT = os.path.join(
    os.environ.get("USERPROFILE", r"C:\Users\Public"),
    r"Documents\doc\script\rbp",
)
TOOL_DIR_FALLBACK = os.path.join(RBP_ROOT, "batch_model_ops")
TOOL_DIR_SHARE = r""
ACTIVE_NAME = "_active_workset_read.txt"


def get_script_dir():
    try:
        code = os.path.dirname(os.path.abspath(__file__))
        # code sits in <tool>\sluzhebnoe; job files are in the tool root one level up
        for folder in (os.path.dirname(code), code):
            if folder and os.path.isfile(os.path.join(folder, "job_paths.csv")):
                return folder
            if folder and os.path.isfile(os.path.join(folder, "mapping.csv")):
                return folder
    except Exception:
        pass
    for folder in (TOOL_DIR_FALLBACK, os.getcwd(), TOOL_DIR_SHARE):
        if folder and (
            os.path.isfile(os.path.join(folder, "job_paths.csv"))
            or os.path.isfile(os.path.join(folder, "mapping.csv"))
        ):
            return folder
    return TOOL_DIR_FALLBACK


SCRIPT_DIR = get_script_dir()  # tool root: job_paths.csv, mapping.csv, job.json
try:
    CODE_DIR = os.path.dirname(os.path.abspath(__file__))  # ops_lib.py, workset_plan.py
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
import workset_plan as wp

_rbp_output = Output


def Output(*args):
    """RBP log is ASCII only. Paths and Cyrillic names go to the CSV."""
    if args:
        _rbp_output(lib.ascii_safe(args[0]))
    else:
        _rbp_output()


def write_text(path, text, append):
    """UTF-8. BOM only when the file is created, never on append."""
    enc = UTF8Encoding(not append)
    mode = FileMode.Append if append else FileMode.Create
    stream = FileStream(path, mode, FileAccess.Write, FileShare.Read)
    writer = StreamWriter(stream, enc)
    try:
        writer.Write(wp.as_text(text))
        writer.Flush()
    finally:
        writer.Dispose()
        stream.Dispose()


def id_text(element_id):
    if element_id is None:
        return u""
    try:
        return wp.as_text(int(element_id.IntegerValue))
    except Exception:
        pass
    try:
        return wp.as_text(int(element_id.Value))
    except Exception:
        return u""


def sort_text(value):
    text = wp.as_text(value)
    try:
        return text.lower()
    except Exception:
        return text


def element_name(element):
    if element is None:
        return u""
    try:
        return Element.Name.__get__(element)
    except Exception:
        try:
            return element.Name
        except Exception:
            return u""


def element_workset_name(document, element):
    if element is None or document is None:
        return u""
    try:
        if not document.IsWorkshared:
            return u""
    except Exception:
        return u""
    try:
        table = document.GetWorksetTable()
        ws = table.GetWorkset(element.WorksetId)
        if ws is None:
            return u""
        return ws.Name
    except Exception:
        return u""


def get_link_path(document, link_type):
    if link_type is None or document is None:
        return u""
    try:
        ext = ExternalFileUtils.GetExternalFileReference(document, link_type.Id)
        if ext is not None:
            path = ModelPathUtils.ConvertModelPathToUserVisiblePath(ext.GetAbsolutePath())
            if path:
                return path
    except Exception:
        pass
    try:
        if ExternalFileUtils.IsExternalFileReference(link_type):
            ext = link_type.GetExternalFileReference()
            return ModelPathUtils.ConvertModelPathToUserVisiblePath(ext.GetAbsolutePath())
    except Exception:
        pass
    try:
        return element_name(link_type)
    except Exception:
        return u""


def is_nested_link(link_type):
    if link_type is None:
        return False
    try:
        return bool(link_type.IsNestedLink)
    except Exception:
        return False


def is_loaded(link_type):
    if link_type is None:
        return False
    try:
        status = str(link_type.GetLinkedFileStatus())
        return status == "Loaded" or status.endswith(".Loaded")
    except Exception:
        return False


def workset_rows(document, host_path):
    col = FilteredWorksetCollector(document).OfKind(WorksetKind.UserWorkset)
    try:
        items = list(col.ToWorksets())
    except Exception:
        items = list(col)
    named = []
    for ws in items:
        try:
            name = ws.Name
        except Exception:
            name = u""
        named.append((sort_text(name), name, ws))
    named.sort(key=lambda item: item[0])
    rows = []
    for _key, name, ws in named:
        try:
            ws_id = id_text(ws.Id)
        except Exception:
            ws_id = u""
        rows.append(wp.read_row(wp.ROW_WORKSET, host_path, ws_id, name))
    return rows


def link_rows(document, host_path):
    types = list(FilteredElementCollector(document).OfClass(RevitLinkType).ToElements())
    instances = list(
        FilteredElementCollector(document).OfClass(RevitLinkInstance).ToElements()
    )
    used = {}
    built = []
    for inst in instances:
        link_type = None
        type_id = u""
        try:
            type_id = id_text(inst.GetTypeId())
            link_type = document.GetElement(inst.GetTypeId())
        except Exception:
            link_type = None
        if type_id:
            used[type_id] = True
        path = get_link_path(document, link_type)
        name = element_name(link_type)
        if not name:
            name = element_name(inst)
        row = wp.read_row(
            wp.ROW_LINK,
            host_path,
            u"",
            u"",
            name,
            path,
            element_workset_name(document, inst),
            element_workset_name(document, link_type),
            u"1" if is_nested_link(link_type) else u"0",
            u"1" if is_loaded(link_type) else u"0",
        )
        built.append((sort_text(name), sort_text(path), row))
    for link_type in types:
        type_id = id_text(link_type.Id)
        if type_id and type_id in used:
            continue
        path = get_link_path(document, link_type)
        name = element_name(link_type)
        row = wp.read_row(
            wp.ROW_LINK,
            host_path,
            u"",
            u"",
            name,
            path,
            u"",
            element_workset_name(document, link_type),
            u"1" if is_nested_link(link_type) else u"0",
            u"1" if is_loaded(link_type) else u"0",
        )
        built.append((sort_text(name), sort_text(path), row))
    built.sort(key=lambda item: item[0] + u"\0" + item[1])
    return [item[2] for item in built]


def collect_rows(document, host_path):
    """Inventory rows. Reads only. A non-workshared model is marked and still
    lists its links; it has no user worksets to invent."""
    if document is None:
        raise Exception("no_document")
    workshared = False
    try:
        workshared = bool(document.IsWorkshared)
    except Exception:
        workshared = False
    rows = []
    if not workshared:
        rows.append(wp.read_row(wp.ROW_NOT_WORKSHARED, host_path))
    else:
        try:
            rows.extend(workset_rows(document, host_path))
        except Exception as ex:
            raise Exception(lib.ascii_safe(u"workset_read_failed: {0}".format(ex)))
    try:
        rows.extend(link_rows(document, host_path))
    except Exception as ex:
        raise Exception(lib.ascii_safe(u"link_read_failed: {0}".format(ex)))
    return rows


def ensure_report():
    folder = lib.shared_reports_dir(SCRIPT_DIR)
    pointer = os.path.join(folder, ACTIVE_NAME)
    if lib.file_exists(pointer):
        try:
            current = lib.strip_bom(lib.read_all_text(pointer).strip()).strip()
            if current and lib.file_exists(current):
                return current
        except Exception:
            pass
    stamp = time.strftime("%Y-%m-%d_%H%M")
    path = os.path.join(folder, "workset_read_{0}.csv".format(stamp))
    write_text(path, wp.header_line() + u"\n", False)
    write_text(pointer, path + u"\n", False)
    return path


def append_rows(path, rows):
    if not rows:
        return
    chunks = []
    for row in rows:
        chunks.append(wp.format_row(row))
        chunks.append(u"\n")
    write_text(path, u"".join(chunks), True)


PROGRESS_NAME = u"_progress_ops.txt"


def progress_mark(status):
    """UTF-8 line for the window: STATUS<TAB>time<TAB>path. The path stays out of Output()."""
    try:
        folder = lib.shared_reports_dir(SCRIPT_DIR)
        line = u"\t".join([wp.as_text(status), time.strftime("%H:%M:%S"), wp.as_text(revitFilePath)]) + u"\n"
        lib.append_all_text(os.path.join(folder, PROGRESS_NAME), line)
    except Exception as pex:
        try:
            Output("progress_mark failed: {0}".format(pex))
        except Exception:
            pass


doc = revit_script_util.GetScriptDocument()
revitFilePath = revit_script_util.GetRevitFilePath()

Output()
Output("=== read_worksets v{0} ===".format(SCRIPT_VERSION))
Output("read only: no transaction, no save, no sync")
progress_mark(u"STARTED")

try:
    report = ensure_report()
    rows = collect_rows(doc, revitFilePath)
    append_rows(report, rows)
    n_ws = 0
    n_link = 0
    marked = 0
    for row in rows:
        if row[0] == wp.ROW_WORKSET:
            n_ws += 1
        elif row[0] == wp.ROW_LINK:
            n_link += 1
        elif row[0] == wp.ROW_NOT_WORKSHARED:
            marked = 1
    Output(
        "rows={0} worksets={1} links={2} not_workshared={3}".format(
            len(rows), n_ws, n_link, marked
        )
    )
    progress_mark(u"OK")
    Output("OK")
except Exception as ex:
    progress_mark(u"ERROR")
    Output("ERROR: {0}".format(ex))
    try:
        report = ensure_report()
        append_rows(
            report,
            [
                wp.read_row(
                    wp.ROW_ERROR,
                    revitFilePath,
                    u"",
                    u"",
                    wp.as_text(ex),
                )
            ],
        )
    except Exception as write_ex:
        Output("ERROR: report_write_failed: {0}".format(write_ex))
