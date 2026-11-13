# -*- coding: utf-8 -*-
# Name: apply_workset_links.py
# Version: 1.1.3
# What it does: See the case README.
# Notes: Paths and project names in this portfolio copy are placeholders.

"""
RBP pass: rename user worksets, create missing ones, insert links.

Per host from workset_plan.csv: renames, then creates, then inserts.
Each change is its own transaction. A link is placed by shared coordinates
only. One sync at the end if anything was committed.
Does not write the model out under another name and does not turn sharing on.

Revit 2021-2025, IronPython 2.7. Output() and raise text are ASCII.
"""

from __future__ import print_function

import clr
import json
import os
import sys
import time

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("System")
import System
from System import EventHandler
from System.IO import FileAccess, FileMode, FileShare, FileStream, StreamWriter
from System.Text import UTF8Encoding
from Autodesk.Revit.DB import (
    BuiltInParameter,
    ExternalFileUtils,
    FailureProcessingResult,
    FailureSeverity,
    FilteredElementCollector,
    FilteredWorksetCollector,
    IFailuresPreprocessor,
    ImportPlacement,
    ModelPathUtils,
    RelinquishOptions,
    RevitLinkInstance,
    RevitLinkOptions,
    RevitLinkType,
    SynchronizeWithCentralOptions,
    TransactWithCentralOptions,
    Transaction,
    TransactionStatus,
    Workset,
    WorksetId,
    WorksetKind,
    WorksetTable,
)
from Autodesk.Revit.UI.Events import DialogBoxShowingEventArgs

import revit_script_util
from revit_script_util import Output

SCRIPT_VERSION = "1.1.3"
RBP_ROOT = os.path.join(
    os.environ.get("USERPROFILE", r"C:\Users\Public"),
    r"Documents\doc\script\rbp",
)
TOOL_DIR_FALLBACK = os.path.join(RBP_ROOT, "batch_model_ops")
TOOL_DIR_SHARE = r""
ACTIVE_NAME = "_active_workset_apply.txt"
SYNC_COMMENT = "worksets-links"
APPLY_HEADER = "host;action;key;target;result;code;message"


def get_script_dir():
    try:
        code = os.path.dirname(os.path.abspath(__file__))
        # code sits in <tool>\sluzhebnoe; job files are in the tool root one level up
        for folder in (os.path.dirname(code), code):
            if folder and os.path.isfile(os.path.join(folder, "job_paths.csv")):
                return folder
            if folder and os.path.isfile(os.path.join(folder, "mapping.csv")):
                return folder
            if folder and os.path.isfile(os.path.join(folder, "workset_plan.csv")):
                return folder
    except Exception:
        pass
    for folder in (TOOL_DIR_FALLBACK, os.getcwd(), TOOL_DIR_SHARE):
        if folder and (
            os.path.isfile(os.path.join(folder, "job_paths.csv"))
            or os.path.isfile(os.path.join(folder, "mapping.csv"))
            or os.path.isfile(os.path.join(folder, "workset_plan.csv"))
        ):
            return folder
    return TOOL_DIR_FALLBACK


SCRIPT_DIR = get_script_dir()  # tool root: job_paths.csv, mapping.csv, job.json
try:
    _THIS_FILE = os.path.abspath(__file__)
except Exception:
    _THIS_FILE = ""  # RBP may run the task without __file__
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
# our folders go to the FRONT of sys.path: a module of the same name elsewhere must not win
_front = 0
for _p in (CODE_DIR, SCRIPT_DIR):
    if not _p:
        continue
    while _p in sys.path:
        sys.path.remove(_p)
    sys.path.insert(_front, _p)
    _front += 1

# modules already loaded before this script (a cached module ignores sys.path): logged below
_PRELOADED = [m for m in ("ops_lib", "workset_plan") if m in sys.modules]

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


def id_text(workset_id):
    if workset_id is None:
        return u""
    try:
        return wp.as_text(int(workset_id.IntegerValue))
    except Exception:
        pass
    try:
        return wp.as_text(int(workset_id.Value))
    except Exception:
        return u""


def name_key(value):
    text = wp.as_text(value).strip()
    try:
        return text.lower()
    except Exception:
        return text


def load_user_worksets(document):
    col = FilteredWorksetCollector(document).OfKind(WorksetKind.UserWorkset)
    try:
        return list(col.ToWorksets())
    except Exception:
        return list(col)


def find_by_id(items, key):
    want = wp.as_text(key).strip()
    if not want:
        return None
    for ws in items:
        if id_text(ws.Id) == want:
            return ws
    return None


def find_by_name(items, name):
    want = name_key(name)
    if not want:
        return None
    for ws in items:
        try:
            current = ws.Name
        except Exception:
            continue
        if name_key(current) == want:
            return ws
    return None


def exc_code(ex, default_code):
    text = name_key(ex)
    if (
        "borrow" in text
        or "editable" in text
        or "owned" in text
        or "another user" in text
        or "checked out" in text
    ):
        return "SKIP_BORROW"
    if "exist" in text or "duplicate" in text or "same name" in text or "already" in text:
        return "SKIP_CONFLICT"
    if "invalid" in text or "empty" in text or "character" in text:
        return "SKIP_BAD_NAME"
    return default_code


def run_txn(document, title, action):
    txn = Transaction(document, title)
    txn.Start()
    try:
        action()
        txn.Commit()
        return True, None
    except Exception as ex:
        try:
            if txn.HasStarted() and not txn.HasEnded():
                txn.RollBack()
        except Exception:
            pass
        return False, ex


def plan_file():
    job_path = os.path.join(SCRIPT_DIR, "workset_job.json")
    name = "workset_plan.csv"
    if lib.file_exists(job_path):
        try:
            data = json.loads(lib.read_all_text(job_path))
            if data.get("plan"):
                name = wp.as_text(data.get("plan"))
        except Exception:
            pass
    if os.path.isabs(name):
        return name
    return os.path.join(SCRIPT_DIR, name)


def log_environment():
    """Where the code and the job files come from (ASCII; other chars as \\uXXXX)."""
    Output("diag: __file__ = {0}".format(wp.ascii_escape(_THIS_FILE) if _THIS_FILE else "(none)"))
    Output("diag: SCRIPT_DIR = {0}".format(wp.ascii_escape(SCRIPT_DIR)))
    Output("diag: CODE_DIR = {0}".format(wp.ascii_escape(CODE_DIR)))
    Output("diag: wp = {0}".format(wp.ascii_escape(getattr(wp, "__file__", "?"))))
    Output("diag: lib = {0}".format(wp.ascii_escape(getattr(lib, "__file__", "?"))))
    Output("diag: preloaded = {0}".format(",".join(_PRELOADED) or "(none)"))
    Output("diag: python = {0}".format(wp.ascii_escape(sys.version.split("\n")[0])))
    for i, p in enumerate(sys.path):
        Output("diag: sys.path[{0}] = {1}".format(i, wp.ascii_escape(p)))


def log_plan_file(path, text):
    try:
        info = System.IO.FileInfo(path)
        Output("diag: plan = {0}".format(wp.ascii_escape(path)))
        Output("diag: plan size = {0}, mtime = {1}".format(
            info.Length, info.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")))
        Output("diag: plan first bytes = {0}".format(wp.hex_bytes(System.IO.File.ReadAllBytes(path), 16)))
        first = wp.as_text(text).replace(u"\r\n", u"\n").split(u"\n")[0]
        Output("diag: plan first line (read) = {0}".format(wp.ascii_escape(first[:120])))
    except Exception as ex:
        Output("diag: plan info failed: {0}".format(wp.ascii_escape(ex)))


def load_plan_rows():
    path = plan_file()
    if not lib.file_exists(path):
        raise Exception("plan_not_found")
    text = lib.read_all_text(path)
    log_plan_file(path, text)
    header, rows = wp.parse_csv(text)
    Output("diag: " + wp.header_log_line(header))
    missing = wp.missing_columns(header)
    if missing:
        raise Exception(lib.ascii_safe("bad_plan_header: missing " + ",".join(missing)))
    return path, rows


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
    path = os.path.join(folder, "workset_apply_{0}.csv".format(stamp))
    write_text(path, APPLY_HEADER + u"\n", False)
    write_text(pointer, path + u"\n", False)
    return path


def append_report(path, host, action, key, target, result, code, message):
    line = wp.format_row([host, action, key, target, result, code, message]) + u"\n"
    write_text(path, line, True)


def progress_mark(status):
    try:
        folder = lib.shared_reports_dir(SCRIPT_DIR)
        line = u"\t".join([
            wp.as_text(status),
            time.strftime("%H:%M:%S"),
            wp.as_text(revitFilePath),
        ]) + u"\n"
        lib.append_all_text(os.path.join(folder, u"_progress_ops.txt"), line)
    except Exception as pex:
        try:
            Output("progress_mark failed: {0}".format(pex))
        except Exception:
            pass


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


def sync_with_relinquish(document):
    is_detached = False
    try:
        is_detached = document.IsDetached
    except Exception:
        is_detached = False
    if is_detached:
        raise Exception("DETACHED - use Create New Local (Detach OFF).")
    if not document.IsWorkshared:
        raise Exception("Expected workshared for Sync.")
    sync_opts = SynchronizeWithCentralOptions()
    sync_opts.SetRelinquishOptions(make_relinquish_all_options())
    try:
        sync_opts.Comment = SYNC_COMMENT
    except Exception:
        pass
    try:
        sync_opts.SaveLocalFile = True
    except Exception:
        pass
    document.SynchronizeWithCentral(TransactWithCentralOptions(), sync_opts)
    Output("SynchronizedWithCentral + relinquish OK")


def record_skip(report, row, code):
    append_report(
        report,
        row.get("host") or revitFilePath,
        row.get("action") or "",
        row.get("key") or "",
        row.get("arg1") or row.get("arg2") or "",
        "SKIP",
        code,
        row.get("code") or "",
    )


def apply_rename(document, items, row):
    new_name = wp.as_text(row.get("arg2") or u"").strip()
    old_name = wp.as_text(row.get("arg1") or u"")
    key = wp.as_text(row.get("key") or u"").strip()
    if not new_name:
        return items, "SKIP", "SKIP_BAD_NAME", "empty"
    ws = find_by_id(items, key)
    if ws is None:
        return items, "SKIP", "SKIP_CONFLICT", "id_missing"
    try:
        current = wp.as_text(ws.Name)
    except Exception:
        current = u""
    if name_key(current) == name_key(new_name):
        return items, "OK", "", "already"
    other = find_by_name(items, new_name)
    if other is not None and id_text(other.Id) != id_text(ws.Id):
        return items, "SKIP", "SKIP_CONFLICT", "name_taken"
    ws_id = ws.Id

    def do_rename():
        WorksetTable.RenameWorkset(document, ws_id, new_name)

    ok, ex = run_txn(document, "ws-rename", do_rename)
    if not ok:
        return items, "SKIP", exc_code(ex, "SKIP_CONFLICT"), "rename_failed"
    return load_user_worksets(document), "OK", "", ""


class _LinkFailures(IFailuresPreprocessor):
    """Roll the insert back. Do not delete a warning and then commit it."""

    def __init__(self):
        self.coords = False
        self.error = False

    def PreprocessFailures(self, accessor):
        rollback = False
        try:
            messages = accessor.GetFailureMessages()
        except Exception:
            return FailureProcessingResult.Continue
        for msg in messages:
            desc = u""
            try:
                desc = wp.as_text(msg.GetDescriptionText())
            except Exception:
                desc = u""
            low = name_key(desc)
            sev = None
            try:
                sev = msg.GetSeverity()
            except Exception:
                sev = None
            if _is_coord_text(low):
                self.coords = True
                rollback = True
            elif sev == FailureSeverity.Error:
                self.error = True
                rollback = True
        if rollback:
            return FailureProcessingResult.ProceedWithRollBack
        return FailureProcessingResult.Continue


_dialog_cancelled = {"on": False}


def _cancel_dialog(sender, args):
    _dialog_cancelled["on"] = True
    try:
        args.OverrideResult(2)  # IDCANCEL
    except Exception:
        pass


def _is_coord_text(low):
    if not low:
        return False
    if "origin" in low:
        return True
    if "shared" in low and "coord" in low:
        return True
    if u"координат" in low:
        return True
    return False


def _link_path(document, link_type):
    if link_type is None or document is None:
        return u""
    try:
        ext = ExternalFileUtils.GetExternalFileReference(document, link_type.Id)
        if ext is not None:
            path = ModelPathUtils.ConvertModelPathToUserVisiblePath(ext.GetAbsolutePath())
            if path:
                return wp.as_text(path)
    except Exception:
        pass
    try:
        return wp.as_text(link_type.Name)
    except Exception:
        return u""


def _existing_links(document):
    found = []
    types = {}
    try:
        for lt in FilteredElementCollector(document).OfClass(RevitLinkType).ToElements():
            types[id_text(lt.Id)] = lt
    except Exception:
        types = {}
    try:
        instances = list(FilteredElementCollector(document).OfClass(RevitLinkInstance).ToElements())
    except Exception:
        instances = []
    for inst in instances:
        tid = u""
        try:
            tid = id_text(inst.GetTypeId())
        except Exception:
            tid = u""
        lt = types.get(tid)
        nested = False
        if lt is not None:
            try:
                nested = bool(lt.IsNestedLink)
            except Exception:
                nested = False
        path = _link_path(document, lt)
        found.append({
            "path": path,
            "name": wp.basename_no_ext(path),
            "nested": nested,
        })
    return found


def _file_missing(path):
    text = wp.as_text(path).strip()
    if not text:
        return True
    up = text.upper()
    if up.startswith("BIM 360://") or up.startswith("BIM360://") or up.startswith("ACC://"):
        return True
    if up.startswith("RSN://"):
        return False
    try:
        return not System.IO.File.Exists(text)
    except Exception:
        return False


def _set_partition(element, workset):
    param = element.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM)
    if param is None or param.IsReadOnly:
        raise Exception("partition_readonly")
    param.Set(int(workset.Id.IntegerValue))
    current = id_text(element.WorksetId)
    if current != id_text(workset.Id):
        raise Exception("partition_not_set")


def _element_present(document, element_id):
    if document is None or element_id is None:
        return False
    try:
        return document.GetElement(element_id) is not None
    except Exception:
        return False


def _created_link_present(document, inst_id, type_id):
    return _element_present(document, inst_id) or _element_present(document, type_id)


def _remove_created_link(document, inst_id, type_id):
    """Delete the instance and the type in their own transaction.

    True only when neither id is left in the model.
    """

    def do_delete():
        if _element_present(document, inst_id):
            document.Delete(inst_id)
        if _element_present(document, type_id):
            document.Delete(type_id)

    ok, _ex = run_txn(document, "ws-insert-undo", do_delete)
    if not ok:
        return False
    return not _created_link_present(document, inst_id, type_id)


def apply_insert(document, items, row):
    path = wp.as_text(row.get("arg1") or u"").strip()
    ws_name = wp.as_text(row.get("arg2") or u"").strip()
    if not path or not ws_name:
        return items, "SKIP", "SKIP_BAD_NAME", "empty"
    up = path.upper()
    if up.startswith("BIM 360://") or up.startswith("BIM360://") or up.startswith("ACC://"):
        return items, "SKIP", "SKIP_NOT_FOUND", "unsupported_path"
    status = wp.classify_link(path, _existing_links(document))
    if status == wp.STATUS_EXISTS:
        return items, "SKIP", "SKIP_EXISTS", ""
    if status == wp.STATUS_OTHER_PATH:
        return items, "SKIP", "SKIP_OTHER_PATH", ""
    if _file_missing(path):
        return items, "SKIP", "SKIP_NOT_FOUND", "file_missing"
    prep = _LinkFailures()
    _dialog_cancelled["on"] = False
    created_inst = [None]
    created_type = [None]
    pin_error = [""]

    def do_insert():
        model_path = ModelPathUtils.ConvertUserVisiblePathToModelPath(path)
        options = RevitLinkOptions(False)
        loaded = RevitLinkType.Create(document, model_path, options)
        type_id = None if loaded is None else loaded.ElementId
        if type_id is None:
            raise Exception("link_type_missing")
        inst = RevitLinkInstance.Create(document, type_id, ImportPlacement.Shared)
        if inst is None:
            raise Exception("link_instance_missing")
        created_type[0] = type_id
        created_inst[0] = inst.Id
        fresh = load_user_worksets(document)
        ws = find_by_name(fresh, ws_name)
        if ws is None:
            Workset.Create(document, ws_name)
            fresh = load_user_worksets(document)
            ws = find_by_name(fresh, ws_name)
        if ws is None:
            raise Exception("workset_missing")
        link_type = document.GetElement(type_id)
        _set_partition(link_type, ws)
        _set_partition(inst, ws)
        # inserted links are always pinned (same transaction); a failed pin keeps the link
        try:
            inst.Pinned = True
        except Exception as pin_ex:
            pin_error[0] = lib.ascii_safe(wp.as_text(pin_ex).strip()) or "unknown"

    def after_commit(committed):
        flagged = prep.coords or _dialog_cancelled["on"]
        link_present = False
        removed = False
        if committed and flagged:
            link_present = _created_link_present(document, created_inst[0], created_type[0])
            if link_present:
                removed = _remove_created_link(document, created_inst[0], created_type[0])
        result, code, why = wp.insert_after_commit(
            committed, prep.coords, _dialog_cancelled["on"], link_present, removed,
        )
        why = wp.with_pin_note(result, why, pin_error[0])
        if result == "OK" and pin_error[0]:
            Output("WARN: link inserted but not pinned: " + lib.ascii_safe(pin_error[0]))
        return load_user_worksets(document), result, code, why

    uiapp = None
    handler = None
    try:
        uiapp = revit_script_util.GetUIApplication()
    except Exception:
        uiapp = None
    hooked = {"on": False}

    def unhook():
        if uiapp is None or handler is None or not hooked["on"]:
            return
        try:
            uiapp.DialogBoxShowing -= handler
        except Exception:
            pass
        hooked["on"] = False

    if uiapp is not None:
        handler = EventHandler[DialogBoxShowingEventArgs](_cancel_dialog)
        uiapp.DialogBoxShowing += handler
        hooked["on"] = True
    txn = Transaction(document, "ws-insert")
    try:
        txn.Start()
        opts = txn.GetFailureHandlingOptions()
        opts.SetFailuresPreprocessor(prep)
        try:
            opts.SetClearAfterRollback(True)
        except Exception:
            pass
        txn.SetFailureHandlingOptions(opts)
        do_insert()
        action, result, code, why = wp.insert_before_commit(
            prep.coords, _dialog_cancelled["on"],
        )
        if action == "rollback":
            if txn.HasStarted() and not txn.HasEnded():
                txn.RollBack()
            return load_user_worksets(document), result, code, why
        status_txn = txn.Commit()
        unhook()
        committed = status_txn == TransactionStatus.Committed
        if not committed and txn.HasStarted() and not txn.HasEnded():
            txn.RollBack()
        return after_commit(committed)
    except Exception as ex:
        unhook()
        ended_committed = False
        try:
            if txn.HasStarted() and not txn.HasEnded():
                txn.RollBack()
            elif txn.HasStarted():
                ended_committed = txn.GetStatus() == TransactionStatus.Committed
        except Exception:
            ended_committed = False
        low = name_key(ex)
        flagged = prep.coords or _dialog_cancelled["on"] or _is_coord_text(low)
        if ended_committed and (prep.coords or _dialog_cancelled["on"]):
            return after_commit(True)
        if flagged:
            coords = prep.coords or _is_coord_text(low)
            result, code, why = wp.insert_after_commit(
                False, coords, _dialog_cancelled["on"], False, False,
            )
            return items, result, code, why
        if "not found" in low or "does not exist" in low or "cannot find" in low or u"не найден" in low:
            return items, "SKIP", "SKIP_NOT_FOUND", "file_missing"
        code = exc_code(ex, "ERROR_LINK")
        result = "ERROR" if code == "ERROR_LINK" else "SKIP"
        return items, result, code, "place_failed"
    finally:
        unhook()


def apply_create(document, items, row):
    name = wp.as_text(row.get("arg1") or u"").strip()
    if not name:
        return items, "SKIP", "SKIP_BAD_NAME", "empty"
    if find_by_name(items, name) is not None:
        return items, "OK", "", "already"

    def do_create():
        Workset.Create(document, name)

    ok, ex = run_txn(document, "ws-create", do_create)
    if not ok:
        code = exc_code(ex, "ERROR_CREATE")
        result = "ERROR" if code == "ERROR_CREATE" else "SKIP"
        return items, result, code, "create_failed"
    return load_user_worksets(document), "OK", "", ""


def _move_workset_elements(document, source, target_name):
    """Move elements onto target_name. True when the transaction commits."""
    from Autodesk.Revit.DB import ElementWorksetFilter, BuiltInParameter as _Bip

    def do_move():
        target = find_by_name(load_user_worksets(document), target_name)
        if target is None:
            Workset.Create(document, target_name)
            target = find_by_name(load_user_worksets(document), target_name)
        if target is None:
            raise Exception("workset_missing")
        filt = ElementWorksetFilter(source.Id)
        elements = list(FilteredElementCollector(document).WherePasses(filt).ToElements())
        for element in elements:
            param = element.get_Parameter(_Bip.ELEM_PARTITION_PARAM)
            if param is None or param.IsReadOnly:
                continue
            param.Set(int(target.Id.IntegerValue))

    ok, _ex = run_txn(document, "ws-move-elements", do_move)
    return bool(ok)


def apply_delete(document, items, row):
    ws_key = wp.as_text(row.get("key") or u"").strip()
    name = wp.as_text(row.get("arg1") or u"").strip()
    mode = wp.as_text(row.get("arg2") or u"move").strip() or u"move"
    target_name = wp.as_text(row.get("arg3") or u"").strip()
    if not target_name:
        target_name = u"30_\u0418\u041e\u0421_\u042d\u043b\u0435\u043c\u0435\u043d\u0442\u044b"
    if name and name_key(name) == name_key(target_name):
        return items, "SKIP", "SKIP_CANNOT_DELETE", "default_workset"
    source = find_by_id(items, ws_key)
    if source is None:
        source = find_by_name(items, name)
    if source is None:
        return items, "SKIP", "SKIP_NOT_FOUND", "workset_missing"
    api = hasattr(WorksetTable, "DeleteWorkset") and hasattr(WorksetTable, "CanDeleteWorkset")
    if not api:
        moved = _move_workset_elements(document, source, target_name)
        result, code, why = wp.delete_outcome(False, mode, False, False, moved, False)
        return load_user_worksets(document), result, code, why

    from Autodesk.Revit.DB import DeleteWorksetOption, DeleteWorksetSettings, WorksharingUtils
    from System.Collections.Generic import List as _NetList

    borrowed = {"on": False}
    cannot = {"on": False}

    def do_delete():
        bag = _NetList[WorksetId]()
        bag.Add(source.Id)
        try:
            WorksharingUtils.CheckoutWorksets(document, bag)
        except Exception:
            borrowed["on"] = True
            raise Exception("borrowed")
        target = find_by_name(load_user_worksets(document), target_name)
        if mode == "elements":
            settings = DeleteWorksetSettings(DeleteWorksetOption.DeleteAllElements)
        else:
            if target is None:
                Workset.Create(document, target_name)
                target = find_by_name(load_user_worksets(document), target_name)
            if target is None:
                raise Exception("workset_missing")
            settings = DeleteWorksetSettings(DeleteWorksetOption.MoveElementsToWorkset, target.Id)
        if not WorksetTable.CanDeleteWorkset(document, source.Id, settings):
            cannot["on"] = True
            raise Exception("cannot_delete")
        WorksetTable.DeleteWorkset(document, source.Id, settings)

    ok, ex = run_txn(document, "ws-delete", do_delete)
    low = name_key(ex)
    if borrowed["on"] or "borrow" in low or "owned" in low or "editable" in low:
        result, code, why = wp.delete_outcome(True, mode, False, True, False, False)
    elif cannot["on"] or "cannot_delete" in low:
        result, code, why = wp.delete_outcome(True, mode, False, False, False, False)
    else:
        result, code, why = wp.delete_outcome(True, mode, True, False, False, bool(ok))
    return load_user_worksets(document), result, code, why


def apply_host(document, report, rows):
    committed = 0
    errors = 0
    if not document.IsWorkshared:
        for row in rows:
            append_report(
                report,
                row.get("host") or revitFilePath,
                row.get("action") or "",
                row.get("key") or "",
                row.get("arg2") or row.get("arg1") or "",
                "SKIP",
                "SKIP_NOT_WORKSHARED",
                "",
            )
        return committed, errors, False
    items = load_user_worksets(document)
    renames = []
    creates = []
    deletes = []
    inserts = []
    rest = []
    for row in rows:
        action = wp.plan_action(row)
        if action == u"rename":
            renames.append(row)
        elif action == u"create":
            creates.append(row)
        elif action == u"delete":
            deletes.append(row)
        elif action == u"insert":
            inserts.append(row)
        else:
            rest.append(row)
    for row in renames:
        items, result, code, message = apply_rename(document, items, row)
        if result == "OK" and code != "already" and message != "already":
            committed += 1
        if result == "OK" and message == "already":
            code = ""
        if result == "ERROR":
            errors += 1
        append_report(
            report, row.get("host") or revitFilePath, "rename",
            row.get("key") or "", row.get("arg2") or "", result, code, message,
        )
    for row in creates:
        items, result, code, message = apply_create(document, items, row)
        if result == "OK" and message != "already":
            committed += 1
        if result == "ERROR":
            errors += 1
        append_report(
            report, row.get("host") or revitFilePath, "create",
            "", row.get("arg1") or "", result, code, message,
        )
    for row in deletes:
        items, result, code, message = apply_delete(document, items, row)
        if result == "OK":
            committed += 1
        if result == "ERROR":
            errors += 1
        append_report(
            report, row.get("host") or revitFilePath, "delete",
            row.get("key") or "", row.get("arg1") or "", result, code, message,
        )
    for row in inserts:
        items, result, code, message = apply_insert(document, items, row)
        if result == "OK":
            committed += 1
        if result == "ERROR":
            errors += 1
        append_report(
            report, row.get("host") or revitFilePath, "insert",
            row.get("key") or "", row.get("arg1") or "", result, code, message,
        )
    for row in rest:
        action = wp.as_text(row.get("action") or u"")
        if wp.plan_action(row) == u"skip":
            code = wp.as_text(row.get("code") or u"") or u"SKIP"
            append_report(
                report, row.get("host") or revitFilePath, action,
                row.get("key") or "", row.get("arg1") or row.get("arg2") or "",
                "SKIP", code, "",
            )
            continue
        # an empty or unknown action is a broken plan row: an error, never a quiet SKIP
        errors += 1
        Output("ERROR_BAD_ACTION: action={0}".format(wp.quoted(action)))
        append_report(
            report, row.get("host") or revitFilePath, action,
            row.get("key") or "", row.get("arg1") or row.get("arg2") or "",
            "ERROR", "ERROR_BAD_ACTION", u"action=" + wp.quoted(action),
        )
    return committed, errors, True


doc = revit_script_util.GetScriptDocument()
revitFilePath = revit_script_util.GetRevitFilePath()

Output()
Output("=== apply_workset_links v{0} ===".format(SCRIPT_VERSION))
Output("rename, create, insert by shared coordinates")
progress_mark(u"STARTED")

try:
    log_environment()
    _plan_path, plan_rows = load_plan_rows()
    mine = wp.rows_for_host(plan_rows, revitFilePath)
    Output("plan rows: {0}, for host: {1}".format(len(plan_rows), len(mine)))
    Output("plan rows for host: {0}".format(len(mine)))
    for _i, _row in enumerate(mine):
        Output("diag: " + wp.row_log_line(_i + 1, _row))
    report = ensure_report()
    committed, errors, workshared = apply_host(doc, report, mine)
    status = "NO_CHANGES"
    if workshared and committed > 0:
        try:
            sync_with_relinquish(doc)
            status = "PARTIAL" if errors else "OK"
        except Exception as ex:
            errors += 1
            status = "ERROR"
            append_report(report, revitFilePath, "sync", "", "", "ERROR", "ERROR_SYNC", "sync_failed")
            Output("ERROR: sync_failed: {0}".format(ex))
    elif errors:
        status = "ERROR"
    progress_mark(status)
    Output("status={0} committed={1} errors={2}".format(status, committed, errors))
except Exception as ex:
    progress_mark(u"ERROR")
    Output("ERROR: {0}".format(ex))
    try:
        report = ensure_report()
        fail_code = "ERROR_SYNC"
        fail_why = "apply_failed"
        if "bad_plan_header" in str(ex):
            fail_code = "ERROR_BAD_PLAN"
            fail_why = lib.ascii_safe(str(ex))
        append_report(report, revitFilePath, "model", "", "", "ERROR", fail_code, fail_why)
    except Exception as write_ex:
        Output("ERROR: report_write_failed: {0}".format(write_ex))
