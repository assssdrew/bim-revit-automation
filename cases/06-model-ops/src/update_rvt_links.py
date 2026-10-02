# -*- coding: utf-8 -*-
"""
RBP Pass 2 — обновление путей RVT Links на новые имена.

Список = НОВЫЕ имена (rvt_list_new.txt после build_new_list.ps1).
Только связи из mapping.csv (old basename).
После успешного батча — удаление старых: delete_old_models.cmd
  (UNC авто; RSN — список to_delete_rsn.txt для ручного удаления в Admin).
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


SCRIPT_DIR = get_script_dir()
if SCRIPT_DIR not in sys.path:
    sys.path.append(SCRIPT_DIR)

import ops_lib as lib

MAPPING_PATH = os.path.join(SCRIPT_DIR, "mapping.csv")
JOB_PATHS = os.path.join(SCRIPT_DIR, "job_paths.csv")
REPORTS_DIR = None
ACTIVE_SUMMARY = None
ACTIVE_DETAILS = None


def ensure_report_files():
    global REPORTS_DIR, ACTIVE_SUMMARY, ACTIVE_DETAILS
    REPORTS_DIR = lib.shared_reports_dir()
    ACTIVE_SUMMARY = os.path.join(REPORTS_DIR, "_active_update_rvt_links.txt")
    ACTIVE_DETAILS = os.path.join(REPORTS_DIR, "_active_update_rvt_links_details.txt")
    Output("Reports dir: {}".format(REPORTS_DIR))

    stamp = time.strftime("%Y-%m-%d_%H%M")
    summary = os.path.join(REPORTS_DIR, "update_rvt_links_{}.csv".format(stamp))
    details = os.path.join(REPORTS_DIR, "update_rvt_links_{}_details.csv".format(stamp))

    if lib.file_exists(ACTIVE_SUMMARY):
        try:
            p = lib.read_all_text(ACTIVE_SUMMARY).strip()
            if p and lib.file_exists(p):
                summary = p
        except Exception:
            pass
    else:
        lib.write_all_text(ACTIVE_SUMMARY, summary)
        lib.write_all_text(
            summary,
            "host;status;updated;skipped;errors;issues;script_version\n",
        )

    if lib.file_exists(ACTIVE_DETAILS):
        try:
            p = lib.read_all_text(ACTIVE_DETAILS).strip()
            if p and lib.file_exists(p):
                details = p
        except Exception:
            pass
    else:
        lib.write_all_text(ACTIVE_DETAILS, details)
        lib.write_all_text(
            details,
            "host;link_name;old_path;new_path;action;result;message\n",
        )

    return summary, details


def get_link_path(link_type):
    try:
        if ExternalFileUtils.IsExternalFileReference(link_type):
            ext = link_type.GetExternalFileReference()
            mp = ext.GetAbsolutePath()
            return ModelPathUtils.ConvertModelPathToUserVisiblePath(mp)
    except Exception:
        pass
    try:
        refs = link_type.GetExternalResourceReferences()
        if refs is not None:
            for k in refs.Keys:
                try:
                    er = refs[k]
                    if er is None:
                        continue
                    try:
                        p = er.InSessionPath
                        if p:
                            return p
                    except Exception:
                        pass
                except Exception:
                    continue
    except Exception:
        pass
    try:
        return Element.Name.__get__(link_type)
    except Exception:
        try:
            return link_type.Name
        except Exception:
            return ""


def get_link_type_name(link_type):
    try:
        return Element.Name.__get__(link_type)
    except Exception:
        try:
            return link_type.Name
        except Exception:
            return "?"


def is_nested(link_type):
    try:
        return bool(link_type.IsNestedLink)
    except Exception:
        return False


def remap_link(old_path, lname, mapping, by_name):
    """Prefer full new_path from job_paths (avoids bare filename -> Program Files)."""
    base = lib.basename_no_ext(old_path) or lib.basename_no_ext(lname)
    key = (base or "").upper()
    if key and key in by_name:
        row = by_name[key]
        return row.get("new_path"), row.get("old_name") or base, row.get("new_name")
    if not old_path:
        return None, None, None
    if key not in mapping:
        name_base = lib.basename_no_ext(lname)
        if name_base.upper() in mapping:
            key = name_base.upper()
            base = name_base
        else:
            return None, None, None
    new_base = mapping[key]
    new_path = lib.replace_filename_in_path(old_path, base, new_base)
    if not new_path:
        if "\\" in old_path or "/" in old_path:
            folder = old_path.replace("\\", "/").rsplit("/", 1)[0]
            if folder and folder != old_path.replace("\\", "/"):
                new_path = folder + "/" + new_base + ".rvt"
        else:
            # basename-only stored path: do not LoadFrom a bare file name
            return None, base, new_base
    return new_path, base, new_base


def load_link_from_path(link_type, new_path):
    """
    Revit 2022 IronPython: LoadFrom requires ModelPath + WorksetConfiguration.
    (1-arg overload often raises: takes exactly 2 arguments (1 given)).
    """
    model_path = ModelPathUtils.ConvertUserVisiblePathToModelPath(new_path)
    try:
        wcfg = WorksetConfiguration(WorksetConfigurationOption.OpenAllWorksets)
    except Exception:
        wcfg = WorksetConfiguration()

    # Primary: 2-arg overload (required on this API/IronPython)
    try:
        return link_type.LoadFrom(model_path, wcfg)
    except Exception as ex1:
        # Fallback attempts
        try:
            return link_type.LoadFrom(model_path)
        except Exception as ex2:
            raise Exception(
                "LoadFrom failed: {0} | fallback: {1}".format(ex1, ex2)
            )


def link_load_ok(result):
    if result is None:
        return True
    try:
        return int(result.LoadResult) == int(LinkLoadResultType.LinkLoaded)
    except Exception:
        try:
            return str(result.LoadResult).endswith("LinkLoaded")
        except Exception:
            return True


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


def save_or_sync(document, path, comment):
    if lib.is_rsn_path(path):
        sync_with_relinquish(document, comment)
        return
    is_detached = False
    try:
        is_detached = document.IsDetached
    except Exception:
        is_detached = False
    if (
        document.IsWorkshared
        and (not is_detached)
        and document.GetWorksharingCentralModelPath() is not None
    ):
        sync_with_relinquish(document, comment)
        return
    opts = SaveAsOptions()
    opts.OverwriteExistingFile = True
    if document.IsWorkshared:
        wopts = WorksharingSaveAsOptions()
        wopts.SaveAsCentral = True
        opts.SetWorksharingOptions(wopts)
    document.SaveAs(path, opts)
    Output("SavedAs in place: {}".format(path))


doc = revit_script_util.GetScriptDocument()
revitFilePath = revit_script_util.GetRevitFilePath()
host_name = lib.basename_no_ext(revitFilePath)

Output()
Output("=== update_rvt_links v{} (Pass 2) ===".format(SCRIPT_VERSION))
Output("Script dir: {}".format(SCRIPT_DIR))
Output("Host file: {}".format(revitFilePath))

try:
    mapping = lib.load_mapping(MAPPING_PATH)
    by_name, by_old, by_stg = lib.load_job_paths(JOB_PATHS)
    Output("Mapping pairs: {}".format(len(mapping)))
    Output("Job path rows: {}".format(len(by_name)))
    summary_csv, details_csv = ensure_report_files()

    updated = 0
    skipped = 0
    errors = 0
    issues = []

    link_types = list(
        FilteredElementCollector(doc).OfClass(RevitLinkType).ToElements()
    )
    Output("RevitLinkType count: {}".format(len(link_types)))

    for lt in link_types:
        lname = get_link_type_name(lt)
        if is_nested(lt):
            skipped += 1
            Output("  skip nested: {}".format(lname))
            continue
        old_path = get_link_path(lt)
        new_path, old_base, new_base = remap_link(
            old_path, lname, mapping, by_name
        )

        if new_path is None:
            skipped += 1
            lib.append_csv(
                details_csv,
                [
                    revitFilePath,
                    lname,
                    old_path,
                    "",
                    "skip",
                    "OK",
                    "not in mapping",
                ],
            )
            continue

        if old_path and old_path.replace("\\", "/").lower() == new_path.replace(
            "\\", "/"
        ).lower():
            skipped += 1
            lib.append_csv(
                details_csv,
                [
                    revitFilePath,
                    lname,
                    old_path,
                    new_path,
                    "skip",
                    "OK",
                    "already new path",
                ],
            )
            continue

        try:
            Output("Remap: {} -> {}".format(old_path, new_path))
            result = load_link_from_path(lt, new_path)
            if link_load_ok(result):
                updated += 1
                lib.append_csv(
                    details_csv,
                    [
                        revitFilePath,
                        lname,
                        old_path,
                        new_path,
                        "LoadFrom",
                        "OK",
                        "",
                    ],
                )
                Output("  OK LoadFrom")
            else:
                errors += 1
                msg = "LoadResult={}".format(
                    getattr(result, "LoadResult", result)
                )
                issues.append("{}: {}".format(lname, msg))
                lib.append_csv(
                    details_csv,
                    [
                        revitFilePath,
                        lname,
                        old_path,
                        new_path,
                        "LoadFrom",
                        "ERROR",
                        msg,
                    ],
                )
                Output("  ERROR {}".format(msg))
        except Exception as ex:
            errors += 1
            msg = str(ex)
            issues.append("{}: {}".format(lname, msg))
            lib.append_csv(
                details_csv,
                [
                    revitFilePath,
                    lname,
                    old_path,
                    new_path,
                    "LoadFrom",
                    "ERROR",
                    msg,
                ],
            )
            Output("  ERROR {}".format(msg))

    status = "OK"
    if errors > 0 and updated > 0:
        status = "PARTIAL"
    elif errors > 0 and updated == 0:
        status = "ERROR"
    elif updated == 0:
        status = "NO_CHANGES"

    comment = "Pass2 update RVT links v{}: upd={} err={}".format(
        SCRIPT_VERSION, updated, errors
    )
    if updated > 0:
        save_or_sync(doc, revitFilePath, comment)
    else:
        Output("No link updates — skip Save/Sync")

    lib.append_csv(
        summary_csv,
        [
            revitFilePath,
            status,
            updated,
            skipped,
            errors,
            " | ".join(issues)[:500],
            SCRIPT_VERSION,
        ],
    )
    Output(
        "Done host={} status={} updated={} skipped={} errors={}".format(
            host_name, status, updated, skipped, errors
        )
    )
    Output("OK")
except Exception as ex:
    Output("ERROR: {}".format(ex))
    try:
        summary_csv, details_csv = ensure_report_files()
        lib.append_csv(
            summary_csv,
            [
                revitFilePath,
                "ERROR",
                0,
                0,
                1,
                str(ex)[:500],
                SCRIPT_VERSION,
            ],
        )
    except Exception:
        pass
    raise
