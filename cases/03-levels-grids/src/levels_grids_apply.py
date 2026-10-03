# -*- coding: utf-8 -*-
# Name: levels_grids_apply.py
# Version: 1.0.2
# What it does: RBP apply safe level/grid edits on base files vs exemplar link (cfg-gated renames/curves).
# Inputs: rvt_list.txt, exemplar link, lg cfg flags, xlsx_out_path.cfg.
# Outputs: Modified hosts (transactional); CSV/Excel session output.
# How to run: RBP Apply script on selected base files only (see case README).
# Notes: Rename/elevation/grid moves controlled by cfg; Sync after changes.

import clr
import math
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


def is_rsn_path(path):
    if not path:
        return False
    return path.upper().startswith("RSN://")


def clear_readonly_attribute(path):
    if is_rsn_path(path):
        return
    try:
        attrs = System.IO.File.GetAttributes(path)
        readonly = System.IO.FileAttributes.ReadOnly
        if (attrs & readonly) == readonly:
            System.IO.File.SetAttributes(path, attrs & ~readonly)
            Output("Cleared Windows ReadOnly on file.")
    except Exception as ex:
        Output("WARN clear ReadOnly: {}".format(ex))


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


def sync_with_relinquish(document, path, comment):
    is_detached = False
    try:
        is_detached = document.IsDetached
    except Exception:
        is_detached = False

    if is_detached:
        raise Exception(
            "Document is DETACHED. For RSN:// and shared centrals use "
            "Create New Local in Batch Processor (Detach = OFF)."
        )

    if not document.IsWorkshared:
        raise Exception(
            "Expected workshared document for sync path: {}".format(path)
        )

    try:
        central = ModelPathUtils.ConvertModelPathToUserVisiblePath(
            document.GetWorksharingCentralModelPath()
        )
        Output("Sync to central: {}".format(central))
    except Exception:
        Output("Sync to central...")

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
    return path


def save_or_sync(document, path, comment):
    is_detached = False
    try:
        is_detached = document.IsDetached
    except Exception:
        is_detached = False

    if is_rsn_path(path):
        Output("Path type: Revit Server (RSN)")
        return sync_with_relinquish(document, path, comment)

    if (
        document.IsWorkshared
        and (not is_detached)
        and document.GetWorksharingCentralModelPath() is not None
    ):
        Output("Path type: workshared local/UNC")
        return sync_with_relinquish(document, path, comment)

    if is_detached and document.IsWorkshared:
        Output(
            "WARN: Detached workshared file — SaveAs in place. "
            "Prefer Create New Local for shared models."
        )

    clear_readonly_attribute(path)

    opts = SaveAsOptions()
    opts.OverwriteExistingFile = True
    opts.Compact = False

    if document.IsWorkshared:
        wopts = WorksharingSaveAsOptions()
        wopts.SaveAsCentral = True
        opts.SetWorksharingOptions(wopts)

    document.SaveAs(path, opts)
    Output("SavedAs in place: {}".format(path))
    return path


def cfg_flag(tolerances, key, default=False):
    v = tolerances.get(key, 1.0 if default else 0.0)
    try:
        return float(v) != 0.0
    except Exception:
        return bool(v)


def allow_apply(row_map, model_path, tolerances):
    """
    Apply safety:
      - by default allowed for any model in rvt_list (user picks only BF)
      - if project_map has the host with role=DISC → block
      - if apply_require_map_bf=1 → require map role=BF (strict)
    """
    if cfg_flag(tolerances, "apply_require_map_bf", False):
        if row_map and row_map.get("role") == "BF":
            return True, "role=BF (strict map)"
        return False, "apply_require_map_bf=1 but host not role=BF in project_map"

    if row_map and row_map.get("role") == "DISC":
        return False, "project_map role=DISC — Apply blocked (use Audit only)"

    return True, "list-selected model (put only BF in rvt_list)"


def apply_level_elevations(document, host_lv, ex_lv, tolerances, dry_actions):
    changed = 0
    for name, e_list in ex_lv.items():
        if name not in host_lv:
            dry_actions.append("SKIP create level '{}' (MVP)".format(name))
            continue
        h_list = host_lv[name]
        if len(h_list) != 1 or len(e_list) != 1:
            dry_actions.append(
                "SKIP ambiguous level '{}' host={} ex={}".format(
                    name, len(h_list), len(e_list)
                )
            )
            continue
        h0 = h_list[0]
        e0 = e_list[0]
        delta_mm = h0["elev_mm"] - e0["elev_mm"]
        if abs(delta_mm) <= float(tolerances["level_mm"]):
            continue
        try:
            h0["element"].Elevation = e0["elev_ft"]
            changed += 1
            dry_actions.append(
                "SET level '{}' elev {:.3f} -> {:.3f} mm".format(
                    name, h0["elev_mm"], e0["elev_mm"]
                )
            )
        except Exception as ex:
            dry_actions.append("FAIL level '{}': {}".format(name, ex))
    for name in host_lv:
        if name not in ex_lv:
            dry_actions.append("SKIP delete extra level '{}' (MVP)".format(name))
    return changed


def apply_grid_curves(document, host_gr, ex_gr, tolerances, dry_actions):
    """Move/rotate host grid to match exemplar mid+angle (no delete/create)."""
    changed = 0
    for name, e_list in ex_gr.items():
        if name not in host_gr:
            dry_actions.append("SKIP create grid '{}' (MVP)".format(name))
            continue
        h_list = host_gr[name]
        if len(h_list) != 1 or len(e_list) != 1:
            dry_actions.append(
                "SKIP ambiguous grid '{}' host={} ex={}".format(
                    name, len(h_list), len(e_list)
                )
            )
            continue
        h0 = h_list[0]
        e0 = e_list[0]
        if h0["mid_xy"] is None or e0["mid_xy"] is None:
            dry_actions.append("SKIP grid '{}' — no mid point".format(name))
            continue

        d_pos = lg_lib.distance_mm_xy(h0["mid_xy"], e0["mid_xy"])
        d_ang = lg_lib.angle_delta_deg(h0["angle_deg"], e0["angle_deg"])
        need_pos = d_pos is not None and d_pos > float(tolerances["grid_mm"])
        need_ang = d_ang is not None and d_ang > float(tolerances["grid_angle_deg"])
        if not need_pos and not need_ang:
            continue

        was_pinned = False
        try:
            was_pinned = h0["element"].Pinned
            if was_pinned:
                h0["element"].Pinned = False
        except Exception:
            was_pinned = False

        try:
            if need_pos:
                dx = e0["mid_xy"][0] - h0["mid_xy"][0]
                dy = e0["mid_xy"][1] - h0["mid_xy"][1]
                ElementTransformUtils.MoveElement(
                    document, h0["id"], XYZ(dx, dy, 0)
                )
            if need_ang and e0["angle_deg"] is not None and h0["angle_deg"] is not None:
                mx, my = e0["mid_xy"]
                axis = Line.CreateBound(XYZ(mx, my, -10), XYZ(mx, my, 10))
                ang = math.radians(float(e0["angle_deg"]) - float(h0["angle_deg"]))
                ElementTransformUtils.RotateElement(document, h0["id"], axis, ang)
            changed += 1
            dry_actions.append(
                "TRANSFORM grid '{}' pos_d={:.3f}mm ang_d={:.4f}deg".format(
                    name,
                    d_pos if d_pos is not None else -1,
                    d_ang if d_ang is not None else -1,
                )
            )
        except Exception as ex2:
            dry_actions.append("FAIL grid '{}': {}".format(name, ex2))
        finally:
            if was_pinned:
                try:
                    h0["element"].Pinned = True
                except Exception:
                    pass

    for name in host_gr:
        if name not in ex_gr:
            dry_actions.append("SKIP delete extra grid '{}' (MVP)".format(name))
    return changed


def run_apply(document, model_path, tolerances, project_rows):
    row_map = lg_lib.find_project_row(project_rows, model_path)
    ok, reason = allow_apply(row_map, model_path, tolerances)
    if not ok:
        raise Exception("Apply blocked: {}".format(reason))

    link_hint = lg_lib.resolve_exemplar_hint(SCRIPT_DIR, tolerances, row_map)
    inst, link_doc, label = lg_lib.find_exemplar_link(document, link_hint)
    if link_doc is None:
        raise Exception("Exemplar link not found: {}".format(label))

    Output("Exemplar link: {}".format(label))
    Output("Apply allowed: {}".format(reason))

    host_lv = lg_lib.collect_levels(document)
    ex_lv = lg_lib.collect_levels(link_doc)
    host_gr = lg_lib.collect_grids(document)
    ex_gr = lg_lib.collect_grids(link_doc)

    actions = []
    changed = 0

    t = Transaction(document, "Levels&Grids Apply v{}".format(SCRIPT_VERSION))
    t.Start()
    try:
        if cfg_flag(tolerances, "apply_elevation", True):
            changed += apply_level_elevations(
                document, host_lv, ex_lv, tolerances, actions
            )
        else:
            actions.append("SKIP elevations (apply_elevation=0)")

        if cfg_flag(tolerances, "apply_grid_curve", True):
            host_gr = lg_lib.collect_grids(document)
            changed += apply_grid_curves(
                document, host_gr, ex_gr, tolerances, actions
            )
        else:
            actions.append("SKIP grids (apply_grid_curve=0)")

        if cfg_flag(tolerances, "apply_rename", False):
            actions.append("SKIP rename — not implemented in MVP")
        else:
            actions.append("SKIP rename (apply_rename=0)")

        t.Commit()
    except Exception:
        if t.HasStarted():
            t.RollBack()
        raise

    for a in actions:
        Output("  " + a)
    Output("Changed elements: {}".format(changed))
    return changed, actions


# --- main ---
TOL = lg_lib.load_tolerances(CONFIG_PATH)
map_file = MAP_PATH if os.path.isfile(MAP_PATH) else None
PROJECT_ROWS = lg_lib.load_project_map(map_file) if map_file else []

Output()
Output("=== Levels & Grids Apply v{} ===".format(SCRIPT_VERSION))
Output("Script dir: {}".format(SCRIPT_DIR))
Output("Config: {}".format(CONFIG_PATH))
Output("Project map: {}".format(map_file or "(optional)"))
Output(
    "Exemplar hint: {!r}".format(
        lg_lib.resolve_exemplar_hint(SCRIPT_DIR, TOL, None)
    )
)
Output(
    "Flags: elev={} grid={} rename={}".format(
        cfg_flag(TOL, "apply_elevation", True),
        cfg_flag(TOL, "apply_grid_curve", True),
        cfg_flag(TOL, "apply_rename", False),
    )
)

doc = None
revitFilePath = ""

try:
    doc = revit_script_util.GetScriptDocument()
    revitFilePath = revit_script_util.GetRevitFilePath()
    Output("File: {}".format(revitFilePath))
    Output(
        "Workshared={}, Detached={}".format(
            doc.IsWorkshared, getattr(doc, "IsDetached", False)
        )
    )

    changed, actions = run_apply(doc, revitFilePath, TOL, PROJECT_ROWS)

    if changed > 0:
        comment = "Batch Levels&Grids Apply v{} ({} changes)".format(
            SCRIPT_VERSION, changed
        )
        saved = save_or_sync(doc, revitFilePath, comment)
        Output("Done: {}".format(saved))
    else:
        Output("No changes — skip Save/Sync")

    Output("OK")
except Exception as ex:
    Output("ERROR: {}".format(ex))
    raise
