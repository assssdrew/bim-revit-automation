# -*- coding: utf-8 -*-
# Name: levels_grids_apply.py
# Version: 1.23.0
# What it does: RBP apply of the operator's fix plan: level/grid edits and optional Project Base Point shift.
# Inputs: rvt_list.txt, fix_plan.csv, exemplar link, tolerances.cfg.
# Outputs: Modified hosts; Sync when the fix-plan checkbox is on.
# How to run: run_cascade.cmd items 2 or 4, then RBP on levels_grids_apply.py.
# Notes: Does not create or delete levels/grids. Survey Point and True North stay untouched.

"""
Revit Batch Processor — Levels & Grids Apply (без UI).

Сдвиг PBP по X/Y/Z из fix_plan. Модели — из пункта 2 (БФ) или пункта 4 (разделы).
Эталон AUTO: БФ→АР, раздел→БФ.

НЕ делает: удаление лишних, создание missing, Survey / True North.
Пакетный CR «Изменить» по осям не используется.

После правок: Sync+Relinquish (workshared/RSN) или SaveAs.
"""

import clr
import codecs
import math
import os
import re
import sys
import System

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
from Autodesk.Revit.DB import *

import revit_script_util
from revit_script_util import Output

SCRIPT_VERSION = "1.23.0"
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
FIX_PLAN_PATH = os.path.join(SCRIPT_DIR, "fix_plan.csv")


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
      - by default any model in rvt_list (BF from menu 2, sections from menu 4)
      - if apply_require_map_bf=1 → require map role=BF (legacy/strict)
    """
    if cfg_flag(tolerances, "apply_require_map_bf", False):
        if row_map and row_map.get("role") == "BF":
            return True, "role=BF (strict map)"
        return False, "apply_require_map_bf=1 but host not role=BF in project_map"

    return True, "list-selected model (rvt_list from fix plan)"


# CopyWatchAlert («Сетка перемещена»): toast without FailureResolution.
GUID_CR_GRID_MOVED = "c3aa4692-975c-4720-8cb4-389532cf43a4"


def _fid_guid(failure):
    try:
        return str(failure.GetFailureDefinitionId().Guid).lower()
    except Exception:
        return ""


def _is_copywatch_warning(failure):
    g = _fid_guid(failure)
    if g == GUID_CR_GRID_MOVED:
        return True
    try:
        g2 = str(BuiltInFailures.CopyMonitorFailures.CopyWatchAlert.Guid).lower()
        if g == g2:
            return True
    except Exception:
        pass
    try:
        desc = lg_lib.u(failure.GetDescriptionText() or "").lower()
    except Exception:
        desc = ""
    needles = (
        u"сетка перемещ",
        u"уровень перемещ",
        u"grid has been moved",
        u"level has been moved",
        u"мониторинг координац",
        u"copywatch",
    )
    for n in needles:
        if n in desc:
            return True
    return False


class CopyWatchDeletePreprocessor(IFailuresPreprocessor):
    """
    Delete CopyWatchAlert so the transaction can commit.
    Does not ResolveFailure — that API has no «Изменить» / «Переместить».
    """

    def __init__(self):
        self.deleted_count = 0
        self.message_count = 0

    def PreprocessFailures(self, failuresAccessor):
        try:
            messages = list(failuresAccessor.GetFailureMessages())
        except Exception:
            messages = []
        self.message_count += len(messages)
        for failure in messages:
            if not _is_copywatch_warning(failure):
                continue
            try:
                failuresAccessor.DeleteWarning(failure)
                self.deleted_count += 1
            except Exception:
                pass
        return FailureProcessingResult.Continue


def find_3d_view(document):
    found = None
    for v in FilteredElementCollector(document).OfClass(View3D):
        try:
            if v.IsTemplate:
                continue
        except Exception:
            continue
        try:
            if v.IsPerspective:
                continue
        except Exception:
            pass
        try:
            name = lg_lib.u(v.Name)
        except Exception:
            name = ""
        if name.startswith("{3D"):
            return v
        if found is None:
            found = v
    return found


def ensure_3d_view(document):
    v = find_3d_view(document)
    if v is not None:
        return v
    vft_id = None
    for vft in FilteredElementCollector(document).OfClass(ViewFamilyType):
        try:
            if vft.ViewFamily == ViewFamily.ThreeDimensional:
                vft_id = vft.Id
                break
        except Exception:
            continue
    if vft_id is None:
        return None
    return View3D.CreateIsometric(document, vft_id)


def _curves_in_view(grid, extent_type, view):
    raw = grid.GetCurvesInView(extent_type, view)
    if raw is None:
        return []
    try:
        return [c for c in raw if c is not None]
    except TypeError:
        return [raw]


def shift_grid_model_z(grid, dz, view3d):
    """
    Shift 3D model extents by dz (feet), keep XY.
    Do not translate Grid.Curve (plan) — unbound / off-plane error.
    """
    if view3d is None:
        raise Exception("no 3D view")
    if abs(float(dz)) < 1e-12:
        return 0
    curves = _curves_in_view(grid, DatumExtentType.Model, view3d)
    n = 0
    t = Transform.CreateTranslation(XYZ(0.0, 0.0, float(dz)))
    for c in curves:
        bound = False
        try:
            bound = bool(c.IsBound)
        except Exception:
            bound = False
        if not bound:
            continue
        new_c = None
        try:
            new_c = c.CreateTransformed(t)
        except Exception:
            new_c = None
        if new_c is None:
            p0 = c.GetEndPoint(0)
            p1 = c.GetEndPoint(1)
            new_c = Line.CreateBound(
                XYZ(p0.X, p0.Y, p0.Z + float(dz)),
                XYZ(p1.X, p1.Y, p1.Z + float(dz)),
            )
        grid.SetCurveInView(DatumExtentType.Model, view3d, new_c)
        n += 1
    if n == 0:
        raise Exception("no bound 3D model curve")
    return n


def apply_level_elevations(document, pairs, tolerances, dry_actions, allow_names=None):
    changed = 0
    for h0, e0, via_mon in pairs:
        if allow_names is not None and lg_lib.u(h0["name"]) not in allow_names:
            continue
        delta_mm = h0["elev_mm"] - e0["elev_mm"]
        if abs(delta_mm) <= 0.01:
            continue
        was_pinned = False
        try:
            was_pinned = bool(h0["element"].Pinned)
            if was_pinned:
                h0["element"].Pinned = False
        except Exception:
            was_pinned = False
        try:
            h0["element"].Elevation = e0["elev_ft"]
            changed += 1
            dry_actions.append(
                "SET level '{}'<-'{}' elev {:.3f} -> {:.3f} mm (monitor={})".format(
                    h0["name"],
                    e0["name"],
                    h0["elev_mm"],
                    e0["elev_mm"],
                    1 if via_mon else 0,
                )
            )
        except Exception as ex:
            dry_actions.append(
                "FAIL level '{}': {}".format(h0["name"], ex)
            )
        if was_pinned:
            try:
                h0["element"].Pinned = True
            except Exception:
                pass
    return changed


def apply_grid_curves(document, pairs, tolerances, dry_actions, allow_names=None):
    """
    Geometric Modify: XY (perp), Z (3D model extents), angle.
    Epsilon skip only — fix_plan Modify must run even if below audit grid_mm.
    """
    eps_mm = 0.01
    eps_ang = 0.001
    changed = 0
    unpin_ids = []
    view3d = None
    try:
        for h0, e0, via_mon in pairs:
            if allow_names is not None and lg_lib.u(h0["name"]) not in allow_names:
                continue
            if h0["mid_xy"] is None or e0["mid_xy"] is None:
                dry_actions.append(
                    "SKIP grid '{}' — no mid point".format(h0["name"])
                )
                continue

            dx, dy, dz = lg_lib.grid_translation_ft(h0, e0)
            d_xy_mm = lg_lib.feet_to_mm(math.hypot(dx, dy))
            d_z_mm = lg_lib.feet_to_mm(dz)
            d_ang = lg_lib.angle_delta_deg(h0["angle_deg"], e0["angle_deg"])
            need_xy = abs(d_xy_mm) > eps_mm
            need_z = abs(d_z_mm) > eps_mm
            need_ang = d_ang is not None and d_ang > eps_ang

            if (not need_xy) and (not need_z) and (not need_ang):
                dry_actions.append(
                    "SKIP grid '{}' — already matches xy={:.3f}mm z={:+.3f}mm".format(
                        h0["name"], d_xy_mm, d_z_mm
                    )
                )
                continue

            try:
                if h0["element"].Pinned:
                    h0["element"].Pinned = False
                    unpin_ids.append(h0["id"])
            except Exception:
                pass

            parts = []
            try:
                if need_xy:
                    ElementTransformUtils.MoveElement(
                        document, h0["id"], XYZ(dx, dy, 0.0)
                    )
                    parts.append("xy={:.3f}mm".format(d_xy_mm))
                if need_z:
                    if view3d is None:
                        view3d = ensure_3d_view(document)
                    try:
                        shift_grid_model_z(h0["element"], dz, view3d)
                        parts.append("z={:+.3f}mm".format(d_z_mm))
                    except Exception as zex:
                        parts.append("z-extents-fail:{}".format(zex))
                if need_ang and e0["angle_deg"] is not None and h0["angle_deg"] is not None:
                    mx, my = e0["mid_xy"]
                    ez = e0.get("z_ft")
                    if ez is None:
                        ez = 0.0
                    axis = Line.CreateBound(
                        XYZ(mx, my, ez - 10.0), XYZ(mx, my, ez + 10.0)
                    )
                    ang = math.radians(
                        float(e0["angle_deg"]) - float(h0["angle_deg"])
                    )
                    ElementTransformUtils.RotateElement(document, h0["id"], axis, ang)
                    parts.append("ang={:.4f}deg".format(d_ang))
                if parts:
                    changed += 1
                    dry_actions.append(
                        "TRANSFORM grid '{}'<-'{}' {} (monitor={})".format(
                            h0["name"],
                            e0["name"],
                            " ".join(parts),
                            1 if via_mon else 0,
                        )
                    )
                else:
                    dry_actions.append(
                        "SKIP grid '{}' — nothing applied".format(h0["name"])
                    )
            except Exception as ex2:
                dry_actions.append("FAIL grid '{}': {}".format(h0["name"], ex2))
    finally:
        _pin_elements(document, unpin_ids)
    return changed


def _pin_elements(document, element_ids):
    n = 0
    for eid in element_ids or []:
        try:
            el = document.GetElement(eid)
            if el is not None and not el.Pinned:
                el.Pinned = True
                n += 1
        except Exception:
            pass
    return n


def apply_renames(document, pairs, dry_actions, allow_map=None):
    """
    Rename host element to exemplar name.
    allow_map: optional dict host_name -> target_name (from fix_plan).
    If allow_map is None — rename all differing pairs (legacy cfg mode).
    """
    changed = 0
    for h0, e0, via_mon in pairs:
        host_name = lg_lib.u(h0["name"])
        if allow_map is not None:
            if host_name not in allow_map:
                continue
            target = allow_map.get(host_name) or e0["name"]
        else:
            target = e0["name"]
        if lg_lib.u(host_name) == lg_lib.u(target):
            continue
        try:
            h0["element"].Name = target
            changed += 1
            dry_actions.append(
                "RENAME '{}' -> '{}' (monitor={})".format(
                    host_name, target, 1 if via_mon else 0
                )
            )
        except Exception as ex:
            dry_actions.append(
                "FAIL rename '{}': {}".format(host_name, ex)
            )
    return changed


def parse_z_shift_mm(text):
    """Plain number or '... Z=-300 ...' from Excel/fix_plan."""
    raw = lg_lib.u(text or "").strip()
    if not raw:
        return None
    m = re.search(r"(?i)\bz\s*=\s*(-?[0-9]+(?:[.,][0-9]+)?)", raw)
    if m:
        try:
            return float(m.group(1).replace(",", "."))
        except Exception:
            return None
    if re.search(r"(?i)\b[xy]\s*=", raw):
        return None
    try:
        return float(raw.replace(",", ".").replace(" ", ""))
    except Exception:
        return None


def parse_axis_shift_mm(text, axis):
    raw = lg_lib.u(text or "").strip()
    if not raw:
        return None
    pat = r"(?i)\b" + re.escape(axis) + r"\s*=\s*(-?[0-9]+(?:[.,][0-9]+)?)"
    m = re.search(pat, raw)
    if not m:
        return None
    try:
        return float(m.group(1).replace(",", "."))
    except Exception:
        return None


def parse_pbp_shift_mm(text):
    """Return (dx, dy, dz) millimetres. Missing axis is 0. All-zero -> (0,0,0)."""
    raw = lg_lib.u(text or "").strip()
    dx = parse_axis_shift_mm(raw, "X")
    dy = parse_axis_shift_mm(raw, "Y")
    dz = parse_axis_shift_mm(raw, "Z")
    if dx is None and dy is None and dz is None:
        z_only = parse_z_shift_mm(raw)
        return 0.0, 0.0, (0.0 if z_only is None else z_only)
    return (
        0.0 if dx is None else dx,
        0.0 if dy is None else dy,
        0.0 if dz is None else dz,
    )


def plan_pbp_shift_mm(model_plan):
    """Signed PBP shift mm from fix_plan. None = do not move PBP."""
    if not model_plan:
        return None
    for r in model_plan:
        action = (r.get("action") or "").lower()
        if action != "modify":
            continue
        code = (r.get("problem_code") or "").strip().lower()
        cat = lg_lib.u(r.get("category") or "")
        el = lg_lib.u(r.get("element") or "").lower()
        is_pbp = (
            code in ("pbp_delta", "pbp_z")
            or cat in ("Coords", u"Координаты")
            or u"базовая точка" in el
            or u"base point" in el
            or el == "pbp"
        )
        if not is_pbp:
            continue
        dx, dy, dz = parse_pbp_shift_mm(r.get("delta"))
        if abs(dx) < 0.001 and abs(dy) < 0.001 and abs(dz) < 0.001:
            continue
        return dx, dy, dz
    return None


def _set_pbp_clipped(el, clipped):
    want = 1 if clipped else 0
    try:
        p = el.LookupParameter("Clipped")
        if p is not None and not p.IsReadOnly:
            p.Set(want)
            return True
    except Exception:
        pass
    try:
        for p in el.Parameters:
            try:
                n = lg_lib.u(p.Definition.Name)
            except Exception:
                continue
            low = n.lower()
            if low == "clipped" or n == u"Скрепка":
                if not p.IsReadOnly:
                    p.Set(want)
                    return True
    except Exception:
        pass
    return False


def apply_pbp_shift(document, dx_mm, dy_mm, dz_mm, dry_actions):
    """Move Project Base Point. Signed millimetres; minus Z = down."""
    try:
        dx_mm = float(dx_mm or 0.0)
        dy_mm = float(dy_mm or 0.0)
        dz_mm = float(dz_mm or 0.0)
    except Exception:
        dry_actions.append("SKIP PBP — bad XYZ value")
        return 0
    if abs(dx_mm) < 0.001 and abs(dy_mm) < 0.001 and abs(dz_mm) < 0.001:
        dry_actions.append("SKIP PBP — XYZ shift is 0")
        return 0

    info = lg_lib.collect_base_point(document, is_survey=False, transform=None)
    if not info or info.get("element") is None:
        dry_actions.append("FAIL PBP — project base point not found")
        return 0
    el = info["element"]
    dx = lg_lib.mm_to_feet(dx_mm)
    dy = lg_lib.mm_to_feet(dy_mm)
    dz = lg_lib.mm_to_feet(dz_mm)

    was_pinned = False
    try:
        was_pinned = el.Pinned
        if was_pinned:
            el.Pinned = False
    except Exception:
        was_pinned = False

    _set_pbp_clipped(el, False)

    z_done = False
    err = None
    if abs(dz_mm) >= 0.001:
        try:
            p = el.get_Parameter(BuiltInParameter.BASEPOINT_ELEVATION_PARAM)
            if p is not None and not p.IsReadOnly:
                p.Set(p.AsDouble() + dz)
                z_done = True
        except Exception as ex:
            err = ex

    mx = dx if abs(dx_mm) >= 0.001 else 0.0
    my = dy if abs(dy_mm) >= 0.001 else 0.0
    mz = 0.0 if z_done else (dz if abs(dz_mm) >= 0.001 else 0.0)
    moved = z_done
    if abs(mx) >= 1e-9 or abs(my) >= 1e-9 or abs(mz) >= 1e-9:
        try:
            ElementTransformUtils.MoveElement(document, el.Id, XYZ(mx, my, mz))
            moved = True
        except Exception as ex2:
            err = ex2
            if not moved:
                try:
                    loc = el.Location
                    if loc is not None:
                        loc.Move(XYZ(mx, my, mz))
                        moved = True
                except Exception as ex3:
                    err = ex3

    if was_pinned:
        try:
            el.Pinned = True
        except Exception:
            pass

    if moved:
        dry_actions.append(
            "MOVE PBP X {:+.3f} Y {:+.3f} Z {:+.3f} mm".format(
                dx_mm, dy_mm, dz_mm
            )
        )
        return 1
    dry_actions.append("FAIL PBP XYZ: {}".format(err))
    return 0


def filter_pairs(pairs, tolerances):
    """Respect apply_use_monitor / apply_name_fallback flags."""
    use_mon = cfg_flag(tolerances, "apply_use_monitor", True)
    use_fb = cfg_flag(tolerances, "apply_name_fallback", True)
    out = []
    for h, e, via in pairs:
        if via and not use_mon:
            continue
        if (not via) and not use_fb:
            continue
        out.append((h, e, via))
    return out


def _split_csv_line(line):
    fields = []
    cur = []
    in_q = False
    i = 0
    s = line
    while i < len(s):
        ch = s[i]
        if in_q:
            if ch == '"':
                if i + 1 < len(s) and s[i + 1] == '"':
                    cur.append('"')
                    i += 2
                    continue
                in_q = False
                i += 1
                continue
            cur.append(ch)
            i += 1
            continue
        if ch == '"':
            in_q = True
            i += 1
            continue
        if ch == ";":
            fields.append("".join(cur))
            cur = []
            i += 1
            continue
        cur.append(ch)
        i += 1
    fields.append("".join(cur))
    return fields


def load_fix_plan(path):
    """
    Load fix_plan.csv from choose_fix_plan.ps1.
    Returns dict:
      {"rows": [...], "sync_after_change": bool}
    """
    if not path or not os.path.isfile(path):
        return None
    rows = []
    sync_after_change = True
    try:
        with codecs.open(path, "r", "utf-8-sig") as f:
            lines = [ln.strip() for ln in f.read().splitlines() if ln.strip()]
    except Exception as ex:
        Output("WARN fix_plan read: {}".format(ex))
        return None
    if not lines:
        return []
    # RU header from choose_fix_plan; English "action" column if present.
    header = _split_csv_line(lines[0])
    # Find action column index
    action_idx = None
    for i, h in enumerate(header):
        hl = (h or "").strip().lower()
        if hl == "action":
            action_idx = i
            break
    if action_idx is None:
        # RU header layout: action is 9th col (index 8)
        action_idx = 8

    # Optional extra column (added by UI): sync_after_change=1/0
    sync_idx = None
    delta_idx = 12
    for i, h in enumerate(header):
        hl = (h or "").strip().lower()
        if hl == "sync_after_change":
            sync_idx = i
        if hl in ("delta",) or hl == u"разница" or (h or "") == u"Разница":
            delta_idx = i

    for line in lines[1:]:
        cols = _split_csv_line(line)
        if len(cols) <= action_idx:
            continue
        def col(i, default=""):
            return cols[i].strip() if i < len(cols) else default

        action = col(action_idx).lower()
        if action not in ("modify", "rename", "accept", "reject"):
            # maybe action_ru column shifted — try known tokens
            continue
        if sync_idx is not None and sync_idx < len(cols):
            sval = (cols[sync_idx] or "").strip().lower()
            if sval in ("0", "false", "no", "off"):
                sync_after_change = False
            elif sval in ("1", "true", "yes", "on"):
                sync_after_change = True

        rows.append(
            {
                "model_name": col(0),
                "model_path": col(1),
                "category": col(2),
                "element": col(3),
                "name_exemplar": col(4),
                "problem_code": col(5),
                "problem": col(6),
                "rule": col(7),
                "action": action,
                "delta": col(delta_idx),
            }
        )
    return {"rows": rows, "sync_after_change": sync_after_change}


def fix_plan_for_model(plan_rows, model_path):
    """Filter plan rows for current host model."""
    if plan_rows is None:
        return None
    base = lg_lib.path_basename(model_path).lower()
    norm = lg_lib.normalize_model_path(model_path)
    out = []
    for r in plan_rows:
        mn = lg_lib.u(r.get("model_name") or "").strip()
        mp = lg_lib.u(r.get("model_path") or "").strip()
        ok = False
        if mn and mn.lower() == base:
            ok = True
        if mp and lg_lib.normalize_model_path(mp) == norm:
            ok = True
        if mp and lg_lib.path_basename(mp).lower() == base:
            ok = True
        if ok:
            out.append(r)
    return out


def plan_allow_sets(model_plan):
    """
    From plan rows for one model build:
      modify_levels, modify_grids: set of host element names
      rename_map: host_name -> target exemplar name (Level+Grid)
    accept/reject are ignored (skip).
    """
    mod_lv = set()
    mod_gr = set()
    ren = {}
    if not model_plan:
        return mod_lv, mod_gr, ren
    for r in model_plan:
        action = (r.get("action") or "").lower()
        el = lg_lib.u(r.get("element") or "").strip()
        if not el:
            continue
        cat = lg_lib.u(r.get("category") or "").strip()
        if action == "modify":
            if cat.lower().startswith("level") or cat == u"Уровни" or cat == "Level":
                mod_lv.add(el)
            elif cat.lower().startswith("grid") or cat == u"Оси" or cat == "Grid":
                mod_gr.add(el)
            else:
                # category_kind written as Level/Grid by wizard
                if cat == "Level":
                    mod_lv.add(el)
                elif cat == "Grid":
                    mod_gr.add(el)
        elif action == "rename":
            target = lg_lib.u(r.get("name_exemplar") or "").strip()
            ren[el] = target
    return mod_lv, mod_gr, ren


def run_apply(document, model_path, tolerances, project_rows, fix_plan_rows=None):
    row_map = lg_lib.find_project_row(project_rows, model_path)
    ok, reason = allow_apply(row_map, model_path, tolerances)
    if not ok:
        raise Exception("Apply blocked: {}".format(reason))

    link_hint = lg_lib.resolve_exemplar_hint(SCRIPT_DIR, tolerances, row_map)
    inst, link_doc, label = lg_lib.find_exemplar_link(
        document, link_hint, host_model_path=model_path
    )
    if link_doc is None:
        raise Exception("Exemplar link not found: {}".format(label))

    Output("Exemplar link: {}".format(label))
    Output("Apply allowed: {}".format(reason))

    model_plan = fix_plan_for_model(fix_plan_rows, model_path)
    use_plan = model_plan is not None and fix_plan_rows is not None

    if use_plan:
        mod_lv, mod_gr, ren_map = plan_allow_sets(model_plan)
        Output(
            "Fix plan for model: modify_levels={} modify_grids={} rename={}".format(
                len(mod_lv), len(mod_gr), len(ren_map)
            )
        )
        if not model_plan:
            Output("No plan rows for this model — nothing to change.")
            return 0, ["SKIP no plan rows for model"]
    else:
        mod_lv = mod_gr = ren_map = None
        Output(
            "No fix_plan.csv — legacy cfg flags: elevation={} grid={} rename={}".format(
                1 if cfg_flag(tolerances, "apply_elevation", True) else 0,
                1 if cfg_flag(tolerances, "apply_grid_curve", True) else 0,
                1 if cfg_flag(tolerances, "apply_rename", False) else 0,
            )
        )

    Output(
        "Pairing: monitor={} name_fallback={}".format(
            1 if cfg_flag(tolerances, "apply_use_monitor", True) else 0,
            1 if cfg_flag(tolerances, "apply_name_fallback", True) else 0,
        )
    )

    def refresh_pairs():
        lv = filter_pairs(
            lg_lib.get_pairing_for_apply(document, link_doc, inst, "Level"),
            tolerances,
        )
        gr = filter_pairs(
            lg_lib.get_pairing_for_apply(document, link_doc, inst, "Grid"),
            tolerances,
        )
        return lv, gr

    lv_pairs, gr_pairs = refresh_pairs()
    Output(
        "Pairs: levels={} grids={} (after option filter)".format(
            len(lv_pairs), len(gr_pairs)
        )
    )

    actions = []
    changed = 0

    cr_pre = CopyWatchDeletePreprocessor()
    t = Transaction(document, "Levels&Grids Apply v{}".format(SCRIPT_VERSION))
    t.Start()
    try:
        opts = t.GetFailureHandlingOptions()
        opts.SetFailuresPreprocessor(cr_pre)
        try:
            opts.SetForcedModalHandling(True)
        except Exception:
            pass
        t.SetFailureHandlingOptions(opts)
        if use_plan:
            # PBP first: shared coords / link appearance can change, then
            # monitored levels/grids are matched to the current exemplar.
            pbp_shift = plan_pbp_shift_mm(model_plan)
            if pbp_shift is not None:
                dx_pbp, dy_pbp, dz_pbp = pbp_shift
                changed += apply_pbp_shift(
                    document, dx_pbp, dy_pbp, dz_pbp, actions
                )
                try:
                    document.Regenerate()
                except Exception:
                    pass
                lv_pairs, gr_pairs = refresh_pairs()
                actions.append(
                    "Re-paired after PBP: levels={} grids={}".format(
                        len(lv_pairs), len(gr_pairs)
                    )
                )
            else:
                actions.append("SKIP PBP (not in fix plan)")

            if mod_lv:
                changed += apply_level_elevations(
                    document, lv_pairs, tolerances, actions, allow_names=mod_lv
                )
            else:
                actions.append("SKIP elevations (not in fix plan)")

            if mod_gr:
                changed += apply_grid_curves(
                    document, gr_pairs, tolerances, actions, allow_names=mod_gr
                )
            else:
                actions.append("SKIP grids (not in fix plan)")

            if ren_map:
                changed += apply_renames(
                    document, lv_pairs + gr_pairs, actions, allow_map=ren_map
                )
            else:
                actions.append("SKIP rename (not in fix plan)")
        else:
            if cfg_flag(tolerances, "apply_elevation", True):
                changed += apply_level_elevations(
                    document, lv_pairs, tolerances, actions
                )
            else:
                actions.append("SKIP elevations (apply_elevation=0)")

            if cfg_flag(tolerances, "apply_grid_curve", True):
                changed += apply_grid_curves(
                    document, gr_pairs, tolerances, actions
                )
            else:
                actions.append("SKIP grids (apply_grid_curve=0)")

            if cfg_flag(tolerances, "apply_rename", False):
                changed += apply_renames(
                    document, lv_pairs + gr_pairs, actions
                )
            else:
                actions.append("SKIP rename (apply_rename=0)")

        try:
            document.Regenerate()
        except Exception:
            pass
        t.Commit()
    except Exception:
        if t.HasStarted():
            t.RollBack()
        raise

    actions.append(
        "CR preprocessor: messages={} copywatch_deleted={}".format(
            cr_pre.message_count, cr_pre.deleted_count
        )
    )
    actions.append(
        "Note: CR dialog Postpone is not in the public API; "
        "geometry Modify runs in background."
    )
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

FIX_PLAN_DATA = load_fix_plan(FIX_PLAN_PATH)
if FIX_PLAN_DATA is None:
    Output("Fix plan: (none) — using tolerances.cfg flags")
    FIX_PLAN_ROWS = None
    SYNC_AFTER_CHANGE = True
else:
    FIX_PLAN_ROWS = FIX_PLAN_DATA.get("rows", [])
    SYNC_AFTER_CHANGE = bool(FIX_PLAN_DATA.get("sync_after_change", True))
    Output(
        "Fix plan: {} ({} rows)".format(FIX_PLAN_PATH, len(FIX_PLAN_ROWS))
    )
    Output("Fix plan sync_after_change={}".format(1 if SYNC_AFTER_CHANGE else 0))

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

    changed, actions = run_apply(
        doc, revitFilePath, TOL, PROJECT_ROWS, fix_plan_rows=FIX_PLAN_ROWS
    )

    if changed > 0 and SYNC_AFTER_CHANGE:
        comment = "Batch Levels&Grids Apply v{} ({} changes)".format(
            SCRIPT_VERSION, changed
        )
        saved = save_or_sync(doc, revitFilePath, comment)
        Output("Done: {}".format(saved))
    elif changed > 0 and not SYNC_AFTER_CHANGE:
        Output("Changes applied. Sync/Save skipped by fix plan option.")
    else:
        Output("No changes — skip Save/Sync")

    Output("OK")
except Exception as ex:
    Output("ERROR: {}".format(ex))
    raise
