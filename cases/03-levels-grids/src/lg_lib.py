# -*- coding: utf-8 -*-
# Name: lg_lib.py
# Version: 1.23.0
# What it does: Shared IronPython helpers for levels/grids audit and apply scripts.
# Inputs: Imported by audit/apply after SCRIPT_DIR on sys.path.
# Outputs: Utility functions only (no standalone run).
# How to run: Not run directly; used from levels_grids_*.py in RBP.
# Notes: IronPython 2.7 / RBP compatible.

"""
Общие хелперы Levels & Grids (IronPython / RBP).
Импортируется audit/apply после добавления SCRIPT_DIR в sys.path.
"""

import codecs
import math
import os
import re

try:
    import System
except ImportError:
    System = None  # syntax checks outside Revit / IronPython

try:
    unicode
except NameError:
    unicode = str  # py3 fallback for syntax checks outside Revit


def u(text):
    if text is None:
        return u""
    try:
        if isinstance(text, unicode):
            return text
    except Exception:
        pass
    try:
        return unicode(text)
    except Exception:
        return unicode(str(text))


def now_stamp():
    if System is None:
        raise Exception("System.DateTime is available only inside Revit")
    return System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")


def now_file_stamp():
    if System is None:
        raise Exception("System.DateTime is available only inside Revit")
    return System.DateTime.Now.ToString("yyyy-MM-dd_HHmm")


def rbp_root():
    home = os.environ.get("USERPROFILE") or r"C:\Users\Public"
    return os.path.join(home, r"Documents\doc\script\rbp")


def shared_reports_dir():
    """Common folder for all RBP reports. Created if missing."""
    path = os.path.join(rbp_root(), "reports")
    if path and not os.path.isdir(path):
        os.makedirs(path)
    return path


def user_reports_dir():
    return shared_reports_dir()


def normalize_model_path(path):
    if not path:
        return ""
    p = path.strip().strip('"').replace("/", "\\")
    while "\\\\" in p[2:]:
        p = p[:2] + p[2:].replace("\\\\", "\\")
    return p.lower()


def is_dummy_rsn_path(path):
    """
    RBP / Revit Server sometimes injects a placeholder like RSN://x/BF_K4.rvt.
    Opening it fails: 'The central model in sourcePath is missing.'
    """
    p = u(path or "").strip().strip('"')
    if not p:
        return True
    return bool(re.match(r"(?i)^RSN://x(/|$)", p))


def path_basename(path):
    if not path:
        return ""
    p = path.replace("\\", "/").rstrip("/")
    if "/" in p:
        return p.split("/")[-1]
    return p


_STRING_KEYS = set(["exemplar_link_default"])


def read_exemplar_link_cfg(script_dir):
    """
    Read exemplar selection.
    Preferred: exemplar_model.cfg (full path from model picker).
    Legacy: exemplar_link.cfg (substring).
    Returns None if neither file exists.
    Values AUTO / empty mean: pick link from each host automatically.
    """
    model_path = os.path.join(script_dir, "exemplar_model.cfg")
    if os.path.isfile(model_path):
        try:
            with open(model_path, "r") as f:
                for raw in f:
                    line = raw.strip().strip('"')
                    if not line or line.startswith("#"):
                        continue
                    return line
            return "AUTO"
        except Exception:
            pass

    path = os.path.join(script_dir, "exemplar_link.cfg")
    if not os.path.isfile(path):
        return None
    try:
        with open(path, "r") as f:
            for raw in f:
                line = raw.strip()
                if not line or line.startswith("#"):
                    continue
                return line
        return "AUTO"
    except Exception:
        return None


def _is_auto_exemplar_hint(hint):
    h = u(hint or "").strip()
    if not h:
        return True
    return h.upper() in ("AUTO", "AUTO_FROM_HOST", "*", "NONE")


def resolve_exemplar_hint(script_dir, tolerances, project_row=None):
    """
    Priority:
      1) project_map row override (optional advanced)
      2) exemplar_model.cfg / exemplar_link.cfg (session chooser)
         AUTO / empty → auto-pick link inside each host
      3) tolerances.cfg exemplar_link_default
      4) empty → auto-pick
    """
    if project_row:
        hint = (project_row.get("exemplar_link_name") or "").strip()
        if hint:
            return hint
    cfg_hint = read_exemplar_link_cfg(script_dir)
    if cfg_hint is not None:
        if _is_auto_exemplar_hint(cfg_hint):
            return ""
        return cfg_hint
    default = (tolerances.get("exemplar_link_default") or "").strip()
    if _is_auto_exemplar_hint(default):
        return ""
    return default


def _hint_looks_like_path(hint):
    h = u(hint or "").strip()
    if not h:
        return False
    low = h.lower()
    if low.startswith("rsn://"):
        return True
    if low.endswith(".rvt"):
        return True
    if "\\" in h or "/" in h:
        return True
    return False


def _get_link_type_path(document, link_type):
    try:
        from Autodesk.Revit.DB import ModelPathUtils

        ext_ref = link_type.GetExternalFileReference()
        if ext_ref:
            model_path = ext_ref.GetPath()
            if model_path:
                return u(
                    ModelPathUtils.ConvertModelPathToUserVisiblePath(model_path)
                )
    except Exception:
        pass
    return u""


def _building_code(name):
    """Extract K1/K4/… from model/link file names."""
    text = path_basename(name or "")
    m = re.search(r"[_\-\s]([KkКк])(\d+)\b", text)
    if not m:
        m = re.search(r"\b([KkКк])(\d+)\b", text)
    if m:
        return "K" + m.group(2)
    return ""


def _looks_like_ar(blob):
    b = u(blob or "").lower()
    return (
        "_ar_" in b
        or b.startswith("ar_")
        or "_ar." in b
        or "/ar_" in b
        or "\\ar_" in b
        or "_ар_" in b
        or "ар_" in b
    )


def _looks_like_bf(blob):
    b = u(blob or "").lower()
    return (
        "_bf_" in b
        or b.startswith("bf_")
        or "_bf." in b
        or "/bf_" in b
        or "\\bf_" in b
    )


def _host_wants_ar(host_model_path):
    """BF hosts compare to AR; everything else prefers BF link."""
    base = path_basename(host_model_path or "").lower()
    return _looks_like_bf(base)


def _display_link_label(c):
    base = path_basename(c.get("type_path") or c.get("doc_path") or "")
    if base:
        return base
    raw = u(c.get("raw_name") or c.get("label") or "")
    if " :" in raw:
        return raw.split(" :", 1)[0].strip()
    return raw or u"?"


def find_exemplar_link(document, link_hint, host_model_path=None):
    """
    Find RevitLinkInstance by:
      - full model path / basename (from exemplar_model.cfg picker), or
      - name/type substring (legacy), or
      - AUTO: AR (for BF host) / BF (for section host) with same building code, or
      - first loaded RVT link if nothing else matches.
    Returns (instance, link_doc, match_label) or (None, None, reason).
    Label is a short file name (not Revit "name : id : позиция …").
    """
    from Autodesk.Revit.DB import FilteredElementCollector, RevitLinkInstance

    hint_raw = u(link_hint or "").strip()
    if _is_auto_exemplar_hint(hint_raw):
        hint_raw = u""
    hint = hint_raw.lower()
    hint_base = path_basename(hint).lower()
    if hint_base.endswith(".rvt"):
        hint_stem = hint_base[:-4]
    else:
        hint_stem = hint_base
    as_path = _hint_looks_like_path(hint_raw)
    host_code = _building_code(host_model_path or "")
    want_ar = _host_wants_ar(host_model_path)

    instances = list(FilteredElementCollector(document).OfClass(RevitLinkInstance))
    candidates = []
    for inst in instances:
        try:
            name = u(inst.Name)
        except Exception:
            name = u""
        type_name = u""
        type_path = u""
        try:
            t = document.GetElement(inst.GetTypeId())
            if t is not None:
                type_name = u(t.Name)
                type_path = _get_link_type_path(document, t)
        except Exception:
            pass
        doc_path = u""
        link_doc = None
        try:
            link_doc = inst.GetLinkDocument()
            if link_doc is not None:
                try:
                    doc_path = u(link_doc.PathName)
                except Exception:
                    doc_path = u""
        except Exception:
            link_doc = None
        label = name or type_name or path_basename(type_path) or path_basename(doc_path)
        candidates.append(
            {
                "inst": inst,
                "link_doc": link_doc,
                "label": label,
                "raw_name": name,
                "name": name.lower(),
                "type_name": type_name.lower(),
                "type_path": normalize_model_path(type_path),
                "doc_path": normalize_model_path(doc_path),
                "base": path_basename(type_path or doc_path or name).lower(),
            }
        )

    def _match(c):
        if not hint:
            return False
        if as_path:
            nh = normalize_model_path(hint_raw)
            if nh and (c["type_path"] == nh or c["doc_path"] == nh):
                return True
            if hint_base and (
                c["base"] == hint_base
                or hint_base in c["name"]
                or hint_base in c["type_name"]
                or hint_base in c["type_path"]
                or hint_base in c["doc_path"]
            ):
                return True
            if hint_stem and (
                hint_stem in c["name"]
                or hint_stem in c["type_name"]
                or hint_stem in c["base"]
            ):
                return True
            return False
        return (
            hint in c["name"]
            or hint in c["type_name"]
            or hint in u(c["label"]).lower()
            or hint in c["base"]
        )

    def _match_auto_from_host(c):
        blob = u" ".join([c["name"], c["type_name"], c["base"], c["type_path"]])
        if host_code and host_code.lower() not in blob.lower():
            return False
        if want_ar:
            return _looks_like_ar(blob)
        return _looks_like_bf(blob)

    def _match_role_only(c):
        """Same role (AR/BF) without building code — used if only one candidate."""
        blob = u" ".join([c["name"], c["type_name"], c["base"], c["type_path"]])
        if want_ar:
            return _looks_like_ar(blob)
        return _looks_like_bf(blob)

    def _return_loaded(matched, note=None):
        for c in matched:
            if c["link_doc"] is not None:
                label = _display_link_label(c)
                if note:
                    label = u"{0} ({1})".format(label, note)
                return c["inst"], c["link_doc"], label
        c0 = matched[0]
        short = _display_link_label(c0)
        return (
            None,
            None,
            u"Связь «{0}» есть в модели, но не загружена. "
            u"В RBP поставьте Worksets = Open All "
            u"(при Close All уровни/оси связи недоступны).".format(short),
        )

    def _try_auto():
        by_bldg = [c for c in candidates if _match_auto_from_host(c)]
        if by_bldg:
            role = u"АР" if want_ar else u"БФ"
            note = u"авто: {0}".format(role)
            if host_code:
                note = u"авто: {0} корпус {1}".format(role, host_code)
            return _return_loaded(by_bldg, note=note)

        by_role = [c for c in candidates if _match_role_only(c)]
        if len(by_role) == 1:
            role = u"АР" if want_ar else u"БФ"
            return _return_loaded(by_role, note=u"авто: единственная связь {0}".format(role))
        return None

    if hint:
        matched = [c for c in candidates if _match(c)]
        if matched:
            return _return_loaded(matched)

        # Wrong global exemplar for multi-BF batch → fall back to auto.
        auto = _try_auto()
        if auto is not None:
            return auto

        want = path_basename(hint_raw) or hint_raw
        role = u"АР" if want_ar else u"БФ"
        if host_code:
            return (
                None,
                None,
                u"В модели нет связи на «{0}». "
                u"Ожидался свой {1} для корпуса {2} (например …{1}_{2}.rvt).".format(
                    want, role, host_code
                ),
            )
        return (
            None,
            None,
            u"В модели нет связи на эталон «{0}». "
            u"Проверьте Manage Links.".format(want),
        )

    auto = _try_auto()
    if auto is not None:
        return auto

    role = u"АР" if want_ar else u"БФ"
    if host_code:
        return (
            None,
            None,
            u"Не найден эталон автоматически: нужна загруженная связь {0} "
            u"с корпусом {1} (имя файла …{0}_{1}.rvt). "
            u"Worksets = Open All.".format(role, host_code),
        )

    for c in candidates:
        if c["link_doc"] is not None:
            return c["inst"], c["link_doc"], _display_link_label(c)
    if candidates:
        return (
            None,
            None,
            u"Связи есть, но ни одна не загружена. В RBP: Worksets = Open All.",
        )
    return None, None, u"В модели нет Revit-связей (RevitLinkInstance)."


def load_key_values(path, defaults):
    values = dict(defaults)
    if not path or not os.path.isfile(path):
        return values
    with open(path, "r") as f:
        for raw in f:
            line = raw.strip()
            if not line or line.startswith("#"):
                continue
            if "=" not in line:
                continue
            key, val = line.split("=", 1)
            key = key.strip().lower()
            val = val.strip()
            if not key:
                continue
            if key in _STRING_KEYS:
                values[key] = val
                continue
            try:
                if val == "":
                    values[key] = 0.0
                else:
                    values[key] = float(val.replace(",", "."))
            except Exception:
                values[key] = val
    return values


def load_tolerances(path):
    defaults = {
        "level_mm": 2.0,
        "grid_mm": 5.0,
        "grid_angle_deg": 0.1,
        "level_delta_yellow_mm": 2.0,
        "level_delta_red_mm": 10.0,
        "grid_delta_yellow_mm": 5.0,
        "grid_delta_red_mm": 20.0,
        "grid_angle_yellow_deg": 0.1,
        "grid_angle_red_deg": 1.0,
        "pbp_delta_yellow_mm": 5.0,
        "pbp_delta_red_mm": 20.0,
        "pbp_z_yellow_mm": 2.0,
        "pbp_z_red_mm": 10.0,
        "exemplar_link_default": "",
        "apply_elevation": 1.0,
        "apply_grid_curve": 1.0,
        "apply_rename": 0.0,
        "apply_use_monitor": 1.0,
        "apply_name_fallback": 1.0,
        "apply_only_role_bf": 0.0,
        "apply_require_map_bf": 0.0,
    }
    return load_key_values(path, defaults)


def _split_csv_line(line):
    result = []
    sb = []
    in_quotes = False
    i = 0
    while i < len(line):
        ch = line[i]
        if ch == '"':
            if in_quotes and i + 1 < len(line) and line[i + 1] == '"':
                sb.append('"')
                i += 1
            else:
                in_quotes = not in_quotes
        elif ch == ";" and not in_quotes:
            result.append("".join(sb).strip())
            sb = []
        else:
            sb.append(ch)
        i += 1
    result.append("".join(sb).strip())
    return result


def load_project_map(path):
    """
    Returns list of dicts:
      building, host_path, role, exemplar_link_name, level_whitelist (list or None)
    """
    rows = []
    if not path or not os.path.isfile(path):
        return rows
    with codecs.open(path, "r", "utf-8-sig") as f:
        lines = f.readlines()
    if not lines:
        return rows
    headers = [h.strip().lower() for h in _split_csv_line(lines[0].strip())]
    for raw in lines[1:]:
        line = raw.strip()
        if not line or line.startswith("#"):
            continue
        vals = _split_csv_line(line)
        row = {}
        for i, h in enumerate(headers):
            row[h] = vals[i] if i < len(vals) else ""
        host = row.get("host_path") or row.get("host") or ""
        if not host:
            continue
        wl_raw = row.get("level_whitelist") or ""
        whitelist = None
        if wl_raw.strip():
            whitelist = [x.strip() for x in wl_raw.split("|") if x.strip()]
        rows.append(
            {
                "building": row.get("building") or "",
                "host_path": host,
                "role": (row.get("role") or "").strip().upper(),
                "exemplar_link_name": row.get("exemplar_link_name")
                or row.get("exemplar")
                or "",
                "level_whitelist": whitelist,
            }
        )
    return rows


def find_project_row(project_rows, model_path):
    if not project_rows:
        return None
    target = normalize_model_path(model_path)
    target_base = path_basename(target).lower()
    for row in project_rows:
        hp = normalize_model_path(row["host_path"])
        if hp and hp == target:
            return row
    for row in project_rows:
        hp = normalize_model_path(row["host_path"])
        if hp and path_basename(hp).lower() == target_base:
            return row
    return None


def feet_to_mm(feet):
    return float(feet) * 304.8


def mm_to_feet(mm):
    return float(mm) / 304.8


def angle_deg_of_curve(curve):
    """Angle of curve XY projection in degrees [0, 180)."""
    try:
        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)
        dx = p1.X - p0.X
        dy = p1.Y - p0.Y
        if abs(dx) < 1e-12 and abs(dy) < 1e-12:
            return 0.0
        ang = math.degrees(math.atan2(dy, dx))
        while ang < 0:
            ang += 180.0
        while ang >= 180.0:
            ang -= 180.0
        return ang
    except Exception:
        return None


def angle_delta_deg(a, b):
    if a is None or b is None:
        return None
    d = abs(float(a) - float(b)) % 180.0
    if d > 90.0:
        d = 180.0 - d
    return d


def midpoint_xy(curve):
    try:
        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)
        return ((p0.X + p1.X) * 0.5, (p0.Y + p1.Y) * 0.5)
    except Exception:
        return None


def distance_mm_xy(a, b):
    if a is None or b is None:
        return None
    dx = a[0] - b[0]
    dy = a[1] - b[1]
    return feet_to_mm(math.sqrt(dx * dx + dy * dy))


def fmt_xy_mm(xy):
    """Human-readable XY in mm for report value columns."""
    if xy is None:
        return ""
    try:
        return "X={:.1f}; Y={:.1f}".format(feet_to_mm(xy[0]), feet_to_mm(xy[1]))
    except Exception:
        return ""


def fmt_xyz_mm(xyz):
    """Human-readable XYZ in mm."""
    if xyz is None or len(xyz) < 3:
        return ""
    try:
        return "X={:.1f}; Y={:.1f}; Z={:.1f}".format(
            feet_to_mm(xyz[0]), feet_to_mm(xyz[1]), feet_to_mm(xyz[2])
        )
    except Exception:
        return ""


def xyz_delta_mm(a, b):
    """Return (xy_mm, z_mm) between two XYZ tuples in feet."""
    if a is None or b is None:
        return None, None
    try:
        dx = a[0] - b[0]
        dy = a[1] - b[1]
        dz = a[2] - b[2]
        xy = feet_to_mm(math.sqrt(dx * dx + dy * dy))
        z = feet_to_mm(dz)
        return xy, z
    except Exception:
        return None, None


def collect_base_point(document, is_survey=False, transform=None):
    """
    Project Base Point (is_survey=False) or Survey Point (is_survey=True).
    pos: internal XYZ, optionally transformed link -> host.
    shared: SharedPosition XYZ (shared coords, not transformed).
    """
    from Autodesk.Revit.DB import FilteredElementCollector, BasePoint, XYZ

    try:
        points = list(FilteredElementCollector(document).OfClass(BasePoint))
    except Exception:
        return None
    for bp in points:
        try:
            shared_flag = bool(bp.IsShared)
        except Exception:
            continue
        if shared_flag != bool(is_survey):
            continue
        pos = None
        try:
            p = bp.Position
            pos = (p.X, p.Y, p.Z)
        except Exception:
            try:
                loc = bp.Location
                p = loc.Point
                pos = (p.X, p.Y, p.Z)
            except Exception:
                pos = None
        if pos is not None and transform is not None:
            try:
                p2 = transform.OfPoint(XYZ(pos[0], pos[1], pos[2]))
                pos = (p2.X, p2.Y, p2.Z)
            except Exception:
                pass
        shared = None
        try:
            sp = bp.SharedPosition
            shared = (sp.X, sp.Y, sp.Z)
        except Exception:
            shared = None
        return {"pos": pos, "shared": shared, "element": bp}
    return None


def point_to_line_distance_ft(px, py, ax, ay, bx, by):
    """Absolute distance from point P to infinite line AB (feet)."""
    dx = bx - ax
    dy = by - ay
    len2 = dx * dx + dy * dy
    if len2 < 1e-18:
        return math.hypot(px - ax, py - ay)
    return abs((px - ax) * dy - (py - ay) * dx) / math.sqrt(len2)


def closest_point_on_line_xy(px, py, ax, ay, bx, by):
    """Closest point on infinite line AB to P (XY feet)."""
    dx = bx - ax
    dy = by - ay
    len2 = dx * dx + dy * dy
    if len2 < 1e-18:
        return ax, ay
    t = ((px - ax) * dx + (py - ay) * dy) / len2
    return ax + t * dx, ay + t * dy


def grid_translation_ft(host_item, ex_item):
    """
    Translation (dx, dy, dz) in feet to put host grid onto exemplar
    in host coordinates (exemplar already link-transformed).
    XY is perpendicular to the exemplar line (not mid-to-mid along the grid).
    """
    dx = dy = dz = 0.0
    hm = host_item.get("mid_xy") if host_item else None
    e_line = grid_line_points(ex_item)
    if hm is not None and e_line is not None:
        (ax, ay), (bx, by) = e_line
        cx, cy = closest_point_on_line_xy(hm[0], hm[1], ax, ay, bx, by)
        dx = cx - hm[0]
        dy = cy - hm[1]
    elif hm is not None and ex_item and ex_item.get("mid_xy") is not None:
        em = ex_item["mid_xy"]
        dx = em[0] - hm[0]
        dy = em[1] - hm[1]
    hz = host_item.get("z_ft") if host_item else None
    ez = ex_item.get("z_ft") if ex_item else None
    if hz is None and host_item and host_item.get("z_mm") is not None:
        hz = mm_to_feet(host_item["z_mm"])
    if ez is None and ex_item and ex_item.get("z_mm") is not None:
        ez = mm_to_feet(ex_item["z_mm"])
    if hz is not None and ez is not None:
        dz = float(ez) - float(hz)
    return dx, dy, dz


def grid_line_points(item):
    """Return ((x0,y0),(x1,y1)) in feet, or None."""
    if not item:
        return None
    p0 = item.get("p0")
    p1 = item.get("p1")
    if p0 is not None and p1 is not None:
        return (p0, p1)
    mid = item.get("mid_xy")
    ang = item.get("angle_deg")
    if mid is None or ang is None:
        return None
    try:
        rad = math.radians(float(ang))
        # synthetic short segment through mid
        c = math.cos(rad)
        s = math.sin(rad)
        return ((mid[0] - c, mid[1] - s), (mid[0] + c, mid[1] + s))
    except Exception:
        return None


def grid_perp_distance_mm(host_item, ex_item):
    """
    Perpendicular distance between two grid lines (mm).
    Midpoint-to-midpoint is wrong for grids with different extents.
    """
    h_line = grid_line_points(host_item)
    e_line = grid_line_points(ex_item)
    if h_line is None or e_line is None:
        return distance_mm_xy(host_item.get("mid_xy"), ex_item.get("mid_xy"))
    (ax, ay), (bx, by) = e_line
    hm = host_item.get("mid_xy")
    if hm is None:
        (hx0, hy0), (hx1, hy1) = h_line
        hm = ((hx0 + hx1) * 0.5, (hy0 + hy1) * 0.5)
    try:
        d_ft = point_to_line_distance_ft(hm[0], hm[1], ax, ay, bx, by)
        return feet_to_mm(d_ft)
    except Exception:
        return distance_mm_xy(host_item.get("mid_xy"), ex_item.get("mid_xy"))


def status_from_delta(delta, yellow, red):
    if delta is None:
        return "YELLOW"
    d = abs(float(delta))
    if d > float(red):
        return "RED"
    if d > float(yellow):
        return "YELLOW"
    return "GREEN"


def worst_status(statuses):
    order = {"GREEN": 0, "YELLOW": 1, "RED": 2, "ERROR": 3}
    worst = "GREEN"
    for s in statuses:
        if not s:
            continue
        su = u(s).upper()
        if order.get(su, 0) > order.get(worst, 0):
            worst = su
    return worst


def _eid_int(eid):
    try:
        return int(eid.IntegerValue)
    except Exception:
        try:
            return int(eid)
        except Exception:
            return None


def _median_num(vals):
    ds = sorted(float(x) for x in vals)
    n = len(ds)
    if n == 0:
        return 0.0
    if n % 2 == 1:
        return ds[n // 2]
    return 0.5 * (ds[n // 2 - 1] + ds[n // 2])


def _level_elevation_ft(lv):
    """Level plane Z in document internal coordinates (not PBP display)."""
    try:
        return float(lv.ProjectElevation)
    except Exception:
        return float(lv.Elevation)


def collect_levels(document, transform=None):
    """
    name -> list of dicts {id, name, elev_ft, elev_mm, element}
    Uses ProjectElevation (internal origin). If transform is set (link -> host),
    Z is mapped into host coordinates.
    """
    from Autodesk.Revit.DB import FilteredElementCollector, Level, XYZ

    by_name = {}
    for lv in FilteredElementCollector(document).OfClass(Level):
        name = u(lv.Name).strip()
        elev_ft = _level_elevation_ft(lv)
        if transform is not None:
            try:
                elev_ft = float(transform.OfPoint(XYZ(0.0, 0.0, elev_ft)).Z)
            except Exception:
                pass
        item = {
            "id": lv.Id,
            "name": name,
            "elev_ft": elev_ft,
            "elev_mm": feet_to_mm(elev_ft),
            "element": lv,
        }
        by_name.setdefault(name, []).append(item)
    return by_name


def collect_grids(document, transform=None):
    """
    name -> list of dicts {id, name, mid_xy, p0, p1, angle_deg, curve, element}
    If transform is set (link -> host), points/angle are in host coordinates.
    """
    from Autodesk.Revit.DB import FilteredElementCollector, Grid

    by_name = {}
    for g in FilteredElementCollector(document).OfClass(Grid):
        name = u(g.Name).strip()
        curve = None
        try:
            curve = g.Curve
        except Exception:
            curve = None
        mid_xy = None
        angle = None
        p0_xy = None
        p1_xy = None
        z_ft = None
        z_mm = None
        if curve is not None:
            try:
                p0 = curve.GetEndPoint(0)
                p1 = curve.GetEndPoint(1)
                if transform is not None:
                    p0 = transform.OfPoint(p0)
                    p1 = transform.OfPoint(p1)
                p0_xy = (p0.X, p0.Y)
                p1_xy = (p1.X, p1.Y)
                mid_xy = ((p0.X + p1.X) * 0.5, (p0.Y + p1.Y) * 0.5)
                z_ft = (p0.Z + p1.Z) * 0.5
                z_mm = feet_to_mm(z_ft)
                angle = math.degrees(math.atan2(p1.Y - p0.Y, p1.X - p0.X))
            except Exception:
                mid_xy = midpoint_xy(curve) if transform is None else None
                angle = angle_deg_of_curve(curve) if transform is None else None
                p0_xy = None
                p1_xy = None
                z_ft = None
                z_mm = None
        item = {
            "id": g.Id,
            "name": name,
            "mid_xy": mid_xy,
            "p0": p0_xy,
            "p1": p1_xy,
            "z_ft": z_ft,
            "z_mm": z_mm,
            "angle_deg": angle,
            "curve": curve,
            "element": g,
        }
        by_name.setdefault(name, []).append(item)
    return by_name


def flatten_by_name(by_name):
    out = []
    for items in by_name.values():
        out.extend(items)
    return out


def index_by_id(items):
    out = {}
    for it in items:
        out[_eid_int(it["id"])] = it
    return out


def host_monitors_link(host_element, link_instance_id):
    """
    True if host element Copy/Monitors the given RevitLinkInstance.

    Revit API: GetMonitoredLinkElementIds() returns IList<ElementId> of
    *link instance* ids — NOT LinkElementId / linked element ids.
    There is no public API for the specific monitored element in the link.
    """
    try:
        if not host_element.IsMonitoringLinkElement():
            return False
    except Exception:
        return False
    try:
        targets = host_element.GetMonitoredLinkElementIds()
    except Exception:
        return False
    if not targets:
        return False
    want = _eid_int(link_instance_id)
    if want is None:
        return False
    for tid in targets:
        try:
            # ElementId of link instance (documented return type)
            if _eid_int(tid) == want:
                return True
        except Exception:
            pass
        # Legacy / mistaken LinkElementId shape — keep tolerant
        try:
            if _eid_int(tid.LinkInstanceId) == want:
                return True
        except Exception:
            pass
    return False


def _pick_ex_by_name(ex_candidates_by_name, name, used_ex):
    lst = [
        e
        for e in ex_candidates_by_name.get(name, [])
        if _eid_int(e["id"]) not in used_ex
    ]
    if len(lst) == 1:
        return lst[0]
    return None


def _pick_ex_by_elev(h, ex_list, used_ex, max_mm=None):
    best = None
    best_d = None
    for e in ex_list:
        eid = _eid_int(e["id"])
        if eid in used_ex:
            continue
        if "elev_mm" not in h or "elev_mm" not in e:
            continue
        d = abs(float(h["elev_mm"]) - float(e["elev_mm"]))
        if best_d is None or d < best_d:
            best_d = d
            best = e
    if best is None:
        return None
    if max_mm is not None and best_d is not None and best_d > float(max_mm):
        return None
    return best


def _pick_ex_by_grid_near(h, ex_list, used_ex, max_mm=None):
    best = None
    best_d = None
    for e in ex_list:
        eid = _eid_int(e["id"])
        if eid in used_ex:
            continue
        d = grid_perp_distance_mm(h, e)
        if d is None:
            continue
        if best_d is None or d < best_d:
            best_d = d
            best = e
    if best is None:
        return None
    if max_mm is not None and best_d is not None and best_d > float(max_mm):
        return None
    return best


def build_monitor_pairs(host_items, ex_by_id, link_instance_id, kind=None):
    """
    Host elements that Monitor the exemplar link instance.

    API does not expose which linked element is monitored, so we pair:
      1) unique name match among unused exemplar elements
      2) levels: closest elevation; grids: closest perpendicular distance

    Returns pairs / broken / duplicate_ex / paired id sets.
    """
    pairs = []
    broken = []
    paired_host = set()
    paired_ex = set()

    ex_list = list(ex_by_id.values()) if isinstance(ex_by_id, dict) else []
    # Also accept flat list via index_by_id callers — ex_by_id is always dict here
    if not ex_list and isinstance(ex_by_id, dict):
        ex_list = list(ex_by_id.values())

    ex_by_name = {}
    for e in ex_list:
        ex_by_name.setdefault(e["name"], []).append(e)

    monitoring_hosts = []
    for h in host_items:
        el = h.get("element")
        if el is None:
            continue
        if host_monitors_link(el, link_instance_id):
            monitoring_hosts.append(h)

    # First pass: unique name matches
    for h in monitoring_hosts:
        e0 = _pick_ex_by_name(ex_by_name, h["name"], paired_ex)
        if e0 is None:
            continue
        pairs.append((h, e0))
        paired_host.add(_eid_int(h["id"]))
        paired_ex.add(_eid_int(e0["id"]))

    # Second pass: closest unused exemplar (elevation / line distance), unique greedy
    pending = [h for h in monitoring_hosts if _eid_int(h["id"]) not in paired_host]
    scored = []
    for h in pending:
        for e in ex_list:
            eid = _eid_int(e["id"])
            if eid in paired_ex:
                continue
            if kind == "Level" or ("elev_mm" in h and "elev_mm" in e):
                try:
                    d = abs(float(h["elev_mm"]) - float(e["elev_mm"]))
                except Exception:
                    continue
            else:
                d = grid_perp_distance_mm(h, e)
                if d is None:
                    continue
            scored.append((d, h, e))
    scored.sort(key=lambda t: t[0])
    for _d, h, e0 in scored:
        hid = _eid_int(h["id"])
        eid = _eid_int(e0["id"])
        if hid in paired_host or eid in paired_ex:
            continue
        pairs.append((h, e0))
        paired_host.add(hid)
        paired_ex.add(eid)

    for h in pending:
        hid = _eid_int(h["id"])
        if hid in paired_host:
            continue
        broken.append(h)
        paired_host.add(hid)

    # Detect many hosts claiming same ex via name (rare with unique names)
    duplicate_ex = {}
    ex_to_hosts = {}
    for h0, e0 in pairs:
        ex_to_hosts.setdefault(_eid_int(e0["id"]), []).append(h0)
    for ex_id, hosts in ex_to_hosts.items():
        if len(hosts) > 1:
            duplicate_ex[ex_id] = hosts

    return {
        "pairs": pairs,
        "broken": broken,
        "duplicate_ex": duplicate_ex,
        "paired_host_ids": paired_host,
        "paired_ex_ids": paired_ex,
    }


def get_pairing_for_apply(host_doc, exemplar_doc, link_instance, kind):
    """
    kind: 'Level' or 'Grid'
    Returns list of (host_item, ex_item, via_monitor bool) for Apply.
    Monitor pairs first, then name fallback for remaining.
    """
    link_tf = None
    link_inst_id = None
    if link_instance is not None:
        try:
            link_inst_id = link_instance.Id
            link_tf = link_instance.GetTotalTransform()
        except Exception:
            link_inst_id = None
            link_tf = None

    if kind == "Level":
        host_map = collect_levels(host_doc)
        ex_map = collect_levels(exemplar_doc, transform=link_tf)
        mon_kind = "Level"
    else:
        host_map = collect_grids(host_doc)
        ex_map = collect_grids(exemplar_doc, transform=link_tf)
        mon_kind = "Grid"

    host_list = flatten_by_name(host_map)
    ex_list = flatten_by_name(ex_map)
    ex_by_id = index_by_id(ex_list)

    result = []
    used_host = set()
    used_ex = set()

    if link_inst_id is not None:
        mon = build_monitor_pairs(host_list, ex_by_id, link_inst_id, kind=mon_kind)
        seen = set()
        for h0, e0 in mon["pairs"]:
            key = (_eid_int(h0["id"]), _eid_int(e0["id"]))
            if key in seen:
                continue
            seen.add(key)
            result.append((h0, e0, True))
            used_host.add(_eid_int(h0["id"]))
            used_ex.add(_eid_int(e0["id"]))

    remain_ex_by_name = {}
    for e in ex_list:
        if _eid_int(e["id"]) in used_ex:
            continue
        remain_ex_by_name.setdefault(e["name"], []).append(e)

    for h in host_list:
        if _eid_int(h["id"]) in used_host:
            continue
        e_list = remain_ex_by_name.get(h["name"], [])
        if len(e_list) == 1:
            e0 = e_list[0]
            result.append((h, e0, False))
            used_host.add(_eid_int(h["id"]))
            used_ex.add(_eid_int(e0["id"]))

    return result


def _tol_num(tolerances, key, default):
    try:
        return float(tolerances[key])
    except Exception:
        try:
            return float(default)
        except Exception:
            return 0.0


def _fmt_mm(value):
    if value is None or value == u"":
        return u""
    try:
        return u"{:.3f}".format(float(value))
    except Exception:
        return u(value)


def _fmt_deg(value):
    if value is None or value == u"":
        return u""
    try:
        return u"{:.4f}".format(float(value))
    except Exception:
        return u(value)


def _geo_detail(
    category, name_host, name_ex, problem, val_host, val_ex, delta, tol, status
):
    return {
        "category": category,
        "name_host": name_host or u"",
        "name_exemplar": name_ex or u"",
        "problem": problem,
        "value_host": val_host if val_host is not None else u"",
        "value_exemplar": val_ex if val_ex is not None else u"",
        "delta": delta if delta is not None else u"",
        "tolerance": tol if tol is not None else u"",
        "status": status,
        "source": "geometry",
    }


def diff_named_levels(host_lv, ex_lv, tolerances, whitelist=None):
    """
    Name-based level compare in host coordinates (link already transformed).

    Same-name pair with |elevation| above the yellow threshold -> elevation_delta.
    Name only in the exemplar -> missing_in_host (leftover).
    Name only in the host -> extra_in_host.
    A count mismatch (10 host vs 19 exemplar) therefore cannot stay GREEN.
    """
    details = []
    wl = None
    if whitelist:
        wl = set([u(x).strip() for x in whitelist if u(x).strip()])

    y_mm = _tol_num(tolerances, "level_delta_yellow_mm", 2.0)
    r_mm = _tol_num(tolerances, "level_delta_red_mm", 10.0)
    tol_mm = _tol_num(tolerances, "level_mm", y_mm)

    names = set(host_lv.keys()) | set(ex_lv.keys())
    for name in sorted(names):
        h_list = list(host_lv.get(name, []) or [])
        e_list = list(ex_lv.get(name, []) or [])

        # Host-only names outside the optional whitelist are ignored.
        if wl is not None and name not in wl and not e_list:
            continue

        if len(h_list) > 1:
            details.append(
                _geo_detail(
                    "Level",
                    name,
                    name,
                    "duplicate_host",
                    len(h_list),
                    u"",
                    u"",
                    u"",
                    "RED",
                )
            )
        if len(e_list) > 1:
            details.append(
                _geo_detail(
                    "Level",
                    name,
                    name,
                    "duplicate_exemplar",
                    u"",
                    len(e_list),
                    u"",
                    u"",
                    "YELLOW",
                )
            )

        if not h_list and e_list:
            details.append(
                _geo_detail(
                    "Level",
                    u"",
                    name,
                    "missing_in_host",
                    u"",
                    _fmt_mm(e_list[0].get("elev_mm")),
                    u"",
                    tol_mm,
                    "RED",
                )
            )
            continue
        if h_list and not e_list:
            if wl is not None and name not in wl:
                continue
            details.append(
                _geo_detail(
                    "Level",
                    name,
                    u"",
                    "extra_in_host",
                    _fmt_mm(h_list[0].get("elev_mm")),
                    u"",
                    u"",
                    tol_mm,
                    "RED",
                )
            )
            continue
        if not h_list or not e_list:
            continue
        if wl is not None and name not in wl:
            continue

        h0 = h_list[0]
        e0 = e_list[0]
        delta = None
        try:
            delta = float(h0["elev_mm"]) - float(e0["elev_mm"])
        except Exception:
            delta = None
        st = status_from_delta(delta, y_mm, r_mm)
        if st != "GREEN":
            details.append(
                _geo_detail(
                    "Level",
                    name,
                    name,
                    "elevation_delta",
                    _fmt_mm(h0.get("elev_mm")),
                    _fmt_mm(e0.get("elev_mm")),
                    _fmt_mm(delta),
                    tol_mm,
                    st,
                )
            )

    if wl is not None:
        for name in sorted(wl):
            if name in host_lv or name in ex_lv:
                continue
            details.append(
                _geo_detail(
                    "Level",
                    u"",
                    name,
                    "whitelist_missing",
                    u"",
                    u"",
                    u"",
                    u"",
                    "YELLOW",
                )
            )
    return details


def diff_named_grids(host_gr, ex_gr, tolerances):
    """
    Name-based grid compare in host coordinates (link already transformed).

    Position is the perpendicular distance between the two lines, so a longer
    grid on the same axis is not a move. Angle is a separate row.
    Missing / extra names are reported the same way as levels.
    """
    details = []
    y_mm = _tol_num(tolerances, "grid_delta_yellow_mm", 5.0)
    r_mm = _tol_num(tolerances, "grid_delta_red_mm", 20.0)
    tol_mm = _tol_num(tolerances, "grid_mm", y_mm)
    y_ang = _tol_num(tolerances, "grid_angle_yellow_deg", 0.1)
    r_ang = _tol_num(tolerances, "grid_angle_red_deg", 1.0)
    tol_ang = _tol_num(tolerances, "grid_angle_deg", y_ang)

    names = set(host_gr.keys()) | set(ex_gr.keys())
    for name in sorted(names):
        h_list = list(host_gr.get(name, []) or [])
        e_list = list(ex_gr.get(name, []) or [])

        if len(h_list) > 1:
            details.append(
                _geo_detail(
                    "Grid",
                    name,
                    name,
                    "duplicate_host",
                    len(h_list),
                    u"",
                    u"",
                    u"",
                    "RED",
                )
            )
        if len(e_list) > 1:
            details.append(
                _geo_detail(
                    "Grid",
                    name,
                    name,
                    "duplicate_exemplar",
                    u"",
                    len(e_list),
                    u"",
                    u"",
                    "YELLOW",
                )
            )

        if not h_list and e_list:
            details.append(
                _geo_detail(
                    "Grid",
                    u"",
                    name,
                    "missing_in_host",
                    u"",
                    name,
                    u"",
                    tol_mm,
                    "RED",
                )
            )
            continue
        if h_list and not e_list:
            details.append(
                _geo_detail(
                    "Grid",
                    name,
                    u"",
                    "extra_in_host",
                    name,
                    u"",
                    u"",
                    tol_mm,
                    "RED",
                )
            )
            continue
        if not h_list or not e_list:
            continue

        h0 = h_list[0]
        e0 = e_list[0]
        d_pos = grid_perp_distance_mm(h0, e0)
        st_pos = status_from_delta(d_pos, y_mm, r_mm)
        if st_pos != "GREEN":
            details.append(
                _geo_detail(
                    "Grid",
                    name,
                    name,
                    "position_delta",
                    fmt_xy_mm(h0.get("mid_xy")),
                    fmt_xy_mm(e0.get("mid_xy")),
                    _fmt_mm(d_pos),
                    tol_mm,
                    st_pos,
                )
            )

        d_ang = angle_delta_deg(h0.get("angle_deg"), e0.get("angle_deg"))
        st_ang = status_from_delta(d_ang, y_ang, r_ang)
        if st_ang != "GREEN":
            details.append(
                _geo_detail(
                    "Grid",
                    name,
                    name,
                    "angle_delta",
                    _fmt_deg(h0.get("angle_deg")),
                    _fmt_deg(e0.get("angle_deg")),
                    _fmt_deg(d_ang),
                    tol_ang,
                    st_ang,
                )
            )
    return details


def compare_levels_grids(
    host_doc,
    exemplar_doc,
    tolerances,
    whitelist=None,
    exemplar_link_instance=None,
    host_model_path=None,
):
    """
    Level elevations, grid position/angle, missing/extra names, and Project
    Base Point vs the exemplar link (link geometry already in host coordinates).

    Copy/Monitor warnings are not read here. The audit merges
    Document.GetWarnings() afterwards. A count or geometry mismatch is an
    issue even when GetWarnings is empty.
    Returns (details_list, summary_counts).
    """
    details = []
    link_tf = None
    if exemplar_link_instance is not None:
        try:
            link_tf = exemplar_link_instance.GetTotalTransform()
        except Exception:
            link_tf = None

    host_lv = collect_levels(host_doc)
    ex_lv = collect_levels(exemplar_doc, transform=link_tf)
    host_gr = collect_grids(host_doc)
    ex_gr = collect_grids(exemplar_doc, transform=link_tf)

    host_lv_list = flatten_by_name(host_lv)
    ex_lv_list = flatten_by_name(ex_lv)
    host_gr_list = flatten_by_name(host_gr)
    ex_gr_list = flatten_by_name(ex_gr)

    def add_detail(
        category,
        name_host,
        name_ex,
        problem,
        val_host,
        val_ex,
        delta,
        tol,
        status,
    ):
        details.append(
            {
                "category": category,
                "name_host": name_host or "",
                "name_exemplar": name_ex or "",
                "problem": problem,
                "value_host": val_host if val_host is not None else "",
                "value_exemplar": val_ex if val_ex is not None else "",
                "delta": delta if delta is not None else "",
                "tolerance": tol if tol is not None else "",
                "status": status,
            }
        )

    def fmt_pbp_delta(xy, z, shared_prefix=False):
        if xy is None:
            return u""
        z_s = 0.0 if z is None else z
        text = u("XY={:.1f}; Z={:.1f}".format(xy, z_s))
        if shared_prefix:
            return u"общ. " + text
        return text

    def compare_coord_point(is_survey):
        """PBP (False) or Survey Point (True). Report-only, never applied."""
        if is_survey:
            name = u"Точка съемки"
            code_delta = "survey_delta"
            code_miss = "survey_missing"
        else:
            name = u"Базовая точка проекта"
            code_delta = "pbp_delta"
            code_miss = "pbp_missing"
        host_bp = collect_base_point(host_doc, is_survey=is_survey, transform=None)
        ex_bp = collect_base_point(
            exemplar_doc, is_survey=is_survey, transform=link_tf
        )
        if host_bp is None and ex_bp is None:
            return
        if host_bp is None or ex_bp is None:
            add_detail(
                "Coords",
                name if host_bp else "",
                name if ex_bp else "",
                code_miss,
                fmt_xyz_mm(host_bp["pos"]) if host_bp else "",
                fmt_xyz_mm(ex_bp["pos"]) if ex_bp else "",
                "",
                tolerances["pbp_delta_yellow_mm"],
                "RED",
            )
            return

        geom_xy, geom_z = (None, None)
        if link_tf is not None:
            geom_xy, geom_z = xyz_delta_mm(host_bp.get("pos"), ex_bp.get("pos"))
        sh_xy, sh_z = xyz_delta_mm(host_bp.get("shared"), ex_bp.get("shared"))

        y_xy = tolerances["pbp_delta_yellow_mm"]
        r_xy = tolerances["pbp_delta_red_mm"]
        y_z = tolerances["pbp_z_yellow_mm"]
        r_z = tolerances["pbp_z_red_mm"]

        statuses = []
        if geom_xy is not None:
            statuses.append(status_from_delta(geom_xy, y_xy, r_xy))
            statuses.append(status_from_delta(geom_z, y_z, r_z))
        if sh_xy is not None:
            statuses.append(status_from_delta(sh_xy, y_xy, r_xy))
            statuses.append(status_from_delta(sh_z, y_z, r_z))
        if not statuses:
            add_detail(
                "Coords",
                name,
                name,
                code_miss,
                fmt_xyz_mm(host_bp.get("pos")),
                fmt_xyz_mm(ex_bp.get("pos")),
                "",
                y_xy,
                "YELLOW",
            )
            return

        st = worst_status(statuses)
        if st == "GREEN":
            return

        parts = []
        geom_bad = False
        shared_bad = False
        if geom_xy is not None:
            st_g = worst_status(
                [
                    status_from_delta(geom_xy, y_xy, r_xy),
                    status_from_delta(geom_z, y_z, r_z),
                ]
            )
            geom_bad = st_g != "GREEN"
        if sh_xy is not None:
            st_s = worst_status(
                [
                    status_from_delta(sh_xy, y_xy, r_xy),
                    status_from_delta(sh_z, y_z, r_z),
                ]
            )
            shared_bad = st_s != "GREEN"
        if geom_bad:
            parts.append(fmt_pbp_delta(geom_xy, geom_z, False))
        if shared_bad:
            parts.append(fmt_pbp_delta(sh_xy, sh_z, True))
        if not parts:
            if geom_xy is not None:
                parts.append(fmt_pbp_delta(geom_xy, geom_z, False))
            elif sh_xy is not None:
                parts.append(fmt_pbp_delta(sh_xy, sh_z, True))

        val_h = fmt_xyz_mm(host_bp.get("pos")) or fmt_xyz_mm(host_bp.get("shared"))
        val_e = fmt_xyz_mm(ex_bp.get("pos")) or fmt_xyz_mm(ex_bp.get("shared"))
        if shared_bad and not geom_bad:
            val_h = fmt_xyz_mm(host_bp.get("shared")) or val_h
            val_e = fmt_xyz_mm(ex_bp.get("shared")) or val_e

        add_detail(
            "Coords",
            name,
            name,
            code_delta,
            val_h,
            val_e,
            u" | ".join([p for p in parts if p]),
            y_xy,
            st,
        )

    details.extend(diff_named_levels(host_lv, ex_lv, tolerances, whitelist))
    details.extend(diff_named_grids(host_gr, ex_gr, tolerances))

    compare_coord_point(False)
    for d in details:
        if not d.get("source"):
            d["source"] = "pbp"

    statuses = [d["status"] for d in details]
    model_status = worst_status(statuses) if details else "GREEN"

    counts = {
        "levels_host": len(host_lv_list),
        "levels_exemplar": len(ex_lv_list),
        "grids_host": len(host_gr_list),
        "grids_exemplar": len(ex_gr_list),
        "issues_total": len(details),
        "issues_red": sum(1 for d in details if d["status"] == "RED"),
        "issues_yellow": sum(1 for d in details if d["status"] == "YELLOW"),
        "status": model_status,
    }
    return details, counts


def _fid_guid_str(fid):
    if fid is None:
        return u""
    try:
        return u(fid.Guid).lower().strip("{}")
    except Exception:
        return u""


# From RBP open-time warnings on SITE_ORG_BF_K4 (2026-08-20):
#   3d983f31-...  «Для экземпляра связи требуется просмотр координаций»
#   c6d534b0-...  «Наблюдаемая связь перемещена»
CR_GUID_LINK_NEEDS_REVIEW = "3d983f31-9ee3-4c3a-bed8-663b32cecec5"
CR_GUID_WATCHED_LINK_MOVED = "c6d534b0-95bd-4047-89c1-32e3878440a6"
CR_KNOWN_GUIDS = set(
    [
        CR_GUID_LINK_NEEDS_REVIEW,
        CR_GUID_WATCHED_LINK_MOVED,
    ]
)


def _failure_ids_match(fid_a, fid_b):
    if fid_a is None or fid_b is None:
        return False
    try:
        if fid_a == fid_b:
            return True
    except Exception:
        pass
    try:
        ga = fid_a.Guid
        gb = fid_b.Guid
        if ga and gb and ga == gb:
            return True
    except Exception:
        pass
    try:
        return int(fid_a.IntegerValue) == int(fid_b.IntegerValue)
    except Exception:
        return False


def _collect_copy_monitor_failure_ids():
    """BuiltInFailures.CopyMonitorFailures + LinkFailures ids (best effort)."""
    ids = []
    try:
        from Autodesk.Revit.DB import BuiltInFailures

        cm = BuiltInFailures.CopyMonitorFailures
        for name in dir(cm):
            if name.startswith("_"):
                continue
            try:
                val = getattr(cm, name)
                if val is not None:
                    ids.append(val)
            except Exception:
                pass
        try:
            lf = BuiltInFailures.LinkFailures
            for name in ("LinkInstanceNeedsReconcile",):
                try:
                    val = getattr(lf, name)
                    if val is not None:
                        ids.append(val)
                except Exception:
                    pass
        except Exception:
            pass
    except Exception:
        pass
    return ids


def _desc_is_copy_monitor(desc):
    low = u(desc or "").lower()
    if not low:
        return False
    needles = (
        u"coordination",
        u"copywatch",
        u"copy/monitor",
        u"copy monitor",
        u"мониторинг",
        u"координац",
        u"просмотр координац",
        u"наблюдаемая связь",
        u"связь перемещ",
        u"watched link",
        u"reconcile",
    )
    for n in needles:
        if n in low:
            return True
    return False


def _parse_coord_review_problem(desc, fid=None):
    """
    Map Revit warning / CR description -> (problem_code, status).
    Wording aligned with Coordination Review UI (RU + EN).
    """
    guid = _fid_guid_str(fid)
    if guid == CR_GUID_WATCHED_LINK_MOVED:
        return "cr_watched_link_moved", "RED"
    if guid == CR_GUID_LINK_NEEDS_REVIEW:
        return "cr_link_needs_review", "YELLOW"

    d = u(desc or "").strip()
    low = d.lower()
    if not low:
        return "cr_alert", "YELLOW"

    if u"сетка перемещ" in low or u"grid has been moved" in low or u"grid moved" in low:
        return "cr_grid_moved", "RED"
    if u"уровень перемещ" in low or u"level has been moved" in low or u"level moved" in low:
        return "cr_level_moved", "RED"
    if u"имя изменено" in low or u"name has been changed" in low or u"name changed" in low:
        return "cr_name_changed", "RED"
    if u"элемент удален" in low or u"element has been deleted" in low or u"element deleted" in low:
        return "cr_element_deleted", "RED"
    if (
        u"наблюдаемая связь" in low
        or u"watched link" in low
        or u"связь перемещ" in low
    ):
        return "cr_watched_link_moved", "RED"
    if (
        u"coordination review" in low
        or u"просмотр координац" in low
        or u"требуется просмотр" in low
        or u"проверка координац" in low
        or u"нужна проверка" in low
        or u"link instance needs" in low
    ):
        return "cr_link_needs_review", "YELLOW"
    if u"coordination monitor" in low or u"copywatch" in low or u"мониторинг" in low:
        return "cr_alert", "YELLOW"
    return "cr_alert", "YELLOW"


def _warning_relates_to_exemplar_link(document, fm, link_inst_id, link_label):
    if link_inst_id is None and not link_label:
        return True
    want_inst = _eid_int(link_inst_id)
    want_type = None
    try:
        inst_el = document.GetElement(link_inst_id)
        if inst_el is not None:
            want_type = _eid_int(inst_el.GetTypeId())
    except Exception:
        want_type = None

    try:
        elems = list(fm.GetFailingElements())
    except Exception:
        elems = []
    if not elems:
        # RU warnings often have no file name in the text — keep if we have an exemplar.
        return want_inst is not None

    from Autodesk.Revit.DB import RevitLinkInstance, RevitLinkType

    for eid in elems:
        eid_n = _eid_int(eid)
        if want_inst is not None and eid_n == want_inst:
            return True
        if want_type is not None and eid_n == want_type:
            return True
        try:
            el = document.GetElement(eid)
        except Exception:
            el = None
        if el is None:
            continue
        try:
            if isinstance(el, RevitLinkInstance) and want_inst is not None:
                if _eid_int(el.Id) == want_inst:
                    return True
                if want_type is not None and _eid_int(el.GetTypeId()) == want_type:
                    return True
            if isinstance(el, RevitLinkType) and want_type is not None:
                if _eid_int(el.Id) == want_type:
                    return True
        except Exception:
            pass
        if want_inst is not None and host_monitors_link(el, link_inst_id):
            return True
    if link_label:
        desc = u(fm.GetDescriptionText() or "").lower()
        lb = path_basename(link_label).lower()
        stem = lb[:-4] if lb.endswith(".rvt") else lb
        if lb and (lb in desc or (stem and stem in desc)):
            return True
    return False


def _element_level_grid_category(document, element):
    if element is None:
        return None, u""
    try:
        from Autodesk.Revit.DB import Grid, Level

        if isinstance(element, Level):
            return "Level", u(element.Name).strip()
        if isinstance(element, Grid):
            return "Grid", u(element.Name).strip()
    except Exception:
        pass
    try:
        cat = element.Category
        if cat is None:
            return None, u""
        nm = u""
        try:
            nm = u(element.Name).strip()
        except Exception:
            nm = u""
        try:
            bic = cat.Id.IntegerValue
            from Autodesk.Revit.DB import BuiltInCategory

            if bic == int(BuiltInCategory.OST_Levels):
                return "Level", nm
            if bic == int(BuiltInCategory.OST_Grids):
                return "Grid", nm
        except Exception:
            pass
    except Exception:
        pass
    return None, u""


def collect_coordination_review_warnings(
    document, link_instance_id=None, link_label=None
):
    """
    Read Copy/Monitor + link-reconcile warnings from Document.GetWarnings().
    Returns (detail_rows, total_warning_count).
    """
    rows = []
    cm_ids = _collect_copy_monitor_failure_ids()
    try:
        warnings = list(document.GetWarnings())
    except Exception:
        warnings = []

    seen = set()
    for fm in warnings:
        if fm is None:
            continue
        try:
            fid = fm.GetFailureDefinitionId()
        except Exception:
            fid = None

        try:
            desc = u(fm.GetDescriptionText() or "").strip()
        except Exception:
            desc = u""

        is_cm = _fid_guid_str(fid) in CR_KNOWN_GUIDS or _desc_is_copy_monitor(desc)
        if not is_cm:
            for cm_fid in cm_ids:
                if _failure_ids_match(fid, cm_fid):
                    is_cm = True
                    break
        if not is_cm:
            continue

        if not _warning_relates_to_exemplar_link(
            document, fm, link_instance_id, link_label
        ):
            continue

        code, status = _parse_coord_review_problem(desc, fid)

        try:
            elem_ids = list(fm.GetFailingElements())
        except Exception:
            elem_ids = []

        if not elem_ids:
            key = ("System", code, desc)
            if key in seen:
                continue
            seen.add(key)
            rows.append(
                {
                    "category": "System",
                    "name_host": u"",
                    "name_exemplar": u"",
                    "problem": code,
                    "value_host": desc,
                    "value_exemplar": u"",
                    "delta": u"",
                    "tolerance": u"",
                    "status": status,
                    "source": "revit_cr",
                }
            )
            continue

        for eid in elem_ids:
            try:
                el = document.GetElement(eid)
            except Exception:
                el = None
            cat, el_name = _element_level_grid_category(document, el)
            if cat is None:
                cat = "System"
                if not el_name:
                    try:
                        el_name = u(el.Name).strip() if el else u""
                    except Exception:
                        el_name = u""
            key = (cat, el_name or desc, code)
            if key in seen:
                continue
            seen.add(key)
            rows.append(
                {
                    "category": cat,
                    "name_host": el_name,
                    "name_exemplar": u"",
                    "problem": code,
                    "value_host": desc or el_name,
                    "value_exemplar": u"",
                    "delta": u"",
                    "tolerance": u"",
                    "status": status,
                    "source": "revit_cr",
                }
            )

    return rows, len(warnings)


def _detail_element_key(detail):
    """(category, lowercased host-or-exemplar name) for CR/geometry merge."""
    cat = detail.get("category") or ""
    el = u(
        detail.get("name_host") or detail.get("name_exemplar") or detail.get("element") or ""
    ).strip().lower()
    return cat, el


def merge_geo_with_coord_review(geo_details, cr_details):
    """
    Keep measured level/grid/PBP rows. Append GetWarnings rows that add a
    signal the geometry pass did not already record for that element.

    Individual Coordination Review lines («Уровень перемещен пользователем»)
    often never appear in Document.GetWarnings(). Dropping a measured delta
    just because a warning exists would hide the millimetre shift. A warning
    with no matching geometry row is still kept, so a real CR hit cannot
    stay GREEN.
    """
    geo_details = list(geo_details or [])
    if not cr_details:
        return geo_details

    move_codes = set(
        [
            "elevation_delta",
            "position_delta",
            "angle_delta",
            "grid_relative_moved",
        ]
    )
    cr_move_codes = set(
        [
            "cr_level_moved",
            "cr_grid_moved",
            "grid_relative_moved",
        ]
    )
    covered_moves = set()
    for d in geo_details:
        code = d.get("problem") or ""
        if code not in move_codes:
            continue
        cat, el = _detail_element_key(d)
        if cat in ("Level", "Grid") and el:
            covered_moves.add((cat, el))

    out = list(geo_details)
    for d in cr_details:
        code = d.get("problem") or ""
        cat, el = _detail_element_key(d)
        if (
            code in cr_move_codes
            and cat in ("Level", "Grid")
            and el
            and (cat, el) in covered_moves
        ):
            continue
        out.append(d)
    return out


def recount_issues(details):
    statuses = [d.get("status") for d in details if d.get("status")]
    model_status = worst_status(statuses) if details else "GREEN"
    return {
        "issues_total": len(details),
        "issues_red": sum(1 for d in details if d.get("status") == "RED"),
        "issues_yellow": sum(1 for d in details if d.get("status") == "YELLOW"),
        "status": model_status,
    }


def csv_escape(value):
    text = u(value)
    if (
        u";" in text
        or u"," in text
        or u'"' in text
        or u"\n" in text
        or u"\r" in text
    ):
        text = u'"' + text.replace(u'"', u'""') + u'"'
    return text


def append_csv(path, headers_en, headers_ru, row_dict):
    write_header = (not os.path.isfile(path)) or (os.path.getsize(path) == 0)
    line_parts = [csv_escape(row_dict.get(h, "")) for h in headers_en]
    line = u";".join(line_parts) + u"\r\n"
    if write_header:
        header = u";".join(headers_ru) + u"\r\n"
        with codecs.open(path, "w", "utf-8-sig") as f:
            f.write(header)
            f.write(line)
    else:
        with codecs.open(path, "a", "utf-8") as f:
            f.write(line)


def append_csv_rows(path, headers_en, headers_ru, rows):
    if not rows:
        return
    write_header = (not os.path.isfile(path)) or (os.path.getsize(path) == 0)
    lines = []
    if write_header:
        lines.append(u";".join(headers_ru))
    for row_dict in rows:
        line_parts = [csv_escape(row_dict.get(h, "")) for h in headers_en]
        lines.append(u";".join(line_parts))
    text = u"\r\n".join(lines) + u"\r\n"
    mode = "w" if write_header else "a"
    enc = "utf-8-sig" if write_header else "utf-8"
    with codecs.open(path, mode, enc) as f:
        f.write(text)
