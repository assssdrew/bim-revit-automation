# -*- coding: utf-8 -*-
# Name: lg_lib.py
# Version: 1.0
# What it does: Shared IronPython helpers for levels/grids audit and apply scripts.
# Inputs: Imported by audit/apply after SCRIPT_DIR on sys.path.
# Outputs: Utility functions only (no standalone run).
# How to run: Not run directly; used from levels_grids_*.py in RBP.
# Notes: IronPython 2.7 / RBP compatible.

import codecs
import math
import os
import System

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
    return System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")


def now_file_stamp():
    return System.DateTime.Now.ToString("yyyy-MM-dd_HHmm")


def normalize_model_path(path):
    if not path:
        return ""
    p = path.strip().strip('"').replace("/", "\\")
    while "\\\\" in p[2:]:
        p = p[:2] + p[2:].replace("\\\\", "\\")
    return p.lower()


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
            return ""
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
        return ""
    except Exception:
        return None


def resolve_exemplar_hint(script_dir, tolerances, project_row=None):
    """
    Priority:
      1) project_map row override (optional advanced)
      2) exemplar_model.cfg / exemplar_link.cfg (session chooser)
      3) tolerances.cfg exemplar_link_default
      4) empty → first loaded RVT link
    """
    if project_row:
        hint = (project_row.get("exemplar_link_name") or "").strip()
        if hint:
            return hint
    cfg_hint = read_exemplar_link_cfg(script_dir)
    if cfg_hint is not None:
        return cfg_hint
    return (tolerances.get("exemplar_link_default") or "").strip()


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


def find_exemplar_link(document, link_hint):
    """
    Find RevitLinkInstance by:
      - full model path / basename (from exemplar_model.cfg picker), or
      - name/type substring (legacy), or
      - first loaded RVT link if hint empty.
    Returns (instance, link_doc, match_label) or (None, None, reason).
    """
    from Autodesk.Revit.DB import FilteredElementCollector, RevitLinkInstance

    hint_raw = u(link_hint or "").strip()
    hint = hint_raw.lower()
    hint_base = path_basename(hint).lower()
    if hint_base.endswith(".rvt"):
        hint_stem = hint_base[:-4]
    else:
        hint_stem = hint_base
    as_path = _hint_looks_like_path(hint_raw)

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

    if hint:
        matched = [c for c in candidates if _match(c)]
        if matched:
            # Prefer loaded
            for c in matched:
                if c["link_doc"] is not None:
                    return c["inst"], c["link_doc"], c["label"]
            c0 = matched[0]
            return None, None, u"связь найдена, но не загружена: " + c0["label"]
        if as_path:
            return (
                None,
                None,
                u"связь-эталон не найдена по модели: " + hint_raw,
            )
        return None, None, u"связь-эталон не найдена по подстроке: " + hint_raw

    for c in candidates:
        if c["link_doc"] is not None:
            return c["inst"], c["link_doc"], c["label"]
    if candidates:
        return None, None, u"нет загруженных RVT-связей"
    return None, None, u"в модели нет RevitLinkInstance"


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
        "exemplar_link_default": "",
        "apply_elevation": 1.0,
        "apply_grid_curve": 1.0,
        "apply_rename": 0.0,
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


def collect_levels(document):
    """name -> list of dicts {id, name, elev_mm, elev_ft, element}"""
    from Autodesk.Revit.DB import FilteredElementCollector, Level

    by_name = {}
    for lv in FilteredElementCollector(document).OfClass(Level):
        name = u(lv.Name).strip()
        elev_ft = float(lv.Elevation)
        item = {
            "id": lv.Id,
            "name": name,
            "elev_ft": elev_ft,
            "elev_mm": feet_to_mm(elev_ft),
            "element": lv,
        }
        by_name.setdefault(name, []).append(item)
    return by_name


def collect_grids(document):
    """name -> list of dicts {id, name, mid_xy, angle_deg, curve, element}"""
    from Autodesk.Revit.DB import FilteredElementCollector, Grid

    by_name = {}
    for g in FilteredElementCollector(document).OfClass(Grid):
        name = u(g.Name).strip()
        curve = None
        try:
            curve = g.Curve
        except Exception:
            curve = None
        item = {
            "id": g.Id,
            "name": name,
            "mid_xy": midpoint_xy(curve) if curve else None,
            "angle_deg": angle_deg_of_curve(curve) if curve else None,
            "curve": curve,
            "element": g,
        }
        by_name.setdefault(name, []).append(item)
    return by_name


def compare_levels_grids(host_doc, exemplar_doc, tolerances, whitelist=None):
    """
    Returns (details_list, summary_counts).
    details: list of dicts for CSV details sheet.
    """
    details = []
    host_lv = collect_levels(host_doc)
    ex_lv = collect_levels(exemplar_doc)
    host_gr = collect_grids(host_doc)
    ex_gr = collect_grids(exemplar_doc)

    wl = None
    if whitelist:
        wl = set([u(x).strip() for x in whitelist])

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

    # --- Levels ---
    all_level_names = set(host_lv.keys()) | set(ex_lv.keys())
    for name in sorted(all_level_names):
        if wl is not None and name not in wl and name not in ex_lv:
            # host-only outside whitelist: ignore as "extra" noise? still report as extra_ignored
            if name in host_lv and name not in ex_lv:
                continue
        h_list = host_lv.get(name, [])
        e_list = ex_lv.get(name, [])

        if len(h_list) > 1:
            add_detail(
                "Level",
                name,
                name,
                "duplicate_host",
                len(h_list),
                "",
                "",
                "",
                "RED",
            )
        if len(e_list) > 1:
            add_detail(
                "Level",
                name,
                name,
                "duplicate_exemplar",
                "",
                len(e_list),
                "",
                "",
                "YELLOW",
            )

        if not h_list and e_list:
            elev = e_list[0]["elev_mm"]
            add_detail(
                "Level",
                "",
                name,
                "missing_in_host",
                "",
                round(elev, 3),
                "",
                tolerances["level_mm"],
                "RED",
            )
            continue
        if h_list and not e_list:
            if wl is not None and name not in wl:
                continue
            elev = h_list[0]["elev_mm"]
            add_detail(
                "Level",
                name,
                "",
                "extra_in_host",
                round(elev, 3),
                "",
                "",
                tolerances["level_mm"],
                "RED",
            )
            continue

        if wl is not None and name not in wl:
            continue

        h0 = h_list[0]
        e0 = e_list[0]
        delta = h0["elev_mm"] - e0["elev_mm"]
        st = status_from_delta(
            delta,
            tolerances["level_delta_yellow_mm"],
            tolerances["level_delta_red_mm"],
        )
        if st != "GREEN":
            add_detail(
                "Level",
                name,
                name,
                "elevation_delta",
                round(h0["elev_mm"], 3),
                round(e0["elev_mm"], 3),
                round(delta, 3),
                tolerances["level_mm"],
                st,
            )

    # whitelist missing in both? — levels required in whitelist not in host
    if wl is not None:
        for name in sorted(wl):
            if name not in host_lv:
                if name in ex_lv:
                    continue  # already missing_in_host
                add_detail(
                    "Level",
                    "",
                    name,
                    "whitelist_missing",
                    "",
                    "",
                    "",
                    "",
                    "YELLOW",
                )

    # --- Grids ---
    all_grid_names = set(host_gr.keys()) | set(ex_gr.keys())
    for name in sorted(all_grid_names):
        h_list = host_gr.get(name, [])
        e_list = ex_gr.get(name, [])

        if len(h_list) > 1:
            add_detail(
                "Grid",
                name,
                name,
                "duplicate_host",
                len(h_list),
                "",
                "",
                "",
                "RED",
            )
        if len(e_list) > 1:
            add_detail(
                "Grid",
                name,
                name,
                "duplicate_exemplar",
                "",
                len(e_list),
                "",
                "",
                "YELLOW",
            )

        if not h_list and e_list:
            add_detail(
                "Grid",
                "",
                name,
                "missing_in_host",
                "",
                name,
                "",
                tolerances["grid_mm"],
                "RED",
            )
            continue
        if h_list and not e_list:
            add_detail(
                "Grid",
                name,
                "",
                "extra_in_host",
                name,
                "",
                "",
                tolerances["grid_mm"],
                "RED",
            )
            continue

        h0 = h_list[0]
        e0 = e_list[0]
        d_pos = distance_mm_xy(h0["mid_xy"], e0["mid_xy"])
        st_pos = status_from_delta(
            d_pos,
            tolerances["grid_delta_yellow_mm"],
            tolerances["grid_delta_red_mm"],
        )
        if st_pos != "GREEN":
            add_detail(
                "Grid",
                name,
                name,
                "position_delta",
                "" if h0["mid_xy"] is None else "mid",
                "" if e0["mid_xy"] is None else "mid",
                "" if d_pos is None else round(d_pos, 3),
                tolerances["grid_mm"],
                st_pos,
            )

        d_ang = angle_delta_deg(h0["angle_deg"], e0["angle_deg"])
        st_ang = status_from_delta(
            d_ang,
            tolerances["grid_angle_yellow_deg"],
            tolerances["grid_angle_red_deg"],
        )
        if st_ang != "GREEN":
            add_detail(
                "Grid",
                name,
                name,
                "angle_delta",
                "" if h0["angle_deg"] is None else round(h0["angle_deg"], 4),
                "" if e0["angle_deg"] is None else round(e0["angle_deg"], 4),
                "" if d_ang is None else round(d_ang, 4),
                tolerances["grid_angle_deg"],
                st_ang,
            )

    statuses = [d["status"] for d in details]
    model_status = worst_status(statuses) if details else "GREEN"

    counts = {
        "levels_host": sum(len(v) for v in host_lv.values()),
        "levels_exemplar": sum(len(v) for v in ex_lv.values()),
        "grids_host": sum(len(v) for v in host_gr.values()),
        "grids_exemplar": sum(len(v) for v in ex_gr.values()),
        "issues_total": len(details),
        "issues_red": sum(1 for d in details if d["status"] == "RED"),
        "issues_yellow": sum(1 for d in details if d["status"] == "YELLOW"),
        "status": model_status,
    }
    return details, counts


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
