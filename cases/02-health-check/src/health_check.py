# -*- coding: utf-8 -*-
# Name: health_check.py
# Version: 1.4.0
# What it does: Read-only RBP audit: health metrics (waves A/B/C) and RVT link inventory per model.
# Inputs: rvt_list.txt, thresholds.cfg, xlsx_out_path.cfg; RBP Create New Local, Detach off.
# Outputs: Per-model CSV append; one Excel report (Summary + RVT Links) after the last model.
# How to run: RBP task on rvt_list.txt; optional color_latest_report.cmd to rebuild Excel.
# Notes: Never Save, Sync, or Relinquish; prefer closing locals without saving.

import clr
import codecs
import os
import System

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
from Autodesk.Revit.DB import *

import revit_script_util
from revit_script_util import Output

SCRIPT_VERSION = "1.4.0"
TOOL_DIR_FALLBACK = r""  # optional: path to deployed copy on share

# Метрики отчёта: волны A/B/C (числа в CSV/XLSX).
# В светофор идут только selected keys из thresholds.cfg — см. compute_status.
CSV_FIELDS = [
    "timestamp",
    "model_path",
    "model_name",
    "revit_version",
    "file_size_mb",
    "is_workshared",
    "warnings_count",
    # links / external
    "links_rvt_total",
    "links_rvt_not_found",
    "links_rvt_unloaded",
    "links_rvt_unpinned",
    "links_cad_link",
    "links_cad_link_not_found",
    "links_cad_import",
    "links_cad_unpinned",
    "linked_models_total",
    "ifc_links_total",
    "images_total",
    # worksets / coords / units
    "worksets_total",
    "worksets_owned_by_others",
    "workset1_elements",
    "length_accuracy",
    "pbp_offset_m",
    "max_geom_distance_m",
    "far_elements_count",
    "project_locations_count",
    "phases_count",
    # content health (wave A/B)
    "inplace_families",
    "groups_types",
    "groups_instances",
    "design_options",
    "views_total",
    "views_on_sheets",
    "views_not_on_sheets",
    "views_no_template",
    "sheets_total",
    "sheets_empty",
    "rooms_unplaced",
    "spaces_unplaced",
    "families_loadable",
    "family_types_total",
    "family_types_unused_est",
    "detail_lines",
    "filled_regions",
    "pinned_instances",
    "schedules_total",
    "materials_total",
    "line_patterns_total",
    "mep_open_connectors_est",
    "warnings_top",
    "status",
    "issues",
    "script_version",
]

CSV_HEADERS_RU = [
    u"Дата/время",
    u"Путь к модели",
    u"Имя файла",
    u"Версия Revit",
    u"Размер файла, МБ",
    u"Рабочая (workshared)",
    u"Количество предупреждений",
    u"Связи RVT, всего",
    u"Связи RVT, not found",
    u"Связи RVT, unloaded",
    u"Связи RVT, без pin",
    u"Связи CAD (link), шт",
    u"Связи CAD (link), not found",
    u"CAD Import, шт",
    u"CAD без pin, шт",
    u"Связанные модели RVT, шт",
    u"IFC-связи, шт",
    u"Изображения, шт",
    u"Рабочие наборы, всего",
    u"Рабочие наборы заняты другими",
    u"Элементы на Workset1, шт",
    u"Точность длины, мм",
    u"Смещение PBP, м",
    u"Макс. дистанция геометрии, м",
    u"Элементы далеко от origin, шт",
    u"Project Locations, шт",
    u"Фазы, шт",
    u"In-place семейства, шт",
    u"Группы (типы), шт",
    u"Группы (экземпляры), шт",
    u"Design Options, шт",
    u"Виды, всего",
    u"Виды на листах",
    u"Виды не на листах",
    u"Виды без шаблона",
    u"Листы, всего",
    u"Листы без видов",
    u"Rooms без размещения",
    u"Spaces без размещения",
    u"Загружаемые семейства, шт",
    u"Типы семейств, шт",
    u"Типы без экземпляров (оценка)",
    u"Detail Lines, шт",
    u"Filled Regions, шт",
    u"Закреплённые экземпляры, шт",
    u"Спецификации, шт",
    u"Материалы, шт",
    u"Типы линий, шт",
    u"Открытые коннекторы MEP (оценка)",
    u"Топ предупреждений",
    u"Статус",
    u"Проблемы",
    u"Версия скрипта",
]


def get_script_dir():
    try:
        return os.path.dirname(os.path.abspath(__file__))
    except Exception:
        pass

    candidates = [TOOL_DIR_FALLBACK, os.getcwd()]
    for folder in candidates:
        if folder and os.path.isfile(os.path.join(folder, "thresholds.cfg")):
            return folder
    return TOOL_DIR_FALLBACK


SCRIPT_DIR = get_script_dir()
CONFIG_PATH = os.path.join(SCRIPT_DIR, "thresholds.cfg")
REPORTS_DIR = os.path.join(SCRIPT_DIR, "reports")
CSV_DIR = os.path.join(REPORTS_DIR, "cvc")
ACTIVE_MARKER = os.path.join(CSV_DIR, "_active_report.txt")
ACTIVE_LINKS_MARKER = os.path.join(CSV_DIR, "_active_links.txt")
XLSX_OUT_CFG = os.path.join(SCRIPT_DIR, "xlsx_out_path.cfg")
RVT_LIST_PATH = os.path.join(SCRIPT_DIR, "rvt_list.txt")
LINKS_CSV_HEADERS = [
    u"Модель-хаб",
    u"Имя связи",
    u"Путь",
    u"Статус связи",
]


def now_stamp():
    return System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")


def now_file_stamp():
    return System.DateTime.Now.ToString("yyyy-MM-dd_HHmm")


def parse_float_list(text):
    result = []
    if not text:
        return result
    for part in text.split(","):
        part = part.strip().replace(",", ".")
        if not part:
            continue
        try:
            result.append(float(part))
        except Exception:
            pass
    return result


def load_thresholds():
    defaults = {
        "warnings_yellow": 50.0,
        "warnings_red": 500.0,
        "links_not_found_yellow": 1.0,
        "links_not_found_red": 3.0,
        "file_size_mb_yellow": 350.0,
        "file_size_mb_red": 500.0,
        "owned_worksets_yellow": 1.0,
        "length_accuracy_ok_max_mm": 1.0,
        "length_accuracy_yellow_max_mm": 10.0,
        "pbp_offset_yellow_m": 1000.0,
        "pbp_offset_red_m": 16000.0,
        # wave A/B soft lights
        "cad_import_yellow": 1.0,
        "cad_import_red": 5.0,
        "inplace_yellow": 5.0,
        "inplace_red": 20.0,
        "links_unpinned_yellow": 1.0,
        "far_geom_yellow_m": 1000.0,
        "far_geom_red_m": 16000.0,
        "far_elements_yellow": 1.0,
        "groups_instances_yellow": 50.0,
        "groups_instances_red": 200.0,
    }

    if not os.path.isfile(CONFIG_PATH):
        Output("No thresholds.cfg — using defaults")
        return defaults

    values = dict(defaults)
    with open(CONFIG_PATH, "r") as f:
        for raw in f:
            line = raw.strip()
            if not line or line.startswith("#"):
                continue
            if "=" not in line:
                continue
            key, val = line.split("=", 1)
            key = key.strip().lower()
            val = val.strip()
            try:
                values[key] = float(val.replace(",", "."))
            except Exception:
                pass
    return values


def ensure_reports_dir():
    if not os.path.isdir(REPORTS_DIR):
        os.makedirs(REPORTS_DIR)
    if not os.path.isdir(CSV_DIR):
        os.makedirs(CSV_DIR)


def get_report_path():
    ensure_reports_dir()
    if os.path.isfile(ACTIVE_MARKER):
        try:
            with open(ACTIVE_MARKER, "r") as f:
                name = f.read().strip()
            if name:
                return os.path.join(CSV_DIR, name)
        except Exception:
            pass

    name = "health_{}.csv".format(now_file_stamp())
    path = os.path.join(CSV_DIR, name)
    with open(ACTIVE_MARKER, "w") as f:
        f.write(name)

    links_name = name.replace(".csv", "_links.csv")
    with open(ACTIVE_LINKS_MARKER, "w") as f:
        f.write(links_name)
    return path


def get_links_report_path(summary_csv_path):
    ensure_reports_dir()
    if os.path.isfile(ACTIVE_LINKS_MARKER):
        try:
            with open(ACTIVE_LINKS_MARKER, "r") as f:
                name = f.read().strip()
            if name:
                return os.path.join(CSV_DIR, name)
        except Exception:
            pass
    base = os.path.basename(summary_csv_path)
    if base.lower().endswith(".csv"):
        name = base[:-4] + "_links.csv"
    else:
        name = base + "_links.csv"
    return os.path.join(CSV_DIR, name)


def read_xlsx_output_dir():
    if not os.path.isfile(XLSX_OUT_CFG):
        return ""
    try:
        with open(XLSX_OUT_CFG, "r") as f:
            for raw in f:
                line = raw.strip()
                if not line or line.startswith("#"):
                    continue
                return line
    except Exception:
        pass
    return ""


def normalize_model_path(path):
    if not path:
        return ""
    p = path.strip().strip('"').replace("/", "\\")
    while "\\\\" in p[2:]:
        p = p[:2] + p[2:].replace("\\\\", "\\")
    return p.lower()


def read_rvt_list_models(list_path=None):
    path = list_path or RVT_LIST_PATH
    if not os.path.isfile(path):
        return []
    models = []
    try:
        with open(path, "r") as f:
            for raw in f:
                line = raw.strip().strip('"')
                if not line or line.startswith("#"):
                    continue
                low = line.lower()
                if low.endswith(".rvt") or low.startswith("rsn://"):
                    models.append(line)
    except Exception as ex:
        Output("WARN read model list: {}".format(ex))
    return models


def count_csv_data_rows(report_path):
    if not report_path or not os.path.isfile(report_path):
        return 0
    n = 0
    try:
        with open(report_path, "r") as f:
            first = True
            for raw in f:
                if first:
                    first = False
                    continue
                if raw.strip():
                    n += 1
    except Exception:
        return 0
    return n


def is_last_model_in_batch(current_path, report_path):
    """
    True только для реальной последней модели текущего прогона RBP.
    Основной источник: revit_script_util.GetProgressNumber/GetProgressMax
    (это тот же счётчик, что 'Processing file (N of M)' в логе RBP).
    """
    # 1) Канон RBP
    try:
        get_num = getattr(revit_script_util, "GetProgressNumber", None)
        get_max = getattr(revit_script_util, "GetProgressMax", None)
        if callable(get_num) and callable(get_max):
            num = int(get_num())
            mx = int(get_max())
            Output("RBP progress: {}/{}".format(num, mx))
            if mx > 0:
                return num >= mx
    except Exception as ex:
        Output("WARN RBP progress API: {}".format(ex))

    # 2) Fallback: число строк CSV >= число моделей в rvt_list.txt
    models = read_rvt_list_models()
    rows = count_csv_data_rows(report_path)
    if models:
        Output(
            "Fallback progress: csv_rows={}/{} (rvt_list)".format(rows, len(models))
        )
        return rows >= len(models)

    Output("WARN cannot detect batch end — building XLSX now")
    return True


def export_xlsx_from_csv(csv_path, wait=True):
    try:
        script_path = os.path.join(SCRIPT_DIR, "csv_to_colored_xlsx.ps1")
        if not os.path.isfile(script_path):
            Output("WARN XLSX export skipped: csv_to_colored_xlsx.ps1 not found")
            return False

        out_dir = read_xlsx_output_dir()
        if not out_dir:
            Output(
                "WARN XLSX export skipped: run reset_report_session.cmd and choose local XLSX folder"
            )
            return False

        links_path = get_links_report_path(csv_path)
        log_path = os.path.join(out_dir, "xlsx_export.log")
        args = (
            u'-NoProfile -ExecutionPolicy Bypass -File "{0}" '
            u'-CsvPath "{1}" -LinksCsvPath "{2}" -OutputDir "{3}" -LogPath "{4}"'
        ).format(script_path, csv_path, links_path, out_dir, log_path)

        psi = System.Diagnostics.ProcessStartInfo()
        psi.FileName = "powershell"
        psi.Arguments = args
        psi.CreateNoWindow = True
        psi.UseShellExecute = False
        psi.RedirectStandardOutput = True
        psi.RedirectStandardError = True

        Output("XLSX export: building -> {}".format(out_dir))
        proc = System.Diagnostics.Process.Start(psi)
        if wait:
            # OpenXML build is fast; wait up to 90s (no Excel UI).
            if not proc.WaitForExit(90000):
                try:
                    proc.Kill()
                except Exception:
                    pass
                Output("WARN XLSX export timeout (>90s)")
                return False
            stdout = proc.StandardOutput.ReadToEnd()
            stderr = proc.StandardError.ReadToEnd()
            if proc.ExitCode == 0:
                if stdout:
                    Output(stdout.strip())
                Output("XLSX export: OK")
                return True
            Output(
                "WARN XLSX export failed (exit {}): {}".format(
                    proc.ExitCode, stderr or stdout
                )
            )
            return False
        else:
            Output("XLSX export: started (async)")
            return True
    except Exception as ex:
        Output("WARN XLSX export exception: {}".format(ex))
        return False


def csv_escape(value):
    if value is None:
        return u""
    try:
        text = value if isinstance(value, unicode) else unicode(value)
    except Exception:
        text = unicode(str(value))
    if u";" in text or u"," in text or u'"' in text or u"\n" in text or u"\r" in text:
        text = u'"' + text.replace(u'"', u'""') + u'"'
    return text


def append_links_rows(links_path, hub_name, link_rows):
    write_header = (not os.path.isfile(links_path)) or (os.path.getsize(links_path) == 0)
    lines = []
    if write_header:
        lines.append(u";".join(LINKS_CSV_HEADERS))
    for item in link_rows:
        row = [
            hub_name or "",
            item.get("name", ""),
            item.get("path", ""),
            item.get("status", ""),
        ]
        lines.append(u";".join([csv_escape(v) for v in row]))

    text = u"\r\n".join(lines) + u"\r\n"
    mode = "w" if write_header else "a"
    enc = "utf-8-sig" if write_header else "utf-8"
    with codecs.open(links_path, mode, enc) as f:
        f.write(text)


def append_csv_row(report_path, row_dict):
    write_header = (not os.path.isfile(report_path)) or (os.path.getsize(report_path) == 0)
    line_parts = [csv_escape(row_dict.get(h, "")) for h in CSV_FIELDS]
    line = u";".join(line_parts) + u"\r\n"

    if write_header:
        header = u";".join(CSV_HEADERS_RU) + u"\r\n"
        with codecs.open(report_path, "w", "utf-8-sig") as f:
            f.write(header)
            f.write(line)
    else:
        with codecs.open(report_path, "a", "utf-8") as f:
            f.write(line)


def safe_float(value, digits=6):
    try:
        return round(float(value), digits)
    except Exception:
        return ""


def get_file_size_mb(path):
    if not path:
        return ""
    if path.upper().startswith("RSN://"):
        return ""
    try:
        if System.IO.File.Exists(path):
            size = System.IO.FileInfo(path).Length
            return int(round(size / (1024.0 * 1024.0)))
    except Exception:
        pass
    return ""


def get_length_accuracy_mm(document):
    opts = document.GetUnits().GetFormatOptions(SpecTypeId.Length)
    acc = opts.Accuracy
    try:
        unit_id = opts.GetUnitTypeId()
        if unit_id == UnitTypeId.Meters:
            acc = acc * 1000.0
        elif unit_id == UnitTypeId.Centimeters:
            acc = acc * 10.0
        elif unit_id == UnitTypeId.Feet:
            acc = UnitUtils.ConvertFromInternalUnits(acc, UnitTypeId.Millimeters)
    except Exception:
        pass
    return safe_float(acc, 8)


def count_warnings(document):
    try:
        return len(list(document.GetWarnings()))
    except Exception:
        try:
            return document.GetWarnings().Count
        except Exception:
            return 0


def link_status_name(status):
    try:
        return str(status)
    except Exception:
        return ""


def collect_rvt_links(document):
    total = 0
    not_found = 0
    unloaded = 0
    unpinned = 0
    ifc_total = 0
    try:
        types = FilteredElementCollector(document).OfClass(RevitLinkType)
        for lt in types:
            total += 1
            try:
                path = get_link_type_path(lt).lower()
                if path.endswith(".ifc"):
                    ifc_total += 1
            except Exception:
                pass
            try:
                st = lt.GetLinkedFileStatus()
                name = link_status_name(st).lower()
                if "notfound" in name or "not found" in name:
                    not_found += 1
                elif "unload" in name:
                    unloaded += 1
            except Exception:
                pass
        for inst in FilteredElementCollector(document).OfClass(RevitLinkInstance):
            try:
                if not inst.Pinned:
                    unpinned += 1
            except Exception:
                pass
    except Exception as ex:
        Output("WARN rvt links: {}".format(ex))
    return total, not_found, unloaded, unpinned, ifc_total


def collect_cad_external(document):
    """CAD Link types vs ImportInstance (not linked)."""
    link_total = 0
    link_nf = 0
    import_total = 0
    unpinned = 0
    try:
        try:
            for ct in FilteredElementCollector(document).OfClass(CADLinkType):
                link_total += 1
                try:
                    st = ct.GetLinkedFileStatus()
                    name = link_status_name(st).lower()
                    if "notfound" in name or "not found" in name:
                        link_nf += 1
                except Exception:
                    pass
        except Exception:
            pass

        for imp in FilteredElementCollector(document).OfClass(ImportInstance):
            try:
                is_linked = False
                try:
                    is_linked = bool(imp.IsLinked)
                except Exception:
                    is_linked = False
                if is_linked:
                    # counted via CADLinkType when available; still check pin
                    pass
                else:
                    import_total += 1
                try:
                    if not imp.Pinned:
                        unpinned += 1
                except Exception:
                    pass
            except Exception:
                pass
    except Exception as ex:
        Output("WARN cad external: {}".format(ex))
    return link_total, link_nf, import_total, unpinned


def get_link_type_path(link_type):
    try:
        ext_ref = link_type.GetExternalFileReference()
        if ext_ref:
            model_path = ext_ref.GetPath()
            if model_path:
                return ModelPathUtils.ConvertModelPathToUserVisiblePath(model_path)
    except Exception:
        pass
    return ""


def basename_from_path(path, fallback=""):
    if path:
        try:
            name = os.path.basename(path.replace("/", "\\"))
            if name:
                return name
        except Exception:
            pass
    return fallback or ""


def collect_rvt_links_detail(document):
    """Прямые RVT-связи хаба (1 уровень): name, path, status."""
    rows = []
    try:
        for lt in FilteredElementCollector(document).OfClass(RevitLinkType):
            try:
                path = get_link_type_path(lt)
                name = ""
                try:
                    name = lt.Name
                except Exception:
                    name = ""
                nice = basename_from_path(path, "")
                if nice:
                    name = nice
                elif name and not name.lower().endswith(".rvt"):
                    name = name + ".rvt"
                elif not name:
                    name = "(rvt link)"

                status = ""
                try:
                    status = link_status_name(lt.GetLinkedFileStatus())
                except Exception:
                    status = ""

                rows.append(
                    {
                        "name": name,
                        "path": path if path else "(path n/a)",
                        "status": status,
                    }
                )
            except Exception:
                pass
    except Exception as ex:
        Output("WARN rvt links detail: {}".format(ex))
    return rows


def collect_warnings_top(document, top_n=3):
    counts = {}
    try:
        warns = list(document.GetWarnings())
    except Exception:
        warns = []

    for w in warns:
        key = ""
        try:
            key = w.GetDescriptionText()
        except Exception:
            key = ""
        if not key:
            key = "(no description)"
        counts[key] = counts.get(key, 0) + 1

    if not counts:
        return ""

    ranked = sorted(counts.items(), key=lambda kv: kv[1], reverse=True)
    top = ranked[:top_n]
    parts = []
    for text, num in top:
        short = text.replace("\r", " ").replace("\n", " ").strip()
        if len(short) > 72:
            short = short[:69] + "..."
        parts.append(u"{0} x{1}".format(short, num))
    return " | ".join(parts)


def collect_worksets(document):
    total = 0
    owned_by_others = 0
    if not document.IsWorkshared:
        return 0, 0
    try:
        collector = FilteredWorksetCollector(document).OfKind(WorksetKind.UserWorkset)
        for ws in collector:
            total += 1
            try:
                info = WorksharingUtils.GetWorksetOwnershipStatus(document, ws.Id)
                name = str(info)
                if "OwnedByOtherUser" in name or "OtherUser" in name:
                    owned_by_others += 1
            except Exception:
                pass
    except Exception as ex:
        Output("WARN worksets: {}".format(ex))
    return total, owned_by_others


def collect_pbp_offset_m(document):
    try:
        points = FilteredElementCollector(document).OfClass(BasePoint)
        for bp in points:
            try:
                if bp.IsShared:
                    continue
                loc = bp.Position
                dist_ft = (loc.X ** 2 + loc.Y ** 2 + loc.Z ** 2) ** 0.5
                meters = UnitUtils.ConvertFromInternalUnits(dist_ft, UnitTypeId.Meters)
                return safe_float(meters, 4)
            except Exception:
                try:
                    ex = bp.get_Parameter(BuiltInParameter.BASEPOINT_EASTWEST_PARAM)
                    ny = bp.get_Parameter(BuiltInParameter.BASEPOINT_NORTHSOUTH_PARAM)
                    elev = bp.get_Parameter(BuiltInParameter.BASEPOINT_ELEVATION_PARAM)
                    if ex and ny and elev:
                        x = ex.AsDouble()
                        y = ny.AsDouble()
                        z = elev.AsDouble()
                        dist_ft = (x ** 2 + y ** 2 + z ** 2) ** 0.5
                        meters = UnitUtils.ConvertFromInternalUnits(
                            dist_ft, UnitTypeId.Meters
                        )
                        return safe_float(meters, 4)
                except Exception:
                    pass
    except Exception as ex:
        Output("WARN PBP: {}".format(ex))
    return ""


def _ft_to_m(dist_ft):
    try:
        return UnitUtils.ConvertFromInternalUnits(dist_ft, UnitTypeId.Meters)
    except Exception:
        return dist_ft * 0.3048


def _elem_distance_ft(elem):
    try:
        loc = elem.Location
        if loc is not None:
            try:
                pt = loc.Point
                return (pt.X ** 2 + pt.Y ** 2 + pt.Z ** 2) ** 0.5
            except Exception:
                pass
            try:
                curve = loc.Curve
                p0 = curve.GetEndPoint(0)
                p1 = curve.GetEndPoint(1)
                d0 = (p0.X ** 2 + p0.Y ** 2 + p0.Z ** 2) ** 0.5
                d1 = (p1.X ** 2 + p1.Y ** 2 + p1.Z ** 2) ** 0.5
                return d0 if d0 > d1 else d1
            except Exception:
                pass
    except Exception:
        pass
    try:
        bb = elem.get_BoundingBox(None)
        if bb is not None:
            cx = 0.5 * (bb.Min.X + bb.Max.X)
            cy = 0.5 * (bb.Min.Y + bb.Max.Y)
            cz = 0.5 * (bb.Min.Z + bb.Max.Z)
            return (cx ** 2 + cy ** 2 + cz ** 2) ** 0.5
    except Exception:
        pass
    return None


def collect_geometry_extent(document, far_yellow_m):
    """Max distance of model elements from internal origin + count beyond yellow."""
    max_m = 0.0
    far_count = 0
    try:
        far_ft = UnitUtils.ConvertToInternalUnits(float(far_yellow_m), UnitTypeId.Meters)
    except Exception:
        far_ft = float(far_yellow_m) / 0.3048

    try:
        collectors = [
            FilteredElementCollector(document).OfClass(FamilyInstance),
            FilteredElementCollector(document).OfClass(HostObject),
        ]
        try:
            collectors.append(FilteredElementCollector(document).OfClass(MEPCurve))
        except Exception:
            pass

        seen = set()
        for col in collectors:
            for elem in col:
                try:
                    eid = elem.Id.IntegerValue
                except Exception:
                    continue
                if eid in seen:
                    continue
                seen.add(eid)
                dist_ft = _elem_distance_ft(elem)
                if dist_ft is None:
                    continue
                dist_m = _ft_to_m(dist_ft)
                if dist_m > max_m:
                    max_m = dist_m
                if dist_ft >= far_ft:
                    far_count += 1
    except Exception as ex:
        Output("WARN geom extent: {}".format(ex))
    return safe_float(max_m, 2), far_count


def collect_views_sheets(document):
    views_total = 0
    on_sheets = 0
    no_template = 0
    sheets_total = 0
    sheets_empty = 0
    schedules = 0
    sheet_view_ids = set()
    try:
        for vp in FilteredElementCollector(document).OfClass(Viewport):
            try:
                sheet_view_ids.add(vp.ViewId.IntegerValue)
            except Exception:
                pass
    except Exception:
        pass

    try:
        for v in FilteredElementCollector(document).OfClass(View):
            try:
                if v.IsTemplate:
                    continue
            except Exception:
                pass
            # skip sheets themselves in view count
            try:
                if isinstance(v, ViewSheet):
                    continue
            except Exception:
                pass
            views_total += 1
            try:
                if v.Id.IntegerValue in sheet_view_ids:
                    on_sheets += 1
            except Exception:
                pass
            try:
                tid = v.ViewTemplateId
                if tid is None or tid.IntegerValue < 0:
                    no_template += 1
            except Exception:
                pass
            try:
                if isinstance(v, ViewSchedule):
                    schedules += 1
            except Exception:
                pass
    except Exception as ex:
        Output("WARN views: {}".format(ex))

    try:
        for sh in FilteredElementCollector(document).OfClass(ViewSheet):
            sheets_total += 1
            try:
                vps = list(FilteredElementCollector(document, sh.Id).OfClass(Viewport))
                if len(vps) == 0:
                    sheets_empty += 1
            except Exception:
                try:
                    if sh.GetAllViewports().Count == 0:
                        sheets_empty += 1
                except Exception:
                    pass
    except Exception as ex:
        Output("WARN sheets: {}".format(ex))

    not_on = views_total - on_sheets
    if not_on < 0:
        not_on = 0
    return views_total, on_sheets, not_on, no_template, sheets_total, sheets_empty, schedules


def collect_spatial_unplaced(document):
    rooms_u = 0
    spaces_u = 0
    try:
        for r in FilteredElementCollector(document).OfCategory(
            BuiltInCategory.OST_Rooms
        ).WhereElementIsNotElementType():
            try:
                if r.Location is None:
                    rooms_u += 1
                    continue
            except Exception:
                pass
            try:
                if r.Area < 1e-9:
                    rooms_u += 1
            except Exception:
                pass
    except Exception as ex:
        Output("WARN rooms: {}".format(ex))
    try:
        for s in FilteredElementCollector(document).OfCategory(
            BuiltInCategory.OST_MEPSpaces
        ).WhereElementIsNotElementType():
            try:
                if s.Location is None:
                    spaces_u += 1
                    continue
            except Exception:
                pass
            try:
                if s.Area < 1e-9:
                    spaces_u += 1
            except Exception:
                pass
    except Exception as ex:
        Output("WARN spaces: {}".format(ex))
    return rooms_u, spaces_u


def collect_families_groups_options(document):
    inplace = 0
    loadable = 0
    types_total = 0
    unused_types = 0
    inst_by_type = {}
    try:
        for fi in FilteredElementCollector(document).OfClass(FamilyInstance):
            try:
                tid = fi.GetTypeId().IntegerValue
                inst_by_type[tid] = inst_by_type.get(tid, 0) + 1
            except Exception:
                pass
    except Exception:
        pass

    try:
        for fam in FilteredElementCollector(document).OfClass(Family):
            try:
                if fam.IsInPlace:
                    inplace += 1
                    continue
            except Exception:
                pass
            loadable += 1
            try:
                for sid in fam.GetFamilySymbolIds():
                    types_total += 1
                    try:
                        iv = sid.IntegerValue
                        if inst_by_type.get(iv, 0) == 0:
                            unused_types += 1
                    except Exception:
                        pass
            except Exception:
                pass
    except Exception as ex:
        Output("WARN families: {}".format(ex))

    g_types = 0
    g_inst = 0
    try:
        g_types = FilteredElementCollector(document).OfClass(GroupType).GetElementCount()
    except Exception:
        try:
            g_types = len(list(FilteredElementCollector(document).OfClass(GroupType)))
        except Exception:
            g_types = 0
    try:
        g_inst = FilteredElementCollector(document).OfClass(Group).GetElementCount()
    except Exception:
        try:
            g_inst = len(list(FilteredElementCollector(document).OfClass(Group)))
        except Exception:
            g_inst = 0

    dopts = 0
    try:
        dopts = FilteredElementCollector(document).OfClass(DesignOption).GetElementCount()
    except Exception:
        try:
            dopts = len(list(FilteredElementCollector(document).OfClass(DesignOption)))
        except Exception:
            dopts = 0

    return inplace, loadable, types_total, unused_types, g_types, g_inst, dopts


def collect_misc_counts(document):
    images = 0
    detail_lines = 0
    filled = 0
    pinned_fi = 0
    materials = 0
    line_patterns = 0
    phases = 0
    locations = 0
    workset1 = 0
    open_conn = 0

    try:
        images = FilteredElementCollector(document).OfClass(ImageType).GetElementCount()
    except Exception:
        try:
            images = len(list(FilteredElementCollector(document).OfClass(ImageType)))
        except Exception:
            images = 0

    try:
        detail_lines = (
            FilteredElementCollector(document).OfClass(DetailLine).GetElementCount()
        )
    except Exception:
        try:
            detail_lines = len(list(FilteredElementCollector(document).OfClass(DetailLine)))
        except Exception:
            detail_lines = 0

    try:
        filled = FilteredElementCollector(document).OfClass(FilledRegion).GetElementCount()
    except Exception:
        try:
            filled = len(list(FilteredElementCollector(document).OfClass(FilledRegion)))
        except Exception:
            filled = 0

    try:
        for fi in FilteredElementCollector(document).OfClass(FamilyInstance):
            try:
                if fi.Pinned:
                    pinned_fi += 1
            except Exception:
                pass
    except Exception:
        pass

    try:
        materials = FilteredElementCollector(document).OfClass(Material).GetElementCount()
    except Exception:
        try:
            materials = len(list(FilteredElementCollector(document).OfClass(Material)))
        except Exception:
            materials = 0

    try:
        line_patterns = (
            FilteredElementCollector(document)
            .OfClass(LinePatternElement)
            .GetElementCount()
        )
    except Exception:
        try:
            line_patterns = len(
                list(FilteredElementCollector(document).OfClass(LinePatternElement))
            )
        except Exception:
            line_patterns = 0

    try:
        phases = document.Phases.Size
    except Exception:
        try:
            phases = len(list(document.Phases))
        except Exception:
            phases = 0

    try:
        locations = document.ProjectLocations.Size
    except Exception:
        try:
            locations = len(list(document.ProjectLocations))
        except Exception:
            locations = 0

    # Workset1 / default workset element count (informational)
    if document.IsWorkshared:
        try:
            ws1_ids = set()
            for ws in FilteredWorksetCollector(document).OfKind(WorksetKind.UserWorkset):
                try:
                    n = ws.Name
                except Exception:
                    n = ""
                nl = (n or "").lower()
                if (
                    nl == "workset1"
                    or "рабочий набор 1" in nl
                    or nl.startswith("workset 1")
                ):
                    ws1_ids.add(ws.Id.IntegerValue)
            if ws1_ids:
                for col in (
                    FilteredElementCollector(document).OfClass(FamilyInstance),
                    FilteredElementCollector(document).OfClass(HostObject),
                ):
                    for e in col:
                        try:
                            wid = e.WorksetId.IntegerValue
                            if wid in ws1_ids:
                                workset1 += 1
                        except Exception:
                            pass
        except Exception as ex:
            Output("WARN workset1: {}".format(ex))

    # MEP open connectors (estimate, may be slow — capped)
    try:
        checked = 0
        max_check = 8000
        for curve in FilteredElementCollector(document).OfClass(MEPCurve):
            checked += 1
            if checked > max_check:
                break
            try:
                cm = curve.ConnectorManager
                if cm is None:
                    continue
                for c in cm.Connectors:
                    try:
                        if not c.IsConnected:
                            open_conn += 1
                    except Exception:
                        pass
            except Exception:
                pass
        for fi in FilteredElementCollector(document).OfClass(FamilyInstance):
            checked += 1
            if checked > max_check * 2:
                break
            try:
                mep = fi.MEPModel
                if mep is None:
                    continue
                cm = mep.ConnectorManager
                if cm is None:
                    continue
                for c in cm.Connectors:
                    try:
                        if not c.IsConnected:
                            open_conn += 1
                    except Exception:
                        pass
            except Exception:
                pass
    except Exception as ex:
        Output("WARN mep connectors: {}".format(ex))

    return (
        images,
        detail_lines,
        filled,
        pinned_fi,
        materials,
        line_patterns,
        phases,
        locations,
        workset1,
        open_conn,
    )


def accuracy_status(accuracy, thr):
    """
    Length rounding (mm):
      OK (None) if <= ok_max
      YELLOW if ok_max < value <= yellow_max
      RED if value > yellow_max
    """
    if accuracy == "" or accuracy is None:
        return None
    try:
        a = float(accuracy)
    except Exception:
        return "YELLOW"
    ok_max = float(thr.get("length_accuracy_ok_max_mm", 1.0))
    yel_max = float(thr.get("length_accuracy_yellow_max_mm", 10.0))
    if a <= ok_max + 1e-12:
        return None
    if a <= yel_max + 1e-12:
        return "YELLOW"
    return "RED"


def _bump_metric(status, issues, thr, key_y, key_r, value, label):
    order = {"GREEN": 0, "YELLOW": 1, "RED": 2, "ERROR": 3}
    try:
        v = float(value)
    except Exception:
        return status
    ry = thr.get(key_y) if key_y else None
    rr = thr.get(key_r) if key_r else None
    new_st = None
    if rr is not None and v >= float(rr):
        new_st = "RED"
    elif ry is not None and v >= float(ry):
        new_st = "YELLOW"
    if new_st:
        if order.get(new_st, 0) > order.get(status, 0):
            status = new_st
        issues.append("{}={}".format(label, value))
    return status


def compute_status(metrics, thr):
    issues = []
    status = "GREEN"

    def bump(new_status):
        order = {"GREEN": 0, "YELLOW": 1, "RED": 2, "ERROR": 3}
        if order.get(new_status, 0) > order.get(status, 0):
            return new_status
        return status

    warnings = metrics.get("warnings_count", 0) or 0
    links_nf = metrics.get("links_rvt_not_found", 0) or 0
    size_mb = metrics.get("file_size_mb", "")
    owned = metrics.get("worksets_owned_by_others", 0) or 0
    acc = metrics.get("length_accuracy", "")
    pbp = metrics.get("pbp_offset_m", "")

    if links_nf >= thr["links_not_found_red"]:
        status = bump("RED")
        issues.append("links_not_found={}".format(links_nf))
    elif links_nf >= thr["links_not_found_yellow"]:
        status = bump("YELLOW")
        issues.append("links_not_found={}".format(links_nf))

    if warnings >= thr["warnings_red"]:
        status = bump("RED")
        issues.append("warnings={}".format(warnings))
    elif warnings >= thr["warnings_yellow"]:
        status = bump("YELLOW")
        issues.append("warnings={}".format(warnings))

    if size_mb != "" and size_mb is not None:
        try:
            sm = float(size_mb)
            if sm >= thr["file_size_mb_red"]:
                status = bump("RED")
                issues.append("file_size_mb={}".format(sm))
            elif sm >= thr["file_size_mb_yellow"]:
                status = bump("YELLOW")
                issues.append("file_size_mb={}".format(sm))
        except Exception:
            pass

    if owned >= thr["owned_worksets_yellow"]:
        status = bump("YELLOW")
        issues.append("owned_worksets={}".format(owned))

    acc_st = accuracy_status(acc, thr)
    if acc_st:
        status = bump(acc_st)
        issues.append("length_accuracy={}".format(acc))

    if pbp != "" and pbp is not None:
        try:
            pm = float(pbp)
            if pm >= thr["pbp_offset_red_m"]:
                status = bump("RED")
                issues.append("pbp_offset_m={}".format(pm))
            elif pm >= thr["pbp_offset_yellow_m"]:
                status = bump("YELLOW")
                issues.append("pbp_offset_m={}".format(pm))
        except Exception:
            pass

    status = _bump_metric(
        status,
        issues,
        thr,
        "cad_import_yellow",
        "cad_import_red",
        metrics.get("links_cad_import", 0),
        "cad_import",
    )
    status = _bump_metric(
        status,
        issues,
        thr,
        "inplace_yellow",
        "inplace_red",
        metrics.get("inplace_families", 0),
        "inplace",
    )
    unpinned = (metrics.get("links_rvt_unpinned", 0) or 0) + (
        metrics.get("links_cad_unpinned", 0) or 0
    )
    status = _bump_metric(
        status, issues, thr, "links_unpinned_yellow", None, unpinned, "links_unpinned"
    )
    status = _bump_metric(
        status,
        issues,
        thr,
        "groups_instances_yellow",
        "groups_instances_red",
        metrics.get("groups_instances", 0),
        "groups_instances",
    )

    max_d = metrics.get("max_geom_distance_m", "")
    if max_d != "" and max_d is not None:
        status = _bump_metric(
            status,
            issues,
            thr,
            "far_geom_yellow_m",
            "far_geom_red_m",
            max_d,
            "max_geom_distance_m",
        )
    status = _bump_metric(
        status,
        issues,
        thr,
        "far_elements_yellow",
        None,
        metrics.get("far_elements_count", 0),
        "far_elements",
    )

    return status, "; ".join(issues)


def get_revit_version(document):
    try:
        return document.Application.VersionNumber
    except Exception:
        try:
            return str(document.Application.VersionName)
        except Exception:
            return ""


def analyze_document(document, path, thr):
    warnings_count = count_warnings(document)
    rvt_total, rvt_nf, rvt_ul, rvt_unpin, ifc_total = collect_rvt_links(document)
    cad_link, cad_nf, cad_imp, cad_unpin = collect_cad_external(document)
    link_rows = collect_rvt_links_detail(document)
    linked_total = len(link_rows)
    ws_total, ws_owned = collect_worksets(document)
    length_acc = get_length_accuracy_mm(document)
    pbp = collect_pbp_offset_m(document)
    size_mb = get_file_size_mb(path)
    warnings_top = collect_warnings_top(document, 3)

    far_y = thr.get("far_geom_yellow_m", 1000.0)
    max_geom, far_n = collect_geometry_extent(document, far_y)

    (
        views_total,
        views_on,
        views_off,
        views_no_tpl,
        sheets_total,
        sheets_empty,
        schedules,
    ) = collect_views_sheets(document)
    rooms_u, spaces_u = collect_spatial_unplaced(document)
    (
        inplace,
        loadable,
        types_total,
        unused_types,
        g_types,
        g_inst,
        dopts,
    ) = collect_families_groups_options(document)
    (
        images,
        detail_lines,
        filled,
        pinned_fi,
        materials,
        line_patterns,
        phases,
        locations,
        workset1,
        open_conn,
    ) = collect_misc_counts(document)

    metrics = {
        "timestamp": now_stamp(),
        "model_path": path or "",
        "model_name": os.path.basename(path) if path else "",
        "revit_version": get_revit_version(document),
        "file_size_mb": size_mb,
        "is_workshared": "yes" if document.IsWorkshared else "no",
        "warnings_count": warnings_count,
        "links_rvt_total": rvt_total,
        "links_rvt_not_found": rvt_nf,
        "links_rvt_unloaded": rvt_ul,
        "links_rvt_unpinned": rvt_unpin,
        "links_cad_link": cad_link,
        "links_cad_link_not_found": cad_nf,
        "links_cad_import": cad_imp,
        "links_cad_unpinned": cad_unpin,
        "linked_models_total": linked_total,
        "ifc_links_total": ifc_total,
        "images_total": images,
        "worksets_total": ws_total,
        "worksets_owned_by_others": ws_owned,
        "workset1_elements": workset1,
        "length_accuracy": length_acc,
        "pbp_offset_m": pbp,
        "max_geom_distance_m": max_geom,
        "far_elements_count": far_n,
        "project_locations_count": locations,
        "phases_count": phases,
        "inplace_families": inplace,
        "groups_types": g_types,
        "groups_instances": g_inst,
        "design_options": dopts,
        "views_total": views_total,
        "views_on_sheets": views_on,
        "views_not_on_sheets": views_off,
        "views_no_template": views_no_tpl,
        "sheets_total": sheets_total,
        "sheets_empty": sheets_empty,
        "rooms_unplaced": rooms_u,
        "spaces_unplaced": spaces_u,
        "families_loadable": loadable,
        "family_types_total": types_total,
        "family_types_unused_est": unused_types,
        "detail_lines": detail_lines,
        "filled_regions": filled,
        "pinned_instances": pinned_fi,
        "schedules_total": schedules,
        "materials_total": materials,
        "line_patterns_total": line_patterns,
        "mep_open_connectors_est": open_conn,
        "warnings_top": warnings_top,
        "script_version": SCRIPT_VERSION,
        "_link_rows": link_rows,
    }

    status, issues = compute_status(metrics, thr)
    metrics["status"] = status
    metrics["issues"] = issues
    return metrics


def write_error_row(path, message, report_path):
    row = dict((h, "") for h in CSV_FIELDS)
    row["timestamp"] = now_stamp()
    row["model_path"] = path or ""
    row["model_name"] = os.path.basename(path) if path else ""
    row["status"] = "ERROR"
    row["issues"] = message
    row["script_version"] = SCRIPT_VERSION
    append_csv_row(report_path, row)


# --- main ---
THRESHOLDS = load_thresholds()
doc = None
revitFilePath = ""
report_path = get_report_path()

Output()
Output("=== Health Check v{} ===".format(SCRIPT_VERSION))
Output("Script dir: {}".format(SCRIPT_DIR))
Output("Config: {}".format(CONFIG_PATH))
Output("Report: {}".format(report_path))

try:
    doc = revit_script_util.GetScriptDocument()
    revitFilePath = revit_script_util.GetRevitFilePath()
    Output("File: {}".format(revitFilePath))
    Output(
        "Workshared={}, Detached={}".format(
            doc.IsWorkshared, getattr(doc, "IsDetached", False)
        )
    )

    row = analyze_document(doc, revitFilePath, THRESHOLDS)
    link_rows = row.pop("_link_rows", [])
    links_path = get_links_report_path(report_path)
    if link_rows:
        append_links_rows(links_path, row.get("model_name", ""), link_rows)
    append_csv_row(report_path, row)

    Output("warnings={}".format(row["warnings_count"]))
    Output(
        "links_rvt: total={}, not_found={}, unloaded={}, unpinned={}".format(
            row["links_rvt_total"],
            row["links_rvt_not_found"],
            row["links_rvt_unloaded"],
            row["links_rvt_unpinned"],
        )
    )
    Output(
        "cad: link={}, link_nf={}, import={}, unpinned={}".format(
            row["links_cad_link"],
            row["links_cad_link_not_found"],
            row["links_cad_import"],
            row["links_cad_unpinned"],
        )
    )
    Output(
        "content: inplace={}, groups_inst={}, views={}, sheets={}".format(
            row["inplace_families"],
            row["groups_instances"],
            row["views_total"],
            row["sheets_total"],
        )
    )
    Output(
        "extent: pbp_m={}, max_geom_m={}, far_elems={}".format(
            row["pbp_offset_m"],
            row["max_geom_distance_m"],
            row["far_elements_count"],
        )
    )
    Output("linked_models_rvt_total={}".format(row["linked_models_total"]))
    Output("links_detail={}".format(links_path))
    if link_rows:
        preview = link_rows[:5]
        for item in preview:
            Output(
                "  rvt: {} | {}".format(item.get("name", ""), item.get("path", ""))
            )
        if len(link_rows) > 5:
            Output("  ... (+{})".format(len(link_rows) - 5))
    if row["warnings_top"]:
        Output("warnings_top={}".format(row["warnings_top"]))
    Output(
        "worksets: total={}, owned_by_others={}, workset1_elems={}".format(
            row["worksets_total"],
            row["worksets_owned_by_others"],
            row["workset1_elements"],
        )
    )
    Output("length_accuracy={}".format(row["length_accuracy"]))
    Output("status={}".format(row["status"]))
    if row["issues"]:
        Output("issues: {}".format(row["issues"]))

    if is_last_model_in_batch(revitFilePath, report_path):
        Output("Last model in RBP batch — building XLSX once")
        export_xlsx_from_csv(report_path, wait=True)
    else:
        Output("XLSX deferred until last model in RBP batch")

    Output("NO SAVE — report only")
    Output("OK")
except Exception as ex:
    msg = unicode(ex)
    Output("ERROR: {}".format(msg))
    try:
        write_error_row(revitFilePath, msg, report_path)
        # If this was the last model, still try to build report
        if is_last_model_in_batch(revitFilePath, report_path):
            export_xlsx_from_csv(report_path, wait=True)
    except Exception as ex2:
        Output("WARN could not write error row: {}".format(ex2))
    raise
