# -*- coding: utf-8 -*-
"""
Revit Batch Processor — пересохранение со сжатием (Compact).

Галочка «Сжать файл» из «Параметры сохранения файла»:
  - обычный .rvt           → Save / SaveAs, Compact = True
  - workshared / RSN://    → SynchronizeWithCentral, Compact = True
    (сжимается хранилище, не только локаль)

Преднастроек нет: Compact всегда включён.

Важно для worksharing / RSN:
  - в RBP: Create New Local, Detach = OFF
  - после сжатия: Relinquish
"""

import clr
import System
import os
import datetime
import re
import codecs
import time
import zipfile

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
from Autodesk.Revit.DB import *

import revit_script_util
from revit_script_util import Output

TOOL_DIR_FALLBACK = r""  # optional deployed copy
TOOL_DIR_FALLBACK_OLD = r""

SYNC_COMMENT = "Batch: compact save"

try:
    _unichr = unichr
except NameError:
    _unichr = chr

try:
    _unicode = unicode
except NameError:
    _unicode = str


def u_from_codes(codes):
    return u"".join(_unichr(c) for c in codes)


REZERV_NAME = u_from_codes([0x0420, 0x0435, 0x0437, 0x0435, 0x0440, 0x0432])


def is_revit_install_dir(folder):
    low = (folder or "").replace("/", "\\").lower()
    return ("\\program files\\autodesk\\revit" in low) or (
        "\\program files (x86)\\autodesk\\revit" in low
    )


def is_writable_dir(folder):
    if not folder or not os.path.isdir(folder):
        return False
    probe = os.path.join(folder, ".compact_write_probe")
    try:
        f = open(probe, "w")
        f.write("ok")
        f.close()
        try:
            os.remove(probe)
        except Exception:
            pass
        return True
    except Exception:
        return False


def is_guid_folder(folder):
    name = os.path.basename((folder or "").rstrip("\\/"))
    return bool(
        re.match(
            r"^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$",
            name,
        )
    )


def ensure_dir(folder):
    if not folder:
        return False
    try:
        if not os.path.isdir(folder):
            os.makedirs(folder)
        return os.path.isdir(folder)
    except Exception:
        return False


def local_appdata():
    p = os.environ.get("LOCALAPPDATA")
    if p and os.path.isdir(p):
        return p
    home = os.environ.get("USERPROFILE") or ""
    if home:
        return os.path.join(home, r"AppData\Local")
    return None


def get_script_dir():
    candidates = []
    try:
        here = os.path.dirname(os.path.abspath(__file__))
        if here and (not is_revit_install_dir(here)) and (not is_guid_folder(here)):
            candidates.append(here)
    except Exception:
        pass

    candidates.append(TOOL_DIR_FALLBACK)
    candidates.append(TOOL_DIR_FALLBACK_OLD)

    home = os.environ.get("USERPROFILE") or os.environ.get("HOME") or ""
    if home:
        candidates.append(
            os.path.join(home, r"Documents\doc\script\rbp\batch_compact_save")
        )

    for folder in candidates:
        if not folder or is_revit_install_dir(folder) or is_guid_folder(folder):
            continue
        py_here = os.path.join(folder, "compact_save.py")
        if os.path.isfile(py_here):
            return folder

    return TOOL_DIR_FALLBACK


OTCHETY_NAME = u_from_codes([0x043E, 0x0442, 0x0447, 0x0435, 0x0442, 0x044B])  # отчеты


def reports_dir(script_dir):
    folder = os.path.join(script_dir, OTCHETY_NAME)
    if not ensure_dir(folder):
        raise Exception(
            "Cannot create reports folder: {}. Need write access next to compact_save.py.".format(
                folder
            )
        )
    if not is_writable_dir(folder):
        raise Exception(
            "Reports folder is not writable: {}. Grant write access on the share.".format(
                folder
            )
        )
    return folder


def report_user_tag():
    raw = as_text(os.environ.get("USERNAME") or os.environ.get("USER") or u"").strip()
    out = []
    for ch in raw:
        o = ord(ch)
        if (
            (48 <= o <= 57)
            or (65 <= o <= 90)
            or (97 <= o <= 122)
            or o > 127
            or ch in (u"-", u"_", u".")
        ):
            out.append(ch)
    tag = u"".join(out).strip(u".")
    if len(tag) > 40:
        tag = tag[:40]
    return tag


def run_id_pointer_path():
    base = local_appdata()
    if not base:
        return None
    return os.path.join(base, "BatchRvt", "batch_compact_save", u"active_run_id.txt")


def get_session_id():
    env_id = os.environ.get("BATCH_COMPACT_RUN_ID")
    if env_id and as_text(env_id).strip():
        return as_text(env_id).strip()
    try:
        p = run_id_pointer_path()
        if p and os.path.isfile(p):
            f = codecs.open(p, "r", "utf-8")
            try:
                text = f.read().strip()
            finally:
                f.close()
            if text:
                return as_text(text)
    except Exception:
        pass
    try:
        return as_text(revit_script_util.GetSessionId()).strip()
    except Exception:
        return u""


def new_report_filename():
    return datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S") + ".csv"


def allocate_report_file(folder, session_id):
    map_path = os.path.join(folder, "_sessions.txt")
    lock_path = os.path.join(folder, "_sessions.lock")
    fs = None
    for _try in range(40):
        try:
            fs = System.IO.File.Open(
                lock_path,
                System.IO.FileMode.OpenOrCreate,
                System.IO.FileAccess.ReadWrite,
                getattr(System.IO.FileShare, "None"),
            )
            break
        except Exception:
            time.sleep(0.1)
            fs = None
    try:
        mapping = {}
        if os.path.isfile(map_path):
            f = codecs.open(map_path, "r", "utf-8")
            try:
                for line in f:
                    line = line.strip()
                    if not line or u"=" not in line:
                        continue
                    sid, name = line.split(u"=", 1)
                    mapping[sid.strip()] = name.strip()
            finally:
                f.close()

        if session_id and session_id in mapping:
            return os.path.join(folder, mapping[session_id])

        name = new_report_filename()
        path = os.path.join(folder, name)
        if os.path.isfile(path):
            n = 2
            stem = name[:-4]
            while os.path.isfile(os.path.join(folder, "{0}_{1}.csv".format(stem, n))):
                n += 1
            name = "{0}_{1}.csv".format(stem, n)
            path = os.path.join(folder, name)

        if not session_id:
            try:
                newest = None
                newest_mtime = 0
                for fn in os.listdir(folder):
                    if not fn.lower().endswith(".csv"):
                        continue
                    fp = os.path.join(folder, fn)
                    mt = os.path.getmtime(fp)
                    if mt > newest_mtime:
                        newest_mtime = mt
                        newest = fp
                if newest and (time.time() - newest_mtime) < 6 * 3600:
                    return newest
            except Exception:
                pass

        if session_id:
            mapping[session_id] = name
            f = codecs.open(map_path, "w", "utf-8")
            try:
                for sid in mapping:
                    f.write(u"{0}={1}\n".format(sid, mapping[sid]))
            finally:
                f.close()
        return path
    finally:
        if fs is not None:
            fs.Close()
        try:
            os.remove(lock_path)
        except Exception:
            pass


def is_backup_copy_path(path):
    if not path:
        return False
    p = path.replace("/", "\\")
    low = p.lower()
    if "\\backup\\" in low:
        return True
    if "\\revit_temp\\" in low:
        return True
    if "_backup\\" in low or low.endswith("_backup"):
        return True
    rezerv_token = u"\\" + REZERV_NAME.lower() + u"\\"
    if rezerv_token in low:
        return True
    if re.search(r"\\RVT\\[0-9]{4}-[0-9]{2}-[0-9]{2}\\", p, re.I):
        return True
    return False


SCRIPT_DIR = get_script_dir()
REPORTS_DIR = None
REPORT_PATH = None

doc = revit_script_util.GetScriptDocument()
revitFilePath = revit_script_util.GetRevitFilePath()


def is_rsn_path(path):
    if not path:
        return False
    return path.upper().startswith("RSN://")


def file_size_bytes(path):
    if not path or is_rsn_path(path):
        return None
    try:
        if os.path.isfile(path):
            return os.path.getsize(path)
    except Exception:
        return None
    return None


# host -> REST service year (same as servers.cfg / start_compact.ps1)
RSN_HOST_YEARS = {
    "revit-server-2021.example.local": "2021",
    "revit-server-2022.example.local": "2022",
    "revit-server-2023.example.local": "2023",
    "revit-server-2024.example.local": "2024",
}


def parse_rsn_path(path):
    """RSN://host/folder/.../model.rvt -> (host, folder_parts, model_name) or None."""
    text = as_text(path).strip()
    if not text.upper().startswith("RSN://"):
        return None
    rest = text[6:].replace("\\", "/")
    parts = [p for p in rest.split("/") if p]
    if len(parts) < 2:
        return None
    return parts[0], parts[1:-1], parts[-1]


def encode_rsn_segment(text):
    raw = as_text(text)
    try:
        data = System.Text.Encoding.UTF8.GetBytes(raw)
    except Exception:
        data = System.Text.Encoding.UTF8.GetBytes(str(raw))
    sb = System.Text.StringBuilder()
    for b in data:
        # unreserved: ALPHA / DIGIT / - . _ ~
        if (
            (48 <= b <= 57)
            or (65 <= b <= 90)
            or (97 <= b <= 122)
            or b in (45, 46, 95, 126)
        ):
            sb.Append(chr(b))
        else:
            sb.Append("%")
            sb.Append(b.ToString("X2"))
    return sb.ToString()


def rsn_year_for_host(host, model_name):
    h = as_text(host).strip()
    if h in RSN_HOST_YEARS:
        return RSN_HOST_YEARS[h]
    cfg = os.path.join(SCRIPT_DIR, "servers.cfg")
    try:
        if os.path.isfile(cfg):
            f = codecs.open(cfg, "r", "utf-8")
            try:
                for line in f:
                    line = line.strip()
                    if not line or line.startswith("#") or "|" not in line:
                        continue
                    bits = line.split("|")
                    if len(bits) >= 3 and bits[0].strip() == h:
                        return bits[2].strip()
            finally:
                f.close()
    except Exception:
        pass
    m = re.search(r"(?i)_R(20|21|22|23|24|25|26)\b", as_text(model_name) or "")
    if m:
        return "20" + m.group(1)
    return None


def http_get_utf8(url):
    req = System.Net.HttpWebRequest.Create(url)
    req.Method = "GET"
    req.Timeout = 45000
    req.ReadWriteTimeout = 45000
    req.Accept = "application/json"
    try:
        req.AutomaticDecompression = (
            System.Net.DecompressionMethods.GZip
            | System.Net.DecompressionMethods.Deflate
        )
    except Exception:
        pass
    req.Headers.Add("User-Name", System.Environment.UserName)
    req.Headers.Add("User-Machine-Name", System.Environment.MachineName)
    req.Headers.Add("Operation-GUID", System.Guid.NewGuid().ToString())
    resp = None
    try:
        resp = req.GetResponse()
        stream = resp.GetResponseStream()
        reader = System.IO.StreamReader(stream, System.Text.Encoding.UTF8)
        try:
            return reader.ReadToEnd()
        finally:
            reader.Close()
    finally:
        if resp is not None:
            resp.Close()


def _json_load(text):
    try:
        clr.AddReference("System.Web.Extensions")
        from System.Web.Script.Serialization import JavaScriptSerializer

        return JavaScriptSerializer().DeserializeObject(text)
    except Exception:
        return None


def _model_size_from_contents(data, model_name):
    if data is None:
        return None
    models = None
    try:
        if hasattr(data, "ContainsKey") and data.ContainsKey("Models"):
            models = data["Models"]
        elif isinstance(data, dict):
            models = data.get("Models") or data.get("models")
    except Exception:
        models = None
    if models is None:
        return None
    want = as_text(model_name).lower()
    try:
        for item in models:
            name = None
            try:
                if hasattr(item, "ContainsKey") and item.ContainsKey("Name"):
                    name = item["Name"]
                elif isinstance(item, dict):
                    name = item.get("Name") or item.get("name")
            except Exception:
                name = None
            if not name:
                continue
            n = as_text(name)
            if not n.lower().endswith(".rvt"):
                n = n + ".rvt"
            if n.lower() != want:
                continue
            size = None
            try:
                if hasattr(item, "ContainsKey") and item.ContainsKey("ModelSize"):
                    size = item["ModelSize"]
                elif isinstance(item, dict):
                    size = item.get("ModelSize") or item.get("modelSize")
            except Exception:
                size = None
            if size is None:
                return None
            return int(size)
    except Exception:
        return None
    return None


def rsn_model_size_bytes(path):
    """Size of central model on Revit Server via Admin REST (ModelSize), bytes."""
    parsed = parse_rsn_path(path)
    if not parsed:
        return None
    host, folder_parts, model = parsed
    year = rsn_year_for_host(host, model)
    if not year:
        Output("WARN rsn size: unknown REST year for host {}".format(host))
        return None

    base = "http://{0}/RevitServerAdminRESTService{1}/AdminRESTService.svc/".format(
        host, year
    )
    variants = []
    if not folder_parts:
        variants.extend(["%7C/contents", "%7C/Contents"])
    else:
        enc = [encode_rsn_segment(p) for p in folder_parts]
        joined = "%7C".join(enc)
        raw = "|".join(folder_parts)
        variants.extend(
            [
                "%7C{0}/contents".format(joined),
                "%7C{0}%7C/contents".format(joined),
                "%7C{0}/Contents".format(joined),
                "%7C{0}%7C/Contents".format(joined),
                "|{0}/contents".format(raw),
                "|{0}|/contents".format(raw),
            ]
        )

    last_err = None
    for svc in variants:
        url = base + svc
        try:
            body = http_get_utf8(url)
            if not body:
                continue
            data = _json_load(body)
            size = _model_size_from_contents(data, model)
            if size is not None:
                return size
            # Regex fallback if serializer failed or structure differs
            if data is None and body:
                pat = (
                    r'"Name"\s*:\s*"'
                    + re.escape(as_text(model).replace(".rvt", ""))
                    + r'(?:\.rvt)?"[^}]*?"ModelSize"\s*:\s*(\d+)'
                )
                m = re.search(pat, body, re.I | re.S)
                if not m:
                    pat2 = (
                        r'"ModelSize"\s*:\s*(\d+)[^}]*?"Name"\s*:\s*"'
                        + re.escape(as_text(model))
                        + r'"'
                    )
                    m = re.search(pat2, body, re.I | re.S)
                if m:
                    return int(m.group(1))
        except Exception as ex:
            last_err = ex
            continue
    if last_err:
        Output("WARN rsn size: {}".format(last_err))
    return None


def fmt_mb(n):
    if n is None:
        return "n/a"
    return "{:.2f} MB".format(n / (1024.0 * 1024.0))


def fmt_delta(before, after):
    if before is None or after is None:
        return "n/a", "n/a"
    delta = after - before
    if before:
        pct = 100.0 * delta / float(before)
        pct_txt = "{:.1f}%".format(pct)
    else:
        pct_txt = "n/a"
    sign = "+" if delta > 0 else ""
    return "{}{}".format(sign, fmt_mb(delta)), pct_txt


def as_text(value):
    if value is None:
        return u""
    if isinstance(value, _unicode):
        return value
    try:
        return _unicode(value)
    except Exception:
        try:
            return _unicode(str(value))
        except Exception:
            return u""


def csv_escape(value):
    text = as_text(value)
    if any(ch in text for ch in (u";", u'"', u"\n", u"\r")):
        return u'"' + text.replace(u'"', u'""') + u'"'
    return text


CSV_SEP = u";"
CSV_COLUMNS = [
    u"имя_модели",
    u"путь",
    u"открыта",
    u"сохранена",
    u"закрыта",
    u"сек",
    u"вес_до_МБ",
    u"вес_после_МБ",
    u"дельта_МБ",
    u"статус",
    u"ошибка",
]
CSV_HEADER = CSV_SEP.join(CSV_COLUMNS)

XLSX_HEADERS = [u"№"] + CSV_COLUMNS
XLSX_COL_WIDTHS = [
    5, 42, 58, 20, 20, 20, 8, 12, 13, 12, 10, 42,
]

SHEET_MODELS = u_from_codes([0x041C, 0x043E, 0x0434, 0x0435, 0x043B, 0x0438])
SHEET_SUMMARY = u_from_codes([0x0421, 0x0432, 0x043E, 0x0434, 0x043A, 0x0430])
SHEET_DYNAMICS = u_from_codes(
    [0x0414, 0x0438, 0x043D, 0x0430, 0x043C, 0x0438, 0x043A, 0x0430]
)
HISTORY_MAX_RUNS = 18


def fmt_mb_cell(n):
    if n is None:
        return u""
    return u"{:.2f}".format(n / (1024.0 * 1024.0)).replace(".", ",")


def model_name_of(path):
    p = as_text(path).replace(u"/", u"\\").rstrip(u"\\")
    if not p:
        return u""
    return p.split(u"\\")[-1]


def now_stamp():
    return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def parse_stamp(stamp):
    text = as_text(stamp).strip()
    if text.startswith(u'="') and len(text) >= 3 and text.endswith(u'"'):
        text = text[2:-1]
    try:
        return datetime.datetime.strptime(text, "%Y-%m-%d %H:%M:%S")
    except Exception:
        return None


def duration_sec(opened_at, finished_at):
    started = parse_stamp(opened_at)
    ended = parse_stamp(finished_at)
    if not started or not ended:
        return u""
    sec = int((ended - started).total_seconds())
    if sec < 0:
        sec = 0
    return u"{}".format(sec)


def parse_mb_cell(text):
    t = as_text(text).strip().replace(u" ", u"").replace(u",", u".")
    if not t:
        return None
    try:
        return float(t)
    except Exception:
        return None


def xml_esc(text):
    return (
        as_text(text)
        .replace(u"&", u"&amp;")
        .replace(u"<", u"&lt;")
        .replace(u">", u"&gt;")
        .replace(u'"', u"&quot;")
    )


def col_letter(index1):
    n = int(index1)
    name = u""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        name = _unichr(65 + rem) + name
    return name


def split_csv_line(line):
    result = []
    buf = []
    in_quotes = False
    text = as_text(line)
    i = 0
    while i < len(text):
        ch = text[i]
        if ch == u'"':
            if in_quotes and i + 1 < len(text) and text[i + 1] == u'"':
                buf.append(u'"')
                i += 1
            else:
                in_quotes = not in_quotes
        elif ch == u";" and not in_quotes:
            result.append(u"".join(buf))
            buf = []
        else:
            buf.append(ch)
        i += 1
    result.append(u"".join(buf))
    return result


def read_utf8_text(path):
    """Binary-safe UTF-8 read for IronPython (codecs utf-8-sig often breaks)."""
    if not path or not os.path.isfile(path):
        return u""
    f = open(path, "rb")
    try:
        data = f.read()
    finally:
        f.close()
    if not data:
        return u""
    bom = codecs.BOM_UTF8
    if data.startswith(bom):
        data = data[len(bom) :]
    try:
        return data.decode("utf-8")
    except Exception:
        try:
            return data.decode("utf-8", "replace")
        except Exception:
            return as_text(data)


def read_report_rows(path):
    if not path or not os.path.isfile(path):
        return []
    raw = read_utf8_text(path)
    lines = raw.splitlines()
    if not lines:
        return []
    headers = split_csv_line(lines[0])
    rows = []
    for line in lines[1:]:
        if not as_text(line).strip():
            continue
        vals = split_csv_line(line)
        rec = {}
        for i, key in enumerate(headers):
            rec[key] = vals[i] if i < len(vals) else u""
        name = rec.get(u"имя_модели") or model_name_of(rec.get(u"путь"))
        rec[u"имя_модели"] = name
        if not rec.get(u"сек"):
            rec[u"сек"] = duration_sec(rec.get(u"открыта"), rec.get(u"закрыта"))
        rows.append(rec)
    return rows


def _xf_xml(num_fmt, font_id, fill_id, border_id, align=None):
    apply = []
    if num_fmt:
        apply.append(u'applyNumberFormat="1"')
    if font_id:
        apply.append(u'applyFont="1"')
    if fill_id:
        apply.append(u'applyFill="1"')
    if border_id:
        apply.append(u'applyBorder="1"')
    if align:
        apply.append(u'applyAlignment="1"')
    body = u""
    if align:
        body = u"<alignment {}/>".format(align)
    return (
        u'<xf numFmtId="{nf}" fontId="{ft}" fillId="{fl}" borderId="{bd}" xfId="0" {ap}>{body}</xf>'
        .format(
            nf=num_fmt,
            ft=font_id,
            fl=fill_id,
            bd=border_id,
            ap=u" ".join(apply),
            body=body,
        )
    )


def build_styles():
    xfs = []

    def add(num_fmt, font_id, fill_id, border_id, align=None):
        xfs.append(_xf_xml(num_fmt, font_id, fill_id, border_id, align))
        return len(xfs) - 1

    styles = {}
    wrap_c = u'vertical="center" wrapText="1" horizontal="center"'
    wrap_l = u'vertical="center" wrapText="1" horizontal="left"'
    mid = u'vertical="center"'
    styles["base"] = add(0, 0, 0, 1, mid)
    styles["header"] = add(0, 1, 2, 1, wrap_c)
    styles["zebra"] = add(0, 0, 3, 1, mid)
    styles["num"] = add(164, 0, 0, 1, mid)
    styles["num_z"] = add(164, 0, 3, 1, mid)
    styles["dt"] = add(165, 0, 0, 1, mid)
    styles["dt_z"] = add(165, 0, 3, 1, mid)
    styles["int"] = add(166, 0, 0, 1, mid)
    styles["int_z"] = add(166, 0, 3, 1, mid)
    styles["ok"] = add(0, 0, 4, 1, mid)
    styles["err"] = add(0, 0, 5, 1, mid)
    styles["dneg"] = add(164, 2, 0, 1, mid)
    styles["dneg_z"] = add(164, 2, 3, 1, mid)
    styles["dpos"] = add(164, 3, 0, 1, mid)
    styles["dpos_z"] = add(164, 3, 3, 1, mid)
    styles["err_row"] = add(0, 0, 5, 1, mid)
    styles["err_num"] = add(164, 0, 5, 1, mid)
    styles["err_dt"] = add(165, 0, 5, 1, mid)
    styles["err_int"] = add(166, 0, 5, 1, mid)
    styles["sum_k"] = add(0, 4, 6, 1, wrap_l)
    styles["sum_v"] = add(0, 0, 0, 1, mid)
    styles["sum_n"] = add(164, 0, 0, 1, mid)
    styles["sum_alert"] = add(166, 0, 5, 1, mid)
    xml = (
        u'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        u'<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        u'<numFmts count="3">'
        u'<numFmt numFmtId="164" formatCode="0.00"/>'
        u'<numFmt numFmtId="165" formatCode="yyyy-mm-dd hh:mm:ss"/>'
        u'<numFmt numFmtId="166" formatCode="0"/>'
        u"</numFmts>"
        u'<fonts count="5">'
        u'<font><sz val="11"/><name val="Calibri"/></font>'
        u'<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>'
        u'<font><sz val="11"/><color rgb="FF006100"/><name val="Calibri"/></font>'
        u'<font><sz val="11"/><color rgb="FF9C0006"/><name val="Calibri"/></font>'
        u'<font><b/><sz val="11"/><name val="Calibri"/></font>'
        u"</fonts>"
        u'<fills count="7">'
        u'<fill><patternFill patternType="none"/></fill>'
        u'<fill><patternFill patternType="gray125"/></fill>'
        u'<fill><patternFill patternType="solid"><fgColor rgb="FF1F4E79"/></patternFill></fill>'
        u'<fill><patternFill patternType="solid"><fgColor rgb="FFF2F2F2"/></patternFill></fill>'
        u'<fill><patternFill patternType="solid"><fgColor rgb="FFC6EFCE"/></patternFill></fill>'
        u'<fill><patternFill patternType="solid"><fgColor rgb="FFFFC7CE"/></patternFill></fill>'
        u'<fill><patternFill patternType="solid"><fgColor rgb="FFD6DCE4"/></patternFill></fill>'
        u"</fills>"
        u'<borders count="2">'
        u"<border><left/><right/><top/><bottom/><diagonal/></border>"
        u'<border>'
        u'<left style="thin"><color rgb="FFB0B0B0"/></left>'
        u'<right style="thin"><color rgb="FFB0B0B0"/></right>'
        u'<top style="thin"><color rgb="FFB0B0B0"/></top>'
        u'<bottom style="thin"><color rgb="FFB0B0B0"/></bottom>'
        u"<diagonal/>"
        u"</border>"
        u"</borders>"
        u'<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        u'<cellXfs count="{n}">{xfs}</cellXfs>'
        u"</styleSheet>"
    ).format(n=len(xfs), xfs=u"".join(xfs))
    return xml, styles


def excel_serial(dt):
    epoch = datetime.datetime(1899, 12, 30)
    delta = dt - epoch
    return delta.days + (delta.seconds + delta.microseconds / 1000000.0) / 86400.0


def cell_inline(ref, text, style):
    t = xml_esc(text)
    return u'<c r="{r}" t="inlineStr" s="{s}"><is><t xml:space="preserve">{t}</t></is></c>'.format(
        r=ref, s=style, t=t
    )


def cell_num(ref, value, style):
    return u'<c r="{r}" s="{s}"><v>{v}</v></c>'.format(r=ref, s=style, v=value)


def build_models_sheet(rows, styles):
    ncols = len(XLSX_HEADERS)
    last_col = col_letter(ncols)
    last_row = max(1, len(rows) + 1)
    parts = [
        u'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        u'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"',
        u' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
        u'<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>',
        u"<sheetViews>",
        u'<sheetView tabSelected="1" workbookViewId="0">',
        u'<pane xSplit="2" ySplit="1" topLeftCell="C2" activePane="bottomRight" state="frozen"/>',
        u'<selection pane="bottomRight" activeCell="C2" sqref="C2"/>',
        u"</sheetView></sheetViews>",
        u"<cols>",
    ]
    for i, width in enumerate(XLSX_COL_WIDTHS):
        parts.append(
            u'<col min="{n}" max="{n}" width="{w}" customWidth="1"/>'.format(
                n=i + 1, w=width
            )
        )
    parts.append(u"</cols><sheetData>")
    parts.append(u'<row r="1" ht="24" customHeight="1">')
    for i, title in enumerate(XLSX_HEADERS):
        ref = col_letter(i + 1) + "1"
        parts.append(cell_inline(ref, title, styles["header"]))
    parts.append(u"</row>")

    num_keys = set([u"вес_до_МБ", u"вес_после_МБ", u"дельта_МБ"])
    dt_keys = set([u"открыта", u"сохранена", u"закрыта"])
    int_keys = set([u"сек"])

    for ridx, rec in enumerate(rows):
        rnum = ridx + 2
        zebra = (ridx % 2) == 1
        is_err = as_text(rec.get(u"статус")).upper() == u"ERROR"
        parts.append(u'<row r="{}" ht="18" customHeight="1">'.format(rnum))
        values = [ridx + 1]
        for key in CSV_COLUMNS:
            values.append(rec.get(key, u""))
        for cidx, val in enumerate(values):
            key = XLSX_HEADERS[cidx]
            ref = col_letter(cidx + 1) + str(rnum)
            status = as_text(rec.get(u"статус")).upper()
            if key == u"статус":
                st = styles["err"] if status == u"ERROR" else styles["ok"]
                parts.append(cell_inline(ref, val, st))
                continue
            if key == u"№" or key in int_keys:
                try:
                    num = int(as_text(val) or 0) if key != u"№" else int(val)
                    if key == u"№":
                        num = int(val)
                    if is_err:
                        st = styles["err_int"]
                    else:
                        st = styles["int_z"] if zebra else styles["int"]
                    parts.append(cell_num(ref, num, st))
                    continue
                except Exception:
                    pass
            if key in dt_keys:
                dt = parse_stamp(val)
                if dt:
                    if is_err:
                        st = styles["err_dt"]
                    else:
                        st = styles["dt_z"] if zebra else styles["dt"]
                    parts.append(cell_num(ref, "{:.10f}".format(excel_serial(dt)), st))
                    continue
            if key in num_keys:
                num = parse_mb_cell(val)
                if num is not None:
                    if key == u"дельта_МБ":
                        if is_err:
                            st = styles["err_num"]
                        elif num < 0:
                            st = styles["dneg_z"] if zebra else styles["dneg"]
                        elif num > 0:
                            st = styles["dpos_z"] if zebra else styles["dpos"]
                        else:
                            st = styles["num_z"] if zebra else styles["num"]
                    else:
                        st = styles["err_num"] if is_err else (
                            styles["num_z"] if zebra else styles["num"]
                        )
                    parts.append(cell_num(ref, "{:.6f}".format(num), st))
                    continue
            if is_err:
                st = styles["err_row"]
            else:
                st = styles["zebra"] if zebra else styles["base"]
            parts.append(cell_inline(ref, val, st))
        parts.append(u"</row>")
    parts.append(u"</sheetData>")
    parts.append(
        u'<autoFilter ref="A1:{lc}{lr}"/>'.format(lc=last_col, lr=last_row)
    )
    parts.append(
        u'<pageSetup orientation="landscape" paperSize="9" fitToWidth="1" fitToHeight="0"/>'
    )
    parts.append(u"</worksheet>")
    return u"".join(parts)


def _sum_row(label, value, note, styles, rnum, value_style):
    cells = [
        cell_inline("A{}".format(rnum), label, styles["sum_k"]),
        value,
        cell_inline("C{}".format(rnum), note or u"", styles["sum_v"]),
    ]
    return u'<row r="{r}">{c}</row>'.format(r=rnum, c=u"".join(cells))


def build_summary_sheet(rows, styles):
    n = len(rows)
    n_ok = 0
    n_err = 0
    before_sum = 0.0
    after_sum = 0.0
    n_before = 0
    n_down = 0
    n_up = 0
    n_flat = 0
    max_sec = -1
    max_sec_name = u""
    min_delta = None
    min_delta_name = u""
    max_delta = None
    max_delta_name = u""
    for rec in rows:
        st = as_text(rec.get(u"статус")).upper()
        if st == u"ERROR":
            n_err += 1
        else:
            n_ok += 1
        before = parse_mb_cell(rec.get(u"вес_до_МБ"))
        after = parse_mb_cell(rec.get(u"вес_после_МБ"))
        if before is not None:
            before_sum += before
            n_before += 1
        if after is not None:
            after_sum += after
        if before is not None and after is not None:
            d = after - before
            if d < -0.005:
                n_down += 1
            elif d > 0.005:
                n_up += 1
            else:
                n_flat += 1
            name = rec.get(u"имя_модели") or u""
            if min_delta is None or d < min_delta:
                min_delta = d
                min_delta_name = name
            if max_delta is None or d > max_delta:
                max_delta = d
                max_delta_name = name
        try:
            sec = int(as_text(rec.get(u"сек") or u"0") or 0)
        except Exception:
            sec = 0
        if sec > max_sec:
            max_sec = sec
            max_sec_name = rec.get(u"имя_модели") or u""

    delta_sum = after_sum - before_sum if n_before else None
    items = [
        (u"Моделей", n, u"", "int"),
        (u"OK", n_ok, u"", "int"),
        (u"ERROR", n_err, u"фильтр по столбцу статус", "alert" if n_err else "int"),
        (u"Вес до, МБ", before_sum if n_before else None, u"", "num"),
        (u"Вес после, МБ", after_sum if n_before else None, u"", "num"),
        (u"Дельта, МБ", delta_sum, u"сумма по прогону", "num"),
        (u"Уменьшились", n_down, u"", "int"),
        (u"Увеличились", n_up, u"сортировка по дельта_МБ", "int"),
        (u"Без изменения", n_flat, u"", "int"),
        (u"Самое долгое, сек", max_sec if max_sec >= 0 else None, max_sec_name, "int"),
        (u"Макс. сжатие, МБ", min_delta, min_delta_name, "num"),
        (u"Макс. рост, МБ", max_delta, max_delta_name, "num"),
    ]

    parts = [
        u'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        u'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        u"<cols>",
        u'<col min="1" max="1" width="28" customWidth="1"/>',
        u'<col min="2" max="2" width="16" customWidth="1"/>',
        u'<col min="3" max="3" width="55" customWidth="1"/>',
        u"</cols><sheetData>",
        u'<row r="1" ht="24" customHeight="1">',
        cell_inline("A1", u"показатель", styles["header"]),
        cell_inline("B1", u"значение", styles["header"]),
        cell_inline("C1", u"модель / подсказка", styles["header"]),
        u"</row>",
    ]
    for i, (label, value, note, kind) in enumerate(items):
        rnum = i + 2
        if value is None or value == u"":
            val_cell = cell_inline("B{}".format(rnum), u"", styles["sum_v"])
        elif kind == "alert":
            val_cell = cell_num("B{}".format(rnum), int(value), styles["sum_alert"])
        elif kind == "int":
            val_cell = cell_num("B{}".format(rnum), int(value), styles["int"])
        else:
            val_cell = cell_num(
                "B{}".format(rnum), "{:.6f}".format(float(value)), styles["sum_n"]
            )
        parts.append(
            u'<row r="{r}">{a}{b}{c}</row>'.format(
                r=rnum,
                a=cell_inline("A{}".format(rnum), label, styles["sum_k"]),
                b=val_cell,
                c=cell_inline("C{}".format(rnum), note, styles["sum_v"]),
            )
        )
    last = 1 + len(items)
    parts.append(u"</sheetData>")
    parts.append(u'<autoFilter ref="A1:C{}"/>'.format(last))
    parts.append(u"</worksheet>")
    return u"".join(parts)


def write_xlsx(path, parts_map):
    if os.path.isfile(path):
        raise Exception("XLSX already exists, refuse overwrite: {}".format(path))
    tmp = path + u".tmp"
    if os.path.isfile(tmp):
        try:
            os.remove(tmp)
        except Exception:
            pass
    compression = zipfile.ZIP_DEFLATED
    try:
        zf = zipfile.ZipFile(tmp, "w", compression)
    except Exception:
        zf = zipfile.ZipFile(tmp, "w", zipfile.ZIP_STORED)
    try:
        for name in parts_map:
            payload = parts_map[name]
            if not isinstance(payload, bytes):
                payload = payload.encode("utf-8")
            zf.writestr(name, payload)
    finally:
        zf.close()
    if os.path.isfile(path):
        try:
            os.remove(tmp)
        except Exception:
            pass
        raise Exception("XLSX already exists, refuse overwrite: {}".format(path))
    os.rename(tmp, path)


def rebuild_xlsx(csv_path):
    xlsx_path = xlsx_path_for(csv_path)
    if os.path.isfile(xlsx_path):
        return xlsx_path
    rows = read_report_rows(csv_path)
    styles_xml, styles = build_styles()
    models_xml = build_models_sheet(rows, styles)
    summary_xml = build_summary_sheet(rows, styles)
    ct = (
        u'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        u'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        u'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        u'<Default Extension="xml" ContentType="application/xml"/>'
        u'<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        u'<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        u'<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        u'<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        u"</Types>"
    )
    rels_root = (
        u'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        u'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        u'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        u"</Relationships>"
    )
    rels_wb = (
        u'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        u'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        u'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
        u'<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>'
        u'<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        u"</Relationships>"
    )
    workbook = (
        u'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        u'<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
        u' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        u"<sheets>"
        u'<sheet name="{m}" sheetId="1" r:id="rId1"/>'
        u'<sheet name="{s}" sheetId="2" r:id="rId2"/>'
        u"</sheets>"
        u"<definedNames>"
        u'<definedName name="_xlnm.Print_Titles">\'{m}\'!$1:$1</definedName>'
        u"</definedNames>"
        u"</workbook>"
    ).format(m=xml_esc(SHEET_MODELS), s=xml_esc(SHEET_SUMMARY))
    write_xlsx(
        xlsx_path,
        {
            "[Content_Types].xml": ct,
            "_rels/.rels": rels_root,
            "xl/workbook.xml": workbook,
            "xl/_rels/workbook.xml.rels": rels_wb,
            "xl/styles.xml": styles_xml,
            "xl/worksheets/sheet1.xml": models_xml,
            "xl/worksheets/sheet2.xml": summary_xml,
        },
    )
    return xlsx_path


def init_report():
    global REPORTS_DIR, REPORT_PATH
    if REPORT_PATH:
        return

    share = os.path.join(SCRIPT_DIR, OTCHETY_NAME)
    local = local_reports_root()
    last_err = None
    for root, is_local in ((share, False), (local, True)):
        if not root:
            continue
        if not try_reports_root(root):
            continue
        try:
            csv_folder = csv_reports_dir(root)
            path = allocate_report_file(csv_folder, root, get_session_id())
            REPORTS_DIR = root
            REPORT_PATH = path
            if is_local:
                Output(
                    "NOTE: share reports not writable; using local: {}".format(root)
                )
            return
        except Exception as ex:
            last_err = ex
            REPORTS_DIR = None
            REPORT_PATH = None
            Output("WARN report init at {}: {}".format(root, ex))

    raise Exception(
        "Cannot write reports. Tried share ({0}) and local ({1}). Last error: {2}".format(
            share, local or u"(no LOCALAPPDATA)", last_err
        )
    )


def append_report_row(
    path,
    opened_at,
    saved_at,
    finished_at,
    size_before,
    size_after,
    status,
    error,
):
    try:
        init_report()
    except Exception as ex:
        Output("ERROR report folder: {}".format(ex))
        return

    name = model_name_of(path)
    delta = u""
    if size_before is not None and size_after is not None:
        delta = fmt_mb_cell(size_after - size_before)
    rec = {
        u"имя_модели": name,
        u"путь": as_text(path),
        u"открыта": as_text(opened_at),
        u"сохранена": as_text(saved_at),
        u"закрыта": as_text(finished_at),
        u"сек": duration_sec(opened_at, finished_at),
        u"вес_до_МБ": fmt_mb_cell(size_before),
        u"вес_после_МБ": fmt_mb_cell(size_after),
        u"дельта_МБ": delta,
        u"статус": as_text(status),
        u"ошибка": as_text(error),
    }
    line = CSV_SEP.join(csv_escape(rec[k]) for k in CSV_COLUMNS) + u"\n"
    new_file = not os.path.isfile(REPORT_PATH)
    try:
        f = open(REPORT_PATH, "ab")
        try:
            if new_file:
                f.write(codecs.BOM_UTF8)
                f.write((CSV_HEADER + u"\n").encode("utf-8"))
            f.write(line.encode("utf-8"))
        finally:
            f.close()
        Output("Report csv: {}".format(REPORT_PATH))
    except Exception as ex:
        Output("WARN report csv: {}".format(ex))
        return

    try:
        xlsx_path = rebuild_xlsx(REPORT_PATH)
        Output("Report xlsx: {}".format(xlsx_path))
    except Exception as ex:
        Output("WARN report xlsx: {}".format(ex))


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


def relinquish_everything(document):
    if not document.IsWorkshared:
        return
    try:
        Output("Relinquishing all worksets/elements...")
        WorksharingUtils.RelinquishOwnership(
            document, make_relinquish_all_options(), TransactWithCentralOptions()
        )
        Output("Relinquish done.")
    except Exception as ex:
        Output("WARN RelinquishOwnership: {}".format(ex))


def central_user_path(document):
    try:
        return ModelPathUtils.ConvertModelPathToUserVisiblePath(
            document.GetWorksharingCentralModelPath()
        )
    except Exception:
        return None


def measure_path(document, path):
    """Размер того файла, который реально сжимаем."""
    if is_rsn_path(path):
        sz = rsn_model_size_bytes(path)
        if sz is not None:
            return sz, "rsn-rest"
        return None, "rsn"
    central = None
    try:
        if document.IsWorkshared and not getattr(document, "IsDetached", False):
            central = central_user_path(document)
    except Exception:
        central = None
    if central and (not is_rsn_path(central)) and os.path.isfile(central):
        return file_size_bytes(central), "central"
    if central and is_rsn_path(central):
        sz = rsn_model_size_bytes(central)
        if sz is not None:
            return sz, "rsn-rest"
        return None, "rsn"
    return file_size_bytes(path), "file"


def sync_compact(document, path):
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

    if is_backup_copy_path(path):
        raise Exception(
            "Opened a backup/copy, not a live model: {}. "
            "Rebuild rvt_list with picker option 1 "
            "(RVT\\*.rvt only, not Резерв / dated folders).".format(path)
        )

    central = central_user_path(document)
    if central:
        Output("Sync+Compact to central: {}".format(central))
        if (not is_rsn_path(central)) and (not os.path.isfile(central)):
            raise Exception(
                "The central model is missing: {}. "
                "Opened file is not the live central. "
                "Pick models from RVT\\ (not Резерв).".format(central)
            )
    else:
        Output("Sync+Compact to central...")

    sync_opts = SynchronizeWithCentralOptions()
    sync_opts.Compact = True
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
    Output("SynchronizedWithCentral Compact=True + relinquish OK")
    return path


def to_model_path(path):
    """Строка пути (UNC / RSN://) -> ModelPath для SaveAs / GetUserWorksetInfo."""
    text = as_text(path).strip()
    if not text:
        raise Exception("Empty path for ModelPath conversion.")
    return ModelPathUtils.ConvertUserVisiblePathToModelPath(text)


def save_compact_in_place(document, path, force_save_as=False):
    """Сохранение со сжатием.

    force_save_as=True (deep): как в UI «Сохранить как» → параметры файла:
    считать моделью из хранилища, Сжать, открыть наборы по умолчанию = Задать.
    Не выходим на Document.Save — иначе галочка «Задать» не применяется.
    """
    clear_readonly_attribute(path)

    if not force_save_as:
        try:
            save_opts = SaveOptions()
            save_opts.Compact = True
            document.Save(save_opts)
            Output("Saved Compact=True: {}".format(path or document.PathName))
            return path
        except Exception as ex:
            Output("WARN Document.Save Compact: {}".format(ex))

    opts = SaveAsOptions()
    opts.OverwriteExistingFile = True
    opts.Compact = True

    if document.IsWorkshared:
        wopts = WorksharingSaveAsOptions()
        wopts.SaveAsCentral = True
        try:
            wopts.ClearTransmitted = False
        except Exception:
            pass
        # UI: «Открыть рабочий набор по умолчанию: Задать...»
        try:
            wopts.OpenWorksetsDefault = SimpleWorksetConfiguration.AskUserToSpecify
            Output("OpenWorksetsDefault = AskUserToSpecify (Задать)")
        except Exception as ex:
            Output("WARN OpenWorksetsDefault AskUserToSpecify: {}".format(ex))
        opts.SetWorksharingOptions(wopts)

    target = path or document.PathName
    model_path = to_model_path(target)
    if force_save_as:
        Output(
            "SaveAs (deep): SaveAsCentral + Compact + AskUserToSpecify -> {}".format(
                target
            )
        )
    document.SaveAs(model_path, opts)
    Output("SavedAs Compact=True SaveAsCentral=True (ModelPath): {}".format(target))

    relinquish_everything(document)
    try:
        save_opts = SaveOptions()
        save_opts.Compact = True
        document.Save(save_opts)
        Output("Saved after relinquish.")
    except Exception as ex:
        Output("WARN save after relinquish: {}".format(ex))
    return target


def deep_save_as_and_sync(document, path):
    """Ручной эталон: Save As ФХ (Compact+Задать) → Sync со сжатием."""
    Output(
        "deep clean: SaveAs as central + Compact + "
        "OpenWorksetsDefault=Задать (AskUserToSpecify)."
    )
    saved = save_compact_in_place(document, path, force_save_as=True)
    verify_worksharing_preserved(document, saved)

    # Шаг эталона: «Синхронизироваться с сжатием и закрыть»
    try:
        if document.IsWorkshared and not getattr(document, "IsDetached", False):
            Output("deep clean: SynchronizeWithCentral Compact=True (final step).")
            sync_compact(document, saved or path)
        else:
            Output(
                "WARN deep: skip final Sync (Detached={} / Workshared={}).".format(
                    getattr(document, "IsDetached", None),
                    getattr(document, "IsWorkshared", None),
                )
            )
    except Exception as ex:
        Output("WARN deep final Sync+Compact: {}".format(ex))
        # SaveAs уже прошёл — не роняем весь файл из‑за Sync
    return saved


def compact_document(document, path):
    if is_backup_copy_path(path):
        raise Exception(
            "Opened a backup/copy, not a live model: {}. "
            "Rebuild rvt_list with picker option 1 "
            "(RVT\\*.rvt only, not Резерв / dated folders).".format(path)
        )

    is_detached = False
    try:
        is_detached = document.IsDetached
    except Exception:
        is_detached = False

    if is_rsn_path(path):
        Output("Path type: Revit Server (RSN)")
        return sync_compact(document, path), "rsn-sync"

    if (
        document.IsWorkshared
        and (not is_detached)
        and document.GetWorksharingCentralModelPath() is not None
    ):
        Output("Path type: workshared local/UNC")
        return sync_compact(document, path), "workshared-sync"

    if is_detached and document.IsWorkshared:
        Output(
            "WARN: Detached workshared file — SaveAs Compact in place. "
            "Prefer Create New Local for shared models."
        )
        return save_compact_in_place(document, path), "detached-saveas"

    Output("Path type: non-workshared file")
    return save_compact_in_place(document, path), "file-save"


opened_at = now_stamp()
size_before = None
try:
    saved_at = u""
    Output()
    Output("Script dir: {}".format(SCRIPT_DIR))
    Output("File: {}".format(revitFilePath))
    write_live_status(revitFilePath, u"start")
    Output(
        "Workshared={}, Detached={}".format(
            doc.IsWorkshared, getattr(doc, "IsDetached", False)
        )
    )
    Output("Compact = True (no extra config)")

    size_before, size_kind = measure_path(doc, revitFilePath)
    Output("Size before ({}): {}".format(size_kind, fmt_mb(size_before)))

    saved, kind = compact_document(doc, revitFilePath)
    saved_at = now_stamp()

    # After Sync+Compact, RS may need a moment to refresh ModelSize.
    if is_rsn_path(revitFilePath) or (
        saved and is_rsn_path(as_text(saved))
    ):
        try:
            time.sleep(1.5)
        except Exception:
            pass

    size_after, size_kind_after = measure_path(doc, saved or revitFilePath)
    delta_txt, pct_txt = fmt_delta(size_before, size_after)
    Output("Size after ({}): {}".format(size_kind_after, fmt_mb(size_after)))

    finished_at = now_stamp()
    Output("Times open/save/end: {} / {} / {}".format(opened_at, saved_at, finished_at))
    append_report_row(
        saved or revitFilePath,
        opened_at,
        saved_at,
        finished_at,
        size_before,
        size_after,
        "OK",
        "",
    )

    Output("Done: {}".format(saved))
    Output("OK")
except Exception as ex:
    Output("ERROR: {}".format(ex))
    try:
        append_report_row(
            revitFilePath,
            opened_at,
            u"",
            now_stamp(),
            size_before,
            None,
            "ERROR",
            as_text(ex),
        )
    except Exception:
        pass
    raise
