# -*- coding: utf-8 -*-
# Name: ops_lib.py
# Version: 3.1.0
# What it does: Shared IronPython helpers for batch_model_ops RBP scripts.
# Inputs: Imported by saveas/detach/update scripts.
# Outputs: Library functions.
# How to run: Not run standalone.
# Notes: IronPython 2.7.

"""Shared helpers for batch_model_ops (IronPython 2.7)."""

from __future__ import print_function

import clr
import json
import os
import System

clr.AddReference("System")


# Byte order mark. IronPython 2.7.3 (RBP) runs str.startswith/endswith through
# culture-sensitive .NET StartsWith/EndsWith, where U+FEFF is ignorable:
# u"action".startswith(u"\ufeff") is True there. BOM checks compare chars
# directly (ordinal), never through startswith/endswith.
BOM = u"\ufeff"


def strip_bom(text):
    """Drop leading BOM chars (ordinal compare; any count). None stays None."""
    if not text:
        return text
    while text[:1] == BOM:
        text = text[1:]
    return text


def is_rsn_path(path):
    if not path:
        return False
    return path.upper().startswith("RSN://")


def basename_no_ext(path_or_name):
    if not path_or_name:
        return ""
    name = path_or_name.replace("\\", "/").split("/")[-1].strip()
    if name.lower().endswith(".rvt"):
        name = name[:-4]
    return name


def rbp_root():
    home = os.environ.get("USERPROFILE") or r"C:\Users\Public"
    return os.path.join(home, r"Documents\doc\script\rbp")


# Reports of this tool live in its own folder: <tool root>\\<REPORTS_FOLDER_NAME>,
# the Russian word for "reports" (escaped: the log side stays ASCII).
REPORTS_FOLDER_NAME = u"\u043e\u0442\u0447\u0451\u0442\u044b"


def legacy_reports_dir():
    """Old common folder rbp\\reports (read-only fallback for older files)."""
    return os.path.join(rbp_root(), "reports")


def reports_dir_from_job_text(text):
    """reports_dir from job.json / workset_job.json text, "" when absent or broken."""
    try:
        text = strip_bom(text)
        data = json.loads(text)
        value = data.get("reports_dir") if isinstance(data, dict) else None
        if value:
            return value
    except Exception:
        pass
    return ""


def tool_reports_dir(tool_dir):
    """The window writes reports_dir into job.json / workset_job.json.
    Without it: <tool_dir>\\<REPORTS_FOLDER_NAME>."""
    for name in ("job.json", "workset_job.json"):
        path = os.path.join(tool_dir, name)
        try:
            if file_exists(path):
                value = reports_dir_from_job_text(read_all_text(path))
                if value:
                    return value
        except Exception:
            pass
    return os.path.join(tool_dir, REPORTS_FOLDER_NAME)


def shared_reports_dir(tool_dir=None):
    """Folder for all reports of this tool (tool_dir\\<REPORTS_FOLDER_NAME>). Created if missing.
    No tool_dir, or the folder cannot be created: the old rbp\\reports."""
    if tool_dir:
        try:
            path = tool_reports_dir(tool_dir)
            ensure_dir(path)
            return path
        except Exception:
            pass
    path = legacy_reports_dir()
    ensure_dir(path)
    return path


def ensure_dir(path):
    if not path:
        return
    if not System.IO.Directory.Exists(path):
        System.IO.Directory.CreateDirectory(path)


def file_exists(path):
    if not path:
        return False
    try:
        return System.IO.File.Exists(path)
    except Exception:
        return os.path.isfile(path)


def read_all_text(path):
    """UTF-8 text without a BOM (ReadAllText drops one; strip_bom drops extra)."""
    return strip_bom(System.IO.File.ReadAllText(path, System.Text.Encoding.UTF8))


def write_all_text(path, text):
    folder = System.IO.Path.GetDirectoryName(path)
    ensure_dir(folder)
    System.IO.File.WriteAllText(path, text, System.Text.Encoding.UTF8)


def append_all_text(path, text):
    folder = System.IO.Path.GetDirectoryName(path)
    ensure_dir(folder)
    System.IO.File.AppendAllText(path, text, System.Text.Encoding.UTF8)


def load_mapping(path):
    """old_basename (no ext, UPPER) -> new_basename (no ext)."""
    mapping = {}
    if not file_exists(path):
        raise Exception("mapping.csv not found: {}".format(path))

    text = read_all_text(path)
    lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")

    for i, raw in enumerate(lines):
        line = raw.strip()
        if not line or line.startswith("#"):
            continue
        low = line.lower()
        if i == 0 and ("old" in low and "new" in low):
            continue
        sep = ";" if ";" in line else ("," if "," in line else None)
        if not sep:
            continue
        parts = [p.strip().strip('"') for p in line.split(sep)]
        if len(parts) < 2:
            continue
        old_b = basename_no_ext(parts[0])
        new_b = basename_no_ext(parts[1])
        if not old_b or not new_b:
            continue
        mapping[old_b.upper()] = new_b

    if not mapping:
        raise Exception("mapping.csv empty: {}".format(path))
    return mapping


def replace_filename_in_path(path, old_base, new_base):
    if not path:
        return None
    path_l = path.lower()
    old_rvt = (old_base + ".rvt").lower()
    idx = path_l.rfind(old_rvt)
    if idx >= 0:
        return path[:idx] + new_base + ".rvt" + path[idx + len(old_rvt) :]
    for a, b in (
        (old_base + ".rvt", new_base + ".rvt"),
        (old_base + ".RVT", new_base + ".RVT"),
        (old_base, new_base),
    ):
        if a in path:
            return path.replace(a, b, 1)
    return None


def map_path(old_path, mapping):
    """Return (new_path, old_base, new_base) or raise."""
    base = basename_no_ext(old_path)
    key = base.upper()
    if key not in mapping:
        raise Exception("Path basename not in mapping: {}".format(base))
    new_base = mapping[key]
    new_path = replace_filename_in_path(old_path, base, new_base)
    if not new_path:
        folder = old_path.replace("\\", "/").rsplit("/", 1)[0]
        if is_rsn_path(old_path) or "/" in old_path:
            new_path = folder + "/" + new_base + ".rvt"
        else:
            sep = "\\" if "\\" in old_path else "/"
            new_path = folder + sep + new_base + ".rvt"
    return new_path, base, new_base


def csv_escape(val):
    if val is None:
        return ""
    try:
        s = unicode(val)
    except Exception:
        s = str(val)
    s = s.replace("\r", " ").replace("\n", " ")
    if ";" in s or '"' in s:
        s = '"' + s.replace('"', '""') + '"'
    return s


def append_csv(path, row_values):
    line = ";".join([csv_escape(v) for v in row_values]) + "\n"
    append_all_text(path, line)


def load_job_paths(path):
    """Rows keyed by old_name UPPER, old_path UPPER, staging_path UPPER."""
    by_name = {}
    by_old = {}
    by_stg = {}
    if not file_exists(path):
        return by_name, by_old, by_stg
    text = read_all_text(path)
    lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    for i, raw in enumerate(lines):
        line = raw.strip()
        if not line or line.startswith("#"):
            continue
        low = line.lower()
        if i == 0 and "old_path" in low:
            continue
        sep = ";" if ";" in line else ("," if "," in line else None)
        if not sep:
            continue
        parts = [p.strip().strip('"') for p in line.split(sep)]
        if len(parts) < 5:
            continue
        row = {
            "old_path": parts[0],
            "new_path": parts[1],
            "staging_path": parts[2],
            "old_name": basename_no_ext(parts[3] or parts[0]),
            "new_name": basename_no_ext(parts[4] or parts[1]),
            "needs_buffer": (parts[5] == "1") if len(parts) > 5 else False,
        }
        if row["old_name"]:
            by_name[row["old_name"].upper()] = row
        if row["old_path"]:
            by_old[row["old_path"].replace("\\", "/").upper()] = row
        if row["staging_path"]:
            by_stg[row["staging_path"].replace("\\", "/").upper()] = row
    return by_name, by_old, by_stg


def find_job_row(revit_path, by_old, by_stg, by_name):
    if not revit_path:
        return None
    key = revit_path.replace("\\", "/").upper()
    if key in by_old:
        return by_old[key]
    if key in by_stg:
        return by_stg[key]
    base = basename_no_ext(revit_path).upper()
    if base in by_name:
        return by_name[base]
    return None


def safe_reports_dir(script_dir=None):
    """Reports folder of the tool (see shared_reports_dir). Created if missing."""
    try:
        primary = shared_reports_dir(script_dir)
        probe = os.path.join(primary, "_write_probe.tmp")
        write_all_text(probe, "ok")
        System.IO.File.Delete(probe)
        return primary
    except Exception:
        local = os.path.join(
            System.Environment.GetFolderPath(
                System.Environment.SpecialFolder.LocalApplicationData
            ),
            "BatchRvt",
            "reports",
        )
        ensure_dir(local)
        return local


def parse_allow_overwrite(job_json_text):
    """True only if job.json explicitly has "allow_overwrite": true."""
    try:
        data = json.loads(job_json_text)
        return data.get("allow_overwrite") is True
    except Exception:
        return False


def job_allow_overwrite(job_json_path):
    """Read allow_overwrite from job.json. Missing/broken file -> False."""
    try:
        if not file_exists(job_json_path):
            return False
        return parse_allow_overwrite(read_all_text(job_json_path))
    except Exception:
        return False


def target_blocks_saveas(new_path, allow_overwrite, exists=None):
    """True if SaveAs must be skipped: target already exists and overwrite
    is not allowed. RSN targets cannot be probed from here (False)."""
    if allow_overwrite or not new_path or is_rsn_path(new_path):
        return False
    return bool((exists or file_exists)(new_path))


def ascii_safe(value):
    """Text for RBP Output() and exception messages: ASCII only, '?' instead of
    anything else (full paths and Cyrillic names belong in the CSV reports)."""
    try:
        text = value if isinstance(value, basestring) else str(value)
    except NameError:  # CPython 3 (tests)
        text = value if isinstance(value, str) else str(value)
    except Exception:
        return "<unprintable>"
    return "".join([ch if ord(ch) < 128 else "?" for ch in text])
