# -*- coding: utf-8 -*-
# Name: relink_prefix.py
# Version: 3.1.0
# What it does: See the case README.
# Notes: Paths and project names in this portfolio copy are placeholders.

"""Folder prefix remap for update_rvt_links.py (IronPython 2.7 / CPython 3).

relink_map.csv (UTF-8, header old_prefix;new_prefix):
    RSN://revit-server-2022.example.local/proj SampleSite/Стадия ПП;RSN://revit-server-2022.example.local/proj SampleSite/Стадия П
A link is remapped only when its path is the old folder itself or lies under it
(folder boundary: 'Стадия П' never matches 'Стадия ПП/...').
"""
from __future__ import unicode_literals

import io
import os

MAP_NAME = "relink_map.csv"


def _u(s):
    if s is None:
        return ""
    try:
        return unicode(s)  # noqa: F821 (IronPython 2.7)
    except NameError:
        return str(s)


def norm(p):
    """Slash-normalized, trimmed, no trailing slash; case kept."""
    s = _u(p).strip().replace("\\", "/")
    while s.endswith("/") and not s.endswith("://"):
        s = s[:-1]
    return s


def _under(path_l, folder_l):
    return path_l == folder_l or path_l.startswith(folder_l + "/")


def load_map(path):
    """List of (old_prefix, new_prefix); [] when file missing/empty."""
    pairs = []
    if not path or not os.path.isfile(path):
        return pairs
    with io.open(path, "r", encoding="utf-8-sig") as f:
        for i, line in enumerate(f):
            line = line.strip()
            if not line or i == 0 and line.lower().startswith("old_prefix"):
                continue
            parts = line.split(";")
            if len(parts) < 2:
                continue
            old, new = norm(parts[0].strip('"')), norm(parts[1].strip('"'))
            if old and new and old.lower() != new.lower():
                pairs.append((old, new))
    return pairs


def _backslash_style(path):
    """True for UNC / drive paths stored with backslashes (RSN:// always uses '/')."""
    s = _u(path).strip()
    return "\\" in s and not s.upper().startswith("RSN://")


def remap(old_path, pairs):
    """Return (new_path, reason): reason = remap | already | outside | empty.

    The new path keeps the slash style of the stored path: a UNC or drive path
    stays with backslashes, an RSN:// path keeps forward slashes.
    """
    p = norm(old_path)
    if not p:
        return None, "empty"
    pl = p.lower()
    for old, new in pairs:
        if _under(pl, new.lower()):
            return None, "already"
    for old, new in pairs:
        if _under(pl, old.lower()):
            result = new + p[len(old):]
            if _backslash_style(old_path):
                result = result.replace("/", "\\")
            return result, "remap"
    return None, "outside"


REASON_TEXT = {
    "already": "already new path",
    "outside": "outside old folder",
    "empty": "no stored path",
}
