# -*- coding: utf-8 -*-
# Name: workset_plan.py
# Version: 3.1.0
# What it does: See the case README.
# Notes: Paths and project names in this portfolio copy are placeholders.

"""CSV and path helpers for the worksets-and-links pass.

No clr and no Revit: CPython tests import this module, and the IronPython
task script uses the same functions. File bytes (UTF-8 BOM) are written by
read_worksets.py through System.IO.
"""

from __future__ import print_function


READ_HEADER = (
    "row_kind",
    "host",
    "workset_id",
    "workset_name",
    "link_name",
    "link_path",
    "instance_workset",
    "type_workset",
    "is_nested",
    "loaded",
)

PLAN_HEADER = (
    "action",
    "host",
    "key",
    "arg1",
    "arg2",
    "arg3",
    "code",
    "create_workset",
)

ROW_WORKSET = "workset"
ROW_LINK = "link"
ROW_NOT_WORKSHARED = "not_workshared"
ROW_ERROR = "error"

STATUS_EXISTS = "exists"
STATUS_OTHER_PATH = "other_path"
STATUS_INSERT = "insert"


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


def _text_type():
    try:
        return unicode  # IronPython 2.7 / Python 2
    except NameError:
        return str


def as_text(val):
    """Unicode text. None becomes an empty string. Bool becomes 1 or 0."""
    if val is None:
        return u""
    if isinstance(val, bool):
        return u"1" if val else u"0"
    kind = _text_type()
    if isinstance(val, kind):
        return val
    try:
        return kind(val)
    except Exception:
        try:
            return kind(val).decode("utf-8", "replace")
        except Exception:
            return u""


def norm_path(path):
    """Compare key: trim, slashes to '/', drop a trailing slash, lower case."""
    text = as_text(path).strip().replace(u"\\", u"/")
    while text.endswith(u"/") and text not in (u"", u"/"):
        text = text[:-1]
    try:
        return text.lower()
    except Exception:
        return text


def same_path(left, right):
    left_key = norm_path(left)
    right_key = norm_path(right)
    return bool(left_key) and left_key == right_key


def basename_no_ext(path_or_name):
    text = as_text(path_or_name).strip()
    if not text:
        return u""
    name = text.replace(u"\\", u"/").split(u"/")[-1].strip()
    if name.lower().endswith(u".rvt") and len(name) >= 4:
        name = name[:-4]
    return name


def csv_escape(val):
    text = as_text(val).replace(u"\r", u" ").replace(u"\n", u" ")
    if u";" in text or u'"' in text:
        return u'"' + text.replace(u'"', u'""') + u'"'
    return text


def format_row(values):
    return u";".join([csv_escape(v) for v in values])


def header_line(header=None):
    cols = READ_HEADER if header is None else header
    return u";".join([as_text(c) for c in cols])


def read_row(
    kind,
    host,
    workset_id=u"",
    workset_name=u"",
    link_name=u"",
    link_path=u"",
    instance_workset=u"",
    type_workset=u"",
    is_nested=u"",
    loaded=u"",
):
    """One inventory row in READ_HEADER order."""
    return [
        as_text(kind),
        as_text(host),
        as_text(workset_id),
        as_text(workset_name),
        as_text(link_name),
        as_text(link_path),
        as_text(instance_workset),
        as_text(type_workset),
        as_text(is_nested),
        as_text(loaded),
    ]


def split_csv_line(line):
    text = as_text(line)
    out = []
    cur = []
    quoted = False
    i = 0
    n = len(text)
    while i < n:
        ch = text[i]
        if quoted:
            if ch == u'"':
                if i + 1 < n and text[i + 1] == u'"':
                    cur.append(u'"')
                    i += 2
                    continue
                quoted = False
                i += 1
                continue
            cur.append(ch)
            i += 1
            continue
        if ch == u'"':
            quoted = True
            i += 1
            continue
        if ch == u";":
            out.append(u"".join(cur))
            cur = []
            i += 1
            continue
        cur.append(ch)
        i += 1
    out.append(u"".join(cur))
    return out


# ---------------------------------------------------------------- plan checks and log lines
# Actions apply_workset_links.py knows; anything else is a broken plan, never a quiet SKIP.
PLAN_ACTIONS = (u"rename", u"create", u"delete", u"insert", u"skip")
REQUIRED_PLAN_COLUMNS = (u"action", u"host")


def plan_action(row):
    """'rename' / 'create' / 'delete' / 'insert' / 'skip', or u"" for an empty or unknown action."""
    try:
        text = as_text(row.get("action") or u"").strip()
    except Exception:
        return u""
    return text if text in PLAN_ACTIONS else u""


def missing_columns(header, required=REQUIRED_PLAN_COLUMNS):
    """Required column names that are not in the header (exact names, as parse_csv gives them)."""
    have = [as_text(h) for h in (header or [])]
    return [c for c in required if c not in have]


def ascii_escape(text):
    """ASCII for the RBP log: other chars as \\uXXXX (a BOM shows as \\ufeff, not as '?')."""
    out = []
    for ch in as_text(text):
        code = ord(ch)
        if code < 32 or code == 127:
            out.append(u"\\x%02x" % code)
        elif code < 128:
            out.append(ch)
        else:
            out.append(u"\\u%04x" % code)
    return u"".join(out)


def quoted(text):
    return u"'" + ascii_escape(text) + u"'"


def hex_bytes(data, count=16):
    """First bytes as 'ef bb bf 61 ...' (bytes, bytearray, a .NET byte[] or a list of ints)."""
    out = []
    try:
        for i, b in enumerate(data):
            if i >= count:
                break
            if not isinstance(b, int):
                try:
                    b = ord(b)
                except Exception:
                    b = int(b)
            out.append(u"%02x" % (int(b) & 0xFF))
    except Exception:
        return u"?"
    return u" ".join(out)


def header_log_line(header):
    return u"plan header ({0}): [{1}]".format(len(header or []), u", ".join(quoted(h) for h in (header or [])))


def row_log_line(index, row):
    """Keys of one parsed plan row and its action / code, escaped."""
    try:
        keys = sorted(row.keys())
    except Exception:
        keys = []
    return u"row {0}: keys=[{1}] action={2} code={3} known_action={4}".format(
        index,
        u", ".join(quoted(k) for k in keys),
        quoted(row.get("action") if hasattr(row, "get") else u""),
        quoted(row.get("code") if hasattr(row, "get") else u""),
        u"yes" if plan_action(row) else u"NO",
    )


def parse_csv(text, header=None):
    """Parse a semicolon CSV. A leading BOM is ignored. Returns (header, rows)."""
    raw = as_text(text)
    raw = strip_bom(raw)
    lines = raw.replace(u"\r\n", u"\n").replace(u"\r", u"\n").split(u"\n")
    found = None
    rows = []
    for line in lines:
        stripped = line.strip()
        if not stripped or stripped.startswith(u"#"):
            continue
        cells = split_csv_line(stripped)
        if found is None:
            found = [c.strip() for c in cells]
            continue
        row = {}
        for i, key in enumerate(found):
            row[key] = cells[i] if i < len(cells) else u""
        rows.append(row)
    if found is None:
        found = list(header) if header is not None else []
    return found, rows


def rows_for_host(rows, host_path):
    """Plan or report rows whose host is the same path (slash and case ignored)."""
    want = norm_path(host_path)
    found = []
    if not rows or not want:
        return found
    for row in rows:
        host = u""
        try:
            host = row.get("host") or u""
        except Exception:
            host = u""
        if norm_path(host) == want:
            found.append(row)
    return found


def classify_link(link_path, existing):
    """How a link to insert relates to links already in the host.

    existing items are mappings with path, name and nested.
    A nested link does not count. The same full path is "exists".
    The same file name under a different path is "other_path" and must not
    be inserted. Anything else is "insert".
    """
    want = norm_path(link_path)
    want_name = basename_no_ext(link_path)
    other = False
    if not existing:
        return STATUS_INSERT
    for item in existing:
        if item is None:
            continue
        nested = False
        try:
            nested = bool(item.get("nested"))
        except Exception:
            nested = False
        if nested:
            continue
        try:
            path = item.get("path") or u""
            name = item.get("name") or u""
        except Exception:
            path = u""
            name = u""
        if not name:
            name = basename_no_ext(path)
        key = norm_path(path)
        if want and key and want == key:
            return STATUS_EXISTS
        if want_name and name and want_name.lower() == as_text(name).lower() and want != key:
            other = True
    if other:
        return STATUS_OTHER_PATH
    return STATUS_INSERT


def _coord_why(coords):
    if coords:
        return "shared_coordinates"
    return "dialog_cancelled"


def insert_before_commit(coords, dialog_cancelled):
    """What to do after the link API returned and before Commit.

    A dialog or failure flag set inside the insert cannot be undone once
    Commit has returned Committed. Roll the open transaction back instead.
    Returns (action, result, code, why). action is "rollback" or "commit".
    """
    if coords or dialog_cancelled:
        return ("rollback", "SKIP", "SKIP_COORDS", _coord_why(coords))
    return ("commit", "OK", "", "")


def insert_after_commit(committed, coords, dialog_cancelled, link_present, removed):
    """Report row after Commit. A Committed transaction can no longer be rolled back.

    If a flag is set and the new instance or type is still in the model,
    the caller has already tried to delete both in a later transaction.
    removed=True means both are gone (SKIP_COORDS, why removed_after_commit).
    removed=False means the link stayed (ERROR_LINK, why left_in_model).
    Returns (result, code, why). result "OK" means the insert stays.
    """
    flagged = bool(coords) or bool(dialog_cancelled)
    if not committed:
        if flagged:
            return ("SKIP", "SKIP_COORDS", _coord_why(coords))
        return ("ERROR", "ERROR_LINK", "place_failed")
    if not flagged:
        return ("OK", "", "")
    if link_present and removed:
        return ("SKIP", "SKIP_COORDS", "removed_after_commit")
    if link_present:
        return ("ERROR", "ERROR_LINK", "left_in_model")
    return ("SKIP", "SKIP_COORDS", _coord_why(coords))


def pin_note(reason):
    """Report message when Pinned = True failed: 'not_pinned: <reason>' (one line, <= 200 chars)."""
    text = as_text(reason)
    for ch in (u"\r", u"\n", u"\t", u";"):
        text = text.replace(ch, u" ")
    text = u" ".join(text.split()) or u"unknown"
    return (u"not_pinned: " + text)[:200]


def with_pin_note(result, why, pin_error):
    """An OK insert keeps result OK; a pin failure is added to the message.
    pin_error is "" / None when the link was pinned (or the insert did not stay)."""
    if result != "OK" or not pin_error:
        return why
    note = pin_note(pin_error)
    return note if not why else as_text(why) + u"; " + note


def delete_outcome(api, mode, can_delete, borrowed, moved, deleted):
    """Status of one workset delete. api is whether DeleteWorkset exists.

    mode is "move" (elements go to another workset) or "elements".
    On a Revit build without the API the workset stays and the code is SKIP_NO_API.
    Returns (result, code, why).
    """
    if borrowed:
        return ("SKIP", "SKIP_BORROW", "")
    if not api:
        if moved:
            return ("SKIP", "SKIP_NO_API", "workset_remains")
        return ("ERROR", "ERROR_DELETE", "move_failed")
    if not can_delete:
        return ("SKIP", "SKIP_CANNOT_DELETE", "")
    if deleted:
        if mode == "elements":
            return ("OK", "", "deleted_with_elements")
        return ("OK", "", "moved_to_default")
    return ("ERROR", "ERROR_DELETE", "delete_failed")
