# -*- coding: utf-8 -*-
# Name: test_workset_plan.py
# Version: 3.1.0
# What it does: See the case README.
# Notes: Paths and project names in this portfolio copy are placeholders.

"""CSV parse, path identity, and the read-only shape of read_worksets.py.

CPython, no Revit:

    python3 -I tests/test_workset_plan.py
"""
from __future__ import unicode_literals

import ast
import os
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
CODE = os.path.dirname(HERE)
import sys

sys.path.insert(0, CODE)

import workset_plan as wp  # noqa: E402


class Paths(unittest.TestCase):
    def test_norm_ignores_case_slashes_and_trailing_slash(self):
        self.assertEqual(
            wp.norm_path(r"RSN://Host/Folder/Model.rvt"),
            wp.norm_path("rsn://host/folder/model.rvt/"),
        )
        self.assertTrue(wp.same_path(r"\\Srv\Share\A.rvt", "//srv/share/a.rvt"))

    def test_empty_is_not_the_same_path(self):
        self.assertEqual(wp.norm_path(None), "")
        self.assertEqual(wp.norm_path("   "), "")
        self.assertFalse(wp.same_path("", ""))
        self.assertFalse(wp.same_path(None, r"\\srv\a.rvt"))

    def test_basename(self):
        self.assertEqual(
            wp.basename_no_ext(r"RSN://h/Папка/PROJ_IOS_VK_ORG_K01_R24.rvt"),
            "PROJ_IOS_VK_ORG_K01_R24",
        )


class Classify(unittest.TestCase):
    def test_same_path_is_exists_even_when_slash_style_differs(self):
        existing = [{"path": r"RSN://Srv/A/VK.rvt", "name": "", "nested": False}]
        self.assertEqual(wp.classify_link("rsn://srv/a/vk.rvt", existing), wp.STATUS_EXISTS)

    def test_same_file_other_path_is_not_inserted(self):
        existing = [{
            "path": r"\\srv\old\PROJ_IOS_VK_ORG_K01_R24.rvt",
            "name": "",
            "nested": False,
        }]
        status = wp.classify_link(r"RSN://srv/new/PROJ_IOS_VK_ORG_K01_R24.rvt", existing)
        self.assertEqual(status, wp.STATUS_OTHER_PATH)

    def test_nested_same_path_does_not_count(self):
        existing = [{"path": r"RSN://srv/a/VK.rvt", "name": "VK", "nested": True}]
        self.assertEqual(
            wp.classify_link(r"RSN://srv/a/VK.rvt", existing),
            wp.STATUS_INSERT,
        )

    def test_exact_path_wins_over_another_copy(self):
        existing = [
            {"path": r"\\other\PROJ_IOS_VK_ORG_K01_R24.rvt", "nested": False},
            {"path": r"RSN://srv/a/PROJ_IOS_VK_ORG_K01_R24.rvt", "nested": False},
        ]
        self.assertEqual(
            wp.classify_link(r"rsn://srv/a/PROJ_IOS_VK_ORG_K01_R24.rvt", existing),
            wp.STATUS_EXISTS,
        )

    def test_new_link(self):
        self.assertEqual(wp.classify_link(r"RSN://srv/a/NEW.rvt", []), wp.STATUS_INSERT)


class InsertStatus(unittest.TestCase):
    def test_flag_before_commit_rolls_back_instead_of_skipping_a_live_link(self):
        self.assertEqual(
            wp.insert_before_commit(False, True),
            ("rollback", "SKIP", "SKIP_COORDS", "dialog_cancelled"),
        )
        self.assertEqual(
            wp.insert_before_commit(True, False),
            ("rollback", "SKIP", "SKIP_COORDS", "shared_coordinates"),
        )
        self.assertEqual(wp.insert_before_commit(False, False), ("commit", "OK", "", ""))

    def test_committed_with_flag_is_not_a_skip_while_the_link_stays(self):
        self.assertEqual(
            wp.insert_after_commit(True, True, False, True, True),
            ("SKIP", "SKIP_COORDS", "removed_after_commit"),
        )
        self.assertEqual(
            wp.insert_after_commit(True, False, True, True, False),
            ("ERROR", "ERROR_LINK", "left_in_model"),
        )

    def test_clean_commit_and_rolled_back_commit_stay_honest(self):
        self.assertEqual(wp.insert_after_commit(True, False, False, True, False), ("OK", "", ""))
        self.assertEqual(
            wp.insert_after_commit(False, True, False, False, False),
            ("SKIP", "SKIP_COORDS", "shared_coordinates"),
        )
        self.assertEqual(
            wp.insert_after_commit(False, False, False, False, False),
            ("ERROR", "ERROR_LINK", "place_failed"),
        )
        self.assertEqual(
            wp.insert_after_commit(True, True, True, False, False),
            ("SKIP", "SKIP_COORDS", "shared_coordinates"),
        )


class PinNote(unittest.TestCase):
    def test_ok_insert_keeps_ok_and_gets_note(self):
        self.assertEqual(wp.with_pin_note("OK", "", "Element is in a group"), "not_pinned: Element is in a group")
        self.assertEqual(wp.with_pin_note("OK", "already", "x"), "already; not_pinned: x")

    def test_pinned_or_not_ok_has_no_note(self):
        self.assertEqual(wp.with_pin_note("OK", "", ""), "")
        self.assertEqual(wp.with_pin_note("OK", "", None), "")
        self.assertEqual(wp.with_pin_note("SKIP", "removed_after_commit", "x"), "removed_after_commit")
        self.assertEqual(wp.with_pin_note("ERROR", "left_in_model", "x"), "left_in_model")

    def test_note_is_one_short_csv_safe_line(self):
        note = wp.pin_note("a;b\r\nc\t d   " + "z" * 300)
        self.assertTrue(note.startswith("not_pinned: a b c d "), note)
        self.assertEqual(len(note), 200)
        self.assertNotIn(";", note)
        self.assertEqual(wp.pin_note("  "), "not_pinned: unknown")
        header, rows = wp.parse_csv(wp.format_row(["h", "insert", "", "t", "OK", "", note]))
        self.assertEqual(header[-1], note)


class DeleteStatus(unittest.TestCase):
    def test_old_revit_only_moves_elements(self):
        self.assertEqual(
            wp.delete_outcome(False, "move", False, False, True, False),
            ("SKIP", "SKIP_NO_API", "workset_remains"),
        )
        self.assertEqual(
            wp.delete_outcome(False, "elements", False, False, False, False),
            ("ERROR", "ERROR_DELETE", "move_failed"),
        )

    def test_api_precheck_and_result(self):
        self.assertEqual(
            wp.delete_outcome(True, "move", False, True, False, False),
            ("SKIP", "SKIP_BORROW", ""),
        )
        self.assertEqual(
            wp.delete_outcome(True, "elements", False, False, False, False),
            ("SKIP", "SKIP_CANNOT_DELETE", ""),
        )
        self.assertEqual(
            wp.delete_outcome(True, "elements", True, False, False, True),
            ("OK", "", "deleted_with_elements"),
        )
        self.assertEqual(
            wp.delete_outcome(True, "move", True, False, False, True),
            ("OK", "", "moved_to_default"),
        )
        self.assertEqual(
            wp.delete_outcome(True, "move", True, False, False, False),
            ("ERROR", "ERROR_DELETE", "delete_failed"),
        )


class Csv(unittest.TestCase):
    def test_read_header_matches_the_plan(self):
        self.assertEqual(
            wp.READ_HEADER,
            (
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
            ),
        )
        self.assertEqual(wp.header_line().split(";"), list(wp.READ_HEADER))

    def test_roundtrip_keeps_cyrillic_type_and_instance_worksets(self):
        host = "RSN://srv/Папка/PROJ_IOS_AR_ORG_K01_R24.rvt"
        rows = [
            wp.read_row(wp.ROW_WORKSET, host, "42", "30_ИОС_Элементы"),
            wp.read_row(wp.ROW_WORKSET, host, "7", "00_Связь_ОВ"),
            wp.read_row(
                wp.ROW_LINK,
                host,
                "",
                "",
                "PROJ_IOS_VK_ORG_K01_R24.rvt",
                "RSN://srv/Папка/PROJ_IOS_VK_ORG_K01_R24.rvt",
                "00_Связь_ВК",
                "00_Связь_ВК",
                "0",
                "1",
            ),
            wp.read_row(wp.ROW_NOT_WORKSHARED, host),
        ]
        text = u"\ufeff" + wp.header_line() + u"\n" + u"\n".join(wp.format_row(r) for r in rows) + u"\n"
        header, parsed = wp.parse_csv(text)
        self.assertEqual(tuple(header), wp.READ_HEADER)
        self.assertEqual(parsed[0]["workset_name"], "30_ИОС_Элементы")
        self.assertEqual(parsed[0]["workset_id"], "42")
        self.assertEqual(parsed[2]["instance_workset"], "00_Связь_ВК")
        self.assertEqual(parsed[2]["type_workset"], "00_Связь_ВК")
        self.assertEqual(parsed[2]["is_nested"], "0")
        self.assertEqual(parsed[2]["loaded"], "1")
        self.assertEqual(parsed[3]["row_kind"], "not_workshared")
        self.assertEqual(parsed[2]["link_path"], "RSN://srv/Папка/PROJ_IOS_VK_ORG_K01_R24.rvt")

    def test_quotes_and_semicolon(self):
        line = wp.format_row(["skip", "h", "k", "a;b", 'say "hi"', "", "SKIP_OTHER_PATH", "0"])
        cells = wp.split_csv_line(line)
        self.assertEqual(cells[3], "a;b")
        self.assertEqual(cells[4], 'say "hi"')

    def test_plan_header_roundtrip(self):
        text = wp.header_line(wp.PLAN_HEADER) + "\n" + wp.format_row(
            ["insert", "RSN://h/a.rvt", "row-1", "RSN://h/b.rvt", "00_Связь_ВК", "VK", "", "1"]
        )
        header, rows = wp.parse_csv(text)
        self.assertEqual(tuple(header), wp.PLAN_HEADER)
        self.assertEqual(rows[0]["create_workset"], "1")
        self.assertEqual(rows[0]["arg2"], "00_Связь_ВК")
        self.assertEqual(rows[0]["code"], "")


class RemotePlanFixture(unittest.TestCase):
    """workset_plan.csv from the remote PC (2026-11-11): the run skipped all 4 rows."""

    HOST = u"RSN://revit-server-2022.example.local/proj SampleSite/Стадия П/04_ЭОМ/SITE_ORG_EOM_K1_R22.rvt"

    def setUp(self):
        path = os.path.join(CODE, "tests", "fixtures", "workset_plan_remote_2026-11-11.csv")
        with open(path, "rb") as handle:
            self.raw = handle.read()

    def test_bytes_are_the_users_file(self):
        self.assertEqual(wp.hex_bytes(self.raw, 10), u"ef bb bf 61 63 74 69 6f 6e 3b")
        self.assertEqual(self.raw.count(b"\xef\xbb\xbf"), 1)

    def test_parses_with_and_without_the_bom(self):
        # System.IO.File.ReadAllText drops the BOM; a plain utf-8 decode keeps it
        for text in (self.raw.decode("utf-8"), self.raw.decode("utf-8-sig")):
            header, rows = wp.parse_csv(text)
            self.assertEqual(tuple(header), wp.PLAN_HEADER)
            self.assertEqual(wp.missing_columns(header), [])
            mine = wp.rows_for_host(rows, self.HOST)
            self.assertEqual(len(mine), 4)
            self.assertEqual([wp.plan_action(r) for r in mine], ["create", "create", "insert", "insert"])
            # code is empty in this plan; HC / AVT are arg3 (section)
            self.assertEqual([r["code"] for r in mine], ["", "", "", ""])
            self.assertEqual([r["arg3"] for r in mine], ["", "", "HC", "AVT"])

    def test_log_lines_are_ascii_and_name_the_keys(self):
        header, rows = wp.parse_csv(self.raw.decode("utf-8"))
        line = wp.header_log_line(header)
        self.assertTrue(all(ord(c) < 128 for c in line), line)
        self.assertIn(u"'action'", line)
        row_line = wp.row_log_line(1, rows[0])
        self.assertTrue(all(ord(c) < 128 for c in row_line), row_line)
        self.assertIn(u"action='create'", row_line)
        self.assertIn(u"known_action=yes", row_line)
        self.assertIn(u"'code'", row_line)

    def test_wrong_decoding_is_visible_in_the_log(self):
        # bytes read as cp1251: the BOM sticks to the first name and the log shows it escaped
        header, rows = wp.parse_csv(self.raw.decode("cp1251"))
        self.assertEqual(wp.missing_columns(header), ["action"])
        self.assertIn(u"\\u043f\\u00bb\\u0457action", wp.header_log_line(header))
        self.assertIn(u"known_action=NO", wp.row_log_line(1, rows[0]))


# .NET chars that culture-sensitive comparison ignores (BOM, zero-width, soft hyphen)
IGNORABLE = "\ufeff\u200b\u200c\u200d\u2060\u00ad"


def _drop_ignorable(text):
    return "".join(c for c in text if c not in IGNORABLE)


class CultureStr(str):
    """str as IronPython 2.7.3 sees it: startswith/endswith go through
    culture-sensitive .NET StartsWith/EndsWith, so ignorable chars do not count
    (u"action".startswith(u"\\ufeff") is True). Other str methods stay ordinal."""

    def startswith(self, prefix, *args):
        if args or isinstance(prefix, tuple):
            return str.startswith(self, prefix, *args)
        return _drop_ignorable(self).startswith(_drop_ignorable(prefix))

    def endswith(self, suffix, *args):
        if args or isinstance(suffix, tuple):
            return str.endswith(self, suffix, *args)
        return _drop_ignorable(self).endswith(_drop_ignorable(suffix))


class IronPython273Bom(unittest.TestCase):
    """Remote PC, RBP IronPython 2.7.3: the header came out as 'ction'."""

    HOST = RemotePlanFixture.HOST

    def setUp(self):
        path = os.path.join(CODE, "tests", "fixtures", "workset_plan_remote_2026-11-11.csv")
        with open(path, "rb") as handle:
            self.raw = handle.read()

    def test_the_emulation_reproduces_the_bug(self):
        text = CultureStr(self.raw.decode("utf-8-sig"))
        self.assertTrue(text.startswith("\ufeff"))       # what 2.7.3 answered
        self.assertFalse(str(text).startswith("\ufeff"))
        old = text[1:] if text.startswith("\ufeff") else text   # the 1.1.1 check
        self.assertTrue(old.startswith("ction;host;"), old[:12])

    def test_parse_csv_keeps_action_with_and_without_bom(self):
        # ReadAllText already drops the BOM (the PC case); utf-8 keeps it
        for text in (self.raw.decode("utf-8-sig"), self.raw.decode("utf-8"), "\ufeff" + self.raw.decode("utf-8")):
            header, rows = wp.parse_csv(CultureStr(text))
            self.assertEqual(header[0], "action", repr(header[:2]))
            self.assertEqual(tuple(header), wp.PLAN_HEADER)
            self.assertEqual(wp.missing_columns(header), [])
            mine = wp.rows_for_host(rows, self.HOST)
            self.assertEqual([wp.plan_action(r) for r in mine], ["create", "create", "insert", "insert"])

    def test_strip_bom_is_ordinal(self):
        self.assertEqual(wp.strip_bom(CultureStr("action;host")), "action;host")
        self.assertEqual(wp.strip_bom(CultureStr("\ufeffaction")), "action")
        self.assertEqual(wp.strip_bom("\ufeff\ufeffaction"), "action")
        self.assertEqual(wp.strip_bom("\u200baction"), "\u200baction")
        self.assertEqual(wp.strip_bom("\ufeff"), "")
        self.assertEqual(wp.strip_bom(""), "")
        self.assertIsNone(wp.strip_bom(None))


class NoCultureBomChecks(unittest.TestCase):
    """Task scripts and helpers in служебное: BOM / zero-width chars are never
    checked with startswith/endswith (culture-sensitive on IronPython 2.7.3)."""

    IGNORABLE_ESC = ("\\ufeff", "\\u200b", "\\u200c", "\\u200d", "\\u2060", "\\u00ad", "\\xad")

    def sources(self):
        for name in sorted(os.listdir(CODE)):
            if name.endswith(".py"):
                with open(os.path.join(CODE, name), "rb") as handle:
                    yield name, handle.read().decode("utf-8")

    def test_no_startswith_endswith_with_ignorable_chars(self):
        found = []
        for name, src in self.sources():
            tree = ast.parse(src)
            for node in ast.walk(tree):
                if not (isinstance(node, ast.Call) and isinstance(node.func, ast.Attribute)):
                    continue
                if node.func.attr not in ("startswith", "endswith"):
                    continue
                arg_src = " ".join(ast.get_source_segment(src, a) or "" for a in node.args)
                values = [n.value for a in node.args for n in ast.walk(a)
                          if isinstance(n, ast.Constant) and isinstance(n.value, str)]
                names = [n.id for a in node.args for n in ast.walk(a) if isinstance(n, ast.Name)]
                if (any(c in IGNORABLE for v in values for c in v)
                        or any(e in arg_src for e in self.IGNORABLE_ESC) or "BOM" in names):
                    found.append("%s:%d %s" % (name, node.lineno, ast.get_source_segment(src, node)))
        self.assertEqual(found, [])

    def test_bom_literal_only_in_the_bom_constant(self):
        found = []
        for name, src in self.sources():
            for i, line in enumerate(src.splitlines(), 1):
                code = line.split("#", 1)[0]
                if "\\ufeff\"" in code or "\ufeff" in code:
                    if code.strip() != 'BOM = u"\\ufeff"':
                        found.append("%s:%d %s" % (name, i, line.strip()))
        self.assertEqual(found, [])

    def test_bom_readers_use_strip_bom(self):
        for name in ("apply_workset_links.py", "read_worksets.py"):
            with open(os.path.join(CODE, name), "rb") as handle:
                src = handle.read().decode("utf-8")
            self.assertIn("current = lib.strip_bom(lib.read_all_text(pointer).strip()).strip()", src, name)
        with open(os.path.join(CODE, "ops_lib.py"), "rb") as handle:
            src = handle.read().decode("utf-8")
        self.assertIn("return strip_bom(System.IO.File.ReadAllText(path, System.Text.Encoding.UTF8))", src)
        self.assertIn("text = strip_bom(text)", src)
        with open(os.path.join(CODE, "workset_plan.py"), "rb") as handle:
            self.assertIn("raw = strip_bom(raw)", handle.read().decode("utf-8"))


class PlanChecks(unittest.TestCase):
    def test_plan_action(self):
        self.assertEqual(wp.plan_action({"action": "insert"}), "insert")
        self.assertEqual(wp.plan_action({"action": " skip "}), "skip")
        self.assertEqual(wp.plan_action({"action": ""}), "")
        self.assertEqual(wp.plan_action({}), "")
        self.assertEqual(wp.plan_action({"action": "вставить + новый набор"}), "")

    def test_missing_columns(self):
        self.assertEqual(wp.missing_columns(["action", "host", "key"]), [])
        self.assertEqual(wp.missing_columns([u"\ufeffaction", "host"]), ["action"])
        self.assertEqual(wp.missing_columns([]), ["action", "host"])

    def test_ascii_escape(self):
        self.assertEqual(wp.ascii_escape(u"\ufeffa\tb"), u"\\ufeffa\\x09b")
        self.assertEqual(wp.ascii_escape(u"ЭОМ"), u"\\u042d\\u041e\\u041c")

    def test_hex_bytes(self):
        self.assertEqual(wp.hex_bytes(b"\xef\xbb\xbfa", 16), u"ef bb bf 61")
        self.assertEqual(wp.hex_bytes(bytearray(b"abc"), 2), u"61 62")
        self.assertEqual(wp.hex_bytes([239, 187], 16), u"ef bb")


class HostFilter(unittest.TestCase):
    def test_rows_for_host_ignores_slash_and_case(self):
        rows = [
            {"host": r"RSN://Srv/A/Model.rvt", "action": "rename"},
            {"host": r"RSN://other/Model.rvt", "action": "create"},
        ]
        mine = wp.rows_for_host(rows, "rsn://srv/a/model.rvt/")
        self.assertEqual(len(mine), 1)
        self.assertEqual(mine[0]["action"], "rename")


class ApplyScript(unittest.TestCase):
    def setUp(self):
        path = os.path.join(CODE, "apply_workset_links.py")
        with open(path, "rb") as handle:
            self.source = handle.read().decode("utf-8-sig")
        self.tree = ast.parse(self.source, "apply_workset_links.py")

    def test_diagnostics_and_loud_errors(self):
        src = self.source
        for need in ("log_environment()", "log_plan_file(path, text)", "wp.header_log_line(header)",
                     "wp.row_log_line(", "wp.hex_bytes(", "diag: __file__", "diag: SCRIPT_DIR", "diag: CODE_DIR",
                     "diag: wp", "diag: lib", "diag: sys.path", "diag: preloaded"):
            self.assertIn(need, src, need)
        self.assertIn("wp.missing_columns(header)", src)
        self.assertIn("ERROR_BAD_PLAN", src)
        self.assertIn("ERROR_BAD_ACTION", src)
        self.assertIn("action = wp.plan_action(row)", src)
        self.assertIn("sys.path.insert(_front, _p)", src)
        self.assertNotIn("sys.path.append(", src)

    def test_unknown_action_is_an_error(self):
        # the rest branch: only a real 'skip' row stays SKIP; anything else counts as an error
        tree = self.tree
        fn = [n for n in tree.body if isinstance(n, ast.FunctionDef) and n.name == "apply_host"][0]
        text = ast.get_source_segment(self.source, fn)
        rest = text[text.index("for row in rest:"):]
        self.assertIn('wp.plan_action(row) == u"skip"', rest)
        self.assertIn("errors += 1", rest)
        self.assertIn('"ERROR", "ERROR_BAD_ACTION"', rest)

    def _fn(self, name, body=None):
        body = self.tree.body if body is None else body
        for node in body:
            if isinstance(node, ast.FunctionDef) and node.name == name:
                return node
        raise AssertionError(name)

    def test_inserted_link_is_pinned_in_the_same_transaction(self):
        insert = self._fn("apply_insert")
        do_insert = self._fn("do_insert", insert.body)
        stmts = do_insert.body
        create = [i for i, s in enumerate(stmts) if "RevitLinkInstance.Create(" in ast.get_source_segment(self.source, s)
                  and "ImportPlacement.Shared" in ast.get_source_segment(self.source, s)]
        part = [i for i, s in enumerate(stmts) if ast.get_source_segment(self.source, s) == "_set_partition(inst, ws)"]
        pin = [i for i, s in enumerate(stmts) if isinstance(s, ast.Try)
               and any(isinstance(b, ast.Assign) and ast.get_source_segment(self.source, b) == "inst.Pinned = True"
                       for b in s.body)]
        self.assertEqual(len(create), 1)
        self.assertEqual(len(part), 1)
        self.assertEqual(len(pin), 1, "inst.Pinned = True inside try in do_insert")
        self.assertLess(create[0], part[0])
        self.assertLess(part[0], pin[0])
        # the pin failure does not raise (no rollback): the handler only records the reason
        handler = stmts[pin[0]].handlers[0]
        self.assertFalse(any(isinstance(n, ast.Raise) for n in ast.walk(handler)))
        self.assertIn("pin_error[0] =", ast.get_source_segment(self.source, handler))
        # do_insert runs between txn.Start() and txn.Commit() of the single ws-insert transaction
        text = ast.get_source_segment(self.source, insert)
        self.assertEqual(text.count("Transaction("), 1)
        self.assertLess(text.index("txn.Start()"), text.index("        do_insert()"))
        self.assertLess(text.index("        do_insert()"), text.index("txn.Commit()"))
        after = ast.get_source_segment(self.source, self._fn("after_commit", insert.body))
        self.assertIn("wp.with_pin_note(result, why, pin_error[0])", after)
        self.assertIn("not pinned", after)
        self.assertEqual(self.source.count(".Pinned = True"), 1)
        self.assertNotIn(".Pinned = False", self.source)

    def test_apply_renames_and_creates_without_saving_as(self):
        banned = {"Save", "SaveAs", "EnableWorksharing"}
        seen = set()
        for node in ast.walk(self.tree):
            if isinstance(node, ast.Call) and isinstance(node.func, ast.Attribute):
                seen.add(node.func.attr)
                self.assertNotIn(node.func.attr, banned, node.func.attr)
        self.assertIn("RenameWorkset", seen)
        self.assertIn("Create", seen)
        self.assertIn("SynchronizeWithCentral", seen)
        self.assertNotIn("EnableWorksharing", self.source)
        self.assertIn("ImportPlacement.Shared", self.source)
        self.assertNotIn("ImportPlacement.Origin", self.source)
        self.assertNotIn("SKIP_NEXT", self.source)
        self.assertIn("SKIP_EXISTS", self.source)
        self.assertIn("SKIP_OTHER_PATH", self.source)
        self.assertIn("SKIP_NOT_FOUND", self.source)
        self.assertIn("ERROR_LINK", self.source)
        self.assertIn("insert_before_commit", self.source)
        self.assertIn("insert_after_commit", self.source)
        self.assertIn("CanDeleteWorkset", self.source)
        self.assertIn("DeleteWorkset", self.source)
        plan_path = os.path.join(CODE, "workset_plan.py")
        with open(plan_path, "rb") as handle:
            plan = handle.read().decode("utf-8-sig")
        self.assertIn("SKIP_COORDS", plan)
        self.assertIn("removed_after_commit", plan)
        self.assertIn("left_in_model", plan)
        self.assertIn("SKIP_NO_API", plan)
        self.assertIn("workset_remains", plan)


class ReadScript(unittest.TestCase):
    def setUp(self):
        path = os.path.join(CODE, "read_worksets.py")
        with open(path, "rb") as handle:
            self.source = handle.read().decode("utf-8-sig")
        self.tree = ast.parse(self.source, "read_worksets.py")

    def test_module_has_no_clr_in_the_plan_helper(self):
        plan = os.path.join(CODE, "workset_plan.py")
        with open(plan, "rb") as handle:
            text = handle.read().decode("utf-8-sig")
        self.assertNotIn("import clr", text)
        self.assertNotIn("RevitAPI", text)

    def test_read_script_does_not_modify_the_model(self):
        banned_attr = {
            "Save",
            "SaveAs",
            "SynchronizeWithCentral",
            "Set",
            "Create",
            "RenameWorkset",
        }
        for node in ast.walk(self.tree):
            if not isinstance(node, ast.Call):
                continue
            func = node.func
            if isinstance(func, ast.Name):
                self.assertNotEqual(func.id, "Transaction", "Transaction() is not a read")
            if isinstance(func, ast.Attribute):
                self.assertNotIn(
                    func.attr,
                    banned_attr,
                    "read_worksets.py calls {0}".format(func.attr),
                )
        self.assertNotIn("EnableWorksharing", self.source)
        self.assertIn("ToWorksets", self.source)
        self.assertIn("ROW_NOT_WORKSHARED", self.source)
        self.assertIn("instance_workset", wp.header_line())
        self.assertIn("type_workset", wp.header_line())


if __name__ == "__main__":
    unittest.main()
