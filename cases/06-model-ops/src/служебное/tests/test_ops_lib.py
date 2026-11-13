# -*- coding: utf-8 -*-
# Name: test_ops_lib.py
# Version: 3.1.0
# What it does: See the case README.
# Notes: Paths and project names in this portfolio copy are placeholders.

"""Tests for the overwrite guards in ops_lib.py (run in CPython, no Revit).

    python3 -I tests/test_ops_lib.py

ops_lib imports clr/System (IronPython); they are stubbed here because the
tested functions only use plain Python.
"""
import os
import sys
import types
import unittest

stub_clr = types.ModuleType("clr")
stub_clr.AddReference = lambda *a, **k: None
sys.modules["clr"] = stub_clr
sys.modules["System"] = types.ModuleType("System")
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import ops_lib as lib  # noqa: E402


class AllowOverwrite(unittest.TestCase):
    def test_default_false(self):
        self.assertFalse(lib.parse_allow_overwrite("{}"))
        self.assertFalse(lib.parse_allow_overwrite('{"relink": true}'))

    def test_true_only_for_json_true(self):
        self.assertTrue(lib.parse_allow_overwrite('{"allow_overwrite": true}'))
        self.assertFalse(lib.parse_allow_overwrite('{"allow_overwrite": false}'))
        self.assertFalse(lib.parse_allow_overwrite('{"allow_overwrite": "true"}'))
        self.assertFalse(lib.parse_allow_overwrite('{"allow_overwrite": 1}'))

    def test_broken_json_is_false(self):
        self.assertFalse(lib.parse_allow_overwrite(""))
        self.assertFalse(lib.parse_allow_overwrite("{not json"))
        self.assertFalse(lib.parse_allow_overwrite("[1, 2]"))

    def test_powershell_pretty_json(self):
        text = '{\r\n  "version":  "1.0.0",\r\n  "allow_overwrite":  true\r\n}'
        self.assertTrue(lib.parse_allow_overwrite(text))


class TargetBlocks(unittest.TestCase):
    def test_existing_target_blocks(self):
        self.assertTrue(lib.target_blocks_saveas("X:/a/b.rvt", False, lambda p: True))

    def test_missing_target_does_not_block(self):
        self.assertFalse(lib.target_blocks_saveas("X:/a/b.rvt", False, lambda p: False))

    def test_allow_overwrite_never_blocks(self):
        self.assertFalse(lib.target_blocks_saveas("X:/a/b.rvt", True, lambda p: True))

    def test_rsn_and_empty_are_not_probed(self):
        def boom(p):
            raise AssertionError("must not probe")

        self.assertFalse(lib.target_blocks_saveas("RSN://h/p/a.rvt", False, boom))
        self.assertFalse(lib.target_blocks_saveas("", False, boom))


class AsciiSafe(unittest.TestCase):
    def test_ascii_unchanged(self):
        self.assertEqual(lib.ascii_safe("Remap: a -> b OK"), "Remap: a -> b OK")

    def test_non_ascii_replaced(self):
        got = lib.ascii_safe("RSN://h/proj \u0421\u0442\u0440\u043e\u0433\u0438\u043d\u043e/x.rvt")
        self.assertEqual(got, "RSN://h/proj ????????/x.rvt")
        self.assertEqual(lib.ascii_safe("a \u2014 b"), "a ? b")

    def test_non_strings(self):
        self.assertEqual(lib.ascii_safe(5), "5")
        self.assertEqual(lib.ascii_safe(None), "None")
        self.assertEqual(lib.ascii_safe(ValueError("\u043e\u0448\u0438\u0431\u043a\u0430")), "??????")

    def test_result_is_always_ascii(self):
        text = "".join(chr(i) for i in range(0, 600))
        self.assertTrue(all(ord(c) < 128 for c in lib.ascii_safe(text)))


class ReportsDir(unittest.TestCase):
    def test_folder_name_is_otchety(self):
        self.assertEqual(lib.REPORTS_FOLDER_NAME, u"\u043e\u0442\u0447\u0451\u0442\u044b")

    def test_reports_dir_from_job_text(self):
        self.assertEqual(lib.reports_dir_from_job_text(u'\ufeff{"reports_dir": "C:\\\\t\\\\r"}'), u"C:\\t\\r")
        self.assertEqual(lib.reports_dir_from_job_text(u'{"relink": true}'), "")
        self.assertEqual(lib.reports_dir_from_job_text(u"{broken"), "")
        self.assertEqual(lib.reports_dir_from_job_text(u"[1]"), "")

    def test_tool_reports_dir_default_and_job(self):
        import tempfile, shutil
        d = tempfile.mkdtemp()
        old_exists = lib.file_exists
        old_read = lib.read_all_text
        try:
            lib.file_exists = os.path.isfile
            lib.read_all_text = lambda p: open(p, encoding="utf-8").read()
            self.assertEqual(lib.tool_reports_dir(d), os.path.join(d, lib.REPORTS_FOLDER_NAME))
            with open(os.path.join(d, "workset_job.json"), "w", encoding="utf-8") as f:
                f.write(u'{"mode": "apply", "reports_dir": "X:/r"}')
            self.assertEqual(lib.tool_reports_dir(d), u"X:/r")
        finally:
            lib.file_exists = old_exists
            lib.read_all_text = old_read
            shutil.rmtree(d)



class StripBom(unittest.TestCase):
    class CultureStr(str):
        """IronPython 2.7.3: startswith(u"\\ufeff") is True for any text."""

        def startswith(self, prefix, *args):
            if prefix == u"\ufeff":
                return True
            return str.startswith(self, prefix, *args)

    def test_ordinal(self):
        self.assertEqual(lib.strip_bom(self.CultureStr(u"{}")), u"{}")
        self.assertEqual(lib.strip_bom(u"\ufeff\ufeffC:/r.csv"), u"C:/r.csv")
        self.assertEqual(lib.strip_bom(u""), u"")
        self.assertIsNone(lib.strip_bom(None))

    def test_job_text_keeps_first_char(self):
        # (CPython json itself calls startswith("\\ufeff"), so the CultureStr is
        # checked on strip_bom only; IronPython 2.7 json does not.)
        self.assertEqual(lib.strip_bom(self.CultureStr(u'{"a": 1}')), u'{"a": 1}')
        self.assertEqual(lib.reports_dir_from_job_text(u'\ufeff\ufeff{"reports_dir": "C:/r"}'), u"C:/r")

    def test_read_all_text_drops_extra_bom(self):
        io = types.SimpleNamespace(File=types.SimpleNamespace(ReadAllText=lambda p, e: u"\ufeffC:/x.csv\r\n"))
        old = getattr(lib.System, "IO", None), getattr(lib.System, "Text", None)
        lib.System.IO = io
        lib.System.Text = types.SimpleNamespace(Encoding=types.SimpleNamespace(UTF8="utf8"))
        try:
            self.assertEqual(lib.read_all_text("p"), u"C:/x.csv\r\n")
        finally:
            lib.System.IO, lib.System.Text = old

if __name__ == "__main__":
    unittest.main()
