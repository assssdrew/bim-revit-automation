# -*- coding: utf-8 -*-
# Name: test_relink_prefix.py
# Version: 3.1.0
# What it does: See the case README.
# Notes: Paths and project names in this portfolio copy are placeholders.

"""Tests for relink_prefix.py (folder-prefix relink). CPython, no Revit:

    python3 -I tests/test_relink_prefix.py
"""
from __future__ import unicode_literals

import io
import os
import shutil
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import relink_prefix as rp  # noqa: E402

RSN_OLD = "RSN://revit-server-2022.example.local/proj SampleSite/Стадия ПП"
RSN_NEW = "RSN://revit-server-2022.example.local/proj SampleSite/Стадия П"


class Norm(unittest.TestCase):
    def test_slashes_and_trailing(self):
        self.assertEqual(rp.norm("\\\\srv\\share\\Old\\"), "//srv/share/Old")
        self.assertEqual(rp.norm("  D:\\a\\b  "), "D:/a/b")
        self.assertEqual(rp.norm("RSN://h/"), "RSN://h")
        self.assertEqual(rp.norm("RSN://"), "RSN://")

    def test_none_and_empty(self):
        self.assertEqual(rp.norm(None), "")
        self.assertEqual(rp.norm("   "), "")


class Remap(unittest.TestCase):
    def setUp(self):
        self.rsn = [(rp.norm(RSN_OLD), rp.norm(RSN_NEW))]

    def test_rsn_cyrillic_remap(self):
        got = rp.remap(RSN_OLD + "/09_ИТП/x.rvt", self.rsn)
        self.assertEqual(got, (RSN_NEW + "/09_ИТП/x.rvt", "remap"))

    def test_folder_boundary(self):
        # 'Стадия П' must not match 'Стадия ПП/...' and vice versa
        self.assertEqual(rp.remap(RSN_OLD + "ПП/x.rvt", self.rsn), (None, "outside"))
        self.assertEqual(rp.remap(RSN_NEW + "/x.rvt", self.rsn), (None, "already"))

    def test_the_old_folder_itself(self):
        self.assertEqual(rp.remap(RSN_OLD, self.rsn), (RSN_NEW, "remap"))

    def test_case_insensitive_match_keeps_tail_case(self):
        got = rp.remap(RSN_OLD.upper() + "/Sub/Model.rvt", self.rsn)
        self.assertEqual(got[1], "remap")
        self.assertTrue(got[0].endswith("/Sub/Model.rvt"))

    def test_outside_and_empty(self):
        self.assertEqual(rp.remap("RSN://other/x.rvt", self.rsn), (None, "outside"))
        self.assertEqual(rp.remap("", self.rsn), (None, "empty"))
        self.assertEqual(rp.remap(None, self.rsn), (None, "empty"))

    def test_unc_keeps_backslashes(self):
        pairs = [(rp.norm("\\\\srv\\share\\Old"), rp.norm("\\\\srv\\share\\New"))]
        got = rp.remap("\\\\srv\\share\\Old\\a\\x.rvt", pairs)
        self.assertEqual(got, ("\\\\srv\\share\\New\\a\\x.rvt", "remap"))

    def test_drive_keeps_backslashes(self):
        pairs = [(rp.norm("D:\\Old"), rp.norm("D:\\New"))]
        self.assertEqual(rp.remap("D:\\Old\\a\\x.rvt", pairs), ("D:\\New\\a\\x.rvt", "remap"))

    def test_forward_slash_unc_stays_forward(self):
        pairs = [(rp.norm("//srv/share/Old"), rp.norm("//srv/share/New"))]
        self.assertEqual(rp.remap("//srv/share/Old/x.rvt", pairs), ("//srv/share/New/x.rvt", "remap"))

    def test_rsn_never_gets_backslashes_even_if_prefix_has_them(self):
        pairs = [(rp.norm("RSN://h/Old"), rp.norm("RSN:\\\\h\\New"))]
        got = rp.remap("RSN://h/Old/x.rvt", pairs)
        self.assertNotIn("\\", got[0])

    def test_first_matching_pair_wins_and_already_checked_first(self):
        pairs = [("//s/a", "//s/b"), ("//s/b", "//s/c")]
        self.assertEqual(rp.remap("\\\\s\\a\\x.rvt", pairs), ("\\\\s\\b\\x.rvt", "remap"))
        self.assertEqual(rp.remap("\\\\s\\b\\x.rvt", pairs), (None, "already"))


class LoadMap(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.dir, ignore_errors=True)

    def write(self, text, bom=True):
        path = os.path.join(self.dir, rp.MAP_NAME)
        with io.open(path, "w", encoding="utf-8-sig" if bom else "utf-8", newline="") as f:
            f.write(text)
        return path

    def test_header_bom_quotes_blank_lines(self):
        path = self.write('old_prefix;new_prefix\r\n\r\n"%s";"%s"\r\n' % (RSN_OLD, RSN_NEW))
        self.assertEqual(rp.load_map(path), [(rp.norm(RSN_OLD), rp.norm(RSN_NEW))])

    def test_without_bom_and_header(self):
        path = self.write("RSN://h/A;RSN://h/B\n", bom=False)
        self.assertEqual(rp.load_map(path), [("RSN://h/A", "RSN://h/B")])

    def test_equal_pair_and_short_rows_ignored(self):
        path = self.write("old_prefix;new_prefix\nRSN://h/A;rsn://h/a\nonlyone\n;\n")
        self.assertEqual(rp.load_map(path), [])

    def test_missing_file_and_none(self):
        self.assertEqual(rp.load_map(os.path.join(self.dir, "nope.csv")), [])
        self.assertEqual(rp.load_map(None), [])


if __name__ == "__main__":
    unittest.main()
