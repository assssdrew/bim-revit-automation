# -*- coding: utf-8 -*-
"""Host-side checks for the levels/grids geometry diff (no Revit).

The 2026-11-06 K4 audit was GREEN with levels 10 vs 19 because compare
only counted elements and read GetWarnings. These cases lock the report
that a re-run must produce.
"""
from __future__ import print_function

import ast
import math
import os
import sys
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
TOOL = os.path.dirname(os.path.dirname(HERE))
if TOOL not in sys.path:
    sys.path.insert(0, TOOL)

import lg_lib  # noqa: E402


TOL = {
    "level_mm": 2.0,
    "level_delta_yellow_mm": 2.0,
    "level_delta_red_mm": 10.0,
    "grid_mm": 5.0,
    "grid_delta_yellow_mm": 5.0,
    "grid_delta_red_mm": 20.0,
    "grid_angle_deg": 0.1,
    "grid_angle_yellow_deg": 0.1,
    "grid_angle_red_deg": 1.0,
}

MM_PER_FT = 304.8


def _mm_ft(mm):
    return float(mm) / MM_PER_FT


def _level(name, elev_mm):
    return {"name": name, "elev_mm": float(elev_mm)}


def _grid(name, x0, y0, x1, y1):
    return {
        "name": name,
        "p0": (x0, y0),
        "p1": (x1, y1),
        "mid_xy": ((x0 + x1) * 0.5, (y0 + y1) * 0.5),
        "angle_deg": math.degrees(math.atan2(y1 - y0, x1 - x0)),
    }


def _axis(name, y_mm=0.0, length_ft=10.0):
    y = _mm_ft(y_mm)
    return _grid(name, 0.0, y, length_ft, y)


def _codes(details, problem):
    return [d for d in details if d.get("problem") == problem]


class GeometryDiffTests(unittest.TestCase):
    def test_k4_levels_moved_and_leftover_are_not_green(self):
        # BF has 10 levels, AR has those 10 moved by 150 mm plus 9 more.
        host = {}
        exemplar = {}
        for i in range(10):
            name = u"L{0}".format(i)
            host[name] = [_level(name, i * 3000.0)]
            exemplar[name] = [_level(name, i * 3000.0 + 150.0)]
        for i in range(10, 19):
            name = u"L{0}".format(i)
            exemplar[name] = [_level(name, i * 3000.0)]

        details = lg_lib.diff_named_levels(host, exemplar, TOL, None)
        moved = _codes(details, "elevation_delta")
        missing = _codes(details, "missing_in_host")
        self.assertEqual(len(moved), 10)
        self.assertEqual(len(missing), 9)
        self.assertTrue(all(d["status"] == "RED" for d in moved))
        self.assertTrue(all(d["source"] == "geometry" for d in details))
        # delta is host minus exemplar: host sits 150 mm below the link.
        self.assertEqual(moved[0]["delta"], "-150.000")

        merged = lg_lib.merge_geo_with_coord_review(details, [])
        counts = lg_lib.recount_issues(merged)
        self.assertEqual(counts["issues_total"], 19)
        self.assertNotEqual(counts["status"], "GREEN")

    def test_matching_elevations_still_flag_leftover_levels(self):
        host = {}
        exemplar = {}
        for i in range(10):
            name = u"L{0}".format(i)
            host[name] = [_level(name, 0.0)]
            exemplar[name] = [_level(name, 0.0)]
        for i in range(10, 19):
            name = u"L{0}".format(i)
            exemplar[name] = [_level(name, 1000.0)]
        details = lg_lib.diff_named_levels(host, exemplar, TOL, None)
        self.assertEqual(_codes(details, "elevation_delta"), [])
        self.assertEqual(len(_codes(details, "missing_in_host")), 9)
        self.assertNotEqual(lg_lib.recount_issues(details)["status"], "GREEN")

    def test_level_thresholds(self):
        def one(delta):
            host = {u"Этаж 01": [_level(u"Этаж 01", 0.0)]}
            ex = {u"Этаж 01": [_level(u"Этаж 01", delta)]}
            return lg_lib.diff_named_levels(host, ex, TOL, None)

        self.assertEqual(one(0.0), [])
        self.assertEqual(one(2.0), [])
        yellow = one(2.1)
        self.assertEqual(len(yellow), 1)
        self.assertEqual(yellow[0]["status"], "YELLOW")
        self.assertEqual(yellow[0]["problem"], "elevation_delta")
        red = one(-10.1)
        self.assertEqual(red[0]["status"], "RED")
        # exemplar is lower, so host minus exemplar is positive.
        self.assertEqual(red[0]["delta"], "10.100")

    def test_host_only_level_and_whitelist(self):
        host = {
            u"Local": [_level(u"Local", 0.0)],
            u"Keep": [_level(u"Keep", 0.0)],
        }
        ex = {
            u"Keep": [_level(u"Keep", 50.0)],
            u"Только АР": [_level(u"Только АР", 1.0)],
        }
        details = lg_lib.diff_named_levels(host, ex, TOL, [u"Keep"])
        problems = [(d["problem"], d["name_host"] or d["name_exemplar"]) for d in details]
        self.assertIn(("elevation_delta", u"Keep"), problems)
        self.assertIn(("missing_in_host", u"Только АР"), problems)
        self.assertNotIn(("extra_in_host", u"Local"), problems)

        skipped = lg_lib.diff_named_levels(
            {u"Skip": [_level(u"Skip", 0.0)]},
            {u"Skip": [_level(u"Skip", 500.0)]},
            TOL,
            [u"Need"],
        )
        self.assertEqual(_codes(skipped, "elevation_delta"), [])
        self.assertEqual(len(_codes(skipped, "whitelist_missing")), 1)

    def test_same_count_shifted_grids_are_not_green(self):
        host = {}
        exemplar = {}
        for i in range(14):
            name = u"G{0}".format(i)
            host[name] = [_axis(name, 0.0)]
            exemplar[name] = [_axis(name, 100.0)]
        details = lg_lib.diff_named_grids(host, exemplar, TOL)
        moved = _codes(details, "position_delta")
        self.assertEqual(len(moved), 14)
        self.assertTrue(all(d["status"] == "RED" for d in moved))
        self.assertEqual(_codes(details, "missing_in_host"), [])
        self.assertEqual(_codes(details, "angle_delta"), [])
        self.assertAlmostEqual(float(moved[0]["delta"]), 100.0, places=2)
        self.assertNotEqual(lg_lib.recount_issues(details)["status"], "GREEN")

    def test_longer_or_reversed_grid_on_same_line_is_not_a_move(self):
        host = {u"1": [_axis(u"1", 0.0, length_ft=10.0)]}
        longer = {u"1": [_axis(u"1", 0.0, length_ft=40.0)]}
        self.assertEqual(lg_lib.diff_named_grids(host, longer, TOL), [])
        reversed_grid = {u"1": [_grid(u"1", 10.0, 0.0, 0.0, 0.0)]}
        self.assertEqual(lg_lib.diff_named_grids(host, reversed_grid, TOL), [])

    def test_grid_angle_only_and_small_shift(self):
        host = {u"A": [_grid(u"A", 0.0, 0.0, 10.0, 0.0)]}
        spun = {u"A": [_grid(u"A", 5.0, -5.0, 5.0, 5.0)]}
        details = lg_lib.diff_named_grids(host, spun, TOL)
        self.assertEqual(_codes(details, "position_delta"), [])
        angles = _codes(details, "angle_delta")
        self.assertEqual(len(angles), 1)
        self.assertEqual(angles[0]["status"], "RED")

        nudge = {u"A": [_axis(u"A", 3.0)]}
        self.assertEqual(lg_lib.diff_named_grids(host, nudge, TOL), [])
        yellow = lg_lib.diff_named_grids(host, {u"A": [_axis(u"A", 8.0)]}, TOL)
        self.assertEqual(yellow[0]["status"], "YELLOW")
        self.assertEqual(yellow[0]["problem"], "position_delta")

    def test_missing_and_extra_grids(self):
        host = {u"1": [_axis(u"1")]}
        ex = {u"2": [_axis(u"2")]}
        details = lg_lib.diff_named_grids(host, ex, TOL)
        self.assertEqual(len(_codes(details, "extra_in_host")), 1)
        self.assertEqual(len(_codes(details, "missing_in_host")), 1)
        self.assertEqual(lg_lib.recount_issues(details)["status"], "RED")

    def test_cr_warning_does_not_hide_measured_delta(self):
        geo = lg_lib.diff_named_levels(
            {u"Этаж 01": [_level(u"Этаж 01", 0.0)]},
            {u"Этаж 01": [_level(u"Этаж 01", 40.0)]},
            TOL,
            None,
        )
        geo.extend(
            lg_lib.diff_named_grids(
                {u"1": [_axis(u"1", 0.0)]},
                {u"1": [_axis(u"1", 50.0)]},
                TOL,
            )
        )
        cr = [
            {
                "category": "Level",
                "name_host": u"этаж 01",
                "name_exemplar": u"",
                "problem": "cr_level_moved",
                "status": "RED",
                "source": "revit_cr",
            },
            {
                "category": "System",
                "name_host": u"",
                "name_exemplar": u"",
                "problem": "cr_link_needs_review",
                "status": "YELLOW",
                "source": "revit_cr",
            },
        ]
        merged = lg_lib.merge_geo_with_coord_review(geo, cr)
        problems = [d["problem"] for d in merged]
        self.assertIn("elevation_delta", problems)
        self.assertIn("position_delta", problems)
        self.assertIn("cr_link_needs_review", problems)
        self.assertNotIn("cr_level_moved", problems)
        self.assertNotEqual(lg_lib.recount_issues(merged)["status"], "GREEN")

    def test_cr_move_without_geometry_is_kept(self):
        cr = [
            {
                "category": "Level",
                "name_host": u"L0",
                "problem": "cr_level_moved",
                "status": "RED",
                "source": "revit_cr",
            }
        ]
        merged = lg_lib.merge_geo_with_coord_review([], cr)
        self.assertEqual(len(merged), 1)
        self.assertEqual(lg_lib.recount_issues(merged)["status"], "RED")

    def test_scripts_have_no_fstrings(self):
        for name in ("lg_lib.py", "levels_grids_audit.py", "levels_grids_apply.py"):
            path = os.path.join(TOOL, name)
            with open(path, "r") as handle:
                tree = ast.parse(handle.read(), path)
            for node in ast.walk(tree):
                self.assertFalse(
                    isinstance(node, ast.JoinedStr),
                    "{0} has an f-string".format(name),
                )


if __name__ == "__main__":
    unittest.main()
