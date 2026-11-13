# -*- coding: utf-8 -*-
# Name: test_script_dir.py
# Version: 3.1.0
# What it does: See the case README.
# Notes: Paths and project names in this portfolio copy are placeholders.

"""The RBP task scripts live in <tool>/sluzhebnoe but read job files from the tool root.

get_script_dir() is taken from each script by AST and run with a fake __file__.
CPython, no Revit:  python3 -I tests/test_script_dir.py
"""
import ast
import os
import shutil
import tempfile
import unittest

CODE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SCRIPTS = ["detach_to_buffer.py", "saveas_central.py", "update_rvt_links.py", "read_worksets.py", "apply_workset_links.py"]


def load_fn(name, fake_file, fallback):
    with open(os.path.join(CODE, name), "rb") as f:
        tree = ast.parse(f.read().decode("utf-8-sig"), name)
    fn = [n for n in tree.body if isinstance(n, ast.FunctionDef) and n.name == "get_script_dir"][0]
    mod = ast.Module(body=[fn], type_ignores=[])
    ns = {"os": os, "__file__": fake_file, "TOOL_DIR_FALLBACK": fallback, "TOOL_DIR_SHARE": ""}
    exec(compile(mod, name, "exec"), ns)
    return ns["get_script_dir"]


class ScriptDir(unittest.TestCase):
    def setUp(self):
        self.root = tempfile.mkdtemp()
        self.code = os.path.join(self.root, u"служебное")
        os.makedirs(self.code)
        self.fallback = os.path.join(self.root, "nowhere")

    def tearDown(self):
        shutil.rmtree(self.root, ignore_errors=True)

    def test_root_one_level_up(self):
        open(os.path.join(self.root, "job_paths.csv"), "w").close()
        for name in SCRIPTS:
            fn = load_fn(name, os.path.join(self.code, name), self.fallback)
            self.assertEqual(fn(), self.root, name)

    def test_root_wins_over_code_dir(self):
        open(os.path.join(self.root, "job_paths.csv"), "w").close()
        open(os.path.join(self.code, "job_paths.csv"), "w").close()
        for name in SCRIPTS:
            fn = load_fn(name, os.path.join(self.code, name), self.fallback)
            self.assertEqual(fn(), self.root, name)

    def test_old_flat_layout_still_works(self):
        open(os.path.join(self.code, "job_paths.csv"), "w").close()
        for name in SCRIPTS:
            fn = load_fn(name, os.path.join(self.code, name), self.fallback)
            self.assertEqual(fn(), self.code, name)

    def test_code_dir_on_sys_path(self):
        for name in SCRIPTS:
            with open(os.path.join(CODE, name), "rb") as f:
                text = f.read().decode("utf-8-sig")
            self.assertIn("CODE_DIR = os.path.dirname(os.path.abspath(__file__))", text, name)
            self.assertIn("for _p in (CODE_DIR, SCRIPT_DIR):", text, name)


if __name__ == "__main__":
    unittest.main()
