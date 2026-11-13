# -*- coding: utf-8 -*-
# Name: test_rbp_ascii.py
# Version: 3.1.0
# What it does: See the case README.
# Notes: Paths and project names in this portfolio copy are placeholders.

"""Static check: the RBP task scripts keep the log ASCII-only.

The RBP log breaks on non-ASCII text (full UNC/RSN paths with Cyrillic), so in the
checked scripts Output() must be the ASCII-safe wrapper and every
raise Exception(...) must carry ASCII text. CPython, no Revit:

    python3 -I tests/test_rbp_ascii.py
"""
import ast
import os
import unittest

TOOL = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CHECKED = ["update_rvt_links.py", "read_worksets.py", "apply_workset_links.py"]


def load(name):
    with open(os.path.join(TOOL, name), "rb") as f:
        return ast.parse(f.read().decode("utf-8-sig"), name)


def is_ascii_safe_call(node):
    return (isinstance(node, ast.Call) and isinstance(node.func, ast.Attribute)
            and node.func.attr == "ascii_safe")


class RbpAscii(unittest.TestCase):
    def test_output_is_the_ascii_wrapper(self):
        for name in CHECKED:
            tree = load(name)
            wrappers = [n for n in tree.body if isinstance(n, ast.FunctionDef) and n.name == "Output"]
            self.assertEqual(len(wrappers), 1, name + ": module-level Output wrapper is missing")
            calls = [c for c in ast.walk(wrappers[0]) if isinstance(c, ast.Call)]
            self.assertTrue(any(is_ascii_safe_call(c) for c in calls),
                            name + ": wrapper must pass text through lib.ascii_safe")
            self.assertTrue(any(isinstance(c.func, ast.Name) and c.func.id == "_rbp_output" for c in calls),
                            name + ": wrapper must call the original Output saved as _rbp_output")

    def test_exceptions_are_ascii(self):
        for name in CHECKED:
            for n in ast.walk(load(name)):
                if not (isinstance(n, ast.Raise) and isinstance(n.exc, ast.Call)):
                    continue
                f = n.exc.func
                if not (isinstance(f, ast.Name) and f.id == "Exception") or not n.exc.args:
                    continue
                arg = n.exc.args[0]
                if isinstance(arg, ast.Constant) and isinstance(arg.value, str):
                    self.assertTrue(all(ord(c) < 128 for c in arg.value),
                                    "%s:%d non-ASCII exception text" % (name, n.lineno))
                else:
                    self.assertTrue(is_ascii_safe_call(arg),
                                    "%s:%d exception text must go through lib.ascii_safe" % (name, n.lineno))

    def test_wrapper_precedes_first_log_call(self):
        # the wrapper must be defined before the first module-level Output() use
        for name in CHECKED:
            tree = load(name)
            wrapper = [n.lineno for n in tree.body if isinstance(n, ast.FunctionDef) and n.name == "Output"]
            self.assertTrue(wrapper, name)
            for n in tree.body:
                if isinstance(n, ast.Expr) and isinstance(n.value, ast.Call):
                    c = n.value.func
                    if isinstance(c, ast.Name) and c.id == "Output":
                        self.assertGreater(n.lineno, wrapper[0], name)


if __name__ == "__main__":
    unittest.main()
