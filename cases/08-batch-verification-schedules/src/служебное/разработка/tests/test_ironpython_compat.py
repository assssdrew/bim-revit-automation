"""Скрипты, которые запускает Revit Batch Processor (IronPython 2.7), не должны использовать то, что там не работает.

Разбор AST (без запуска): re.sub/re.split/re.subn с flags=/count= по ключу, f-строки, nonlocal, yield from, async/await,
аннотации, ключевые аргументы только по имени, str.casefold/removeprefix/removesuffix/isascii, os.makedirs(exist_ok=),
open(..., encoding=), print(..., end=/sep=/file=). Реальный сбой: «sub() got an unexpected keyword argument 'flags'»
в export_schedules.py (коммит 1df0be9) остановил выгрузку на живом Revit, а тесты на CPython 3 его не заметили.
"""
import ast
import glob
import os
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(os.path.dirname(HERE)))
TOOL = os.path.join(ROOT, "служебное", "скрипты")      # папка с .ps1/.py окна
REVIT_SCRIPTS = (
    [os.path.join(TOOL, "export_schedules.py"), os.path.join(TOOL, "insert_schedules.py"),
     os.path.join(TOOL, "read_template_specs.py"), os.path.join(TOOL, "set_assembly_code.py"),
     os.path.join(TOOL, "bvs_sections.py")]
    + sorted(glob.glob(os.path.join(ROOT, "служебное", "*.py")))
    + sorted(glob.glob(os.path.join(ROOT, "служебное", "разработка", "инструменты", "*.py")))
)
RE_FUNCS = {"sub", "subn", "split"}
BAD_METHODS = {"casefold", "removeprefix", "removesuffix", "isascii"}
BAD_KEYWORDS_BY_FUNC = {"makedirs": {"exist_ok"}, "open": {"encoding", "newline", "errors"}, "print": {"end", "sep", "file", "flush"}}


def problems(path):
    with open(path, encoding="utf-8") as f:
        tree = ast.parse(f.read(), path)
    found = []
    for node in ast.walk(tree):
        where = "%s:%d" % (os.path.basename(path), getattr(node, "lineno", 0))
        if isinstance(node, ast.Call):
            fn = node.func
            name = fn.attr if isinstance(fn, ast.Attribute) else (fn.id if isinstance(fn, ast.Name) else "")
            kws = {k.arg for k in node.keywords if k.arg}
            if name in RE_FUNCS and isinstance(fn, ast.Attribute) and isinstance(fn.value, ast.Name) and fn.value.id == "re":
                if kws & {"flags", "count"}:
                    found.append("%s: re.%s(..., %s=...) — в IronPython 2.7 нельзя, флаг в compile() или (?u) в шаблоне" % (where, name, sorted(kws & {"flags", "count"})[0]))
            # lookbehind в шаблонах: в IronPython 2.7 на живом Revit (13_/тема спеки) не сработал, хотя в CPython работает
            if name in ("search", "match", "sub", "subn", "split", "findall", "finditer", "compile") and isinstance(fn, ast.Attribute) \
                    and isinstance(fn.value, ast.Name) and fn.value.id == "re" and node.args:
                pat = node.args[0]
                if isinstance(pat, ast.Constant) and isinstance(pat.value, str) and ("(?<=" in pat.value or "(?<!" in pat.value):
                    found.append("%s: lookbehind в re.%s — в IronPython 2.7 ненадёжен, разберите строку без regex" % (where, name))
            if name in BAD_METHODS and isinstance(fn, ast.Attribute):
                found.append("%s: .%s() есть только в Python 3" % (where, name))
            if name in BAD_KEYWORDS_BY_FUNC and kws & BAD_KEYWORDS_BY_FUNC[name]:
                found.append("%s: %s(%s=...) — в IronPython 2.7 нет такого аргумента" % (where, name, sorted(kws & BAD_KEYWORDS_BY_FUNC[name])[0]))
        elif isinstance(node, ast.JoinedStr):
            found.append("%s: f-строка" % where)
        elif isinstance(node, (ast.Nonlocal, ast.YieldFrom, ast.AsyncFunctionDef, ast.Await, ast.AsyncFor, ast.AsyncWith)):
            found.append("%s: конструкция только для Python 3 (%s)" % (where, type(node).__name__))
        elif isinstance(node, ast.FunctionDef) or isinstance(node, ast.Lambda):
            a = node.args
            if a.kwonlyargs or getattr(a, "posonlyargs", None):
                found.append("%s: аргументы только по имени / только позиционные" % where)
            if isinstance(node, ast.FunctionDef):
                if node.returns is not None or any(x.annotation is not None for x in a.args + a.kwonlyargs):
                    found.append("%s: аннотации типов" % where)
        elif isinstance(node, ast.AnnAssign):
            found.append("%s: аннотация переменной" % where)
        elif isinstance(node, ast.Starred) and not isinstance(getattr(node, "ctx", None), ast.Load):
            found.append("%s: распаковка со * в присваивании" % where)
    return found


class IronPythonCompatTests(unittest.TestCase):
    def test_scripts_found(self):
        names = [os.path.basename(p) for p in REVIT_SCRIPTS]
        self.assertIn("export_schedules.py", names)
        self.assertIn("dump_schedule_filters.py", names)
        self.assertTrue(all(os.path.exists(p) for p in REVIT_SCRIPTS))

    def test_no_python3_only_constructs_in_revit_scripts(self):
        found = []
        for p in REVIT_SCRIPTS:
            found += problems(p)
        self.assertEqual(found, [], "\n".join(found))

    def test_checker_catches_the_original_defect(self):
        import tempfile
        d = tempfile.mkdtemp(prefix="bvs_ipy_")
        p = os.path.join(d, "bad.py")
        with open(p, "w", encoding="utf-8") as f:
            f.write("import re\nre.sub(u'a', u'b', u'c', flags=re.UNICODE)\nre.split(u'a', u'b', flags=1)\nx = f'{1}'\n"
                    "'a'.casefold()\nimport os\nos.makedirs('x', exist_ok=True)\n")
        found = problems(p)
        self.assertEqual(len(found), 5, found)
        with open(p, "w", encoding="utf-8") as f:
            f.write("import re\nre.search(u'(?<=a)b', u'ab')\n")
        self.assertEqual(len(problems(p)), 1)


if __name__ == "__main__":
    unittest.main()
