"""Режим «Всё» должен давать результат, совпадающий с эталоном tests/baseline (снят до работ над выборочной выгрузкой).

Сравниваются: структурный дамп xlsx, google_push_payload.json, задание_история.txt (как множество строк),
задание_журнал.txt, итоговая строка сборки. Допустимые отличия перечислены в ALLOWED (сейчас — нет).
"""
import os
import shutil
import sys
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import baseline_lib as bl  # noqa: E402

PWSH = os.environ.get("PWSH") or shutil.which("pwsh")

# {шаг: [описание допустимого отличия]} — заполняется явно, если режим «Всё» намеренно меняется
ALLOWED = {}


def first_diff(a, b, path=""):
    if type(a) != type(b):
        return "%s: тип %s != %s" % (path, type(a).__name__, type(b).__name__)
    if isinstance(a, dict):
        for k in sorted(set(a) | set(b)):
            if k not in a:
                return "%s/%s: нет в результате" % (path, k)
            if k not in b:
                return "%s/%s: лишнее в результате" % (path, k)
            d = first_diff(a[k], b[k], "%s/%s" % (path, k))
            if d:
                return d
        return None
    if isinstance(a, list):
        if len(a) != len(b):
            return "%s: длина %d != %d" % (path, len(a), len(b))
        for i, (x, y) in enumerate(zip(a, b)):
            d = first_diff(x, y, "%s[%d]" % (path, i))
            if d:
                return d
        return None
    return None if a == b else "%s: %r != %r" % (path, a, b)


@unittest.skipUnless(PWSH, "нет pwsh")
class AllModeMatchesBaseline(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.results = bl.all_scenarios(PWSH)

    def check_step(self, step):
        got = self.results[step]
        exp = bl.load_baseline(step)
        for key in ("xlsx", "payload"):
            d = first_diff(got[key], exp[key], key)
            self.assertIsNone(d, "%s: %s" % (step, d))
        for key in ("history", "ledger"):
            self.assertEqual(got[key], exp[key], "%s: %s отличается от эталона" % (step, key))
        self.assertEqual(got["stdout"], exp["stdout"], step)

    def test_scenario_a_steps(self):
        for step in ("A1", "A2", "A3", "A4"):
            with self.subTest(step=step):
                self.check_step(step)

    def test_legacy_files_without_false_new(self):
        self.check_step("L1")
        self.assertIn("NEW=0", self.results["L1"]["stdout"])      # старые журнал и история: «Новое» нет
        self.assertIn("FIXED=0", self.results["L1"]["stdout"])    # и «исправлено» нет


if __name__ == "__main__":
    unittest.main()
