"""Веб-страница для просмотра отчёта (служебное/google_push: doGet в Code.gs + Viewer.html).

Серверная часть гоняется на подставных листах (node, только чтение), страница — в Chromium через Playwright.
Нет node / playwright / Chromium — соответствующий тест пропускается.
"""
import os
import shutil
import subprocess
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(os.path.dirname(HERE)))
TOOL = os.path.join(ROOT, "служебное", "скрипты")      # папка с .ps1/.py окна
PUSH = os.path.join(ROOT, "служебное", "google_push")
NODE = shutil.which("node")


def run_node(script, timeout=180):
    r = subprocess.run([NODE, os.path.join(HERE, script)], capture_output=True, text=True, encoding="utf-8", timeout=timeout)
    return r, r.stdout + r.stderr


@unittest.skipUnless(NODE, "нет node")
class ViewerTests(unittest.TestCase):
    def test_server_side_mock(self):
        r, out = run_node("viewer_mock_test.js")
        self.assertEqual(r.returncode, 0, out)
        self.assertIn("all ok", out)

    def test_browser(self):
        r, out = run_node("viewer_browser_test.js")
        self.assertEqual(r.returncode, 0, out)
        if "SKIP" in out:
            self.skipTest(out.strip())
        self.assertIn("all ok", out)

    def test_existing_push_mock_still_passes(self):
        """doPost и отправка данных не затронуты: прежний мок-тест Code.gs."""
        r, out = run_node("code_gs_mock_test.js")
        self.assertEqual(r.returncode, 0, out)


class ViewerStaticTests(unittest.TestCase):
    def test_do_post_unchanged_and_no_secrets(self):
        code = open(os.path.join(PUSH, "Code.gs"), encoding="utf-8").read()
        self.assertIn("function doPost(e)", code)
        self.assertIn("REPLACE_ME_WITH_LONG_RANDOM_SECRET", code)
        # настоящий google_push.cfg на рабочем ПК лежит на диске — это нормально; важно, что он не в git и закрыт .gitignore
        tracked = subprocess.run(["git", "ls-files", "--error-unmatch", "google_push.cfg"], cwd=PUSH, capture_output=True)
        self.assertNotEqual(tracked.returncode, 0, "google_push.cfg в git")
        with open(os.path.join(ROOT, ".gitignore"), encoding="utf-8") as f:
            self.assertIn("google_push.cfg", f.read())

    def test_readme_describes_viewer(self):
        text = open(os.path.join(PUSH, "README.md"), encoding="utf-8").read()
        for needle in ("Viewer.html", "?section=", "Выполнять как: Я", "У кого есть доступ: Все", "вход в Google не нужен",
                       "LAST_PUSH_MS", "viewerGetVersion", "Данные обновились", "Новая версия» в обоих"):
            self.assertIn(needle, text)


if __name__ == "__main__":
    unittest.main()
