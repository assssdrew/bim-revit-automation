"""Переносимость папки: все пути — от расположения файлов (в т. ч. UNC, кириллица, пробелы); настройки пользователя — в профиле; один диалог выбора папки;
плоская раскладка корня (сверху только запуск, инструкция и рабочие скрипты, остальное — в служебное\\)."""
import os
import re
import shutil
import subprocess
import sys
import tempfile
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from test_chain import pwsh, PWSH  # noqa: E402

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(HERE)))                    # корень: то, что нужно пользователю
SERVICE = os.path.join(ROOT, "служебное")
TOOL = os.path.join(SERVICE, "скрипты")                                             # все .ps1/.py окна и задач Revit
LIB = os.path.join(SERVICE, "lib")
ROOT_FILES = {"Проверочные спецификации.vbs", "ИНСТРУКЦИЯ.txt", "servers.cfg", ".gitignore"}
ROOT_DIRS = {"служебное", "отчеты"}                                                 # отчеты создаётся сама (в git её нет)
IGNORED_DIRS = {"__pycache__", ".git"}


IP_RE = r"\b(?:\d{1,3}\\?\.){3}\d{1,3}\b"          # и «revit-server-2022.example.local», и в экранированном regex «revit-server-2022\.example\.local»


def read(p):
    with open(p, encoding="utf-8-sig") as f:
        return f.read()


def sources():
    for d, dirs, files in os.walk(ROOT):
        dirs[:] = [x for x in dirs if x not in ("tests", "__pycache__", ".git", "отчеты")]
        for f in files:
            if f.endswith((".ps1", ".py", ".bat", ".vbs", ".cmd")):
                yield os.path.join(d, f)


class LayoutTests(unittest.TestCase):
    def test_root_contains_exactly_what_the_user_needs(self):
        names = {n for n in os.listdir(ROOT) if n not in IGNORED_DIRS}
        files = {n for n in names if os.path.isfile(os.path.join(ROOT, n))}
        dirs = {n for n in names if os.path.isdir(os.path.join(ROOT, n))}
        self.assertEqual(files, ROOT_FILES)                                   # ни одного .ps1/.py в корне
        self.assertTrue(dirs <= ROOT_DIRS, dirs - ROOT_DIRS)                  # lib, tests и прочее — только в служебное
        self.assertIn("служебное", dirs)

    def test_service_first_level(self):
        names = set(os.listdir(SERVICE))
        for must in ("скрипты", "lib", "google_push", "rbp_rsn_patch", "разработка", "choose_models.ps1"):
            self.assertIn(must, names)
        # .bat и диагностические скрипты — в разработка\инструменты, в служебное только нужное окну
        self.assertEqual(names, {"скрипты", "lib", "google_push", "rbp_rsn_patch", "разработка", "choose_models.ps1"}
                         | (names & {"__pycache__", "excel_blank.xlsx"}))
        self.assertFalse(any(n.endswith((".bat", ".md")) for n in names))
        inst = os.listdir(os.path.join(SERVICE, "разработка", "инструменты"))
        self.assertEqual(len([n for n in inst if n.endswith(".bat")]), 2)   # дамп и сравнение фильтров
        self.assertTrue(os.path.exists(os.path.join(LIB, "DocumentFormat.OpenXml.dll")))
        scripts = os.listdir(TOOL)
        self.assertIn("ui_schedules.ps1", scripts)
        self.assertIn("insert_schedules.py", scripts)
        for n in scripts:
            if n == "__pycache__":
                continue
            self.assertTrue(n.endswith((".ps1", ".py")) or n == "разделы.txt", n)
        for gone in ("lib", "tests", "README.md", "АРХИТЕКТУРА_цепочки.md", "ФАЙЛЫ_ОТЧЕТОВ.md", "BUILDER.txt", "пример_списка_спецификаций.txt", "schedule_sets.txt",
                     "build_release.ps1", "build_release.bat", "_система"):
            self.assertFalse(os.path.exists(os.path.join(ROOT, gone)), gone)
        dev = os.path.join(SERVICE, "разработка")
        for kept in ("tests", "README.md", "АРХИТЕКТУРА_цепочки.md", "ФАЙЛЫ_ОТЧЕТОВ.md"):
            self.assertTrue(os.path.exists(os.path.join(dev, kept)), kept)
        self.assertTrue(os.path.exists(os.path.join(dev, "tests", "fixtures", "schedule_sets.txt")))   # только тестовые данные

    def test_gitignore_and_secret_example(self):
        gi = read(os.path.join(ROOT, ".gitignore"))
        self.assertIn("__pycache__/", gi)
        self.assertIn("служебное/google_push/google_push.cfg", gi)
        ex = read(os.path.join(ROOT, "служебное", "google_push", "google_push.cfg.example"))
        self.assertRegex(ex, r"(?m)^webapp_url=\s*$")
        self.assertRegex(ex, r"(?m)^secret=\s*$")
        self.assertFalse(os.path.exists(os.path.join(ROOT, "служебное", "google_push", "google_push.cfg")) and subprocess.run(
            ["git", "ls-files", "--error-unmatch", "служебное/google_push/google_push.cfg"], cwd=ROOT, capture_output=True).returncode == 0)

    def test_instruction_describes_flat_layout(self):
        t = read(os.path.join(ROOT, "ИНСТРУКЦИЯ.txt"))
        self.assertIn("Установка на общий компьютер", t)
        for must in ("google_push.cfg", "UNC", "Unblock-File", "ExecutionPolicy", "Проверочные спецификации.vbs", "служебное\\разработка"):
            self.assertIn(must, t)
        self.assertNotIn("_система", t)
        self.assertNotIn("build_release", t)


@unittest.skipUnless(PWSH, "нет pwsh")
class SelfCheckTests(unittest.TestCase):
    def run_selfcheck(self, tool_dir):
        """Реальный блок самопроверки ui_schedules.ps1 (до загрузки WinForms: на Linux типы Windows Forms не компилируются, на Windows — `ui_schedules.ps1 -SelfCheck`)."""
        text = read(os.path.join(tool_dir, "ui_schedules.ps1"))
        cut = text.index("Add-Type -AssemblyName System.Windows.Forms")
        tmp = os.path.join(tool_dir, "_selfcheck_tmp.ps1")
        with open(tmp, "w", encoding="utf-8-sig") as f:
            f.write(text[:cut])
        try:
            return subprocess.run([PWSH, "-NoProfile", "-File", tmp, "-SelfCheck"], capture_output=True, text=True, encoding="utf-8",
                                  env=dict(os.environ, DOTNET_SYSTEM_GLOBALIZATION_INVARIANT="1"), timeout=120)
        finally:
            os.remove(tmp)

    def test_selfcheck_finds_all_files_from_own_location(self):
        r = self.run_selfcheck(TOOL)
        self.assertEqual(r.returncode, 0, r.stdout + r.stderr)
        self.assertIn("SELFCHECK OK", r.stdout)
        self.assertIn("TOOLDIR=" + TOOL, r.stdout)
        self.assertIn("REPORTDIR=" + os.path.join(ROOT, "отчеты"), r.stdout)       # отчёты — в корне, не внутри служебное

    def test_selfcheck_works_from_copy_in_cyrillic_path_with_spaces_and_reports_missing(self):
        tmp = tempfile.mkdtemp(prefix="bvs_port_")
        self.addCleanup(shutil.rmtree, tmp, True)
        dst = os.path.join(tmp, "Общая папка с пробелами", "batch_verification_schedules")
        shutil.copytree(ROOT, dst, ignore=shutil.ignore_patterns("tests", "__pycache__", ".git", "*.pyc", "разработка", "отчеты"))
        tool = os.path.join(dst, "служебное", "скрипты")
        r = self.run_selfcheck(tool)
        self.assertEqual(r.returncode, 0, r.stdout + r.stderr)
        self.assertIn("TOOLDIR=" + tool, r.stdout)
        self.assertIn("REPORTDIR=" + os.path.join(dst, "отчеты"), r.stdout)
        os.remove(os.path.join(tool, "project_config.ps1"))
        os.remove(os.path.join(dst, "servers.cfg"))
        r2 = self.run_selfcheck(tool)
        self.assertEqual(r2.returncode, 3)
        self.assertIn("MISSING=project_config.ps1", r2.stdout)
        self.assertIn("MISSING=..\\..\\servers.cfg", r2.stdout)                         # servers.cfg читается из корня

    def test_layout_helpers(self):
        tmp = tempfile.mkdtemp(prefix="bvs_rr_")
        self.addCleanup(shutil.rmtree, tmp, True)
        script = os.path.join(tmp, "t.ps1")
        root = os.path.join(tmp, "Общая папка", "AppX")
        tool = os.path.join(root, "служебное", "скрипты")
        with open(script, "w", encoding="utf-8-sig") as f:
            f.write(". '%s'\n(Get-BvsRootDir '%s')\n(Get-BvsServiceDir '%s/')\n(Get-BvsLibDir '%s')\n(Get-ReportRootDir '%s')\n" % (os.path.join(TOOL, "report_layout.ps1"), tool, tool, tool, tool))
        r = pwsh("-File", script)
        self.assertEqual(r.returncode, 0, r.stderr)
        lines = r.stdout.splitlines()
        self.assertEqual(lines, [root, os.path.join(root, "служебное"), os.path.join(root, "служебное", "lib"), os.path.join(root, "отчеты")])

    def test_every_literal_relative_reference_exists(self):
        """Все `Join-Path $ToolDir "файл"` / `$PSScriptRoot "файл"` указывают на существующие файлы (необязательные — в списке)."""
        optional = {"google_push.cfg", "list_paths.ps1", "rvt_list.txt", "rbp_rsn_patch\\apply_rsn_support.ps1", "служебное\\google_push\\google_push.cfg"}
        missing = []
        for p in sources():
            if not p.endswith(".ps1"):
                continue
            d = os.path.dirname(p)
            for var, rel in re.findall(r'Join-Path \$(ToolDir|PSScriptRoot|here|tool|root|__pushDir) +"([^"$]+\.(?:ps1|py|dll|cfg|txt))"', read(p)):
                if rel in optional:
                    continue
                bases = [TOOL, SERVICE, ROOT, d]
                if not any(os.path.exists(os.path.join(b, rel.replace("\\", os.sep))) for b in bases):
                    missing.append((os.path.basename(p), rel))
        self.assertEqual(missing, [])


class PortabilityStaticTests(unittest.TestCase):
    def test_vbs_is_portable_and_ascii(self):
        raw = open(os.path.join(ROOT, "Проверочные спецификации.vbs"), "rb").read()
        t = raw.decode("ascii")                                                              # WSH читает .vbs в ANSI: кириллицы в тексте нет
        self.assertIn("ScriptFullName", t)
        self.assertIn("ui_schedules.ps1", t)
        self.assertIsNone(re.search(r"[A-Za-z]:\\", t))
        self.assertIn("-ExecutionPolicy Bypass", t)

    def test_bat_files_use_pushd_and_own_location(self):
        n = 0
        bat_dir = os.path.join(ROOT, "служебное", "разработка", "инструменты")
        for f in os.listdir(bat_dir):
            if f.endswith(".bat"):
                n += 1
                t = read(os.path.join(bat_dir, f))
                self.assertIn('pushd "%~dp0"', t, f)
                self.assertIn("%~dp0", t)
                self.assertIn("-ExecutionPolicy Bypass", t)
                self.assertIsNone(re.search(r"[A-Za-z]:\\", t), f)
        self.assertEqual(n, 2)   # «Выгрузить фильтры» и «Сравнить фильтры»

    def test_no_user_profiles_or_server_addresses_in_code(self):
        # Единственный разрешённый адрес — UNC классификатора PROJ: запасная константа в set_assembly_code.py
        # и значение по умолчанию для проекта PROJ в project_config.ps1 (Get-ProjectClassifierPath).
        allowed_ip = {
            os.path.normcase(os.path.normpath(os.path.join(TOOL, "set_assembly_code.py"))): "fileserver.example.local",
            os.path.normcase(os.path.normpath(os.path.join(TOOL, "project_config.ps1"))): "fileserver.example.local",
        }
        for p in sources():
            t = read(p)
            rel = os.path.relpath(p, ROOT)
            self.assertIsNone(re.search(r"(?i)[A-Z]:\\Users\\", t), rel)
            self.assertIsNone(re.search(r"(?i)pota[p]ov", t), rel)
            found = re.findall(IP_RE, t)
            allowed = allowed_ip.get(os.path.normcase(os.path.normpath(p)))
            if allowed:
                self.assertEqual(set(found), {allowed}, rel)
            else:
                self.assertEqual(found, [], rel)

    def test_only_one_folder_dialog_and_it_is_the_address_bar_one(self):
        for p in sources():
            self.assertNotIn("FolderBrowserDialog", read(p), p)
        self.assertIn("function Show-FolderPathDialog", read(os.path.join(TOOL, "folder_dialog.ps1")))
        self.assertIn("folder_dialog.ps1", read(os.path.join(TOOL, "ui_schedules.ps1")))

    def test_user_settings_live_in_user_profile_not_in_code(self):
        ui = read(os.path.join(TOOL, "ui_schedules.ps1"))
        self.assertIn('$AppDir = Join-Path $env:LOCALAPPDATA "BatchRvt\\batch_verification_schedules"', ui)
        self.assertNotIn('Join-Path $ToolDir "run.cfg"', ui)
        self.assertIn("$ReportDir = Get-ReportRootDir $ToolDir", ui)
        for name in ("insert_schedules.py", "export_schedules.py", "set_assembly_code.py"):
            t = read(os.path.join(TOOL, name))
            self.assertIn('LOCAL_APPDATA = os.environ.get("LOCALAPPDATA")', t)
            self.assertNotIn("TOOL_DIR_SHARE", t)
        self.assertNotIn("BUILDER.txt", read(os.path.join(TOOL, "bvs_assignment_guard.ps1")))

    def test_all_scripts_take_report_root_from_one_function(self):
        for p in sources():
            if p.endswith(".ps1"):
                self.assertNotIn('Join-Path $ToolDir "отчеты"', read(p), p)

    def test_no_build_release_leftovers(self):
        for p in sources():
            self.assertNotIn("_система", read(p), p)


if __name__ == "__main__":
    unittest.main()
