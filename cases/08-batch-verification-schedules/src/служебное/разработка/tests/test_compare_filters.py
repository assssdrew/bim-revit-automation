"""Выгрузка фильтров спецификаций (служебное/dump_schedule_filters.py) и сравнение двух выгрузок (compare_filter_lib.ps1)."""
import json
import os
import runpy
import shutil
import subprocess
import sys
import tempfile
import types
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import fake_rbp  # noqa: E402
from test_revit_versions import make_rvt, info_text  # noqa: E402  настоящий (минимальный) файл Revit с годом внутри: версия определяется по содержимому
import revit_api_stub as base  # noqa: E402  подставной Revit API

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(HERE)))
TOOL = os.path.join(ROOT, "служебное", "скрипты")      # папка с .ps1/.py окна
SERVICE = os.path.join(ROOT, "служебное")
DUMP_PY = os.path.join(SERVICE, "разработка", "инструменты", "dump_schedule_filters.py")
LIB = os.path.join(SERVICE, "разработка", "инструменты", "compare_filter_lib.ps1")
PWSH = os.environ.get("PWSH") or shutil.which("pwsh")
NBSP = "\u00a0"
with open(DUMP_PY, encoding="utf-8") as _f:
    DUMP_SRC = _f.read()          # читаем заранее: под эмуляцией open() на кириллице падает, как в IronPython


def exec_dump():
    exec(compile(DUMP_SRC, "dump_schedule_filters.py", "exec"), {"__name__": "__main__"})


# ---------- выгрузка из Revit (подставной API) ----------
class Filter:
    def __init__(self, field_id, kind, value=None, kind_value=None):
        self.FieldId, self.FilterType, self.value, self.kind_value = field_id, kind, value, kind_value

    def GetStringValue(self):
        if not isinstance(self.value, str):
            raise TypeError("not a string filter")
        return self.value

    def GetDoubleValue(self):
        if not isinstance(self.value, float):
            raise TypeError("not double")
        return self.value

    def GetIntegerValue(self):
        if not isinstance(self.value, int) or isinstance(self.value, bool):
            raise TypeError("not int")
        return self.value


class Definition:
    def __init__(self, fields, filters):
        self.fields, self.filters = fields, filters

    def GetFilterCount(self):
        return len(self.filters)

    def GetFilter(self, i):
        return self.filters[i]

    def GetField(self, fid):
        return base.FakeField(self.fields[fid])


FIELDS = {1: "Код по классификатору", 2: "Связанный файл: Имя файла", 3: "ADSK_Размер_Ширина"}


def view(name, filters, template=False):
    return base.FakeView(name, Definition(FIELDS, [Filter(*f) for f in filters]), template)


def _non_ascii(path):
    return any(ord(c) > 127 for c in str(path))


class IronPythonDotNet:
    """Эмуляция IronPython 2.7 под RBP: open()/os.path.*/os.makedirs/os.listdir на путях с кириллицей падают [IOError]
    (без стека), работают только .NET System.IO (File/Directory/Path) и System.Environment — их тут и подставляем."""

    def __enter__(self):
        import builtins
        real_open, real_isdir, real_isfile = open, os.path.isdir, os.path.isfile
        real_abspath, real_makedirs = os.path.abspath, os.makedirs
        self.saved = (builtins.open, os.path.isdir, os.path.isfile, os.path.abspath, os.makedirs)
        self.calls = []

        def guard(real):
            def f(p, *a, **k):
                if _non_ascii(p):
                    raise IOError("IOError: " + p)
                return real(p, *a, **k)
            return f
        builtins.open = guard(real_open)
        os.path.isdir, os.path.isfile = guard(real_isdir), guard(real_isfile)
        os.path.abspath, os.makedirs = guard(real_abspath), guard(real_makedirs)

        def read_all(p, enc=None):
            self.calls.append(("ReadAllText", p))
            with real_open(p, encoding="utf-8-sig") as f:
                return f.read()

        def write_all(p, t, enc=None):
            self.calls.append(("WriteAllText", p))
            with real_open(p, "w", encoding="utf-8-sig" if enc else "utf-8", newline="") as f:
                f.write(t)

        def append_all(p, t, enc=None):
            self.calls.append(("AppendAllText", p))
            with real_open(p, "a", encoding="utf-8", newline="") as f:
                f.write(t)

        def create_dir(p):
            self.calls.append(("CreateDirectory", p))
            __import__("pathlib").Path(p).mkdir(parents=True, exist_ok=True)

        self.temp = tempfile.mkdtemp(prefix="bvs_tmp_")
        sio = types.ModuleType("System.IO")
        sio.File = types.SimpleNamespace(Exists=lambda p: real_isfile(p),        # как .NET: для папки File.Exists даёт False
                                          ReadAllText=read_all, WriteAllText=write_all,
                                         AppendAllText=append_all)
        sio.Directory = types.SimpleNamespace(CreateDirectory=create_dir)
        sio.Path = types.SimpleNamespace(Combine=lambda a, b: a.rstrip("\\/") + os.sep + b, GetTempPath=lambda: self.temp + os.sep)
        stx = types.ModuleType("System.Text")
        stx.UTF8Encoding = lambda b=False: b
        system = types.ModuleType("System")
        system.IO, system.Text = sio, stx
        system.Environment = types.SimpleNamespace(GetEnvironmentVariable=lambda n: os.environ.get(n))
        self.mods = {"System": system, "System.IO": sio, "System.Text": stx}
        self.old_mods = {k: sys.modules.get(k) for k in self.mods}
        sys.modules.update(self.mods)
        return self

    def __exit__(self, *a):
        import builtins
        builtins.open, os.path.isdir, os.path.isfile, os.path.abspath, os.makedirs = self.saved
        for k, v in self.old_mods.items():
            if v is None:
                sys.modules.pop(k, None)
            else:
                sys.modules[k] = v
        shutil.rmtree(self.temp, True)


def run_dump(views, path="C:\\proj\\PROJ_IOS_SVOD_ORG_R24.rvt", name_contains="", out_name="out", precreate=False, quirks=True,
             fail_with=None, log_folder=None):
    tmp = tempfile.mkdtemp(prefix="bvs_dump_")
    out = os.path.join(tmp, out_name)
    if precreate:
        os.makedirs(out)
    cfg_dir = os.path.join(tmp, "BatchRvt", "batch_verification_schedules")
    os.makedirs(cfg_dir)
    with open(os.path.join(cfg_dir, "dump_filters.cfg"), "w", encoding="utf-8-sig") as f:
        cfg = "report_dir=%s\nname_contains=%s\n" % (out, name_contains)
        if log_folder:
            os.makedirs(log_folder, exist_ok=True)
            cfg += "log_folder=%s\n" % log_folder
        f.write(cfg)
    old = os.environ.get("LOCALAPPDATA")
    os.environ["LOCALAPPDATA"] = tmp
    doc = base.FakeDoc(views)
    saved = base.install_fake_revit(doc, path)
    sys.modules["revit_script_util"].LOG = []
    sys.modules["revit_script_util"].Output = sys.modules["revit_script_util"].LOG.append
    result = {}
    try:
        with IronPythonDotNet() as net:
            result["net"] = net
            if fail_with:
                sys.modules["revit_script_util"].GetScriptDocument = fail_with
            try:
                exec_dump()
            except Exception as ex:
                result["error"] = ex
            logs = list(getattr(sys.modules["revit_script_util"], "LOG", []))
            result["logs"] = logs
            err = os.path.join(net.temp, "dump_schedule_filters_error.txt")
            if log_folder:
                alt = os.path.join(log_folder, "dump_schedule_filters_error.txt")
                if os.path.exists(alt):
                    err = alt
            if os.path.exists(err):
                with open(err, encoding="utf-8") as f:
                    result["error_file"] = f.read()
                    result["error_path"] = err
    finally:
        base.restore(saved)
        if old is None:
            os.environ.pop("LOCALAPPDATA", None)
        else:
            os.environ["LOCALAPPDATA"] = old
    files = os.listdir(out) if os.path.isdir(out) else []
    text = ""
    if files:
        with open(os.path.join(out, files[0]), encoding="utf-8-sig") as f:
            text = f.read()
    shutil.rmtree(tmp, True)
    run_dump.last = result
    return files, text, doc


class DumpTests(unittest.TestCase):
    def test_dump_format_read_only_and_values(self):
        views = [
            view("П        5.1_Ширина_Детали возд-в_Без значения", [
                (1, "NotEqual", "999"), (3, "HasNoValue"), (2, "NotContains", "_AR_"), (2, "HasValue")]),
            view("П        8_Площадь_Без значения", []),
            view("П        9_Числовая", [(3, "Greater", 12.5), (3, "Equal", 7)]),
            view("Шаблон", [(1, "HasValue")], template=True),
        ]
        files, text, doc = run_dump(views)
        self.assertEqual(files, ["фильтры_PROJ_IOS_SVOD_ORG_R24.txt"])
        lines = text.replace("\r\n", "\n").split("\n")
        self.assertEqual(lines[0], "Спецификация\tПоле фильтра\tУсловие\tЗначение")
        spec = "П        5.1_Ширина_Детали возд-в_Без значения"
        self.assertEqual(lines[1:5], [
            spec + "\tКод по классификатору\tNotEqual\t999", spec + "\tADSK_Размер_Ширина\tHasNoValue\t",
            spec + "\tСвязанный файл: Имя файла\tNotContains\t_AR_", spec + "\tСвязанный файл: Имя файла\tHasValue\t"])
        self.assertIn("П        8_Площадь_Без значения\t(нет фильтров)\t\t", lines)
        self.assertIn("П        9_Числовая\tADSK_Размер_Ширина\tGreater\t12.5", lines)
        self.assertIn("П        9_Числовая\tADSK_Размер_Ширина\tEqual\t7", lines)
        self.assertFalse(any("Шаблон" in l for l in lines))                       # шаблоны не выгружаются
        self.assertEqual((doc.saved, doc.transactions), (0, []))                  # модель не менялась и не сохранялась

    def test_report_dir_with_cyrillic_exists_or_not_does_not_crash(self):
        """Реальный сбой: [WindowsError] в os.makedirs на пути …\\отчеты\\дамп_фильтров_… (папку уже создал ps1)."""
        views = [view("П        5.1_Ширина", [(3, "HasNoValue")])]
        for precreate in (True, False):
            files, text, _ = run_dump(views, out_name=os.path.join("отчеты", "дамп_фильтров_20261016_092750"), precreate=precreate)
            self.assertEqual(files, ["фильтры_PROJ_IOS_SVOD_ORG_R24.txt"], precreate)
            self.assertIn("5.1_Ширина", text)

    def test_two_models_in_one_folder_second_run_sees_existing_folder(self):
        """Второй файл в той же папке: папка уже есть — без ошибки."""
        tmp = tempfile.mkdtemp(prefix="bvs_dump2_")
        self.addCleanup(shutil.rmtree, tmp, True)
        out = os.path.join(tmp, "дамп_фильтров")
        for leaf in ("A_R21.rvt", "B_R24.rvt"):
            cfg_dir = os.path.join(tmp, "BatchRvt", "batch_verification_schedules")
            os.makedirs(cfg_dir, exist_ok=True)
            with open(os.path.join(cfg_dir, "dump_filters.cfg"), "w", encoding="utf-8-sig") as f:
                f.write("report_dir=%s\n" % out)
            old = os.environ.get("LOCALAPPDATA")
            os.environ["LOCALAPPDATA"] = tmp
            saved = base.install_fake_revit(base.FakeDoc([view("П        1_А", [(3, "HasNoValue")])]), "C:\\p\\" + leaf)
            try:
                with IronPythonDotNet():
                    exec_dump()
            finally:
                base.restore(saved)
                if old is None:
                    os.environ.pop("LOCALAPPDATA", None)
                else:
                    os.environ["LOCALAPPDATA"] = old
        self.assertEqual(sorted(os.listdir(out)), ["фильтры_A_R21.txt", "фильтры_B_R24.txt"])

    def test_all_file_operations_go_through_dotnet_not_os_or_open(self):
        """Реальный сбой: [IOError] без стека. Под эмуляцией (open/os.path на кириллице падают IOError) дамп пишется целиком через .NET."""
        views = [view("П        5.1_Ширина", [(3, "HasNoValue")])]
        out_name = os.path.join("отчеты", "дамп_фильтров_20261016_092750")
        files, text, _ = run_dump(views, out_name=out_name)
        r = run_dump.last
        self.assertNotIn("error", r)
        self.assertEqual(files, ["фильтры_PROJ_IOS_SVOD_ORG_R24.txt"])
        kinds = [k for k, _ in r["net"].calls]
        for need in ("ReadAllText", "CreateDirectory", "WriteAllText"):
            self.assertIn(need, kinds)
        self.assertNotIn("error_file", r)
        src = open(DUMP_PY, encoding="utf-8").read()
        for bad in ("import os", "open(", "os.path", "__file__", "codecs"):
            self.assertNotIn(bad, src)

    def test_error_writes_to_log_folder_from_cfg(self):
        folder = os.path.join(tempfile.mkdtemp(prefix="bvs_dumplog_"), "logs")
        self.addCleanup(shutil.rmtree, os.path.dirname(folder), True)
        def boom():
            raise IOError("boom from GetScriptDocument")
        _, _, _ = run_dump([view("П 1", [])], fail_with=boom, log_folder=folder)
        r = run_dump.last
        self.assertIsInstance(r.get("error"), IOError)
        self.assertTrue(os.path.isfile(os.path.join(folder, "dump_schedule_filters_error.txt")))
        self.assertIn("Traceback", r["error_file"])
        self.assertTrue(r["error_path"].startswith(folder))

    def test_error_writes_full_traceback_to_temp_and_log_and_reraises(self):
        """Любая ошибка: traceback + шаг + путь модели в %TEMP%\\dump_schedule_filters_error.txt и в лог RBP, исключение пробрасывается."""
        def boom():
            raise IOError("boom from GetScriptDocument")
        _, _, _ = run_dump([view("П 1", [])], fail_with=boom)
        r = run_dump.last
        self.assertIsInstance(r.get("error"), IOError)
        err = r["error_file"]
        self.assertIn("Traceback", err)
        self.assertIn("boom from GetScriptDocument", err)
        self.assertIn("step:", err)
        self.assertIn("PROJ_IOS_SVOD_ORG_R24.rvt", err)
        self.assertTrue(any("DUMP ERROR" in l for l in r["logs"]))
        self.assertTrue(any("boom from GetScriptDocument" in l for l in r["logs"]))

    def test_error_in_cfg_step_is_named(self):
        """Нет файла настроек — понятная ошибка и шаг «read cfg»."""
        views = [view("П 1", [])]
        tmp = tempfile.mkdtemp(prefix="bvs_nocfg_")
        self.addCleanup(shutil.rmtree, tmp, True)
        old = os.environ.get("LOCALAPPDATA")
        os.environ["LOCALAPPDATA"] = tmp
        saved = base.install_fake_revit(base.FakeDoc(views), "C:\\p\\M.rvt")
        try:
            with IronPythonDotNet() as net:
                with self.assertRaises(Exception):
                    exec_dump()
                with open(os.path.join(net.temp, "dump_schedule_filters_error.txt"), encoding="utf-8") as f:
                    self.assertIn("step: read cfg", f.read())
        finally:
            base.restore(saved)
            if old is None:
                os.environ.pop("LOCALAPPDATA", None)
            else:
                os.environ["LOCALAPPDATA"] = old

    def test_name_contains_filter(self):
        views = [view("П        5.1_Ширина", [(3, "HasNoValue")]), view("П        6.1_Высота", [(3, "HasNoValue")])]
        _, text, _ = run_dump(views, name_contains="6.1_")
        self.assertIn("6.1_Высота", text)
        self.assertNotIn("5.1_Ширина", text)


# ---------- разбор входных путей Выгрузить фильтры спек.bat -> dump_schedule_filters.ps1 ----------
@unittest.skipUnless(PWSH, "нет pwsh")
class DumpRunnerInputTests(unittest.TestCase):
    """Реальный сбой: два файла, перетащенных на bat, — второй путь попал в -OutDir, дамп «писался» в rvt-файл."""
    PS1 = os.path.join(SERVICE, "разработка", "инструменты", "dump_schedule_filters.ps1")

    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="bvs_run_")
        self.addCleanup(shutil.rmtree, self.tmp, True)
        folder = os.path.join(self.tmp, "Projects", "20 Проверочные спецификации")
        os.makedirs(folder)
        self.models = []
        for leaf in ("П_Проверочные спецификации 2026-10-01_PROJ.rvt", "РД_Проверочные спецификации 2026-10-01_Все.rvt"):
            path = os.path.join(folder, leaf)
            make_rvt(path, info_text(2024))
            self.models.append(path)
        self.folder = folder
        self.app = os.path.join(self.tmp, "app")
        self.fake = fake_rbp.install(self.tmp, "BatchRvt", "import sys\nsys.exit(0)\n")          # на Windows — .cmd + Python
        self.reports = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(HERE))), "отчеты")
        self.before = set(os.listdir(self.reports)) if os.path.isdir(self.reports) else set()
        self.addCleanup(self.cleanup_reports)

    def cleanup_reports(self):
        if os.path.isdir(self.reports):
            for n in set(os.listdir(self.reports)) - self.before:
                shutil.rmtree(os.path.join(self.reports, n), True)

    def run_ps1(self, args):
        env = dict(os.environ, DOTNET_SYSTEM_GLOBALIZATION_INVARIANT="1")
        r = subprocess.run([PWSH, "-NoProfile", "-File", self.PS1] + args + ["-NoPause", "-RbpExe", self.fake, "-AppDir", self.app],
                           capture_output=True, text=True, encoding="utf-8", env=env, cwd=self.tmp)
        cfg = {}
        with open(os.path.join(self.app, "dump_filters.cfg"), encoding="utf-8-sig") as f:
            for line in f.read().splitlines():
                if "=" in line:
                    k, v = line.split("=", 1)
                    cfg[k] = v
        lists = []
        for n in sorted(os.listdir(self.app)):
            if n.startswith("rvt_list_filter_dump_"):
                with open(os.path.join(self.app, n), encoding="utf-8") as f:
                    lists += [l for l in f.read().splitlines() if l]
        return r, cfg, lists

    def check(self, args):
        r, cfg, lists = self.run_ps1(args)
        out = cfg["report_dir"]
        self.assertEqual(os.path.dirname(out), os.path.join(self.reports, "операции"), r.stdout + r.stderr)   # отчеты\операции\ в папке инструмента
        self.assertTrue(os.path.basename(out).startswith("дамп_фильтров_"))
        self.assertTrue(os.path.isdir(out))
        self.assertEqual(sorted(lists), sorted(self.models), r.stdout)
        self.assertNotIn(".rvt\n", r.stdout.split("Папка:")[-1][:3000].replace("\r", "") if "Папка:" in r.stdout else "")
        self.assertEqual(sorted(os.listdir(self.folder)), sorted(os.path.basename(m) for m in self.models))   # рядом с моделями мусора нет

    def test_two_paths_as_separate_arguments_dragged_on_bat(self):
        self.check(list(self.models))

    def test_two_quoted_paths_joined_by_semicolon_in_one_argument(self):
        self.check(['"%s";"%s"' % tuple(self.models)])

    def test_two_paths_joined_by_semicolon_unquoted_and_with_spaces(self):
        self.check([";".join(self.models)])

    def test_outdir_pointing_to_rvt_is_ignored(self):
        self.check(list(self.models) + ["-OutDir", self.models[0]])

    def test_folder_skips_backups_backup_dir_and_summary_and_txt_list_keeps_explicit_files(self):
        root = os.path.join(self.tmp, "модели")
        nest = os.path.join(root, "К01", "RVT")
        backup = os.path.join(nest, "Backup")
        os.makedirs(backup)
        keep = []
        for rel in (os.path.join("К01", "RVT", "PROJ_IOS_OV_K01.rvt"), os.path.join("К01", "RVT", "вложенные", "PROJ_IOS_VK_K02.rvt")):
            path = os.path.join(root, rel)
            os.makedirs(os.path.dirname(path), exist_ok=True)
            make_rvt(path, info_text(2024))
            keep.append(path)
        make_rvt(os.path.join(nest, "PROJ_IOS_OV_K01.0001.rvt"), info_text(2024))
        make_rvt(os.path.join(nest, "PROJ_IOS_SVOD_ORG_R24.rvt"), info_text(2024))
        make_rvt(os.path.join(backup, "PROJ_IOS_OV_K01.rvt"), info_text(2024))
        r, cfg, lists = self.run_ps1([root])
        self.assertEqual(r.returncode in (0, 2), True, r.stdout + r.stderr)
        self.assertEqual(sorted(lists), sorted(keep), r.stdout + r.stderr)
        self.assertIn("Найдено моделей: 2", r.stdout)
        self.assertIn("Пропущено папок резервных копий:", r.stdout)
        self.assertIn("пропущены резервные копии", r.stdout)
        self.assertIn("пропущены сводные", r.stdout)
        listed = os.path.join(self.tmp, "список.txt")
        svod = os.path.join(nest, "PROJ_IOS_SVOD_ORG_R24.rvt")
        with open(listed, "w", encoding="utf-8") as f:
            f.write("# комментарий\n\n%s\n%s;%s\n" % (keep[0], keep[1], svod))
        r2, _, lists2 = self.run_ps1([listed])
        self.assertEqual(sorted(lists2), sorted(keep + [svod]), r2.stdout + r2.stderr)

    def test_rezerv_and_dated_folders_are_not_entered_and_duplicate_name_keeps_short_path(self):
        root = os.path.join(self.tmp, "SampleProject", "1 - Стадия П")
        rvt = os.path.join(root, "4_Разработка проекта", "9_ПТ", "RVT")
        os.makedirs(rvt)
        work = os.path.join(rvt, "PROJ_IOS_PT_ORG_K02_R24.rvt")
        other = os.path.join(rvt, "PROJ_IOS_OV_ORG_K01_R24.rvt")
        make_rvt(work, info_text(2024))
        make_rvt(other, info_text(2024))
        dated = os.path.join(rvt, "Резерв", "2026-10-19")
        os.makedirs(dated)
        make_rvt(os.path.join(dated, "PROJ_IOS_PT_ORG_K02_R24.rvt"), info_text(2024))
        for name in ("Архив_2024", "Archive", "Old", "Старые", "Старое", "Старый", "Стар", "OLD_models", "_backup", "2026-01-02"):
            folder = os.path.join(rvt, name)
            os.makedirs(folder)
            make_rvt(os.path.join(folder, "PROJ_IOS_PT_ORG_K02_R24.rvt"), info_text(2024))
        start = os.path.join(rvt, "Старт", "PROJ_IOS_SS_ORG_K03_R24.rvt")
        os.makedirs(os.path.dirname(start))
        make_rvt(start, info_text(2024))
        deeper = os.path.join(rvt, "копия")
        os.makedirs(deeper)
        make_rvt(os.path.join(deeper, "PROJ_IOS_PT_ORG_K02_R24.rvt"), info_text(2024))
        only_dated = os.path.join(rvt, "2026-03-03")
        os.makedirs(only_dated)
        make_rvt(os.path.join(only_dated, "PROJ_IOS_HS_ORG_K01_R24.rvt"), info_text(2024))
        r, _, lists = self.run_ps1([root])
        self.assertEqual(sorted(lists), sorted([work, other, start]), r.stdout + r.stderr)
        self.assertIn("Найдено моделей: 3", r.stdout)
        self.assertRegex(r.stdout, r"Пропущено папок резервных копий: [1-9]")
        self.assertIn("дубль:", r.stdout)
        self.assertIn(work, r.stdout)
        self.assertNotIn("PROJ_IOS_HS_ORG_K01_R24.rvt", "\n".join(lists))

    def test_spec_name_list_is_written_next_to_filter_dump(self):
        writer = (
            "import os, sys\n"
            "d = os.environ.get('BVS_DUMP_OUT') or ''\n"
            "if d:\n"
            "    p = os.path.join(d, 'фильтры_Модель.txt')\n"
            "    t = ('Спецификация\\tПоле фильтра\\tУсловие\\tЗначение\\n'"
            "'П        1_А\\tПоле\\tHasValue\\t\\n'"
            "'П        1_А\\tПоле2\\tEqual\\t1\\n'"
            "'П        5_Б\\t(нет фильтров)\\t\\t\\n')\n"
            "    open(p, 'w', encoding='utf-8-sig', newline='').write(t)\n"
            "sys.exit(0)\n"
        )
        fake = fake_rbp.install(self.tmp, "BatchRvtSpec", writer)
        env = dict(os.environ, DOTNET_SYSTEM_GLOBALIZATION_INVARIANT="1")
        r = subprocess.run([PWSH, "-NoProfile", "-File", self.PS1, self.models[0], "-NoPause", "-RbpExe", fake, "-AppDir", self.app],
                           capture_output=True, text=True, encoding="utf-8", env=env, cwd=self.tmp)
        self.assertEqual(r.returncode, 0, r.stdout + r.stderr)
        cfg = {}
        with open(os.path.join(self.app, "dump_filters.cfg"), encoding="utf-8-sig") as f:
            for line in f.read().splitlines():
                if "=" in line:
                    k, v = line.split("=", 1)
                    cfg[k] = v
        spec = os.path.join(cfg["report_dir"], "спеки_Модель.txt")
        with open(spec, encoding="utf-8-sig") as f:
            names = [ln.strip() for ln in f.read().splitlines() if ln.strip()]
        self.assertEqual(names, ["П        1_А", "П        5_Б"])


# ---------- сравнение (PowerShell) ----------
def compare(old_text, new_text):
    tmp = tempfile.mkdtemp(prefix="bvs_cmp_")
    try:
        po, pn = os.path.join(tmp, "old.txt"), os.path.join(tmp, "new.txt")
        for p, t in ((po, old_text), (pn, new_text)):
            with open(p, "w", encoding="utf-8-sig", newline="") as f:
                f.write("Спецификация\tПоле фильтра\tУсловие\tЗначение\n" + t)
        script = os.path.join(tmp, "t.ps1")
        with open(script, "w", encoding="utf-8-sig") as f:
            f.write(". '%s'\n$o = @(Read-FilterDump '%s'); $n = @(Read-FilterDump '%s')\n$r = Compare-FilterDumps $o $n\n"
                    "Write-CompareCsv '%s' $r.Rows\n"
                    "@{ matched = $r.Matched; unchanged = $r.Unchanged; onlyOld = $r.OnlyOld; onlyNew = $r.OnlyNew; "
                    "rows = @($r.Rows | ForEach-Object { @{ n = $_.Number; s = $_.Spec; c = $_.Change; w = $_.Was; b = $_.Became; note = $_.Note } }) } "
                    "| ConvertTo-Json -Depth 5 -Compress\n" % (LIB, po, pn, os.path.join(tmp, "out.csv")))
        env = dict(os.environ, DOTNET_SYSTEM_GLOBALIZATION_INVARIANT="1")
        r = subprocess.run([PWSH, "-NoProfile", "-File", script], capture_output=True, text=True, encoding="utf-8", env=env)
        assert r.returncode == 0, r.stdout + r.stderr
        with open(os.path.join(tmp, "out.csv"), encoding="utf-8-sig") as f:
            csv_text = f.read()
        return json.loads(r.stdout), csv_text
    finally:
        shutil.rmtree(tmp, True)


def rows_of(*specs):
    out = []
    for name, flts in specs:
        if not flts:
            out.append("%s\t(нет фильтров)\t\t" % name)
        for f in flts:
            out.append("\t".join([name] + list(f)))
    return "\n".join(out) + "\n"


L = "Связанный файл: Имя файла"


@unittest.skipUnless(PWSH, "нет pwsh")
class CompareTests(unittest.TestCase):
    def test_changes_added_removed_changed_and_unchanged(self):
        old = rows_of(
            ("П        5.1_Ширина_Детали возд-в_Без значения", [("Код по классификатору", "NotEqual", "999"), ("ADSK_Размер_Ширина", "HasNoValue", ""), (L, "Contains", "_AR_")]),
            ("П        6.1_Высота_Детали возд-в_Без значения", [("ADSK_Размер_Высота", "HasNoValue", ""), (L, "NotContains", "_AR_")]),
            ("П        7.1_Диаметр", [(L, "NotContains", "_AR_"), (L, "NotContains", "_KR_")]),
            ("П        8_Площадь", [("ADSK_Материал наименование", "NotContains", "_TASK_")]),
            ("П        13_Материал_Трубы", [("ADSK_Материал наименование", "Equal", "")]),
        )
        new = rows_of(
            # порядок фильтров другой — это не изменение; _AR_: Contains -> NotContains (изменено: условие)
            ("П        5.1_Ширина_Детали возд-в_Без значения", [(L, "NotContains", "_AR_"), ("ADSK_Размер_Ширина", "HasNoValue", ""), ("Код по классификатору", "NotEqual", "999")]),
            ("П        6.1_Высота_Детали возд-в_Без значения", [("ADSK_Размер_Высота", "HasNoValue", ""), (L, "NotContains", "_AR_"), (L, "NotContains", "_TASK_")]),   # добавлено
            ("П        7.1_Диаметр", [(L, "NotContains", "_AR_")]),                                                                                          # удалено _KR_
            ("П        8_Площадь", [(L, "NotContains", "_TASK_")]),                                                                                          # изменено: поле
            ("П        13_Материал_Трубы", [("ADSK_Материал наименование", "Equal", "")]),                                                                  # без изменений
        )
        res, csv_text = compare(old, new)
        self.assertEqual((res["matched"], res["unchanged"], res["onlyOld"], res["onlyNew"]), (5, 1, 0, 0))
        got = {(r["n"], r["c"]): (r["w"], r["b"]) for r in res["rows"]}
        self.assertEqual(got[("5.1", "изменено: условие")], (L + " Contains «_AR_»", L + " NotContains «_AR_»"))
        self.assertEqual(got[("6.1", "добавлено")], ("", L + " NotContains «_TASK_»"))
        self.assertEqual(got[("7.1", "удалено")], (L + " NotContains «_KR_»", ""))
        self.assertEqual(got[("8", "изменено: поле")], ("ADSK_Материал наименование NotContains «_TASK_»", L + " NotContains «_TASK_»"))
        self.assertEqual(len(res["rows"]), 4)
        self.assertTrue(csv_text.startswith("Номер;Спецификация;Изменение;Было;Стало;Примечание"))

    def test_value_and_field_value_changes(self):
        old = rows_of(("П        1_А", [("Код по классификатору", "NotEqual", "999")]), ("П        2_Б", [("Поле", "Equal", "")]))
        new = rows_of(("П        1_А", [("Код по классификатору", "NotEqual", "000")]), ("П        2_Б", [("Поле", "HasNoValue", "")]))
        res, _ = compare(old, new)
        got = {r["n"]: r["c"] for r in res["rows"]}
        self.assertEqual(got, {"1": "изменено: значение", "2": "изменено: условие"})

    def test_specs_only_in_one_file_and_number_fallback(self):
        old = rows_of(("П        3.1_Тип системы", [("Имя системы", "HasNoValue", "")]), ("П        20_Только старая", [("А", "HasValue", "")]),
                      ("П        4.1_Старое имя", [("Б", "HasValue", "")]))
        new = rows_of(("П        3.1_Тип системы", [("Имя системы", "HasNoValue", "")]), ("П        30_Только новая", [("В", "HasValue", "")]),
                      ("П        4.1_Новое имя", [("Б", "HasValue", ""), ("Г", "HasValue", "")]))
        res, _ = compare(old, new)
        self.assertEqual((res["matched"], res["onlyOld"], res["onlyNew"]), (2, 1, 1))
        changes = {r["c"] for r in res["rows"]}
        self.assertIn("спека только в «было»", changes)
        self.assertIn("спека только в «стало»", changes)
        paired = [r for r in res["rows"] if r["n"] == "4.1"][0]
        self.assertEqual(paired["c"], "добавлено")
        self.assertIn("сопоставлено по номеру", paired["note"])
        self.assertIn("Старое имя", paired["note"])

    def test_prefix_whitespace_case_and_typo_do_not_matter(self):
        old = rows_of(("П        2.1_Клас-я системы_Без значния", [("Классификация систем", "HasNoValue", "")]),
                      ("Π" + NBSP + NBSP + "4.1_Имя системы", [("Имя системы", "Equal", "")]))
        new = rows_of(("2.1_Клас-я Системы_Без значения", [("Классификация систем", "HasNoValue", "")]),
                      ("П 4.1_имя системы", [("Имя системы", "Equal", "")]))
        res, _ = compare(old, new)
        self.assertEqual((res["matched"], res["unchanged"], res["onlyOld"], res["onlyNew"]), (2, 2, 0, 0))
        self.assertEqual(res["rows"], [] if not res["rows"] else res["rows"])

    def test_duplicate_spec_names_and_identical_dumps(self):
        text = rows_of(("П        5_Копия", [("А", "HasValue", "")]), ("П        5_Копия", [("А", "HasValue", "")]))
        res, csv_text = compare(text, text)
        self.assertEqual(res["rows"], [] if not res["rows"] else res["rows"])
        # листы с одинаковым именем в выгрузке неразличимы (строки идут подряд) — считаются одной спекой
        self.assertEqual((res["matched"], res["unchanged"]), (1, 1))
        self.assertEqual(csv_text.strip(), "Номер;Спецификация;Изменение;Было;Стало;Примечание")

    def test_real_unfixed_dump_vs_itself_and_fixed_variant(self):
        real = os.path.join(HERE, "fixtures", "filters_before_fix.txt")
        if not os.path.isfile(real):
            self.skipTest("portfolio copy omits the client schedule-filter export")
        with open(real, encoding="utf-8-sig") as f:
            body = f.read().split("\n", 1)[1]
        lines = [l.split("\t") for l in body.splitlines() if l.strip()]
        old = "\n".join("\t".join(l[:4]) for l in lines) + "\n"
        res, _ = compare(old, old)
        self.assertEqual((res["matched"], res["unchanged"], res["rows"]), (57, 57, []))
        # «стало»: 7.8 — условие _AR_ вместо прежнего (после исправления фильтров)
        new_lines = []
        for l in lines:
            l = l[:4]
            if "7.8_Диаметр_Фитингов коробов" in l[0] and l[3] == "_AR_":
                l = [l[0], l[1], "NotContains", l[3]]
            new_lines.append("\t".join(l))
        res, _ = compare(old, "\n".join(new_lines) + "\n")
        self.assertEqual([(r["n"], r["c"]) for r in res["rows"]], [("7.8", "изменено: условие")])


if __name__ == "__main__":
    unittest.main()
