# -*- coding: utf-8 -*-
"""Код по классификатору: абсолютный UNC, пропуск сводной и служебных, без сохранения если путь уже тот."""
import os
import re
import shutil
import sys
import tempfile
import types
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from test_chain import PWSH  # noqa: E402
from test_two_tabs import run  # noqa: E402

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(HERE)))
TOOL = os.path.join(ROOT, "служебное", "скрипты")
PY = os.path.join(TOOL, "set_assembly_code.py")
UI = os.path.join(TOOL, "ui_schedules.ps1")
EXPECTED = "\\\\server\\share\\BIM\\classifier.txt"


class ExtRef(object):
    def __init__(self, stored, kind):
        self.PathType = kind
        self._stored = stored

    def GetPath(self):
        return self._stored


class Table(object):
    def __init__(self, stored="", kind="", present=False, behavior="absolute"):
        self.present = present
        self.behavior = behavior
        self.loads = []
        self.ref = ExtRef(stored, kind) if present else None

    def IsExternalFileReference(self):
        return self.present and self.ref is not None

    def GetExternalFileReference(self):
        if self.ref is None:
            raise Exception("none")
        return self.ref

    def LoadFrom(self, reference, results):
        self.loads.append(reference)
        if self.behavior == "fail":
            return "Failure"
        if self.behavior == "relative":
            self.present = True
            self.ref = ExtRef("UniformatClassifications.txt", "Relative")
            return "Success"
        self.present = True
        self.ref = ExtRef(reference["path"], "Absolute")
        return "Success"


class Doc(object):
    def __init__(self, table):
        self.table = table
        self.IsWorkshared = True
        self.IsDetached = False
        self.synced = 0
        self.saved = 0
        self.commits = 0
        self.rollbacks = 0

    def GetWorksharingCentralModelPath(self):
        return "central"

    def SynchronizeWithCentral(self, transact, opts):
        self.synced += 1

    def SaveAs(self, path, opts):
        self.saved += 1


class Transaction(object):
    def __init__(self, doc, name):
        self.doc = doc
        self.name = name

    def Start(self):
        pass

    def Commit(self):
        self.doc.commits += 1

    def RollBack(self):
        self.doc.rollbacks += 1


class Bag(object):
    def __init__(self, *args, **kwargs):
        self.Comment = ""
        self.SaveLocalFile = False
        self.OverwriteExistingFile = False
        self.SaveAsCentral = False

    def SetRelinquishOptions(self, opts):
        pass

    def SetWorksharingOptions(self, opts):
        pass


def install_runtime(local_appdata):
    clr = types.ModuleType("clr")
    clr.AddReference = lambda name: None
    sys.modules["clr"] = clr

    rsu = types.ModuleType("revit_script_util")
    rsu.Output = lambda text: None
    rsu._doc = None
    rsu._path = ""
    rsu.GetScriptDocument = lambda: rsu._doc
    rsu.GetRevitFilePath = lambda: rsu._path
    sys.modules["revit_script_util"] = rsu

    io = types.ModuleType("System.IO")

    class StreamReader(object):
        def __init__(self, path, detect):
            self._f = open(path, "rb")

        def ReadToEnd(self):
            data = self._f.read()
            if data.startswith(b"\xef\xbb\xbf"):
                data = data[3:]
            return data.decode("utf-8")

        def Close(self):
            self._f.close()

    class StreamWriter(object):
        def __init__(self, path, append, encoding):
            mode = "ab" if append else "wb"
            folder = os.path.dirname(path)
            if folder and not os.path.isdir(folder):
                os.makedirs(folder)
            self._f = open(path, mode)

        def Write(self, text):
            if not isinstance(text, str):
                text = str(text)
            self._f.write(text.encode("utf-8"))

        def Close(self):
            self._f.close()

    class FileObj(object):
        @staticmethod
        def Exists(path):
            return os.path.isfile(path)

    io.StreamReader = StreamReader
    io.StreamWriter = StreamWriter
    io.File = FileObj
    sys.modules["System.IO"] = io
    system = types.ModuleType("System")
    system.IO = io
    text = types.ModuleType("System.Text")

    class UTF8Encoding(object):
        def __init__(self, bom=False):
            self.bom = bom

    text.UTF8Encoding = UTF8Encoding
    system.Text = text
    sys.modules["System"] = system
    sys.modules["System.Text"] = text

    created = []

    class ExternalResourceReference(object):
        @staticmethod
        def CreateLocalResource(doc, resource_type, path, path_type):
            ref = {"path": path, "path_type": path_type, "resource": resource_type}
            created.append(ref)
            return ref

    class ModelPathUtils(object):
        @staticmethod
        def ConvertUserVisiblePathToModelPath(path):
            return path

        @staticmethod
        def ConvertModelPathToUserVisiblePath(path):
            return path

    class AssemblyCodeTable(object):
        @staticmethod
        def GetAssemblyCodeTable(doc):
            return doc.table

    class PathType(object):
        Absolute = "Absolute"
        Relative = "Relative"

    class BuiltIn(object):
        AssemblyCodeTable = "AssemblyCodeTable"

    class ExternalResourceTypes(object):
        BuiltInExternalResourceTypes = BuiltIn

    class KeyBasedTreeEntriesLoadResults(object):
        pass

    class WorksharingUtils(object):
        @staticmethod
        def RelinquishOwnership(*args):
            pass

    db = types.ModuleType("Autodesk.Revit.DB")
    db.ExtRef = ExtRef
    db.Table = Table
    db.Transaction = Transaction
    db.ExternalResourceReference = ExternalResourceReference
    db.ModelPathUtils = ModelPathUtils
    db.AssemblyCodeTable = AssemblyCodeTable
    db.PathType = PathType
    db.ExternalResourceTypes = ExternalResourceTypes
    db.KeyBasedTreeEntriesLoadResults = KeyBasedTreeEntriesLoadResults
    db.WorksharingUtils = WorksharingUtils
    db.created = created
    for name in ("RelinquishOptions", "SynchronizeWithCentralOptions", "TransactWithCentralOptions", "SaveAsOptions", "WorksharingSaveAsOptions"):
        setattr(db, name, Bag)
    revit = types.ModuleType("Autodesk.Revit")
    revit.DB = db
    autodesk = types.ModuleType("Autodesk")
    autodesk.Revit = revit
    sys.modules["Autodesk"] = autodesk
    sys.modules["Autodesk.Revit"] = revit
    sys.modules["Autodesk.Revit.DB"] = db

    os.environ["LOCALAPPDATA"] = local_appdata
    with open(PY, encoding="utf-8") as handle:
        source = handle.read()
    mod = types.ModuleType("set_assembly_code")
    mod.__file__ = PY
    exec(compile(source, PY, "exec"), mod.__dict__)
    return mod, rsu, db


class DispositionTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="bvs_asm_")
        self.addCleanup(shutil.rmtree, self.tmp, True)
        self.mod, self.rsu, self.db = install_runtime(self.tmp)

    def test_constant_is_the_proj_unc(self):
        self.assertEqual(self.mod.CLASSIFIER_PROJECT, "PROJ")
        self.assertEqual(self.mod.CLASSIFIER_PATH, EXPECTED)
        self.assertTrue(self.mod.is_unc(self.mod.CLASSIFIER_PATH))

    def test_same_absolute_ignores_case_and_slashes_but_not_relative(self):
        target = self.mod.CLASSIFIER_PATH
        folded = target.replace("\\", "/").replace("ПРОЕКТНЫЙ", "проектный")
        self.assertTrue(self.mod.same_absolute(folded, "Absolute", target))
        self.assertFalse(self.mod.same_absolute("UniformatClassifications.txt", "Relative", target))
        self.assertFalse(self.mod.same_absolute(target, "Relative", target))
        self.assertFalse(self.mod.same_absolute(target, "Content", target))
        self.assertFalse(self.mod.same_absolute(r"\\other\share\file.txt", "Absolute", target))
        self.assertFalse(self.mod.same_absolute("", "Absolute", target))

    def test_which_models_are_sections(self):
        m = self.mod
        section = r"\\srv\ОВ\RVT\PROJ_IOS_OV_ORG_K01_R24.rvt"
        self.assertEqual(m.model_disposition(section), ("section", ""))
        self.assertEqual(m.model_disposition(r"C:\m\Старт\RVT\PROJ_IOS_SS_START.rvt")[0], "section")
        self.assertEqual(m.model_disposition(r"C:\proj\ОВ\RVT\PROJ_SOMETHING.rvt")[0], "section")
        for path in (
            r"C:\m\PROJ_IOS_SVOD_ORG_R24.rvt",
            r"C:\m\PROJ_IOS_TASK_R24.rvt",
            r"C:\m\PROJ_IOS_BASE_R24.rvt",
            r"C:\m\Резерв\PROJ_IOS_OV_ORG_K01_R24.rvt",
            r"C:\m\Backup\PROJ_IOS_VK_X.rvt",
            r"C:\m\Архив_2024\PROJ_IOS_VK_X.rvt",
            r"C:\m\Archive\PROJ_IOS_SS_X.rvt",
            r"C:\m\Old\PROJ_IOS_OT_X.rvt",
            r"C:\m\Старые\PROJ_IOS_OT_X.rvt",
            r"C:\m\OLD_models\PROJ_IOS_HS_X.rvt",
            r"C:\m\2026-01-02\PROJ_IOS_OV_X.rvt",
            r"C:\m\PROJ_IOS_OV_ORG_K01_R24.0001.rvt",
        ):
            self.assertEqual(m.model_disposition(path)[0], "skipped_not_section", path)
        self.assertEqual(m.model_disposition(r"C:\m\SITE_IOS_OV_ORG_K01.rvt"), ("error", "not_target_project"))


class ApplyTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="bvs_asm_")
        self.addCleanup(shutil.rmtree, self.tmp, True)
        self.report_dir = os.path.join(self.tmp, "work")
        os.makedirs(self.report_dir)
        self.mod, self.rsu, self.db = install_runtime(self.tmp)
        # На Linux «\» в LOCAL_CFG — не разделитель: писать ровно в тот путь, который видит скрипт.
        cfg_path = self.mod.LOCAL_CFG
        folder = os.path.dirname(cfg_path)
        if folder and not os.path.isdir(folder):
            os.makedirs(folder)
        with open(cfg_path, "w", encoding="utf-8") as f:
            f.write("report_dir=%s\r\nrun_id=run1\r\n" % self.report_dir)
        self.mod.classifier_readable = lambda path: True

    def _run(self, path, table):
        doc = Doc(table)
        self.rsu._doc = doc
        self.rsu._path = path
        self.mod.main()
        csv_path = os.path.join(self.report_dir, "run1_классификатор.csv")
        text = open(csv_path, encoding="utf-8-sig").read().strip().splitlines()
        return doc, text

    def _write_cfg(self, extra):
        with open(self.mod.LOCAL_CFG, "w", encoding="utf-8") as f:
            f.write("report_dir=%s\r\nrun_id=run1\r\n%s" % (self.report_dir, extra))

    def test_path_and_project_from_run_cfg(self):
        own = "\\\\srv\\share\\Классификатор SITE.txt"
        self._write_cfg("classifier_path=%s\r\nclassifier_project=SITE\r\n" % own)
        doc, text = self._run(r"\\srv\ОВ\RVT\SITE_IOS_OV_ORG_K01.rvt", Table())
        cells = text[1].split(";")
        self.assertEqual(cells[2], own)                                       # путь из run.cfg, не встроенный
        self.assertEqual(cells[3], "updated")
        self.assertEqual(self.db.created[-1]["path"], own)
        self.assertEqual(self.mod.classifier_settings({}), (self.mod.CLASSIFIER_PATH, "PROJ"))    # нет строк — встроенные значения
        self.assertEqual(self.mod.classifier_settings({"classifier_path": " ", "classifier_project": ""}), (self.mod.CLASSIFIER_PATH, "PROJ"))

    def test_without_run_cfg_lines_falls_back_to_builtin(self):
        self._write_cfg("")
        doc, text = self._run(r"\\srv\ОВ\RVT\PROJ_IOS_OV_ORG_K01_R24.rvt", Table())
        cells = text[1].split(";")
        self.assertEqual(cells[2], self.mod.CLASSIFIER_PATH)
        self.assertEqual(cells[3], "updated")
        doc, text = self._run(r"\\srv\ОВ\RVT\SITE_IOS_OV_ORG_K01.rvt", Table())
        self.assertEqual(text[-1].split(";")[3:5], ["error", "not_target_project"])   # чужой префикс без classifier_project — как раньше

    def test_already_absolute_does_not_load_or_sync(self):
        folded = self.mod.CLASSIFIER_PATH.replace("\\", "/").replace("ПРОЕКТНЫЙ", "проектный")
        table = Table(stored=folded, kind="Absolute", present=True)
        doc, lines = self._run(r"C:\m\PROJ_IOS_OV_ORG_K01_R24.rvt", table)
        self.assertEqual(lines[-1].split(";")[3], "already_set")
        self.assertEqual(table.loads, [])
        self.assertEqual(doc.synced, 0)
        self.assertEqual(doc.commits, 0)
        self.assertEqual(self.db.created, [])

    def test_relative_path_is_replaced_and_synced(self):
        table = Table(stored="UniformatClassifications.txt", kind="Relative", present=True)
        doc, lines = self._run(r"C:\m\PROJ_IOS_VK_ORG_K01_R24.rvt", table)
        cells = lines[-1].split(";")
        self.assertEqual(cells[3], "updated")
        self.assertEqual(cells[1], "UniformatClassifications.txt")
        self.assertEqual(cells[2], self.mod.CLASSIFIER_PATH)
        self.assertEqual(len(table.loads), 1)
        self.assertEqual(self.db.created[-1]["path_type"], "Absolute")
        self.assertEqual(self.db.created[-1]["path"], self.mod.CLASSIFIER_PATH)
        self.assertEqual(doc.commits, 1)
        self.assertEqual(doc.synced, 1)
        self.assertEqual(doc.saved, 0)

    def test_api_that_stores_relative_is_rolled_back(self):
        table = Table(stored="", kind="", present=False, behavior="relative")
        doc, lines = self._run(r"C:\m\PROJ_IOS_OV_ORG_K01_R24.rvt", table)
        cells = lines[-1].split(";")
        self.assertEqual(cells[3], "error")
        self.assertEqual(cells[4], "stored_not_absolute")
        self.assertEqual(doc.commits, 0)
        self.assertGreaterEqual(doc.rollbacks, 1)
        self.assertEqual(doc.synced, 0)

    def test_summary_and_backup_are_not_saved(self):
        table = Table()
        doc, lines = self._run(r"C:\m\PROJ_IOS_SVOD_ORG_R24.rvt", table)
        self.assertEqual(lines[-1].split(";")[3], "skipped_not_section")
        self.assertEqual(table.loads, [])
        self.assertEqual(doc.synced, 0)
        table2 = Table()
        doc2, lines2 = self._run(r"C:\m\Резерв\PROJ_IOS_OV_ORG_K01_R24.rvt", table2)
        self.assertEqual(lines2[-1].split(";")[3], "skipped_not_section")
        self.assertEqual(doc2.synced, 0)

    def test_missing_classifier_does_not_open_the_table(self):
        # setUp подменяет проверку файла; здесь нужна настоящая — UNC на этой машине не читается.
        self.mod, self.rsu, self.db = install_runtime(self.tmp)
        calls = []
        real_get = self.db.AssemblyCodeTable.GetAssemblyCodeTable

        def counting(doc):
            calls.append(doc)
            return real_get(doc)

        self.db.AssemblyCodeTable.GetAssemblyCodeTable = staticmethod(counting)
        doc, lines = self._run(r"C:\m\PROJ_IOS_OV_ORG_K01_R24.rvt", Table())
        self.assertEqual(lines[-1].split(";")[3], "error")
        self.assertIn("classifier_file_missing", lines[-1])
        self.assertEqual(calls, [])
        self.assertEqual(doc.synced, 0)


class SourceTests(unittest.TestCase):
    def test_script_and_launcher_contract(self):
        import ast
        with open(PY, encoding="utf-8") as handle:
            py = handle.read()
        with open(UI, encoding="utf-8-sig") as handle:
            ui = handle.read()
        calls = []
        for node in ast.walk(ast.parse(py)):
            if isinstance(node, ast.Call) and isinstance(node.func, ast.Attribute) and node.func.attr in ("lower", "upper"):
                calls.append(node.lineno)
        self.assertEqual(calls, [])
        self.assertNotIn("codecs.open", py)
        self.assertIn("PathType.Absolute", py)
        self.assertIn(".LoadFrom(", py)
        self.assertNotIn("ViewSchedule", py)
        self.assertIn('elseif ($Mode -eq "classifier")', ui)
        self.assertIn("Test-ClassifierReadable $classifierPath", ui)
        self.assertIn("Модели не открываются.", ui)
        self.assertIn('if ($Mode -eq "summary" -or $Mode -eq "sections") { $verTemplate = $source }', ui)
        self.assertIn("--create_new_local", ui)
        self.assertIn('$worksets = "close_all"', ui)
        # Та же распаковка, что ConvertFrom-PyString: строка константы читается как один UNC.
        match = re.search(r'(?m)^CLASSIFIER_PATH\s*=\s*u"(.*)"\s*$', py)
        self.assertIsNotNone(match)
        raw = match.group(1)
        chars = []
        i = 0
        while i < len(raw):
            if raw[i] == "\\" and i + 1 < len(raw) and raw[i + 1] == "\\":
                chars.append("\\")
                i += 2
                continue
            chars.append(raw[i])
            i += 1
        self.assertEqual("".join(chars), EXPECTED)
        for name in ("ui_schedules.ps1", "project_dist_ui.ps1"):
            raw = open(os.path.join(TOOL, name), "rb").read(3)
            self.assertEqual(raw, b"\xef\xbb\xbf", name)


@unittest.skipUnless(PWSH, "нет pwsh")
class LauncherPathTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="bvs_asm_ps_")
        self.addCleanup(shutil.rmtree, self.tmp, True)

    def test_classifier_path_comes_from_project_settings_with_proj_default(self):
        folder = os.path.join(self.tmp, "Классификатор")
        os.makedirs(folder)
        sample = os.path.join(folder, "MR_Классификатор (для Revit).txt")
        with open(sample, "wb") as f:
            f.write(b"1\tname\n")
        cfg_path = os.path.join(self.tmp, "проект", "распределение_SITE.txt")
        out = run(self.tmp, r"""
$c = New-ProjectConfig 'П'
'proj default: ' + (Get-ProjectClassifierPath $c 'PROJ')
'proj lower: ' + (Get-ProjectClassifierPath $c 'proj')
'other empty: [' + (Get-ProjectClassifierPath $c 'SITE') + ']'
$c.Classifier = '\\srv\share\Классификатор SITE.txt'
'own: ' + (Get-ProjectClassifierPath $c 'SITE')
'own beats default: ' + (Get-ProjectClassifierPath $c 'PROJ')
Write-ProjectConfig $c '%s'
$r = Read-ProjectConfig '%s'
'roundtrip: ' + $r.Classifier
'copy: ' + (Copy-ProjectConfig $r).Classifier
$e = New-ProjectConfig 'П'; Write-ProjectConfig $e '%s'
'empty not written: ' + (-not ([IO.File]::ReadAllText('%s').Contains('classifier=')))
'no: ' + (Test-ClassifierReadable 'C:\no\such\classifier.txt')
'yes: ' + (Test-ClassifierReadable '%s')
""" % (cfg_path, cfg_path, cfg_path, cfg_path, sample.replace("'", "''")), ["Test-ClassifierReadable"])
        self.assertEqual(out[0], "proj default: " + EXPECTED)
        self.assertEqual(out[1], "proj lower: " + EXPECTED)
        self.assertEqual(out[2], "other empty: []")
        self.assertEqual(out[3], "own: \\\\srv\\share\\Классификатор SITE.txt")
        self.assertEqual(out[4], "own beats default: \\\\srv\\share\\Классификатор SITE.txt")
        self.assertEqual(out[5], "roundtrip: \\\\srv\\share\\Классификатор SITE.txt")
        self.assertEqual(out[6], "copy: \\\\srv\\share\\Классификатор SITE.txt")
        self.assertEqual(out[7], "empty not written: True")                     # прежние файлы без строки classifier= не меняются по формату
        self.assertEqual(out[8], "no: False")
        self.assertEqual(out[9], "yes: True")


class ClassifierSettingsWiringTests(unittest.TestCase):
    def test_ui_takes_path_from_project_and_passes_it_in_run_cfg(self):
        with open(UI, encoding="utf-8-sig") as handle:
            ui = handle.read()
        block = ui[ui.index('elseif ($Mode -eq "classifier") {'):ui.index("if (($Mode -eq \"summary\" -or $Mode -eq \"sections\") -and (Get-InsertModeValue) -eq \"replace\")")]
        self.assertIn("Get-ProjectClassifierPath $script:Project $wantProject", block)
        self.assertNotIn("Get-ClassifierScriptSetting", ui)                    # код set_assembly_code.py окно больше не читает
        self.assertIn('"classifier_path={0}" -f $classifierPath', block)
        self.assertIn('"classifier_project={0}" -f $wantProject', block)
        self.assertLess(block.index("Test-ClassifierReadable $classifierPath"), block.index("Get-DistProjectModelPaths"))   # проверка файла до моделей
        with open(os.path.join(TOOL, "project_config.ps1"), encoding="utf-8-sig") as handle:
            cfg = handle.read()
        # значение по умолчанию в окне и запасное в скрипте — один и тот же путь
        m = re.search(r'"PROJ"\s*=\s*"([^"]+)"', cfg)
        self.assertIsNotNone(m)
        self.assertEqual(m.group(1), EXPECTED)


if __name__ == "__main__":
    unittest.main()
