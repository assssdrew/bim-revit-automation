"""Шаг 5: «В проекте» только в [inproject]. Перенос прежнего выбор_<проект>.txt (только выбор, совпадают, различаются, повторный запуск),
резервная копия отчеты\\проект\\_backup_<время> до записи и вне автоочистки, открытие проекта в окне."""
import hashlib
import os
import shutil
import sys
import tempfile
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from test_chain import pwsh, PWSH  # noqa: E402

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(HERE)))
TOOL = os.path.join(ROOT, "служебное", "скрипты")
CFG = os.path.join(TOOL, "project_config.ps1")
DIST = os.path.join(TOOL, "project_dist_ui.ps1")
LAYOUT = os.path.join(TOOL, "report_layout.ps1")
UI = os.path.join(TOOL, "ui_schedules.ps1")

SPECS = ["1_Классификатор_Все", "3.1_Имя_Все", "7.8_Диаметр_Все", "13_Материал_Трубы_Без значения"]


def q(p):
    return p.replace("'", "''")


def read(p):
    with open(p, encoding="utf-8-sig") as f:
        return f.read()


def sha(p):
    with open(p, "rb") as f:
        return hashlib.sha256(f.read()).hexdigest()


@unittest.skipUnless(PWSH, "нет pwsh")
class SelectionMigrationTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="bvs_mig5_")
        self.addCleanup(shutil.rmtree, self.tmp, True)
        self.rep = os.path.join(self.tmp, "отчеты")
        self.proj = os.path.join(self.rep, "проект")
        self.hist = os.path.join(self.rep, "история")
        os.makedirs(self.proj)
        os.makedirs(self.hist)
        self.cfg = os.path.join(self.proj, "распределение_PROJ.txt")
        self.sel = os.path.join(self.proj, "выбор_PROJ.txt")

    def write_cfg(self, inproject=None, name="PROJ", extra=()):
        lines = ["# настройка", "[project]", "stage=П", "template=C:\\t\\общий.rvt", "[specs]"] + SPECS
        if inproject is not None:
            lines += ["[inproject]"] + list(inproject)
        lines += ["[cells]"] + list(extra)
        path = os.path.join(self.proj, "распределение_%s.txt" % name)
        with open(path, "w", encoding="utf-8-sig", newline="") as f:
            f.write("\r\n".join(lines) + "\r\n")
        return path

    def write_sel(self, names, name="PROJ"):
        path = os.path.join(self.proj, "выбор_%s.txt" % name)
        with open(path, "w", encoding="utf-8-sig", newline="") as f:
            f.write("\r\n".join(names) + ("\r\n" if names else ""))
        return path

    def put(self, rel, text="x"):
        path = os.path.join(self.rep, *rel.split("/"))
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, "w", encoding="utf-8") as f:
            f.write(text)
        return path

    def run_ps(self, body):
        script = os.path.join(self.tmp, "t.ps1")
        with open(script, "w", encoding="utf-8-sig") as f:
            f.write(". '%s'\n. '%s'\n$rd = '%s'\n$cfgPath = '%s'\n"
                    "function InFile([string]$p) { $x = Read-ProjectConfig $p; $n = Get-ProjectInProjectNames $x; if ($null -eq $n) { return 'ALL' }; return ($n -join ' | ') }\n%s"
                    % (q(CFG), q(LAYOUT), q(self.rep), q(self.cfg), body))
        r = pwsh("-File", script)
        self.assertEqual(r.returncode, 0, r.stdout + r.stderr)
        return [l for l in r.stdout.splitlines() if l.strip()]

    def backups(self):
        return sorted(n for n in os.listdir(self.proj) if n.startswith("_backup_"))

    def migrated(self):
        return sorted(n for n in os.listdir(self.proj) if ".migrated_" in n)

    # ---- случаи переноса ----
    def test_only_selection_file_is_imported_into_inproject_and_renamed(self):
        self.write_cfg(inproject=None)
        self.write_sel(["# комментарий", "П        3.1_Имя_Все", "7.8_Диаметр_Все", "99_Нет в общем файле"])
        orig_cfg = sha(self.cfg)
        self.put("история/история_PROJ.txt", "история")
        self.put("история/задание_журнал_PROJ.txt", "журнал")
        self.put("проект/google_PROJ.cfg", "webapp_url=x\nsecret=y\n")
        out = self.run_ps(r"""
$r = Invoke-ProjectStateMigration $rd ([datetime]'2026-11-10T14:05:00')
'action: ' + $r.Results[0].Action
'note: ' + ($r.Notes -join ' ')
'in: ' + (InFile $cfgPath)
'unknown: ' + ((Read-ProjectConfig $cfgPath).SelectionUnknown -join ',')
'backup: ' + (Split-Path -Leaf $r.Backup)
""")
        self.assertEqual(out[0], "action: imported")
        self.assertEqual(out[1], "note: Выбор «В проекте» проекта PROJ перенесён в настройку проекта (технический шаг, один раз).")
        self.assertEqual(out[2], "in: 3.1_Имя_Все | 7.8_Диаметр_Все | 99_Нет в общем файле")     # имя не из общего файла не потеряно
        self.assertEqual(out[3], "unknown: 99_Нет в общем файле")
        self.assertEqual(out[4], "backup: _backup_2026-11-10_14-05")
        self.assertFalse(os.path.exists(self.sel))
        self.assertEqual(self.migrated(), ["выбор_PROJ.migrated_2026-11-10.txt"])           # переименован, не удалён
        self.assertIn("99_Нет в общем файле", read(os.path.join(self.proj, "выбор_PROJ.migrated_2026-11-10.txt")))
        bdir = os.path.join(self.proj, "_backup_2026-11-10_14-05")
        # копия сделана ДО записи: в ней прежняя настройка без [inproject] и сам файл выбора
        self.assertEqual(sha(os.path.join(bdir, "распределение_PROJ.txt")), orig_cfg)
        self.assertTrue(os.path.exists(os.path.join(bdir, "выбор_PROJ.txt")))
        self.assertTrue(os.path.exists(os.path.join(bdir, "google_PROJ.cfg")))
        self.assertEqual(read(os.path.join(bdir, "история", "история_PROJ.txt")), "история")
        self.assertEqual(read(os.path.join(bdir, "история", "задание_журнал_PROJ.txt")), "журнал")
        log = read(os.path.join(bdir, "перенос_выбора.txt"))
        self.assertIn("[inproject] был пуст — перенесено из выбор_PROJ.txt", log)
        self.assertIn("выбор_PROJ.txt -> выбор_PROJ.migrated_2026-11-10.txt", log)
        self.assertNotIn("secret=y", log)
        self.assertEqual(read(os.path.join(self.hist, "история_PROJ.txt")), "история")       # оригиналы на месте

    def test_empty_inproject_section_is_filled_from_selection(self):
        self.write_cfg(inproject=[])
        self.write_sel(["1_Классификатор_Все"])
        out = self.run_ps(r"""
$r = Invoke-ProjectStateMigration $rd
'action: ' + $r.Results[0].Action
'in: ' + (InFile $cfgPath)
""")
        self.assertEqual(out, ["action: imported", "in: 1_Классификатор_Все"])

    def test_equal_selection_is_only_renamed_and_settings_are_untouched(self):
        self.write_cfg(inproject=["3.1_Имя_Все", "7.8_Диаметр_Все"])
        self.write_sel(["РД        3.1_имя_ВСЕ", "7.8_Диаметр_Все"])            # префикс стадии и регистр не мешают сравнению
        before = sha(self.cfg)
        out = self.run_ps(r"""
$r = Invoke-ProjectStateMigration $rd
'action: ' + $r.Results[0].Action
'note: [' + ($r.Notes -join ' ') + ']'
""")
        self.assertEqual(out, ["action: renamed", "note: []"])
        self.assertEqual(sha(self.cfg), before)
        self.assertEqual(len(self.migrated()), 1)
        self.assertEqual(len(self.backups()), 1)                                    # копия и перед переименованием

    def test_different_selection_keeps_inproject_and_logs_the_difference(self):
        self.write_cfg(inproject=["3.1_Имя_Все", "7.8_Диаметр_Все"])
        self.write_sel(["3.1_Имя_Все", "13_Материал_Трубы_Без значения"])
        before = sha(self.cfg)
        out = self.run_ps(r"""
$r = Invoke-ProjectStateMigration $rd
'action: ' + $r.Results[0].Action
'in: ' + (InFile $cfgPath)
""")
        self.assertEqual(out, ["action: kept", "in: 3.1_Имя_Все | 7.8_Диаметр_Все"])          # остаётся то, что читали скрипты Revit
        self.assertEqual(sha(self.cfg), before)
        self.assertFalse(os.path.exists(self.sel))
        log = read(os.path.join(self.proj, self.backups()[0], "перенос_выбора.txt"))
        self.assertIn("различаются, оставлен [inproject] (2 спек)", log)
        self.assertIn("Только в файле выбора (1): 13_Материал_Трубы_Без значения", log)
        self.assertIn("Только в [inproject] (1): 7.8_Диаметр_Все", log)

    def test_second_start_does_nothing_and_old_migrated_file_is_never_overwritten(self):
        self.write_cfg(inproject=None)
        self.write_sel(["3.1_Имя_Все"])
        self.put("проект/выбор_PROJ.migrated_2026-11-10.txt", "прежняя копия")     # уже есть файл с таким именем
        out = self.run_ps(r"""
$now = [datetime]'2026-11-10T09:00:00'
$r1 = Invoke-ProjectStateMigration $rd $now
'first: ' + $r1.Results[0].Action + ' -> ' + (Split-Path -Leaf $r1.Results[0].Renamed)
$h = (Get-FileHash -LiteralPath $cfgPath).Hash
$r2 = Invoke-ProjectStateMigration $rd $now
'second: ' + $r2.Results.Count + ' backup=[' + $r2.Backup + ']'
'cfg same: ' + ($h -eq (Get-FileHash -LiteralPath $cfgPath).Hash)
'in: ' + (InFile $cfgPath)
""")
        self.assertEqual(out[0], "first: imported -> выбор_PROJ.migrated_2026-11-10_2.txt")
        self.assertEqual(out[1:], ["second: 0 backup=[]", "cfg same: True", "in: 3.1_Имя_Все"])
        self.assertEqual(read(os.path.join(self.proj, "выбор_PROJ.migrated_2026-11-10.txt")), "прежняя копия")
        self.assertEqual(len(self.backups()), 1)                                    # повторный запуск копий не плодит

    def test_several_projects_share_one_backup_and_a_lone_selection_file_is_left_alone(self):
        self.write_cfg(inproject=None)
        self.write_sel(["3.1_Имя_Все"])
        site = self.write_cfg(inproject=["1_Классификатор_Все"], name="SITE")
        self.write_sel(["1_Классификатор_Все"], name="SITE")
        lone = self.write_sel(["3.1_Имя_Все"], name="OLD")                          # нет распределение_OLD.txt
        out = self.run_ps(r"""
$r = Invoke-ProjectStateMigration $rd
($r.Results | ForEach-Object { $_.Project + '=' + $_.Action }) -join ','
'site: ' + (InFile '%s')
""" % q(site))
        self.assertEqual(out[0], "PROJ=imported,OLD=none,SITE=renamed")
        self.assertEqual(out[1], "site: 1_Классификатор_Все")
        self.assertEqual(len(self.backups()), 1)
        self.assertTrue(os.path.exists(lone))
        self.assertEqual([n.split(".migrated_")[0] for n in self.migrated()], ["выбор_PROJ", "выбор_SITE"])

    def test_without_backup_nothing_is_written(self):
        self.write_cfg(inproject=None)
        self.write_sel(["3.1_Имя_Все"])
        before = sha(self.cfg)
        out = self.run_ps(r"""
try { [void](Invoke-ProjectSelectionMigration $cfgPath { throw 'диск полон' }); 'no error' } catch { 'error: ' + $_.Exception.Message }
try { [void](Invoke-ProjectSelectionMigration $cfgPath { return '' }); 'no error' } catch { 'error: ' + $_.Exception.Message }
""")
        self.assertEqual(out, ["error: диск полон", "error: нет резервной копии — перенос не выполнен"])
        self.assertEqual(sha(self.cfg), before)
        self.assertTrue(os.path.exists(self.sel))
        self.assertEqual(self.migrated(), [])

    def test_no_selection_files_no_backup_no_writes(self):
        self.write_cfg(inproject=["3.1_Имя_Все"])
        before = sha(self.cfg)
        out = self.run_ps(r"""
$r = Invoke-ProjectStateMigration $rd
'results: ' + $r.Results.Count + ' backup=[' + $r.Backup + ']'
""")
        self.assertEqual(out, ["results: 0 backup=[]"])
        self.assertEqual(sha(self.cfg), before)
        self.assertEqual(sorted(os.listdir(self.proj)), ["распределение_PROJ.txt"])

    # ---- резервная копия ----
    def test_backup_skips_previous_backups_and_is_never_cleaned_up(self):
        self.write_cfg(inproject=None)
        self.write_sel(["3.1_Имя_Все"])
        self.put("проект/_backup_2025-01-01_10-00/распределение_PROJ.txt", "старая копия")
        self.put("проект/_backup_2025-01-01_10-00/битый.bad.xlsx", "x")
        self.put("проект/справочник_разделов.txt", "справочник")
        self.put("служебные/20240101_000000/a.txt", "старое")
        out = self.run_ps(r"""
$r = Invoke-ProjectStateMigration $rd ([datetime]'2026-11-10T14:05:00')
$b = Split-Path -Leaf $r.Backup
'backup: ' + $b
'nested old backup: ' + (Test-Path -LiteralPath (Join-Path $r.Backup '_backup_2025-01-01_10-00'))
$hk = Invoke-BvsHousekeeping -ReportDir $rd -KeepDays 1 -Now ([datetime]'2030-01-01')
'aux removed: ' + (@($hk.Aux) -join ',')
'bad removed: [' + (@($hk.Bad) -join ',') + ']'
$null = Initialize-ReportLayout $rd 1 '' ([datetime]'2030-01-01')
'backups: ' + (@(Get-ChildItem -LiteralPath (Join-Path $rd 'проект') -Directory -Filter '_backup_*' | ForEach-Object Name | Sort-Object) -join ',')
""")
        self.assertEqual(out[0], "backup: _backup_2026-11-10_14-05")
        self.assertEqual(out[1], "nested old backup: False")                           # прежние копии в новую не копируются
        self.assertEqual(out[2], "aux removed: 20240101_000000")                        # автоочистка работает, но…
        self.assertEqual(out[3], "bad removed: []")                                     # …в _backup_ ничего не трогает
        self.assertEqual(out[4], "backups: _backup_2025-01-01_10-00,_backup_2026-11-10_14-05")
        self.assertTrue(os.path.exists(os.path.join(self.proj, "_backup_2025-01-01_10-00", "битый.bad.xlsx")))
        self.assertEqual(read(os.path.join(self.proj, "_backup_2026-11-10_14-05", "справочник_разделов.txt")), "справочник")

    def test_backup_folder_is_not_a_project_and_not_a_cache(self):
        self.write_cfg(inproject=None)
        self.write_sel(["3.1_Имя_Все"])
        out = self.run_ps(r"""
$r = Invoke-ProjectStateMigration $rd
'projects: ' + ((Get-ProjectNames $rd) -join ',')
'cache: ' + (Split-Path -Leaf (Find-ProjectCachePath $rd '' 'C:\t\общий.rvt'))
""")
        self.assertEqual(out, ["projects: PROJ", "cache: распределение_PROJ.txt"])


@unittest.skipUnless(PWSH, "нет pwsh")
class WindowStartTests(unittest.TestCase):
    """Окно: перенос при открытии проекта (Load-DistProject), строка в статусе, «Взять выбор из сводной» пишет только [inproject] и только по «Сохранить»."""

    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="bvs_mig5w_")
        self.addCleanup(shutil.rmtree, self.tmp, True)
        self.rep = os.path.join(self.tmp, "отчеты")
        self.proj = os.path.join(self.rep, "проект")
        os.makedirs(self.proj)
        self.cfg = os.path.join(self.proj, "распределение_PROJ.txt")
        with open(self.cfg, "w", encoding="utf-8-sig", newline="") as f:
            f.write("\r\n".join(["[project]", "stage=П", "[specs]"] + SPECS) + "\r\n")
        with open(os.path.join(self.proj, "выбор_PROJ.txt"), "w", encoding="utf-8-sig", newline="") as f:
            f.write("3.1_Имя_Все\r\n")

    def run_ps(self, body):
        script = os.path.join(self.tmp, "t.ps1")
        with open(script, "w", encoding="utf-8-sig") as f:
            f.write(r"""
. '%s'
. '%s'
. '%s'
$ReportDir = '%s'
$cfgPath = '%s'
function Show-DistGrid { }
function Save-AsgSelection { }
class FakeBox { [string]$Text = ''; [bool]$Checked = $false; [int]$SelectedIndex = 0; [bool]$Enabled = $true; $Items = (New-Object System.Collections.ArrayList); $SelectedItem = $null }
$ui = @{}
foreach ($k in 'DistStage', 'DistTemplate', 'DistModelsDir', 'Summary', 'Status') { $ui[$k] = [FakeBox]::new() }
$script:PartialUiReady = $false; $script:SummaryDone = $false; $script:DistRsnModels = @()
function InFile([string]$p) { $x = Read-ProjectConfig $p; $n = Get-ProjectInProjectNames $x; if ($null -eq $n) { return 'ALL' }; return ($n -join ' | ') }
""" % (q(CFG), q(DIST), q(LAYOUT), q(self.rep), q(self.cfg)) + body)
        r = pwsh("-File", script)
        self.assertEqual(r.returncode, 0, r.stdout + r.stderr)
        return [l for l in r.stdout.splitlines() if l.strip()]

    def test_opening_the_project_migrates_once_and_take_from_summary_waits_for_save(self):
        out = self.run_ps(r"""
$ui.Status.Text = 'Проект PROJ.'
Load-DistProject $cfgPath
Add-DistStateMigrationNote
'status: ' + $ui.Status.Text
'mem: ' + (@($script:Project.Specs | Where-Object { Test-ProjectSpecIn $script:Project $_ }) -join ' | ')
'note cleared: ' + ($script:StateMigrationNote -eq '')
$h = (Get-FileHash -LiteralPath $cfgPath).Hash
Load-DistProject $cfgPath                                            # повторное открытие: переносить нечего
'second note: [' + $script:StateMigrationNote + ']'
'same after reopen: ' + ($h -eq (Get-FileHash -LiteralPath $cfgPath).Hash)
# «Взять выбор из сводной» (часть после чтения Revit): правка в памяти, файл ждёт «Сохранить»
Enter-DistDirty
$r = Set-ProjectSelectionFromNames $script:Project @('П        1_Классификатор_Все', '5.5_Только в сводной')
'dirty: ' + $script:DistDirty
'file same: ' + ($h -eq (Get-FileHash -LiteralPath $cfgPath).Hash)
'commit: ' + (Commit-DistDirty)
'in file: ' + (InFile $cfgPath)
'selection files: ' + (@(Get-ChildItem -LiteralPath (Join-Path $ReportDir 'проект') -Filter 'выбор_*.txt' | Where-Object { $_.Name -notmatch 'migrated' }).Count)
""")
        self.assertTrue(out[0].startswith("status: Проект PROJ. Выбор «В проекте» проекта PROJ перенесён в настройку проекта (технический шаг, один раз). Резервная копия: "), out[0])
        self.assertEqual(out[1], "mem: 3.1_Имя_Все")
        self.assertEqual(out[2], "note cleared: True")
        self.assertEqual(out[3:5], ["second note: []", "same after reopen: True"])
        self.assertEqual(out[5:9], ["dirty: True", "file same: True", "commit: True", "in file: 1_Классификатор_Все | 5.5_Только в сводной"])
        self.assertEqual(out[9], "selection files: 0")

    def test_ui_start_shows_the_note_and_runs_before_the_project_is_read(self):
        ui = read(UI)
        start = ui[ui.index("$projectNames = @(Get-ProjectNames $ReportDir)"):]
        self.assertLess(start.index("Load-DistProject"), start.index("Add-DistStateMigrationNote"))
        self.assertLess(start.index("Add-DistStateMigrationNote"), start.index("[void]$form.ShowDialog()"))
        d = read(DIST)
        sw = d[d.index("function Switch-Project"):d.index("function Invoke-NewProject")]
        self.assertIn("Add-DistStateMigrationNote", sw)
        for btn in ('"Взять выбор из сводной"', '"Сохранить как новый проект"', '"Посчитать элементы"', '"Классификатор"', '"Отменить"'):
            self.assertIn(btn, d)                                               # кнопки на месте


if __name__ == "__main__":
    unittest.main()
