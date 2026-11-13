"""«В проекте» только в [inproject] настройки проекта (прежний выбор_<проект>.txt больше не пишется): сохранение по «Сохранить», восстановление при загрузке, неизвестные имена, «Взять выбор из сводной»."""
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
P = "П        "


def read(p):
    with open(p, encoding="utf-8-sig") as f:
        return f.read()


SETUP = r"""
function New-Proj {
    $c = New-ProjectConfig 'П'
    Update-ProjectSpecs $c @('П        1_Классификатор_Все', 'П        3.1_Имя_Все', 'П        3.1_Имя_Пустые', 'П        7.8_Диаметр_Все', 'П        13_Материал_Трубы_Без значения')
    $c.InProject = $null
    $c.Summary = $null
    return $c
}
function InNames($c) { (@($c.Specs | Where-Object { Test-ProjectSpecIn $c $_ }) -join ' | ') }
function InFile([string]$p) { $x = Read-ProjectConfig $p; $n = Get-ProjectInProjectNames $x; if ($null -eq $n) { return @() }; return @($n) }
"""


@unittest.skipUnless(PWSH, "нет pwsh")
class SelectionFileTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="bvs_sel_")
        self.addCleanup(shutil.rmtree, self.tmp, True)
        self.proj_dir = os.path.join(self.tmp, "отчеты", "проект")
        os.makedirs(self.proj_dir)
        self.sel = os.path.join(self.proj_dir, "выбор_PROJ.txt")

    def run_ps(self, body):
        script = os.path.join(self.tmp, "t.ps1")
        with open(script, "w", encoding="utf-8-sig") as f:
            f.write(". '%s'\n. '%s'\n$P = '%s'\n$cfgPath = '%s'\n$selPath = '%s'\n%s\n%s" % (CFG, DIST, P, os.path.join(self.proj_dir, "распределение_PROJ.txt").replace("'", "''"),
                                                                                    self.sel.replace("'", "''"), SETUP, body))
        r = pwsh("-File", script)
        self.assertEqual(r.returncode, 0, r.stdout + r.stderr)
        return [l for l in r.stdout.splitlines() if l.strip()]

    def write_sel(self, lines):
        with open(self.sel, "w", encoding="utf-8-sig", newline="") as f:
            f.write("\r\n".join(lines) + "\r\n")

    def digest(self):
        with open(self.sel, "rb") as f:
            return hashlib.sha256(f.read()).hexdigest()

    def test_path_is_next_to_settings_and_named_by_project(self):
        # путь прежнего файла выбора нужен только переносу (Invoke-ProjectStateMigration)
        out = self.run_ps(r"""
(Get-ProjectSelectionPath $cfgPath)
(Get-ProjectSelectionPath '')
""")
        self.assertEqual(out[0], self.sel)
        self.assertEqual(len(out), 1)                                          # пустой путь — пустая строка

    def test_save_and_restore_roundtrip_through_inproject(self):
        out = self.run_ps(r"""
function Show-DistGrid { }
$ui = @{
    DistStage = [pscustomobject]@{ Text = 'П' }
    DistTemplate = [pscustomobject]@{ Text = '' }
    DistModelsDir = [pscustomobject]@{ Text = '' }
    Summary = [pscustomobject]@{ Text = '' }
    Status = [pscustomobject]@{ Text = '' }
}
$script:SummaryDone = $false
$script:DistRsnModels = @()
$script:Project = New-Proj
$script:ProjectPath = $cfgPath
Set-ProjectInProjectBulk $script:Project @($script:Project.Specs) $false
[void](Save-DistProject)
'file after clear: ' + @(InFile $cfgPath).Count
Invoke-ProjectSpecInClick $script:Project ($P + '3.1_Имя_Пустые')
Invoke-ProjectSpecInClick $script:Project ($P + '7.8_Диаметр_Все')
[void](Save-DistProject)
'file after clicks: ' + ((InFile $cfgPath) -join ' | ')
$before = (Get-FileHash -LiteralPath $cfgPath -Algorithm SHA256).Hash
Invoke-DistBulk $true
'bulk dirty: ' + $script:DistDirty
'bulk did not write: ' + ((Get-FileHash -LiteralPath $cfgPath -Algorithm SHA256).Hash -eq $before)
[void](Commit-DistDirty)
'file after commit: ' + @(InFile $cfgPath).Count
'clean after commit: ' + (-not $script:DistDirty)
Invoke-DistBulk $false
[void](Commit-DistDirty)
Invoke-ProjectSpecInClick $script:Project ($P + '1_Классификатор_Все'); [void](Save-DistProject)
$script:Project = New-Proj
'restore: ' + (Restore-DistSelection (InFile $cfgPath))
'in project: ' + (InNames $script:Project)
'no selection file: ' + (-not (Test-Path -LiteralPath $selPath))
""")
        self.assertEqual(out[0], "file after clear: 0")
        self.assertEqual(out[1], "file after clicks: 3.1_Имя_Пустые | 7.8_Диаметр_Все")        # имена без префикса стадии, по строке на спеку
        self.assertEqual(out[2], "bulk dirty: True")
        self.assertEqual(out[3], "bulk did not write: True")
        self.assertEqual(out[4], "file after commit: 5")
        self.assertEqual(out[5], "clean after commit: True")
        self.assertEqual(out[6], "restore: Выбор восстановлен из настройки проекта: 1 спек.")
        self.assertEqual(out[7], "in project: 1_Классификатор_Все")
        self.assertEqual(out[8], "no selection file: True")                                       # второго файла больше нет

    def test_unknown_specs_are_listed_and_kept_in_inproject(self):
        with open(os.path.join(self.proj_dir, "распределение_PROJ.txt"), "w", encoding="utf-8-sig", newline="") as f:
            f.write("\r\n".join(["[project]", "stage=П", "[specs]", "1_Классификатор_Все", "3.1_Имя_Все", "7.8_Диаметр_Все",
                                 "[inproject]", "3.1_Имя_Все", "7.8_Диаметр_Все", "99_Нет такой_Все", "20_Тоже нет"]) + "\r\n")
        out = self.run_ps(r"""
$c = Read-ProjectConfig $cfgPath
'unknown: ' + ($c.SelectionUnknown -join '; ')
$script:Project = New-Proj
'restore: ' + (Restore-DistSelection (Get-ProjectInProjectNames $c))
'in project: ' + (InNames $script:Project)
Write-ProjectConfig $script:Project $cfgPath
'file: ' + ((InFile $cfgPath) -join ' | ')
""")
        self.assertEqual(out[0], "unknown: 99_Нет такой_Все; 20_Тоже нет")
        self.assertEqual(out[1], "restore: Выбор восстановлен из настройки проекта: 2 спек. нет в общем файле: 99_Нет такой_Все; 20_Тоже нет.")
        self.assertEqual(out[2], "in project: 3.1_Имя_Все | 7.8_Диаметр_Все")
        self.assertEqual(out[3], "file: 3.1_Имя_Все | 7.8_Диаметр_Все | 99_Нет такой_Все | 20_Тоже нет")   # не теряются при записи

    def test_unknown_names_survive_the_next_save(self):
        out = self.run_ps(r"""
$script:Project = New-Proj
[void](Restore-DistSelection @('3.1_Имя_Все', '99_Нет такой_Все'))
Invoke-ProjectSpecInClick $script:Project ($P + '7.8_Диаметр_Все')      # пользователь отметил ещё одну
Write-ProjectConfig $script:Project $cfgPath
(InFile $cfgPath) -join ' | '
""")
        self.assertEqual(out[0], "3.1_Имя_Все | 7.8_Диаметр_Все | 99_Нет такой_Все")

    def test_no_inproject_gives_hint_and_keeps_marks(self):
        out = self.run_ps(r"""
$script:Project = New-Proj
Invoke-ProjectSpecInClick $script:Project ($P + '3.1_Имя_Все')
'text: ' + (Restore-DistSelection $null)
'marks kept: ' + (InNames $script:Project)
'file exists: ' + (Test-Path -LiteralPath $selPath)
""")
        self.assertEqual(out[0], "text: Выбор не найден. Отметьте спеки или нажмите «Взять выбор из сводной».")
        self.assertEqual(out[1], "marks kept: 1_Классификатор_Все | 3.1_Имя_Пустые | 7.8_Диаметр_Все | 13_Материал_Трубы_Без значения")
        self.assertEqual(out[2], "file exists: False")                                           # файл выбора не создаётся

    def test_take_from_summary_marks_by_names_and_counts_unknown(self):
        out = self.run_ps(r"""
$c = New-Proj
'has selection (all in project): ' + (Test-ProjectHasSelection $c)
$names = @('П        1_Классификатор_Все', 'РД        13_Материал_Трубы_Без значния', '5.5_Нет в общем файле_Все', '5.5_Нет в общем файле_Все', '')
$r = Set-ProjectSelectionFromNames $c $names
'marked: ' + $r.Marked + ' unknown: ' + @($r.Unknown).Count + ' (' + ($r.Unknown -join ',') + ')'
'in project: ' + (InNames $c)
'specs unchanged: ' + $c.Specs.Count          # таблица под сводную не перестраивается
'has selection: ' + (Test-ProjectHasSelection $c)
$c2 = New-ProjectConfig 'П'; Update-ProjectSpecs $c2 @('1_А_Все'); $c2.InProject = @{}
'empty selection: ' + (Test-ProjectHasSelection $c2)
Write-ProjectConfig $c $cfgPath
'file: ' + ((InFile $cfgPath) -join ' | ')
'no selection file: ' + (-not (Test-Path -LiteralPath $selPath))
""")
        self.assertEqual(out[0], "has selection (all in project): True")
        self.assertEqual(out[1], "marked: 2 unknown: 1 (5.5_Нет в общем файле_Все)")
        self.assertEqual(out[2], "in project: 1_Классификатор_Все | 13_Материал_Трубы_Без значения")
        self.assertEqual(out[3], "specs unchanged: 5")
        self.assertEqual(out[4], "has selection: True")
        self.assertEqual(out[5], "empty selection: False")
        self.assertEqual(out[6], "file: 1_Классификатор_Все | 13_Материал_Трубы_Без значения | 5.5_Нет в общем файле_Все")
        self.assertEqual(out[7], "no selection file: True")

    def test_project_file_survives_layout_migration_aux_cleanup_and_cache_reset(self):
        app = os.path.join(self.tmp, "AppData")                        # «кэш окна»: %LOCALAPPDATA%\BatchRvt\batch_verification_schedules
        os.makedirs(app)
        for n in ("ui_settings.txt", "run.cfg", "template_read.cfg", "rvt_list.txt"):
            open(os.path.join(app, n), "w").write("x")
        out = self.run_ps(r"""
. '%s'
$rd = '%s'
$app = '%s'
$script:Project = New-Proj
Set-ProjectInProjectBulk $script:Project @($P + '7.8_Диаметр_Все') $false
Write-ProjectConfig $script:Project $cfgPath
$h0 = (Get-FileHash -LiteralPath $cfgPath -Algorithm SHA256).Hash
$null = Initialize-ReportLayout $rd 30           # перенос старых файлов и очистка служебных папок
[void](Remove-OldAuxDirs (Join-Path $rd 'служебные') 1)
Remove-Item -LiteralPath $app -Recurse -Force    # сброс кэша окна
'cfg same: ' + ($h0 -eq (Get-FileHash -LiteralPath $cfgPath -Algorithm SHA256).Hash)
'in file: ' + @(InFile $cfgPath).Count
""" % (LAYOUT.replace("'", "''"), os.path.join(self.tmp, "отчеты").replace("'", "''"), app.replace("'", "''")))
        self.assertEqual(out[-2:], ["cfg same: True", "in file: 4"])
        self.assertFalse(os.path.exists(app))

    def test_bulk_revert_save_as_and_later_clicks_do_not_touch_the_open_project(self):
        out = self.run_ps(r"""
function Show-DistGrid { }
$ui = @{
    DistStage = [pscustomobject]@{ Text = 'П' }
    DistTemplate = [pscustomobject]@{ Text = 'C:\t\общий.rvt' }
    DistModelsDir = [pscustomobject]@{ Text = 'C:\t\модели' }
    Summary = [pscustomobject]@{ Text = '' }
    Status = [pscustomobject]@{ Text = '' }
}
$script:SummaryDone = $false
$script:DistRsnModels = @()
$ReportDir = Split-Path (Split-Path $cfgPath)
$script:Project = New-Proj
$script:ProjectPath = $cfgPath
[void](Save-DistProject)
function Sha([string]$p) { return (Get-FileHash -LiteralPath $p -Algorithm SHA256).Hash }
$cfg0 = Sha $cfgPath
Invoke-DistBulk $false
'dirty: ' + $script:DistDirty
'status: ' + $ui.Status.Text
'cfg same: ' + ($cfg0 -eq (Sha $cfgPath))
'snap null: ' + ($null -eq $script:DistDirtySnapshot.InProject)
'live empty: ' + $script:Project.InProject.Count
Invoke-ProjectSpecInClick $script:Project ($P + '3.1_Имя_Пустые')
Request-DistSave
'click kept file: ' + ($cfg0 -eq (Sha $cfgPath))
'timer off: ' + ($null -eq $script:DistSaveTimer)
Restore-DistDirty
'after revert dirty: ' + $script:DistDirty
'after revert null: ' + ($null -eq $script:Project.InProject)
'after revert file: ' + ($cfg0 -eq (Sha $cfgPath))
Invoke-ProjectSpecInClick $script:Project ($P + '7.8_Диаметр_Все')
[void](Save-DistProject)
'outside dirty wrote: ' + ((InFile $cfgPath) -join ' | ')
# снова массовое снятие и «сохранить как» — исходный файл не меняется
$cfg1 = Sha $cfgPath
Invoke-DistBulk $false
'as: ' + (Save-DistAsNewProject 'COPY1')
'open: ' + (Split-Path -Leaf $script:ProjectPath)
'dirty after as: ' + $script:DistDirty
'orig cfg: ' + ($cfg1 -eq (Sha $cfgPath))
'copy in count: ' + @(InFile (Get-ProjectConfigPath $ReportDir 'COPY1')).Count
'no selection files: ' + (@(Get-ChildItem (Join-Path $ReportDir 'проект') -Filter 'выбор_*').Count -eq 0)
""")
        self.assertEqual(out[0], "dirty: True")
        self.assertEqual(out[1], "status: Есть несохранённые изменения")
        self.assertEqual(out[2:5], ["cfg same: True", "snap null: True", "live empty: 0"])
        self.assertEqual(out[5:8], ["click kept file: True", "timer off: True", "after revert dirty: False"])
        self.assertEqual(out[8:10], ["after revert null: True", "after revert file: True"])
        # снятие 7.8 со «все в проекте»: в [inproject] остаются остальные четыре
        self.assertEqual(out[10], "outside dirty wrote: 1_Классификатор_Все | 3.1_Имя_Все | 3.1_Имя_Пустые | 13_Материал_Трубы_Без значения")
        self.assertEqual(out[11:15], ["as: True", "open: распределение_COPY1.txt", "dirty after as: False", "orig cfg: True"])
        self.assertEqual(out[15:17], ["copy in count: 0", "no selection files: True"])

    def test_distribute_stays_in_memory_and_keeps_manual_cells_until_commit(self):
        out = self.run_ps(r"""
function Show-DistGrid { }
$ui = @{
    DistStage = [pscustomobject]@{ Text = 'П' }
    DistTemplate = [pscustomobject]@{ Text = '' }
    DistModelsDir = [pscustomobject]@{ Text = '' }
    Summary = [pscustomobject]@{ Text = '' }
    Status = [pscustomobject]@{ Text = '' }
}
$script:SummaryDone = $false
$script:DistRsnModels = @()
$ReportDir = Split-Path (Split-Path $cfgPath)
$script:Project = New-ProjectConfig 'П'
Update-ProjectModels $script:Project @('A_IOS_OV_X.rvt')
Update-ProjectSpecs $script:Project @('1_Классификатор_Все', '5.1_Ширина_Воздуховоды_Все')
$script:Project.Categories[(Get-ProjectSpecKey '5.1_Ширина_Воздуховоды_Все')] = 'Воздуховоды'
$script:Project.InProject = $null
$script:Project.Summary = $null
Set-ProjectCell $script:Project 'A_IOS_OV_X' '1_Классификатор_Все' 0 'm' -1
$script:ProjectPath = $cfgPath
[void](Save-DistProject)
function Sha([string]$p) { return (Get-FileHash -LiteralPath $p -Algorithm SHA256).Hash }
$cfg0 = Sha $cfgPath
Invoke-DistAssign
'dirty: ' + $script:DistDirty
'cfg same: ' + ($cfg0 -eq (Sha $cfgPath))
$m = Get-ProjectCell $script:Project 'A_IOS_OV_X' '1_Классификатор_Все'
$a = Get-ProjectCell $script:Project 'A_IOS_OV_X' '5.1_Ширина_Воздуховоды_Все'
'manual: ' + $m.State + '/' + $m.Origin
'auto: ' + $a.State + '/' + $a.Origin
$disk = [IO.File]::ReadAllText($cfgPath)
'disk auto: ' + $disk.Contains("5.1_Ширина_Воздуховоды_Все`t1`ta")
Restore-DistDirty
$m2 = Get-ProjectCell $script:Project 'A_IOS_OV_X' '1_Классификатор_Все'
'restored manual: ' + $m2.State + '/' + $m2.Origin
'restored auto: ' + ($null -eq (Get-ProjectCell $script:Project 'A_IOS_OV_X' '5.1_Ширина_Воздуховоды_Все'))
'clean: ' + (-not $script:DistDirty)
Invoke-DistAssign
[void](Commit-DistDirty)
$disk2 = [IO.File]::ReadAllText($cfgPath)
'committed auto: ' + $disk2.Contains("5.1_Ширина_Воздуховоды_Все`t1`ta")
'committed manual: ' + $disk2.Contains("1_Классификатор_Все`t0`tm")
'clean after: ' + (-not $script:DistDirty)
""")
        self.assertEqual(out[:6], [
            "dirty: True", "cfg same: True", "manual: 0/m", "auto: 1/a", "disk auto: False", "restored manual: 0/m",
        ])
        self.assertEqual(out[6:], ["restored auto: True", "clean: True", "committed auto: True", "committed manual: True", "clean after: True"])

    def test_a_single_click_stays_in_memory_until_save_or_revert(self):
        out = self.run_ps(r"""
function Show-DistGrid { }
$ui = @{
    DistStage = [pscustomobject]@{ Text = 'П' }
    DistTemplate = [pscustomobject]@{ Text = 'C:\t\общий.rvt' }
    DistModelsDir = [pscustomobject]@{ Text = 'C:\t\модели' }
    Summary = [pscustomobject]@{ Text = '' }
    Status = [pscustomobject]@{ Text = '' }
}
$script:SummaryDone = $false
$script:DistRsnModels = @()
$script:Project = New-Proj
$script:ProjectPath = $cfgPath
[void](Save-DistProject)
function Sha([string]$p) { return (Get-FileHash -LiteralPath $p -Algorithm SHA256).Hash }
$cfg0 = Sha $cfgPath
Enter-DistDirty
Invoke-ProjectSpecInClick $script:Project ($P + '7.8_Диаметр_Все')
'dirty: ' + $script:DistDirty
'status: ' + $ui.Status.Text
'sel same: ' + ($cfg0 -eq (Sha $cfgPath))
'cfg same: ' + ($cfg0 -eq (Sha $cfgPath))
'mem off: ' + (-not (Test-ProjectSpecIn $script:Project ($P + '7.8_Диаметр_Все')))
Request-DistSave
'timer: ' + ($null -eq $script:DistSaveTimer -or -not $script:DistSaveTimer.Enabled)
Restore-DistDirty
'restored on: ' + (Test-ProjectSpecIn $script:Project ($P + '7.8_Диаметр_Все'))
'clean: ' + (-not $script:DistDirty)
Enter-DistDirty
Invoke-ProjectSpecInClick $script:Project ($P + '7.8_Диаметр_Все')
'commit: ' + (Commit-DistDirty)
'file off: ' + (-not ((InFile $cfgPath) -contains '7.8_Диаметр_Все'))
'clean after: ' + (-not $script:DistDirty)
""")
        self.assertEqual(out[:6], ["dirty: True", "status: Есть несохранённые изменения", "sel same: True", "cfg same: True", "mem off: True", "timer: True"])
        self.assertEqual(out[6:], ["restored on: True", "clean: True", "commit: True", "file off: True", "clean after: True"])


class WiringTests(unittest.TestCase):
    def setUp(self):
        self.d, self.ui = read(DIST), read(UI)

    def test_inproject_is_the_only_source_and_restore_on_load(self):
        click = self.d[self.d.index("function Invoke-DistCellClick"):self.d.index("function Invoke-DistBulk")]
        self.assertLess(click.index("Enter-DistDirty"), click.index("Invoke-ProjectGridClick"))  # снимок до правки, файл не пишется
        bulk = self.d[self.d.index("function Invoke-DistBulk"):self.d.index("function Invoke-DistToggleGroup")]
        self.assertLess(bulk.index("Enter-DistDirty"), bulk.index("Set-ProjectInProjectBulk"))
        for gone in ("Save-DistSelection", "Write-ProjectSelection", "$script:SelectionPath"):
            self.assertNotIn(gone, self.d)
            self.assertNotIn(gone, self.ui)
            self.assertNotIn(gone, read(CFG))
        load = self.d[self.d.index("function Invoke-DistLoad"):self.d.index("function Invoke-DistAssign")]
        self.assertLess(load.index("Get-ProjectInProjectNames $cfgObj"), load.index("Update-ProjectSpecs $cfgObj $names"))   # до перечитывания общего файла
        self.assertIn("Restore-DistSelection $prevIn", load)
        restore = self.d[self.d.index("function Restore-DistSelection"):self.d.index("function Invoke-DistTakeFromSummary")]
        self.assertNotIn("Write-ProjectConfig", restore)                                       # при загрузке файл не пишется
        self.assertNotIn("Read-ProjectSelection", restore)
        self.assertIn("Выбор не найден. Отметьте спеки или нажмите ", restore)
        self.assertIn("Выбор восстановлен из настройки проекта: {0} спек.", restore)
        self.assertIn("нет в общем файле: {0}.", restore)
        loadp = self.d[self.d.index("function Load-DistProject"):self.d.index("function Update-ProjectList")]
        self.assertLess(loadp.index("Invoke-DistStateMigration"), loadp.index("Read-ProjectConfig $Path"))     # перенос — до чтения настройки

    def test_take_from_summary_button_and_flow(self):
        page = self.d[self.d.index("function New-DistPage"):]
        self.assertIn("$flow3.Controls.AddRange(@($ui.DistHide, $ui.DistAll, $ui.DistNone, $ui.DistFromSummary))", page)    # рядом с «Выделить всё»/«Снять всё»
        self.assertIn('New-DistButton "Взять выбор из сводной"', page)
        fn = self.d[self.d.index("function Invoke-DistTakeFromSummary"):self.d.index("# «Загрузить спецификации и модели»: спецификации")]
        ask = fn.index("Заменить текущий выбор выбором из сводной?")
        self.assertIn("Test-ProjectHasSelection", fn[:ask])                                      # спрашиваем, только если галочки уже есть
        self.assertIn("MessageBoxDefaultButton]::Button2", fn)                                   # по умолчанию «Нет»
        self.assertLess(ask, fn.index("Resolve-RevitRun"))                                       # вопрос — до запуска Revit
        self.assertLess(ask, fn.index("Invoke-RbpAndWait"))
        self.assertIn("Отмечено по сводной: {0}; в сводной есть, но нет в общем файле: {1}", fn)
        self.assertNotIn("Save-DistSelection", fn)
        self.assertNotIn("Request-DistSave", fn)
        self.assertIn("Format-SkippedModelsStatus", fn)
        self.assertIn("links_path=", fn)
        self.assertIn("Read-SummaryLinksFile", fn)
        self.assertIn("Set-ProjectModelsFromLinkNames", fn)
        self.assertIn("нет файла в папке:", self.d)
        self.assertIn("Связи сводной не найдены, колонки моделей оставлены из папки.", fn)
        self.assertIn("будут удалены при «Сохранить»", fn)
        self.assertIn("«Отменить» вернёт всё", fn)
        self.assertIn("Get-DistColumnRebuildPrompt", fn)
        self.assertIn("Get-DistThinLinksPrompt", fn)
        self.assertIn("Update-AsgSelForProjectModels", fn)
        self.assertLess(fn.index("Get-DistThinLinksPrompt"), fn.index("Set-ProjectModelsFromLinkNames"))
        self.assertLess(fn.index("Confirm-DistYes"), fn.index("Enter-DistDirty"))
        enter = self.d[self.d.index("function Enter-DistDirty"):self.d.index("function Commit-DistDirty")]
        self.assertIn("$script:DistDirtyPaths = Copy-DistModelPaths $script:DistModelPaths", enter)
        restore = self.d[self.d.index("function Restore-DistDirty"):self.d.index("function Invoke-DistCommit")]
        self.assertLess(restore.index("Set-ProjectFields"), restore.index("$script:DistModelPaths = $paths"))
        self.assertLess(fn.index("Enter-DistDirty"), fn.index("Set-ProjectSelectionFromNames"))
        self.assertLess(fn.index("Set-ProjectSelectionFromNames"), fn.index("Set-ProjectModelsFromLinkNames"))
        self.assertLess(fn.index("Set-ProjectModelsFromLinkNames"), fn.index("Show-DistGrid"))
        for gone in ("Update-ProjectSpecs", "Invoke-DistLoad", "Update-ProjectModels"):          # список спек не пересобираем; модели — Set-ProjectModelsFromLinkNames
            self.assertNotIn(gone, fn)
        self.assertIn("DistFromSummary = ((-not $busy) -and $loaded)", self.ui)
        self.assertIn("Пропущено {0} моделей без целевого раздела", read(CFG))
        self.assertNotIn("function Get-ModelUnknownSectionCode", read(CFG))

    def test_selection_file_is_only_touched_by_the_migration(self):
        for name in os.listdir(TOOL):
            if not name.endswith(".ps1"):
                continue
            t = read(os.path.join(TOOL, name))
            self.assertNotIn("Write-ProjectSelection", t, name)
            if name != "project_config.ps1":
                self.assertNotIn("Read-ProjectSelection", t, name)
                self.assertNotIn("выбор_", t.replace("выбор_<проект>", ""), name)


if __name__ == "__main__":
    unittest.main()
