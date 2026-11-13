"""Проекты: список вверху окна, настройка проекта хранит пути/Google/режим/«Выборочно», перенос старого ui_settings.txt, журнал на проект, смена PROJ↔SITE без смешивания."""
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
UI = os.path.join(TOOL, "ui_schedules.ps1")
DIST = os.path.join(TOOL, "project_dist_ui.ps1")
CFG = os.path.join(TOOL, "project_config.ps1")
GCFG = os.path.join(TOOL, "google_push_cfg.ps1")
LAYOUT = os.path.join(TOOL, "report_layout.ps1")
P = "П        "


def read(p):
    with open(p, encoding="utf-8-sig") as f:
        return f.read()


def q(s):
    return str(s).replace("'", "''")


UI_FUNCS = ["Get-CurrentProjectName", "Get-ProjectCode", "Get-ModelLeaf", "Get-LedgerPath", "Get-HistoryPath", "Normalize-InsertModeValue", "Get-InsertModeValue",
            "Set-InsertModeBox", "Get-CombinedInsertMode", "Get-GoogleStatusText", "Show-ProjectGoogle", "Save-GoogleConnection", "Read-Settings", "Invoke-StartupSettingsMigration",
            "Read-LegacyModelList", "Get-RunTargetText"]

PRELUDE = r"""
. '%s'
. '%s'
. '%s'
. '%s'
. '%s'
$errs = $null; $tok = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile('%s', [ref]$tok, [ref]$errs)
$texts = foreach ($n in @(%s)) {
    $fn = $ast.FindAll({ param($x) $x -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $x.Name -eq $n }, $true) | Select-Object -First 1
    if (-not $fn) { throw "нет функции $n" }
    $fn.Extent.Text
}
$tmpFile = Join-Path ([IO.Path]::GetTempPath()) ('bvs_fn_' + [guid]::NewGuid().ToString('N') + '.ps1')
[IO.File]::WriteAllText($tmpFile, ($texts -join "`n"), (New-Object System.Text.UTF8Encoding $true))
. $tmpFile
Remove-Item $tmpFile
""" % (q(CFG), q(GCFG), q(LAYOUT), q(DIST), q(DIST), q(UI), ",".join("'%s'" % f for f in UI_FUNCS))


class Base(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="bvs_pr_")
        self.addCleanup(shutil.rmtree, self.tmp, True)
        self.rep = os.path.join(self.tmp, "отчеты")
        self.app = os.path.join(self.tmp, "App")
        os.makedirs(os.path.join(self.rep, "проект"))
        os.makedirs(self.app)

    def run_ps(self, body, stubs=True):
        head = ""
        if stubs:
            head = r"""
$ReportDir = '%s'; $AppDir = '%s'; $SettingsPath = Join-Path $AppDir 'ui_settings.txt'; $ListPath = Join-Path $AppDir 'rvt_list.txt'; $ToolDir = '%s'
$script:Layout = Get-ReportLayout $ReportDir
$script:AuxKeepDays = 30; $script:PartialUiReady = $false; $script:SummaryDone = $false; $script:ProjectListBusy = $false
$script:InsertModeValues = @("add_diff", "replace", "initial"); $script:InsertModeLabels = @("Добавить недостающие", "Заменить все")
function Show-DistGrid { }
function Update-ActionState { }
function Save-Settings { }
function Reset-PartialLists { $script:resets++ }
function Refresh-PartialLists { }
function Save-AsgSelection { }
function Save-DistProject-Stub { }
$script:resets = 0
function InFile([string]$p) { $x = Read-ProjectConfig $p; if ($null -eq $x) { return @() }; $n = Get-ProjectInProjectNames $x; if ($null -eq $n) { return @() }; return @($n) }
class FakeBox { [string]$Text = ''; [bool]$Checked = $false; [int]$SelectedIndex = 0; [bool]$Enabled = $true; $Items = (New-Object System.Collections.ArrayList); $SelectedItem = $null }
function New-FakeUi {
    $u = @{}
    foreach ($k in 'DistStage', 'DistTemplate', 'DistModelsDir', 'Summary', 'Google', 'GoogleInfo', 'GoogleConnect', 'GoogleCheck', 'Status', 'AsgPart', 'ProjectBox') { $u[$k] = [FakeBox]::new() }
    $u.InsertModeBox = [FakeBox]::new(); [void]$u.InsertModeBox.Items.Add('a'); [void]$u.InsertModeBox.Items.Add('b')
    $u.Source = $u.DistTemplate
    return $u
}
$ui = New-FakeUi
""" % (q(self.rep), q(self.app), q(TOOL))
        script = os.path.join(self.tmp, "t.ps1")
        with open(script, "w", encoding="utf-8-sig") as f:
            f.write(PRELUDE + head + body)
        r = pwsh("-File", script)
        self.assertEqual(r.returncode, 0, r.stdout + r.stderr)
        return [l for l in r.stdout.splitlines() if l.strip()]


@unittest.skipUnless(PWSH, "нет pwsh")
class SwitchProjectTests(Base):
    def test_switching_projects_does_not_mix_paths_selection_google_journal(self):
        out = self.run_ps(r"""
function Set-Fields($tpl, $dir, $sum, $link, $mode, $url, $secret, $done) {
    $ui.DistTemplate.Text = $tpl; $ui.DistModelsDir.Text = $dir; $ui.Summary.Text = $sum; $ui.Google.Text = $link
    $ui.InsertModeBox.SelectedIndex = $mode; $script:SummaryDone = $done
    Write-ProjectGoogleConfig (Get-ProjectGoogleConfigPath $script:ProjectPath) $url $secret            # подключение пишется в google_<проект>.cfg, полей в окне нет
}
function Dump { ($ui.DistTemplate.Text, $ui.DistModelsDir.Text, $ui.Summary.Text, $ui.Google.Text, (Get-InsertModeValue), (Read-ProjectGoogleConfig (Get-ProjectGoogleConfigPath $script:ProjectPath)).WebAppUrl, (Read-ProjectGoogleConfig (Get-ProjectGoogleConfigPath $script:ProjectPath)).Secret, $script:SummaryDone, (Get-LedgerPath $ui.Summary.Text | Split-Path -Leaf), (Get-HistoryPath '' | Split-Path -Leaf)) -join ' | ' }
# проект PROJ
[void](New-ProjectFile 'PROJ')
'open: ' + (Get-CurrentProjectName)
Update-ProjectSpecs $script:Project @('1_А_Все', '2_Б_Все', '3_В_Все')
$script:Project.InProject = $null
Set-ProjectInProjectBulk $script:Project @('2_Б_Все', '3_В_Все') $false        # в проекте только 1_А_Все
Set-Fields 'C:\avt\Спеки.rvt' 'C:\avt\модели' 'C:\avt\PROJ_IOS_SVOD.rvt' 'https://docs/proj' 1 'https://script.google.com/macros/s/PROJ/exec' 'secretAVT' $true
$script:AsgSel = @{ Models = @{ 'PROJ_IOS_OV_1' = $true }; Specs = $null; Pairs = @{ 'PROJ_IOS_OV_1|1_А_Все' = $true } }
$ui.AsgPart.Checked = $true
[void](Save-DistProject)
'PROJ dump: ' + (Dump)
# проект SITE: новый — пустой, ничего от PROJ
[void](New-ProjectFile 'SITE')
'open: ' + (Get-CurrentProjectName)
'SITE fresh: ' + (Dump)
'SITE asg: ' + ($null -eq $script:AsgSel.Models) + ' ' + ($null -eq $script:AsgSel.Pairs) + ' partial ' + $ui.AsgPart.Checked
Update-ProjectSpecs $script:Project @('1_А_Все', '2_Б_Все', '3_В_Все')
$script:Project.InProject = $null
Set-ProjectInProjectBulk $script:Project @('1_А_Все') $false                          # в проекте 2_Б и 3_В
Set-Fields 'D:\site\Общий.rvt' 'RSN://srv/site' 'D:\site\SITE_IOS_SVOD.rvt' 'https://docs/site' 0 'https://script.google.com/macros/s/SITE/exec' 'secretSITE' $false
$script:AsgSel = @{ Models = @{ 'SITE_IOS_VK_1' = $true }; Specs = $null; Pairs = @{} }
$ui.AsgPart.Checked = $false
[void](Save-DistProject)
'SITE dump: ' + (Dump)
# обратно на PROJ
Switch-Project 'PROJ'
'back: ' + (Get-CurrentProjectName)
'PROJ again: ' + (Dump)
'PROJ asg: ' + (@($script:AsgSel.Models.Keys) -join ',') + ' ' + (@($script:AsgSel.Pairs.Keys) -join ',') + ' partial ' + $ui.AsgPart.Checked
'PROJ selection file: ' + ((InFile $script:ProjectPath) -join ',')
Switch-Project 'SITE'
'SITE again: ' + (Dump)
'SITE asg: ' + (@($script:AsgSel.Models.Keys) -join ',') + ' partial ' + $ui.AsgPart.Checked
'SITE selection file: ' + ((InFile $script:ProjectPath) -join ',')
'resets: ' + $script:resets
'projects: ' + ((Get-ProjectNames $ReportDir) -join ',')
'google files: ' + (@(Get-ChildItem (Join-Path $ReportDir 'проект') -Filter 'google_*.cfg' | ForEach-Object Name | Sort-Object) -join ',')
'selection files: ' + (@(Get-ChildItem (Join-Path $ReportDir 'проект') -Filter 'выбор_*.txt' | ForEach-Object Name | Sort-Object) -join ',')
""")
        lines = out
        self.assertEqual(lines[0], "open: PROJ")
        avt = "C:\\avt\\Спеки.rvt | C:\\avt\\модели | C:\\avt\\PROJ_IOS_SVOD.rvt | https://docs/proj | replace | https://script.google.com/macros/s/PROJ/exec | secretAVT | True | задание_журнал_PROJ.txt | история_PROJ.txt"
        self.assertEqual(lines[1], "PROJ dump: " + avt)
        self.assertEqual(lines[2], "open: SITE")
        self.assertTrue(lines[3].startswith("SITE fresh:  |  |  |  | add_diff |  |  | False | задание_журнал_SITE.txt | история_SITE.txt"), lines[3])    # пустое: ничего от PROJ
        self.assertEqual(lines[4], "SITE asg: True True partial False")
        site = "D:\\site\\Общий.rvt | RSN://srv/site | D:\\site\\SITE_IOS_SVOD.rvt | https://docs/site | add_diff | https://script.google.com/macros/s/SITE/exec | secretSITE | False | задание_журнал_SITE.txt | история_SITE.txt"
        self.assertEqual(lines[5], "SITE dump: " + site)
        self.assertEqual(lines[6], "back: PROJ")
        self.assertEqual(lines[7], "PROJ again: " + avt)                                        # все значения PROJ вернулись, ничего от SITE
        self.assertEqual(lines[8], "PROJ asg: PROJ_IOS_OV_1 PROJ_IOS_OV_1|1_А_Все partial False")
        self.assertEqual(lines[9], "PROJ selection file: 1_А_Все")
        self.assertEqual(lines[10], "SITE again: " + site)
        self.assertEqual(lines[11], "SITE asg: SITE_IOS_VK_1 partial False")
        self.assertEqual(lines[12], "SITE selection file: 2_Б_Все,3_В_Все")
        self.assertEqual(lines[13], "resets: 4")                                               # списки «Выборочно» сбрасывались при каждой смене
        self.assertEqual(lines[14], "projects: PROJ,SITE")
        self.assertEqual(lines[15], "google files: google_PROJ.cfg,google_SITE.cfg")
        self.assertEqual(lines[16], "selection files: ")                                      # «В проекте» — в [inproject], отдельных файлов выбора нет
        self.assertNotIn("secretSITE", read(os.path.join(self.rep, "проект", "распределение_PROJ.txt")))        # секрет — только в google_<проект>.cfg
        self.assertNotIn("secretAVT", read(os.path.join(self.rep, "проект", "распределение_SITE.txt")))

    def test_new_project_name_validation(self):
        out = self.run_ps(r"""
[void](New-ProjectFile 'PROJ')
'empty: ' + (Test-NewProjectName $ReportDir '  ')
'bad: ' + (Test-NewProjectName $ReportDir 'A/B')
'dup: ' + (Test-NewProjectName $ReportDir 'PROJ')
'ok: [' + (Test-NewProjectName $ReportDir 'SITE') + ']'
""")
        self.assertEqual(out[0], "empty: Укажите имя проекта.")
        self.assertTrue(out[1].startswith("bad: В имени проекта нельзя"), out[1])
        self.assertEqual(out[2], "dup: Проект «PROJ» уже есть.")
        self.assertEqual(out[3], "ok: []")

    def test_unsaved_bulk_asks_before_leaving_the_project(self):
        out = self.run_ps(r"""
[void](New-ProjectFile 'PROJ')
Update-ProjectSpecs $script:Project @('1_А_Все', '2_Б_Все')
$script:Project.InProject = $null
$script:Project.Summary = $null
[void](Save-DistProject)
$sel = $script:ProjectPath
function Names([string]$p) { return ((InFile $p) -join ',') }
$before = Names $sel
Invoke-DistBulk $false
'dirty: ' + $script:DistDirty
'during: ' + ((Names $sel) -eq $before)
'mem: ' + (@($script:Project.Specs | Where-Object { Test-ProjectSpecIn $script:Project $_ })).Count
$script:DistDirtyChoice = 'cancel'
'new-cancel: ' + @((New-ProjectFile 'SITE'))[-1]
'still: ' + (Get-CurrentProjectName)
'site: ' + (Test-Path -LiteralPath (Get-ProjectConfigPath $ReportDir 'SITE'))
'still-dirty: ' + $script:DistDirty
'file-same: ' + ((Names $sel) -eq $before)
$script:DistDirtyChoice = 'save'
'new-save: ' + @((New-ProjectFile 'SITE'))[-1]
'open: ' + (Get-CurrentProjectName)
'dirty-after-save: ' + $script:DistDirty
'proj-after-save: [' + (Names (Get-ProjectConfigPath $ReportDir 'PROJ')) + ']'
""")
        self.assertEqual(out[:5], ["dirty: True", "during: True", "mem: 0", "new-cancel: False", "still: PROJ"])
        self.assertEqual(out[5:8], ["site: False", "still-dirty: True", "file-same: True"])
        self.assertEqual(out[8:], ["new-save: True", "open: SITE", "dirty-after-save: False", "proj-after-save: []"])

    def test_discard_on_leave_restores_the_open_project_files(self):
        out = self.run_ps(r"""
[void](New-ProjectFile 'PROJ')
Update-ProjectSpecs $script:Project @('1_А_Все', '2_Б_Все')
$script:Project.InProject = $null
$script:Project.Summary = $null
Set-ProjectInProjectBulk $script:Project @('1_А_Все', '2_Б_Все') $true
[void](Save-DistProject)
Invoke-DistBulk $false
$script:DistDirtyChoice = 'discard'
'new: ' + @((New-ProjectFile 'SITE'))[-1]
'open: ' + (Get-CurrentProjectName)
'dirty: ' + $script:DistDirty
'proj: ' + ((InFile (Get-ProjectConfigPath $ReportDir 'PROJ')) -join ',')
""")
        self.assertEqual(out, ["new: True", "open: SITE", "dirty: False", "proj: 1_А_Все,2_Б_Все"])


@unittest.skipUnless(PWSH, "нет pwsh")
class MigrationTests(Base):
    LEGACY = ("source=C:\\общий\\Спеки.rvt\nprojstage=РД\nprojfile=%s\nprojmodelsdir=C:\\модели\nsummary=C:\\m\\PROJ_IOS_SVOD_ORG_R24.rvt\n"
              "googleurl=https://docs/proj\nsummarydone=1\ninsertmodesummary=replace\ninsertmodesections=replace\nasgscope=partial\nasgmodels=PROJ_IOS_OV_1|PROJ_IOS_VK_1\n"
              "asgpairsfmt=1\nasgpair=PROJ_IOS_OV_1|1_А_Все\ndyndays=14\ncollapseday=1\n")

    def write_legacy(self, projfile=""):
        with open(os.path.join(self.app, "ui_settings.txt"), "w", encoding="utf-8-sig") as f:
            f.write(self.LEGACY % projfile)

    def test_old_settings_move_into_their_project_and_nothing_is_lost(self):
        self.write_legacy(os.path.join(self.rep, "проект", "распределение_PROJ.txt").replace("\\", "/"))
        out = self.run_ps(r"""
$saved = Read-Settings
'migrated flag before: ' + $saved.Migrated
$name = Invoke-StartupSettingsMigration $saved
'project: ' + $name
$c = Read-ProjectConfig (Get-ProjectConfigPath $ReportDir $name)
'template: ' + $c.Template
'models: ' + $c.ModelsDir
'stage: ' + $c.Stage
'summary: ' + $c.SummaryModel
'link: ' + $c.GoogleLink
'done: ' + $c.SummaryDone
'mode: ' + $c.InsertMode
'scope: ' + $c.AsgScope
'asg models: ' + (@($c.AsgModels.Keys | Sort-Object) -join ',')
'asg pairs: ' + (@($c.AsgPairs.Keys) -join ',')
'backup: ' + (Test-Path -LiteralPath (Join-Path $AppDir 'ui_settings.до_проектов.txt'))
'again: [' + (Invoke-StartupSettingsMigration @{ Migrated = $true }) + ']'
# повтор не затирает значения проекта
$c.Template = 'C:\другой.rvt'; Write-ProjectConfig $c (Get-ProjectConfigPath $ReportDir $name)
$null = Invoke-StartupSettingsMigration $saved
'template kept: ' + (Read-ProjectConfig (Get-ProjectConfigPath $ReportDir $name)).Template
""")
        self.assertEqual(out[0], "migrated flag before: False")
        self.assertEqual(out[1], "project: PROJ")
        self.assertEqual(out[2:12], ["template: C:\\общий\\Спеки.rvt", "models: C:\\модели", "stage: РД", "summary: C:\\m\\PROJ_IOS_SVOD_ORG_R24.rvt", "link: https://docs/proj",
                                     "done: True", "mode: replace", "scope: partial", "asg models: PROJ_IOS_OV_1,PROJ_IOS_VK_1", "asg pairs: PROJ_IOS_OV_1|1_А_Все"])
        self.assertEqual(out[12], "backup: True")
        self.assertEqual(out[13], "again: []")
        self.assertEqual(out[14], "template kept: C:\\другой.rvt")
        self.assertIn("source=C:\\общий\\Спеки.rvt", read(os.path.join(self.app, "ui_settings.до_проектов.txt")))     # копия исходного файла сохранена

    def test_project_name_from_models_prefix_when_no_projfile(self):
        self.write_legacy("")
        out = self.run_ps(r"""
$saved = Read-Settings
'project: ' + (Invoke-StartupSettingsMigration $saved)
""")
        self.assertEqual(out[0], "project: PROJ")                                                  # по префиксу моделей сводной

    def test_old_model_list_becomes_models_dir_when_empty(self):
        with open(os.path.join(self.app, "ui_settings.txt"), "w", encoding="utf-8-sig") as f:
            f.write("source=C:\\x\\Спеки.rvt\nsummary=C:\\x\\SITE_IOS_SVOD_ORG_R24.rvt\n")
        with open(os.path.join(self.app, "rvt_list.txt"), "w", encoding="utf-8") as f:
            f.write("\\\\srv\\проект\\ОВ\\RVT\\a.rvt\n\\\\srv\\проект\\ВК\\RVT\\b.rvt\n")
        out = self.run_ps(r"""
$name = Invoke-StartupSettingsMigration (Read-Settings)
'project: ' + $name
'models: ' + (Read-ProjectConfig (Get-ProjectConfigPath $ReportDir $name)).ModelsDir
""")
        self.assertEqual(out, ["project: SITE", "models: \\\\srv\\проект"])

    def test_settings_file_keeps_only_last_project_and_general_params(self):
        ui = read(UI)
        save = ui[ui.index("function Save-Settings"):ui.index("function Get-CurrentProjectName")]
        for kept in ("lastproject=", "migrated=1", "dyndays=", "dyncustom=", "collapseday=", "auxkeepdays=", "dynfrom=", "dynto="):
            self.assertIn(kept, save)
        for gone in ("source=", "summary=", "googleurl", "projstage", "projmodelsdir", "projrsn", "asgpair", "asgmodels", "insertmode", "summarydone", "asgscope", "projfile"):
            self.assertNotIn(gone, save)                                                          # проектное — в настройке проекта


@unittest.skipUnless(PWSH, "нет pwsh")
class LedgerTests(Base):
    def test_ledger_is_per_project_and_old_shared_ledger_moves_to_proj(self):
        hist = os.path.join(self.rep, "история")
        os.makedirs(hist)
        with open(os.path.join(hist, "задание_журнал.txt"), "w", encoding="utf-8-sig") as f:
            f.write("old ledger\n")
        with open(os.path.join(hist, "задание_журнал_SITE.txt"), "w", encoding="utf-8-sig") as f:
            f.write("site ledger\n")
        out = self.run_ps(r"""
$m = Invoke-LedgerPerProjectMigration $ReportDir
'moves: ' + @($m).Count
'old exists: ' + (Test-Path -LiteralPath (Join-Path $script:Layout.History 'задание_журнал.txt'))
'proj: ' + ([IO.File]::ReadAllText((Join-Path $script:Layout.History 'задание_журнал_PROJ.txt')).Trim())
'site untouched: ' + ([IO.File]::ReadAllText((Join-Path $script:Layout.History 'задание_журнал_SITE.txt')).Trim())
'again: ' + @(Invoke-LedgerPerProjectMigration $ReportDir).Count
$script:ProjectPath = Join-Path $ReportDir 'проект\распределение_PROJ.txt'
'ledger proj: ' + (Split-Path -Leaf (Get-LedgerPath 'C:\m\SITE_IOS_SVOD.rvt'))       # открыт проект — код по проекту, а не по сводной
$script:ProjectPath = Join-Path $ReportDir 'проект\распределение_SITE.txt'
'ledger site: ' + (Split-Path -Leaf (Get-LedgerPath ''))
""")
        self.assertEqual(out, ["moves: 1", "old exists: False", "proj: old ledger", "site untouched: site ledger", "again: 0",
                               "ledger proj: задание_журнал_PROJ.txt", "ledger site: задание_журнал_SITE.txt"])

    def test_existing_proj_ledger_is_not_overwritten(self):
        hist = os.path.join(self.rep, "история")
        os.makedirs(hist)
        for name, text in (("задание_журнал.txt", "old\n"), ("задание_журнал_PROJ.txt", "current proj\n")):
            with open(os.path.join(hist, name), "w", encoding="utf-8-sig") as f:
                f.write(text)
        self.run_ps(r"[void](Invoke-LedgerPerProjectMigration $ReportDir)")
        self.assertEqual(read(os.path.join(hist, "задание_журнал_PROJ.txt")).strip(), "current proj")
        conflicts = os.listdir(os.path.join(hist, "_конфликты_переноса"))
        self.assertEqual(conflicts, ["задание_журнал.txt"])                                           # старый не пропал, лежит в конфликтах

    def test_build_assignment_uses_per_project_ledger_name(self):
        text = read(os.path.join(TOOL, "build_assignment.ps1"))
        self.assertIn("'^история_(.+)\\.txt$'", text)
        self.assertIn('"задание_журнал_" + $Matches[1] + ".txt"', text)


@unittest.skipUnless(PWSH, "нет pwsh")
class RunTargetAndWiringTests(Base):
    def test_status_before_run_names_project_and_google_tail(self):
        out = self.run_ps(r"""
[void](New-ProjectFile 'PROJ')
'no google: ' + (Get-RunTargetText)
Write-ProjectGoogleConfig (Get-ProjectGoogleConfigPath $script:ProjectPath) 'https://script.google.com/macros/s/AKfycbXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX/exec' 'TOPSECRET'
'with google: ' + (Get-RunTargetText)
""")
        # на ПК с настоящим служебное\google_push\google_push.cfg вместо «не задан» пишется про общую таблицу — оба ответа верны
        self.assertIn(out[0], ("no google: Проект PROJ. Google: не задан",
                               "no google: Проект PROJ. Google: у проекта нет своего адреса (общая таблица — только по подтверждению)"), out[0])
        self.assertTrue(out[1].startswith("with google: Проект PROJ. Google: …") and out[1].endswith("/exec"), out[1])
        self.assertNotIn("TOPSECRET", "\n".join(out))

    def test_wiring(self):
        ui, dist = read(UI), read(DIST)
        for part in ('$ui.ProjectBox = New-Object System.Windows.Forms.ComboBox', '$ui.ProjectNew = New-DistButton "Новый проект"', "$ui.ProjectNew.Add_Click({ Invoke-NewProject })",
                     "Switch-Project ([string]$ui.ProjectBox.SelectedItem)", "$top.Dock = \"Top\"", "$form.Controls.Add($top)"):
            self.assertIn(part, ui)
        start = ui[ui.index("function Start-Batch"):]
        self.assertIn("$script:Run.TargetText = Get-RunTargetText", start)                      # «В сводную», «В разделы», «Собрать задание»
        self.assertIn('{0}. Идёт загрузка в сводную.', ui)
        self.assertIn('{0}. Идёт загрузка в разделы.', ui)
        self.assertIn('{0}. Идёт выгрузка из сводной.', ui)
        self.assertIn("Update-ProjectList", ui[ui.index("Load-DistProject $(if ($open)"):])
        for key in ("Run = ((-not $busy) -and $hasProject)", "DistLoad = ((-not $busy) -and $hasProject)"):
            self.assertIn(key, ui)
        self.assertIn("Выберите проект в списке вверху окна", dist)                              # загрузка без проекта не стартует
        load = dist[dist.index("function Invoke-DistLoad"):dist.index("function Invoke-DistAssign")]
        self.assertNotIn("Get-ProjectPrefix", load)                                              # путь настройки — открытый проект, а не префикс моделей


if __name__ == "__main__":
    unittest.main()
