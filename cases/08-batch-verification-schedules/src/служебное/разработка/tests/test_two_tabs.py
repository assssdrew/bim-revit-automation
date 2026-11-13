"""Окно из двух вкладок («Спецификации», «Задание»): настройки окна (в т.ч. старые), независимое включение «В сводную» и «В разделы», режим вставки, источник моделей из прежнего списка."""
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
CFG = os.path.join(TOOL, "project_config.ps1")

# функции окна берутся из ui_schedules.ps1 как есть (по AST), без запуска самого окна; подключаются точкой в области скрипта, как в окне
PRELUDE = r"""
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
"""


def run(tmp, body, funcs=()):
    script = os.path.join(tmp, "t.ps1")
    names = ",".join("'%s'" % f for f in funcs) or "@()"
    with open(script, "w", encoding="utf-8-sig") as f:
        f.write(PRELUDE % (CFG, UI, names) + body)
    r = pwsh("-File", script)
    assert r.returncode == 0, r.stdout + r.stderr
    return [l for l in r.stdout.splitlines() if l.strip()]


@unittest.skipUnless(PWSH, "нет pwsh")
class TwoTabsTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="bvs_tabs_")
        self.addCleanup(shutil.rmtree, self.tmp, True)

    def test_old_settings_are_read_nothing_lost(self):
        settings = os.path.join(self.tmp, "ui_settings.txt")
        with open(settings, "w", encoding="utf-8-sig") as f:
            f.write("source=C:\\шаблон\\Спецификации.rvt\nprojstage=РД\nprojfile=C:\\o\\проект\\распределение_PROJ.txt\nprojmodelsdir=RSN://srv/Проект\n"
                    "projrsn=RSN://srv/Проект/A_IOS_OV_1.rvt\nprojrsn=RSN://srv/Проект/A_IOS_VK_1.rvt\nsummary=C:\\m\\PROJ_IOS_SVOD_ORG_R24.rvt\n"
                    "googleurl=https://docs.google.com/x\ninsertmodesummary=replace\ninsertmodesections=initial\nasgscope=partial\nasgmodels=A|B\nauxkeepdays=10\n")
        out = run(self.tmp, r"""
$SettingsPath = '%s'; $script:AuxKeepDays = 30
$m = Read-Settings
'source: ' + $m.Source
'stage: ' + $m.ProjStage
'file: ' + $m.ProjFile
'dir: ' + $m.ProjModelsDir
'rsn: ' + ($m.ProjRsn -join ';')
'summary: ' + $m.Summary
'google: ' + $m.Google
'modes: ' + $m.InsertModeSummary + ' ' + $m.InsertModeSections
'scope: ' + $m.AsgScope + ' ' + $m.AsgModels + ' keep ' + $script:AuxKeepDays
""" % settings, ["Read-Settings"])
        self.assertEqual(out, ["source: C:\\шаблон\\Спецификации.rvt", "stage: РД", "file: C:\\o\\проект\\распределение_PROJ.txt", "dir: RSN://srv/Проект",
                               "rsn: RSN://srv/Проект/A_IOS_OV_1.rvt;RSN://srv/Проект/A_IOS_VK_1.rvt", "summary: C:\\m\\PROJ_IOS_SVOD_ORG_R24.rvt",
                               "google: https://docs.google.com/x", "modes: replace initial", "scope: partial A|B keep 10"])

    def test_single_insert_mode_from_old_two_modes(self):
        out = run(self.tmp, r"""
$script:InsertModeValues = @("add_diff", "replace", "initial")
'both replace: ' + (Get-CombinedInsertMode 'replace' 'replace')
'one replace: ' + (Get-CombinedInsertMode 'replace' 'initial')
'initial: ' + (Get-CombinedInsertMode 'initial' 'initial')
'add_diff: ' + (Get-CombinedInsertMode 'add_diff' 'add_diff')
'empty: ' + (Get-CombinedInsertMode '' '')
""", ["Normalize-InsertModeValue", "Get-CombinedInsertMode"])
        self.assertEqual(out, ["both replace: replace", "one replace: add_diff", "initial: add_diff", "add_diff: add_diff", "empty: add_diff"])

    def test_sections_button_is_independent_of_summary(self):
        out = run(self.tmp, r"""
$ui = @{}
foreach ($k in 'Tabs', 'Run', 'DistLoad', 'DynBuild', 'DistAssign', 'DistCount', 'RunSummary', 'RunSections') { $ui[$k] = [pscustomobject]@{ Enabled = $true } }
$ui.Summary = [pscustomobject]@{ Text = 'RSN://srv/SITE/SITE_ORG_SVOD_022.rvt' }
# порядок цифр: Загрузить, Распределить, Посчитать, В сводную, В разделы | Задание, вкладки
function State { ($ui.DistLoad.Enabled, $ui.DistAssign.Enabled, $ui.DistCount.Enabled, $ui.RunSummary.Enabled, $ui.RunSections.Enabled, $ui.Run.Enabled, $ui.Tabs.Enabled | ForEach-Object { [int]$_ }) -join '' }
$script:SummaryDone = $false
$script:Project = $null; $script:ProjectPath = ''; $script:ActionsBusy = $false
Update-ActionState
'start: ' + (State)                                  # только вкладки
$c = New-ProjectConfig 'П'
$script:Project = $c
Update-ActionState
'no project path: ' + (State)                        # без проекта кнопки шагов неактивны (даже если данные в памяти)
$script:ProjectPath = 'C:\o\проект\распределение_SITE.txt'
Update-ActionState
'empty project: ' + (State)
Update-ProjectSpecs $c @('П        3.1_Имя_Все')
Update-ProjectModels $c @('SITE_AVT_ORG_K01_R24')
Update-ActionState
'loaded: ' + (State)                                 # Распределить и Посчитать — после загрузки
Set-ProjectCell $c 'SITE_AVT_ORG_K01_R24' 'П        3.1_Имя_Все' 1 'c' 5
Update-ActionState
'count only: ' + (State)                             # числа без распределения шаг не открывают
Set-ProjectCell $c 'SITE_AVT_ORG_K01_R24' 'П        3.1_Имя_Все' 1 'a' 5
Update-ActionState
'distributed: ' + (State)                            # путь сводной заполнен, SummaryDone = false — обе кнопки включены
$script:SummaryDone = $true
Update-ActionState
'summary done: ' + (State)                           # флаг «В сводную» кнопки не меняет
$ui.Summary.Text = ''; $script:SummaryDone = $false
Update-ActionState
'no summary model: ' + (State)                       # пустой путь сводной тоже не нужен
$ui.Summary.Text = 'RSN://srv/SITE/SITE_ORG_SVOD_022.rvt'
Set-ActionsEnabled $false
'busy: ' + (State)
Set-ActionsEnabled $true
'idle: ' + (State)                                   # путь сводной заполнен, SummaryDone = false — «В разделы» снова включена
""", ["Test-DistLoaded", "Update-ActionState", "Set-ActionsEnabled"])
        self.assertEqual(out, ["start: 0000001", "no project path: 0000001", "empty project: 1000011", "loaded: 1110011", "count only: 1110011", "distributed: 1111111",
                               "summary done: 1111111", "no summary model: 1111111", "busy: 0000000", "idle: 1111111"])

    def test_legacy_model_list_becomes_models_source(self):
        out = run(self.tmp, r"""
'unc: ' + (Get-LegacyModelsSource @('\\srv\проект\ОВ\RVT\a.rvt', '\\srv\проект\ВК\RVT\b.rvt')).Source
'one: ' + (Get-LegacyModelsSource @('C:\p\ОВ\RVT\a.rvt')).Source
'mixed case: ' + (Get-LegacyModelsSource @('C:\P\ОВ\a.rvt', 'c:\p\ВК\b.rvt')).Source
$r = Get-LegacyModelsSource @('RSN://srv/Проект/ИОС/A_OV_.rvt', 'RSN://srv/Проект/ИОС/A_VK_.rvt')
'rsn: ' + $r.Source + ' ' + @($r.Rsn).Count
'empty: [' + (Get-LegacyModelsSource @()).Source + ']'
""")
        self.assertEqual(out, ["unc: \\\\srv\\проект", "one: C:\\p\\ОВ\\RVT", "mixed case: C:\\P", "rsn: RSN://srv/Проект/ИОС 2", "empty: []"])


class WindowSourceTests(unittest.TestCase):
    """Проверки текста окна: что осталось по-старому (проверка версий, итог «готова / не готово», FolderBrowserDialog не используется)."""
    def setUp(self):
        with open(UI, encoding="utf-8-sig") as f:
            self.ui = f.read()

    def test_start_batch_modes_keep_checks_and_failure_text(self):
        start = self.ui[self.ui.index("function Start-Batch"):]
        for mode in ('"summary"', '"sections"', '"assign"'):
            self.assertIn("$Mode -eq " + mode, self.ui[self.ui.index("function Start-Batch"):self.ui.index("$ui = @{}")] + "") if mode != '"sections"' else None
        self.assertIn("Resolve-RevitRun @($chosen) $verTemplate", start)
        self.assertIn("Get-RunFailureText $code $runStats", self.ui)
        self.assertIn("Get-DistProjectModelPaths", start)                      # ④ — модели из распределения
        self.assertIn("$summary = $ui.Summary.Text.Trim()", start)              # ③ и «Задание» берут сводную с вкладки «Спецификации»

    def test_selfcheck_lists_files_the_window_needs(self):
        head = self.ui[:self.ui.index("SELFCHECK OK")]
        for f in ("project_dist_ui.ps1", "project_config.ps1", "folder_dialog.ps1", "revit_versions.ps1", "выбрать_с_сервера.ps1"):
            self.assertIn('"%s"' % f, head)

    def test_old_sections_and_files_pages_code_is_gone_but_models_list_file_is_never_overwritten(self):
        self.assertNotIn("Save-ModelList", self.ui)                            # rvt_list.txt только читается (Read-LegacyModelList)
        self.assertIn("Read-LegacyModelList", self.ui)
        self.assertNotIn("FolderBrowserDialog", self.ui)


if __name__ == "__main__":
    unittest.main()
