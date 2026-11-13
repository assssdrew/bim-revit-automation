"""project_dist_ui.ps1: чтение результатов RBP (шаблон, предзаполнение) и итоги; статическая проверка подключения вкладки «Распределение»."""
import os
import re
import shutil
import sys
import tempfile
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from test_chain import pwsh, PWSH  # noqa: E402

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(HERE)))
TOOL = os.path.join(ROOT, "служебное", "скрипты")      # папка с .ps1/.py окна
CFG = os.path.join(TOOL, "project_config.ps1")
DIST = os.path.join(TOOL, "project_dist_ui.ps1")
UI = os.path.join(TOOL, "ui_schedules.ps1")
P = "П        "


def read(p):
    with open(p, encoding="utf-8-sig") as f:
        return f.read()


def run(tmp, body):
    script = os.path.join(tmp, "t.ps1")
    with open(script, "w", encoding="utf-8-sig") as f:
        f.write(". '%s'\n. '%s'\n$P = '%s'\n%s\n" % (CFG, DIST, P, body))
    r = pwsh("-File", script)
    assert r.returncode == 0, r.stdout + r.stderr
    return [l for l in r.stdout.splitlines() if l.strip()]


@unittest.skipUnless(PWSH, "нет pwsh")
class DistLogicTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="bvs_dist_")
        self.addCleanup(shutil.rmtree, self.tmp, True)

    def write(self, name, text):
        p = os.path.join(self.tmp, name)
        with open(p, "w", encoding="utf-8-sig", newline="") as f:
            f.write(text)
        return p

    def test_template_and_prefill_files(self):
        tpl = self.write("tpl.tsv", "П        5.1_Ширина_Воздуховоды_Все\tВоздуховоды\n\n13_Материал_Трубы_Пустые\tТрубы\n")
        pre = self.write("pre.tsv",
                         "A_IOS_OV_X\t5.1_Ширина_Воздуховоды_Все\t12\t0.5\tколлектор\n"
                         "A_IOS_OV_X\t13_Материал_Трубы_Пустые\t0\t0.4\tколлектор\n"
                         "A_IOS_OV_X\t7_Нет\t-1\t0.1\tошибка\n"
                         "#МОДЕЛЬ\tA_IOS_OV_X\t1.5\t3\t1\tколлектор\tпачка\t0.2\t0.6\t0.1\t0.5\t0.1\n"
                         "#ОШИБКА\tA_IOS_VK_X\tне открылась\n")
        out = run(self.tmp, r"""
$n = @(Read-TemplateSpecsFile '%s'); 'tpl: ' + $n.Count + ' ' + $n[1]
$r = Read-PrefillResults '%s'
'counts: ' + $r.Counts.Count + ' models: ' + $r.Models.Count + ' errors: ' + $r.Errors.Count
'model: ' + $r.Models[0].Model + ' ' + $r.Models[0].Seconds + ' ' + $r.Models[0].Specs + ' ' + $r.Models[0].Errors + ' ' + $r.Models[0].Mode + ' ' + $r.Models[0].Open + '/' + $r.Models[0].Paste + '/' + $r.Models[0].Count
'err: ' + $r.Errors[0]
$c = New-ProjectConfig 'П'
Update-ProjectModels $c @('A_IOS_OV_X')
Update-ProjectSpecs $c @('5.1_Ширина_Воздуховоды_Все', '13_Материал_Трубы_Пустые', '7_Нет')
$t = Merge-PrefillIntoProject $c $r 3725
$t
(Get-ProjectCell $c 'A_IOS_OV_X' '5.1_Ширина_Воздуховоды_Все').State
(Get-ProjectCell $c 'A_IOS_OV_X' '13_Материал_Трубы_Пустые').State
(Get-ProjectCell $c 'A_IOS_OV_X' '7_Нет').Origin
Get-DistSummaryText $c
(Get-DistLogLines $r 3725) -join '||'
""" % (tpl, pre))
        self.assertEqual(out[0], "tpl: 2 13_Материал_Трубы_Пустые")
        self.assertEqual(out[1], "counts: 3 models: 1 errors: 1")
        self.assertEqual(out[2], "model: A_IOS_OV_X 1.5 3 1 пачка 0.2/0.6/0.5")
        self.assertIn("не открылась", out[3])
        self.assertTrue(out[4].startswith("Посчитано: пар 2, не определено 1, моделей 1. Время прогона 01:02:05"), out[4])
        self.assertIn("A_IOS_OV_X: 1.5 с", out[4])
        self.assertEqual(out[5:8], ["0", "0", "c"])                        # «Посчитать элементы» меняет только числа: галочек нет, признак «только число»
        self.assertIn("Спецификаций 3", out[8])
        log = out[9].split("||")
        self.assertIn("пачка", log[0])
        for part in ("открытие источника 0.2 с", "вставка 0.6 с", "профиль 0.1 с", "подсчёт 0.5 с", "откат 0.1 с", "всего 1.5 с"):
            self.assertIn(part, log[0])
        self.assertTrue(log[-1].startswith("Общее время прогона: 01:02:05"), log[-1])
        self.assertTrue(any(x.startswith("ОШИБКА A_IOS_VK_X") for x in log))

    def test_missing_files_are_empty(self):
        out = run(self.tmp, r"""
(@(Read-TemplateSpecsFile 'нет.tsv')).Count
$r = Read-PrefillResults 'нет.tsv'
$r.Counts.Count; $r.Models.Count; $r.Errors.Count
""")
        self.assertEqual(out, ["0", "0", "0", "0"])


class DistWiringTests(unittest.TestCase):
    def setUp(self):
        self.ui = read(UI)
        self.dist = read(DIST)

    def test_no_hardcoded_packs_or_list_file_left(self):
        for bad in ("sectionPacks", "Build-ScheduleListFile", "Read-ListBlock", "$ui.Sets", "schedule_sets", "$SetsFile", "ScheduleListFile"):
            self.assertNotIn(bad, self.ui, bad)

    def test_sources_and_pages(self):
        self.assertIn('. (Join-Path $ToolDir "project_config.ps1")', self.ui)
        self.assertIn('. (Join-Path $ToolDir "project_dist_ui.ps1")', self.ui)
        self.assertIn("$ui.PageDist = New-DistPage", self.ui)
        # две вкладки: «Спецификации» и «Задание»; страниц «Файлы», «Сводная», «Разделы» и боковой навигации больше нет
        self.assertIn('$tabSpecs.Text = "Спецификации"', self.ui)
        self.assertIn('$tabAsg.Text = "Задание"', self.ui)
        self.assertIn("$tabSpecs.Controls.Add($ui.PageDist)", self.ui)
        self.assertIn("$tabAsg.Controls.Add($ui.PageAssign)", self.ui)
        self.assertEqual(self.ui.count("New-Object System.Windows.Forms.TabPage"), 2)
        for gone in ("PageFiles", "PageSummary", "PageSections", "NavDist", "NavFiles", "NavSummary", "NavSections", "NavAssign", "Show-Step", "Add-FileRow"):
            self.assertNotIn(gone, self.ui + self.dist, gone)
        # лента: ① → ② → ③ → ④, ③ и ④ запускают прежние режимы вставки
        self.assertIn('$ui.RunSummary.Add_Click({ Start-Batch -Mode "summary" })', self.dist)
        self.assertIn('$ui.RunSections.Add_Click({ Start-Batch -Mode "sections" })', self.dist)
        self.assertIn('$ui.RunClassifier.Add_Click({ Start-Batch -Mode "classifier" })', self.dist)
        self.assertIn('$ui.Run.Add_Click({ Start-Batch -Mode "assign" })', self.ui)
        # «Распределить» — после загрузки; «В сводную» и «В разделы» — после распределения, независимо; «Посчитать элементы» — после загрузки
        state = self.ui[self.ui.index("function Update-ActionState"):self.ui.index("function Set-ActionsEnabled")]
        self.assertIn("DistAssign = ((-not $busy) -and $loaded)", state)
        self.assertIn("RunSummary = ((-not $busy) -and $distributed)", state)
        self.assertIn("RunSections = ((-not $busy) -and $distributed)", state)
        self.assertIn("RunClassifier = ((-not $busy) -and $loaded)", state)
        self.assertNotIn("SummaryDone", state)
        self.assertNotIn("noSummary", state)
        self.assertIn("DistCount = ((-not $busy) -and $loaded)", state)
        self.assertIn("Update-ActionState", self.dist[self.dist.index("function Show-DistGrid"):self.dist.index("function Save-DistProject")])

    def test_run_uses_project_file_and_refuses_empty_path(self):
        self.assertIn('("project_path={0}" -f $ProjectPath)', self.ui)
        self.assertIn("-ProjectPath $projectPath", self.ui)
        start = self.ui[self.ui.index("function Start-Batch"):self.ui.index("$exe = Find-BatchRvt")]
        self.assertIn("Get-ProjectRunBlocker $script:ProjectPath", start)           # пустой путь — понятное сообщение, запуска нет
        self.assertLess(start.index("Get-ProjectRunBlocker"), start.index("$projectPath = $script:ProjectPath"))
        self.assertIn("return", start[start.index("if ($blocker)"):start.index("$projectPath = $script:ProjectPath")])
        self.assertIn("Get-PartialModelSpecInfo $script:AsgAllModels $script:Project", self.ui)
        self.assertIn("Get-ProjectSummarySpecs $Cfg", self.dist)
        assign = self.dist[self.dist.index("function Invoke-DistAssign"):self.dist.index("# Щелчок по таблице")]
        self.assertIn("Get-ProjectRunBlocker $script:ProjectPath", assign)
        self.assertLess(assign.index("Get-ProjectRunBlocker"), assign.index("Invoke-RbpAndWait"))

    def test_legacy_import_removed(self):
        for text in (self.ui, self.dist, read(os.path.join(TOOL, "project_config.ps1"))):
            for bad in ("Import-LegacySets", "schedule_sets", "PcLegacyPacks"):
                self.assertNotIn(bad, text, bad)

    def test_folder_dialogs_use_address_bar_dialog_only(self):
        dialog = read(os.path.join(TOOL, "folder_dialog.ps1"))
        self.assertIn("Show-FolderPathDialog", self.dist)
        self.assertIn("function Show-FolderPathDialog", dialog)
        self.assertIn("folder_dialog.ps1", self.ui)
        self.assertNotIn("function Show-FolderPathDialog", self.ui)                       # одно определение, в общем файле
        self.assertIn("OpenFileDialog", dialog)
        for root, _, files in os.walk(TOOL):
            if os.sep + "tests" in root or ".git" in root:
                continue
            for f in files:
                if f.endswith((".ps1", ".py", ".bat", ".vbs", ".cmd")):
                    self.assertNotIn("FolderBrowserDialog", read(os.path.join(root, f)), f)

    def test_dist_controls_order_and_fields(self):
        d = self.dist
        for c in ("DistTemplate", "DistModelsDir", "DistStage", "DistLoad", "DistHide", "DistAssign", "DistGrid", "DistCount", "RunSummary", "RunSections", "InsertModeBox"):
            self.assertRegex(d, r"\$ui\.%s = " % c, c)
        self.assertNotIn("DistInfo", d + self.ui)                                   # нижней информационной строки нет
        for gone in ("DistCode", "$ui.DistSave ", "DistRead", "DistPrefill", "Сохранить настройку проекта", "Шифр", "шифр"):
            self.assertNotIn(gone, d, gone)
            self.assertNotIn("DistCode", self.ui)
        page = d[d.index("function New-DistPage"):]
        order = [page.index(x) for x in ('"Общий файл" $ui.DistTemplate', '"Модели" $ui.DistModelsDir', '"Сводная модель" $ui.Summary', 'Text = "Стадия"', 'DistHide.Text = "Скрыть неучаствующие"',
                                         'DistAll = New-DistButton "Выделить всё"', "DistGrid = $g", 'DistLoad = New-DistButton "Загрузить спецификации и модели"',
                                         'DistAssign = New-DistButton "Распределить"', 'RunSummary = New-DistButton "В сводную"', 'RunSections = New-DistButton "В разделы"',
                                         'DistCount = New-DistButton "Посчитать элементы"', '$ui.InsertModeBox = $combo')]
        self.assertEqual(order, sorted(order))                                      # порядок сверху вниз
        # кнопки у путей: общий файл — «Обзор…»; модели и сводная — «Обзор…» и «С сервера…»
        self.assertIn("Add-PathRow $paths 0 \"Общий файл\" $ui.DistTemplate @($bTpl)", page)
        self.assertIn("Add-PathRow $paths 1 \"Модели\" $ui.DistModelsDir @($bDir, $bDirSrv)", page)
        self.assertIn("Add-PathRow $paths 2 \"Сводная модель\" $ui.Summary @($bSum, $bSumSrv)", page)
        self.assertIn("$ui.Source = $ui.DistTemplate", page)                        # одно поле общего файла (source= в настройках)
        # все пути сохраняются сами
        for box in ("DistTemplate", "Summary", "DistModelsDir", "DistStage"):
            self.assertIn("$ui.%s.Add_TextChanged" % box, page)
        # папка принимает только папку
        self.assertIn("Get-ModelsDirProblem", d)
        self.assertIn("Confirm-DistModelsDir", d)

    def test_clicks_mark_dirty_and_redraw_only_touched_rows(self):
        d = self.dist
        click = d[d.index("function Invoke-DistCellClick"):d.index("function Invoke-DistBulk")]
        self.assertLess(click.index("Enter-DistDirty"), click.index("Invoke-ProjectGridClick"))
        self.assertNotIn("Save-DistSelection", click)
        self.assertNotIn("Save-DistProject", click)               # клик не пишет файл
        self.assertIn("Update-DistGridRows $row.Group", click)
        self.assertNotIn("Show-DistGrid", click.split("Update-DistGridRows")[0].split("Invoke-ProjectGridClick")[1])   # без полной перерисовки
        req = d[d.index("function Request-DistSave"):d.index("function Flush-DistSave")]
        self.assertNotIn("Interval = 700", d)                     # отложенного автосохранения больше нет
        self.assertNotIn("Save-DistProject", req)
        self.assertIn("Enter-DistDirty", req)
        self.assertIn("$script:DistSaveTimer.Stop()", d[d.index("function Save-DistProject"):d.index("function Request-DistSave")])
        self.assertIn("$ui.DistStage.Add_TextChanged({ Request-DistSave })", d)
        load = d[d.index("function Invoke-DistLoad"):d.index("function Invoke-DistAssign")]
        self.assertLess(load.index("Confirm-DistDirtyLeave"), load.index("Enter-DistDirty"))
        self.assertLess(load.index("Enter-DistDirty"), load.index("Update-ProjectSpecs"))
        self.assertNotIn("[void](Save-DistProject)", load)         # загрузка сама файл не переписывает
        close = self.ui[self.ui.index("$form.Add_FormClosing"):]
        self.assertLess(close.index("Confirm-DistDirtyLeave"), close.index("Flush-DistSave"))
        count = d[d.index("function Invoke-DistCount"):d.index("function Invoke-DistCellClick")]
        self.assertLess(count.index("Confirm-DistDirtyLeave"), count.index("Get-DistProjectModelPaths"))
        self.assertLess(count.index("Get-DistProjectModelPaths"), count.index("Write-RunFiles"))
        # одна двойная буферизация и SuspendLayout вокруг обновления
        self.assertIn('GetProperty("DoubleBuffered"', d)
        self.assertIn("SuspendLayout()", d[d.index("function Update-DistGridRows"):d.index("function Show-DistGrid")])
        self.assertIn("ResumeLayout()", d[d.index("function Update-DistGridRows"):d.index("function Show-DistGrid")])

    def test_bulk_buttons_above_in_project_column(self):
        d = self.dist
        for c in ("DistAll", "DistNone"):
            self.assertRegex(d, r"\$ui\.%s = " % c)
        self.assertIn('"Выделить всё"', d)
        self.assertIn('"Снять всё"', d)
        page = d[d.index("function New-DistPage"):]
        self.assertLess(page.index("$ui.DistAll = "), page.index("$ui.DistGrid = $g"))       # над таблицей
        bulk = d[d.index("function Invoke-DistBulk"):d.index("function Invoke-DistToggleGroup")]
        self.assertIn("$script:DistExpand.Hide", bulk)                                       # с учётом «Скрыть не участвующие»
        self.assertNotIn("Request-DistSave", bulk)
        self.assertNotIn("Save-DistSelection", bulk)
        self.assertEqual(bulk.count("Show-DistGrid"), 1)
        self.assertLess(bulk.index("Enter-DistDirty"), bulk.index("Set-ProjectInProjectBulk"))

    def test_bulk_actions_hold_a_dirty_copy_until_the_user_saves(self):
        d, ui = self.dist, self.ui
        self.assertIn('DistDirtyNote.Text = "Есть несохранённые изменения"', d)
        self.assertIn('New-DistButton "Сохранить"', d)
        self.assertIn('New-DistButton "Сохранить как новый проект"', d)
        self.assertIn('New-DistButton "Отменить"', d)
        req = d[d.index("function Request-DistSave"):d.index("function Flush-DistSave")]
        self.assertIn("Enter-DistDirty", req)
        self.assertNotIn("Save-DistProject", req)
        save = d[d.index("function Save-DistProject"):d.index("function Request-DistSave")]
        self.assertIn("$script:DistDirty", save)
        self.assertNotIn("function Save-DistSelection", d)          # «В проекте» — только [inproject], пишется вместе с настройкой
        click = d[d.index("function Invoke-DistCellClick"):d.index("function Invoke-DistBulk")]
        self.assertLess(click.index("Enter-DistDirty"), click.index("Invoke-ProjectGridClick"))
        self.assertNotIn("Request-DistSave", click)
        assign = d[d.index("function Invoke-DistAssign"):d.index("function Invoke-DistCount")]
        self.assertLess(assign.index("Enter-DistDirty"), assign.index("Set-ProjectDistribution"))
        sw = d[d.index("function Switch-Project"):d.index("function Invoke-NewProject")]
        self.assertLess(sw.index("Confirm-DistDirtyLeave"), sw.index("Save-DistProject"))
        close = ui[ui.index("$form.Add_FormClosing"):]
        self.assertLess(close.index("Confirm-DistDirtyLeave"), close.index("Flush-DistSave"))
        start = ui[ui.index("function Start-Batch"):ui.index("function Start-Batch") + 500]
        self.assertLess(start.index("Confirm-DistDirtyLeave"), start.index("Save-Settings"))

    def test_only_engineering_models_are_taken(self):
        files = self.dist[self.dist.index("function Get-DistModelFiles"):self.dist.index("function Confirm-DistModelsDir")]
        self.assertIn("Select-DistModelPaths", files)
        self.assertIn("Test-ProjectEngineeringModel", self.cfg_text if hasattr(self, "cfg_text") else read(CFG))
        self.assertNotIn("IOS", files.replace("Get-IosModelsFromFolder", ""))
        self.assertNotIn("Test-IsIosRvtName", files)
        self.assertNotIn("Get-SectionHeader", self.ui)


class GridLookAndButtonsTests(unittest.TestCase):
    """Внешний вид таблицы и кнопок вкладки «Спецификации» (проверка текста окна; визуально — на Windows)."""
    def setUp(self):
        self.ui = read(UI)
        self.dist = read(DIST)
        self.cfg = read(CFG)

    def test_no_plus_minus_boxes_and_chevron_is_drawn_thin(self):
        for ch in ("\u229e", "\u229f"):                                           # ⊞ ⊟
            self.assertNotIn(ch, self.dist + self.cfg + self.ui)
        self.assertNotIn("PcPlus", self.cfg)
        self.assertIn("function Write-DistChevron", self.dist)
        self.assertIn("New-Object System.Drawing.Pen($Color, 1.3)", self.dist)          # тонкая линия, не залитый треугольник
        self.assertNotIn("FillPolygon", self.dist)
        self.assertIn("$g.Add_CellPainting", self.dist)
        # раскрытие: значок слева и двойной щелчок по названию
        self.assertIn("$g.Add_CellMouseDoubleClick", self.dist)
        self.assertIn("function Invoke-DistToggleGroup", self.dist)

    def test_thin_light_grid_and_no_separator_columns_drawn(self):
        d = self.dist
        for line in ("$g.GridColor = $grid", "$g.CellBorderStyle = [System.Windows.Forms.DataGridViewCellBorderStyle]::Single",
                     "$g.ColumnHeadersBorderStyle = [System.Windows.Forms.DataGridViewHeaderBorderStyle]::Single",
                     "$g.AdvancedCellBorderStyle.All = [System.Windows.Forms.DataGridViewAdvancedCellBorderStyle]::Single",
                     "$g.AdvancedColumnHeadersBorderStyle.All = [System.Windows.Forms.DataGridViewAdvancedCellBorderStyle]::Single"):
            self.assertIn(line, d)
        self.assertNotIn("DimGray", d)                                                  # толстых тёмных разделителей между разделами нет
        self.assertIn('Where-Object { $_.Kind -ne "sep" }', d)                          # столбцы-разделители в таблицу не попадают
        self.assertIn("HeaderCell.Style.BackColor = Get-DistTint", d)                    # цветные заголовки разделов остались

    def test_buttons_are_plain_same_height_in_order(self):
        d = self.dist
        page = d[d.index("function New-DistPage"):]
        texts = ['"Загрузить спецификации и модели"', '"Распределить"', '"В сводную"', '"В разделы"', '"Посчитать элементы"']
        pos = [page.index("New-DistButton " + t) for t in texts]
        self.assertEqual(pos, sorted(pos))                                              # слева направо
        for ch in "\u2460\u2461\u2462\u2463\u2192":                                  # ① ② ③ ④ →
            self.assertNotIn(ch, d + self.ui)
        btn = d[d.index("function New-DistButton"):d.index("function New-DistPage")]
        self.assertIn("FlatStyle]::Standard", btn)
        self.assertEqual(btn.count(", 30)"), 2)                                         # одна высота 30
        # режим вставки — выпадающий список рядом с «В разделы»
        self.assertIn("$ribbon.Controls.AddRange(@($ui.DistLoad, $ui.DistAssign, $ui.RunSummary, $ui.RunSections, $ui.RunClassifier, $combo))", page)
        box = self.ui[self.ui.index("function New-InsertModeBox"):self.ui.index("function Get-CombinedInsertMode")]
        self.assertIn("DropDownList", box)
        self.assertIn('@("Добавить недостающие", "Заменить все")', self.ui)

    def test_font_is_segoe_ui_9_and_table_not_bold(self):
        import re as _re
        for text in (self.ui, self.dist):
            for m in _re.finditer(r'Font\("([^"]+)",\s*([0-9.]+)', text):
                self.assertEqual((m.group(1), m.group(2)), ("Segoe UI", "9"), m.group(0))
        grid = self.dist[self.dist.index("function New-DistPage"):]
        self.assertNotIn("Bold", self.dist)
        self.assertIn("$g.DefaultCellStyle.Font = $plain", grid)
        self.assertIn("$g.ColumnHeadersDefaultCellStyle.Font = $plain", grid)


if __name__ == "__main__":
    unittest.main()


@unittest.skipUnless(PWSH, "нет pwsh")
class DistRsnTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="bvs_rsn_")
        self.addCleanup(shutil.rmtree, self.tmp, True)

    def test_rsn_label_filter_and_problem(self):
        out = run(self.tmp, r"""
$p = @('RSN://srv/Проект/ИОС/A_IOS_OV_1.rvt', 'RSN://srv/Проект/ИОС/A_IOS_ВК_1.rvt', 'RSN://srv/Проект/Сводная.rvt', 'RSN://srv/Проект/ИОС/A_AR_1.rvt')
'label: ' + (Get-RsnSourceLabel $p)
'label1: ' + (Get-RsnSourceLabel @($p[0]))
'sel: ' + (@(Select-DistModelPaths $p 'C:/x/Сводная.rvt' '') -join ';')
'problem: [' + (Get-ModelsDirProblem 'RSN://srv/Проект') + ']'
'rsn: ' + (Test-RsnModelsSource ' rsn://a') + ' ' + (Test-RsnModelsSource 'C:\a')
""")
        self.assertIn("label: RSN://srv/Проект", out)
        self.assertIn("label1: RSN://srv/Проект/ИОС", out)
        self.assertIn("sel: RSN://srv/Проект/ИОС/A_IOS_OV_1.rvt;RSN://srv/Проект/ИОС/A_IOS_ВК_1.rvt", out)
        self.assertIn("problem: []", out)
        self.assertIn("rsn: True False", out)

    def test_rsn_models_roundtrip(self):
        path = os.path.join(self.tmp, "p.cfg")
        out = run(self.tmp, r"""
$c = New-ProjectConfig 'П'
$c.ModelsDir = 'RSN://srv/Проект'
$c.ModelsRsn.Add('RSN://srv/Проект/A_IOS_OV_1.rvt'); $c.ModelsRsn.Add('RSN://srv/Проект/A_IOS_VK_1.rvt')
Write-ProjectConfig $c '%s'
$r = Read-ProjectConfig '%s'
'dir: ' + $r.ModelsDir
'rsn: ' + ($r.ModelsRsn -join ';')
""" % (path, path))
        self.assertIn("dir: RSN://srv/Проект", out)
        self.assertIn("rsn: RSN://srv/Проект/A_IOS_OV_1.rvt;RSN://srv/Проект/A_IOS_VK_1.rvt", out)

    def test_ui_wiring(self):
        s = read(DIST)
        self.assertIn('"С сервера…"', s)
        self.assertIn("Get-ServerModels", s)
        self.assertNotIn("FolderBrowserDialog", s)
        self.assertIn("Test-RsnModelsSource $dir", s)


@unittest.skipUnless(PWSH, "нет pwsh")
class SummaryLinkTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="bvs_links_")
        self.addCleanup(shutil.rmtree, self.tmp, True)

    def test_links_build_sections_and_do_not_change_file_format(self):
        path = os.path.join(self.tmp, "распределение_T.txt")
        links = os.path.join(self.tmp, "summary_links.txt")
        with open(links, "w", encoding="utf-8-sig", newline="\n") as f:
            f.write("# comment\n\nA_IOS_OV_1.rvt\nB_IOS_VK_1.RVT : 1\nA_IOS_SVOD_T.rvt\nA_IOS_TASK_1\nA_IOS_XYZ_K1\nA_IOS_QQ_K1\na_ios_ov_1\nAR\nСводная модель\n")
        out = run(self.tmp, r"""
$c = New-ProjectConfig 'П'
Update-ProjectSpecs $c @('1_А_Все')
$c.InProject = @{}; $c.InProject[(Get-ProjectSpecKey '1_А_Все')] = $true
Update-ProjectModels $c @('A_IOS_OV_1','A_IOS_VK_1','A_IOS_PD_1')
Set-ProjectCell $c 'A_IOS_OV_1' '1_А_Все' 1 'm' 3
Set-ProjectCell $c 'A_IOS_VK_1' '1_А_Все' 1 'm' 4
$snap = Copy-ProjectConfig $c
$idle = Set-ProjectModelsFromLinkNames $c @('A_IOS_SVOD_T','A_TASK_1','BASE','')
'idle: ' + $idle.Applied + ' ' + (($c.Models | ForEach-Object { $_.Name }) -join ',')
$names = @(Read-SummaryLinksFile '%s')
'read: ' + ($names -join ',')
'missing file: ' + @(Read-SummaryLinksFile 'нет_такого.txt').Count
$r = Set-ProjectModelsFromLinkNames $c $names
'applied: ' + $r.Applied
'models: ' + (($c.Models | ForEach-Object { $_.Name + '=' + $_.Section }) -join ',')
'secs: ' + (($c.Sections | ForEach-Object { $_.Code }) -join ',')
'unknown: ' + $r.Unknown
'in: ' + $c.InProject.ContainsKey((Get-ProjectSpecKey '1_А_Все'))
'cell ov: ' + ($(if (Get-ProjectCell $c 'A_IOS_OV_1' '1_А_Все') { 'yes' } else { 'no' }))
'cell vk: ' + ($(if (Get-ProjectCell $c 'A_IOS_VK_1' '1_А_Все') { 'yes' } else { 'no' }))
'one: ' + (Format-DistMissingFiles 'A_IOS_PD_1')
$miss = @(Get-NamesMissingInPaths $r.Names @('C:\m\A_IOS_OV_1.rvt'))
'miss: ' + ($miss -join ',')
Write-ProjectConfig $c '%s'
$txt = [IO.File]::ReadAllText('%s')
'fmt ov: ' + $txt.Contains("A_IOS_OV_1`tOV")
'fmt unk: ' + $txt.Contains("A_IOS_XYZ_K1")
'fmt qq: ' + $txt.Contains("A_IOS_QQ_K1")
'fmt oldvk: ' + $txt.Contains('A_IOS_VK_1')
'fmt links: ' + $txt.Contains('[links]')
'fmt pd: ' + $txt.Contains('A_IOS_PD_1')
'back: ' + (($snap.Models | ForEach-Object { $_.Name }) -join ',')
'skipped: ' + $r.Skipped + ' ' + (Format-SkippedModelsStatus $r.Skipped)
""" % (links, path, path))
        self.assertEqual(out[0], "idle: False A_IOS_OV_1,A_IOS_VK_1,A_IOS_PD_1")
        self.assertEqual(out[1], "read: A_IOS_OV_1,B_IOS_VK_1,A_IOS_SVOD_T,A_IOS_TASK_1,A_IOS_XYZ_K1,A_IOS_QQ_K1,a_ios_ov_1,AR,Сводная модель")
        self.assertEqual(out[2], "missing file: 0")
        self.assertEqual(out[3], "applied: True")
        self.assertEqual(out[4], "models: A_IOS_OV_1=OV,B_IOS_VK_1=VK")
        self.assertEqual(out[5], "secs: OV,VK")
        self.assertEqual(out[6], "unknown: ")
        self.assertEqual(out[7:12], ["in: True", "cell ov: yes", "cell vk: no", "one: нет файла в папке: A_IOS_PD_1.", "miss: B_IOS_VK_1"])
        self.assertEqual(out[12:18], ["fmt ov: True", "fmt unk: False", "fmt qq: False", "fmt oldvk: False", "fmt links: False", "fmt pd: False"])
        self.assertEqual(out[18], "back: A_IOS_OV_1,A_IOS_VK_1,A_IOS_PD_1")
        self.assertEqual(out[19], "skipped: 2 Пропущено 2 моделей без целевого раздела")

    def test_paths_come_from_folder_only_for_summary_links(self):
        out = run(self.tmp, r"""
$script:Project = New-ProjectConfig 'П'
Update-ProjectModels $script:Project @('A_IOS_OV_1','A_IOS_VK_1','A_IOS_PD_1')
$r = Set-ProjectModelsFromLinkNames $script:Project @('A_IOS_OV_1.rvt','A_IOS_PD_1','A_IOS_SVOD_T.rvt','A_IOS_XYZ_K1')
function Get-DistModelFiles { return @('C:\m\A_IOS_OV_1.rvt','C:\m\A_IOS_VK_1.rvt') }
$script:DistModelPaths = @{ 'A_IOS_VK_1' = 'C:\old\should-not-use.rvt' }
$miss = @(Update-DistModelPathsForNames $r.Names)
'umiss: ' + ($miss -join ',')
$mp = Get-DistProjectModelPaths
'paths: ' + ($mp.Paths -join ',')
'missing: ' + ($mp.Missing -join ',')
'cache vk: ' + $script:DistModelPaths.ContainsKey('A_IOS_VK_1')
""")
        self.assertEqual(out, [
            "umiss: A_IOS_PD_1",
            "paths: C:\\m\\A_IOS_OV_1.rvt",
            "missing: A_IOS_PD_1",
            "cache vk: False",
        ])

    def test_link_suffix_keeps_model_cells_and_asg_and_reports_missing_file(self):
        out = run(self.tmp, r"""
$env:USERNAME = 'petrov'
$env:USER = 'petrov'
$c = New-ProjectConfig 'П'
Update-ProjectSpecs $c @('1_А')
Update-ProjectModels $c @('A_IOS_OV_1','A_IOS_VK_1')
Set-ProjectCell $c 'A_IOS_OV_1' '1_А' 1 'm' 4
Set-ProjectCell $c 'A_IOS_VK_1' '1_А' 1 'm' 5
$c.AsgModels = @{ 'A_IOS_OV_1' = $true; 'A_IOS_VK_1' = $true }
$c.AsgPairs = @{ 'A_IOS_OV_1|1_А' = $true; 'A_IOS_OV_1|*' = $true; 'A_IOS_VK_1|1_А' = $true; 'A_IOS_VK_1|*' = $true }
$script:AsgSel = @{ Models = @{ 'A_IOS_OV_1' = $true; 'A_IOS_VK_1' = $true }; Pairs = @{ 'A_IOS_VK_1|1_А' = $true; 'A_IOS_OV_1|*' = $true }; Specs = $null }
$r = Set-ProjectModelsFromLinkNames $c @('a_ios_ov_1_ОТСОЕДИНЕНО.rvt : 2', 'A_IOS_OV_1_Petrov', 'A_IOS_PD_1') @('C:\m\A_IOS_OV_1.rvt')
Update-AsgSelForProjectModels $c
'applied: ' + $r.Applied
'models: ' + (($c.Models | ForEach-Object { $_.Name + '=' + $_.Section }) -join ',')
'leaving: ' + ($r.Leaving -join ',')
'missing: ' + ($r.Missing -join ',')
'cell ov: ' + ($(if (Get-ProjectCell $c 'A_IOS_OV_1' '1_А') { 'yes' } else { 'no' }))
'cell vk: ' + ($(if (Get-ProjectCell $c 'A_IOS_VK_1' '1_А') { 'yes' } else { 'no' }))
'asg: ' + ((@($c.AsgModels.Keys) | Sort-Object) -join ',')
'pairs: ' + ((@($c.AsgPairs.Keys) | Sort-Object) -join ',')
'sel models: ' + ((@($script:AsgSel.Models.Keys) | Sort-Object) -join ',')
'sel pairs: ' + ((@($script:AsgSel.Pairs.Keys) | Sort-Object) -join ',')
'key: ' + (Get-LinkMatchKey 'A_IOS_OV_1.rvt : 4') + ' ' + (Get-LinkMatchKey 'C:\m\a_ios_ov_1_отсоединено.rvt')
'thin: ' + (Test-SummaryLinksThin 5 2) + ' ' + (Test-SummaryLinksThin 4 2) + ' ' + (Test-SummaryLinksThin 0 3)
'ask no: ' + (Confirm-DistYes 'no' 'текст' 'заголовок')
$many = 1..12 | ForEach-Object { 'M{0}' -f $_ }
'leave: ' + (Format-LeavingModels $many)
'col: ' + ((Get-DistColumnRebuildPrompt @('A_IOS_VK_1')) -replace "`r`n", ' | ')
function Get-DistModelFiles { return @('C:\m\a_ios_ov_1_отсоединено.rvt') }
$script:Project = $c
$script:DistModelPaths = @{}
$mp = Get-DistProjectModelPaths
'fallback: ' + ($mp.Paths -join ',')
'fallback miss: ' + ($mp.Missing -join ',')
""")
        self.assertEqual(out[0], "applied: True")
        self.assertEqual(out[1], "models: A_IOS_OV_1=OV,A_IOS_PD_1=PD")
        self.assertEqual(out[2], "leaving: A_IOS_VK_1")
        self.assertEqual(out[3], "missing: A_IOS_PD_1")
        self.assertEqual(out[4:8], ["cell ov: yes", "cell vk: no", "asg: A_IOS_OV_1", "pairs: A_IOS_OV_1|*,A_IOS_OV_1|1_А"])
        self.assertEqual(out[8:10], ["sel models: A_IOS_OV_1", "sel pairs: A_IOS_OV_1|*"])
        self.assertEqual(out[10], "key: a_ios_ov_1 a_ios_ov_1")
        self.assertEqual(out[11], "thin: True False False")
        self.assertEqual(out[12], "ask no: False")
        self.assertEqual(out[13], "leave: Уйдут модели (12): M1, M2, M3, M4, M5, M6, M7, M8, M9, M10 и ещё 2.")
        self.assertIn("будут удалены при «Сохранить»", out[14])
        self.assertIn("«Отменить» вернёт всё", out[14])
        self.assertIn("Уйдут модели (1): A_IOS_VK_1.", out[14])
        self.assertEqual(out[15], "fallback: C:\\m\\a_ios_ov_1_отсоединено.rvt")
        self.assertEqual(out[16], "fallback miss: A_IOS_PD_1")

    def test_cancel_restores_model_paths_and_thin_refusal_is_separate(self):
        out = run(self.tmp, r"""
function Show-DistGrid { }
$ui = @{
    DistStage = [pscustomobject]@{ Text = 'П' }
    DistTemplate = [pscustomobject]@{ Text = '' }
    DistModelsDir = [pscustomobject]@{ Text = '' }
    Summary = [pscustomobject]@{ Text = '' }
    Status = [pscustomobject]@{ Text = '' }
}
$script:Project = New-ProjectConfig 'П'
Update-ProjectModels $script:Project @('A_IOS_OV_1','A_IOS_VK_1','A_IOS_PD_1','A_IOS_SS_1')
$script:ProjectPath = ''
$script:DistModelPaths = @{ 'A_IOS_OV_1' = 'C:\keep\OV.rvt' }
Enter-DistDirty
$script:DistModelPaths = @{ 'A_IOS_VK_1' = 'C:\changed\VK.rvt' }
Restore-DistDirty
'restored: ' + $script:DistModelPaths['A_IOS_OV_1']
'changed gone: ' + (-not $script:DistModelPaths.ContainsKey('A_IOS_VK_1'))
'clean: ' + (-not $script:DistDirty)
$plan = Get-SummaryLinkPlan $script:Project @('A_IOS_OV_1') @('C:\m\A_IOS_OV_1.rvt')
'thin plan: ' + (Test-SummaryLinksThin $script:Project.Models.Count @($plan.Names).Count)
'prompt: ' + ((Get-DistThinLinksPrompt @($plan.Names).Count $script:Project.Models.Count $plan.Leaving) -replace "`r`n", ' | ')
""")
        self.assertEqual(out[0], "restored: C:\\keep\\OV.rvt")
        self.assertEqual(out[1:3], ["changed gone: True", "clean: True"])
        self.assertEqual(out[3], "thin plan: True")
        self.assertIn("меньше половины моделей таблицы (1 из 4)", out[4])
        self.assertIn("Уйдут модели (3): A_IOS_VK_1, A_IOS_PD_1, A_IOS_SS_1.", out[4])


@unittest.skipUnless(PWSH, "нет pwsh")
class ModelFolderSkipTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="bvs_walk_")
        self.addCleanup(shutil.rmtree, self.tmp, True)

    def test_archive_old_and_dated_folders_are_not_walked(self):
        root = os.path.join(self.tmp, "модели")
        keep = {
            os.path.join("work", "RVT", "PROJ_IOS_OV_K01.rvt"),
            os.path.join("nested", "deep", "RVT", "PROJ_IOS_PD_K01.rvt"),
            os.path.join("oldmodels", "RVT", "PROJ_IOS_SS_KEEP.rvt"),
            os.path.join("FolderOld", "RVT", "PROJ_IOS_OT_KEEP.rvt"),
            os.path.join("notes_2026-10-19", "RVT", "PROJ_IOS_HC_KEEP.rvt"),
            os.path.join("Старт", "RVT", "PROJ_IOS_SS_START.rvt"),
        }
        skip = {
            os.path.join("work", "RVT", "PROJ_IOS_OV_K01.0001.rvt"),
            os.path.join("work", "RVT", "PROJ_IOS_OV_backup.rvt"),
            os.path.join("Архив_2024", "RVT", "PROJ_IOS_VK_OLD.rvt"),
            os.path.join("Archive", "RVT", "PROJ_IOS_SS_OLD.rvt"),
            os.path.join("Old", "RVT", "PROJ_IOS_EOM_OLD.rvt"),
            os.path.join("Старые", "RVT", "PROJ_IOS_OT_OLD.rvt"),
            os.path.join("OLD_models", "RVT", "PROJ_IOS_HS_OLD.rvt"),
            os.path.join("2026-10-19", "RVT", "PROJ_IOS_ITP_OLD.rvt"),
            os.path.join("2026-03-03_копия", "RVT", "PROJ_IOS_PT_OLD.rvt"),
            os.path.join("Backup", "RVT", "PROJ_IOS_OV_BAK.rvt"),
            os.path.join("Резерв", "RVT", "PROJ_IOS_VK_REZ.rvt"),
            os.path.join("_backup", "RVT", "PROJ_IOS_SS_BAK.rvt"),
            os.path.join("Revit_temp", "RVT", "PROJ_IOS_PT_TMP.rvt"),
            os.path.join("Семейства", "RVT", "PROJ_IOS_OV_FAM.rvt"),
            os.path.join("Хлам", "RVT", "PROJ_IOS_OV_TRASH.rvt"),
            os.path.join("Стар", "RVT", "PROJ_IOS_OT_STAR.rvt"),
            os.path.join("Старое", "RVT", "PROJ_IOS_OT_OLD2.rvt"),
            os.path.join("Старый", "RVT", "PROJ_IOS_OT_OLD3.rvt"),
        }
        for rel in keep | skip:
            path = os.path.join(root, rel)
            os.makedirs(os.path.dirname(path), exist_ok=True)
            with open(path, "wb") as f:
                f.write(b"rvt")
        ui = os.path.join(TOOL, "ui_schedules.ps1")
        dialog = os.path.join(TOOL, "folder_dialog.ps1")
        script = os.path.join(os.path.dirname(root), "walk.ps1")
        with open(script, "w", encoding="utf-8-sig") as f:
            f.write(r"""
function Take-Fns([string]$Path, [string[]]$Names) {
    $errs = $null; $tok = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$tok, [ref]$errs)
    $texts = foreach ($n in $Names) {
        $fn = $ast.FindAll({ param($x) $x -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $x.Name -eq $n }, $true) | Select-Object -First 1
        if (-not $fn) { throw "нет функции $n" }
        $fn.Extent.Text
    }
    return ($texts -join "`n")
}
$src = (Take-Fns '%s' @('Get-NativePath','Resolve-PickedFolderPath')) + "`n" + (Take-Fns '%s' @('Test-SkippedArchiveFolderName','Test-ExcludedModelPath','Test-IsIosRvtName','Test-SkipWalkFolderName','Add-IosRvtFromFolder','Get-IosModelsFromFolder'))
$fn = Join-Path ([IO.Path]::GetTempPath()) ('bvs_walkfn_' + [guid]::NewGuid().ToString('N') + '.ps1')
[IO.File]::WriteAllText($fn, $src, (New-Object System.Text.UTF8Encoding $true))
. $fn
Remove-Item $fn
$found = @(Get-IosModelsFromFolder '%s')
($found | ForEach-Object { [IO.Path]::GetFileName($_) } | Sort-Object) -join ','
""" % (dialog, ui, root))
        r = pwsh("-File", script)
        self.assertEqual(r.returncode, 0, r.stdout + r.stderr)
        got = [l for l in r.stdout.splitlines() if l.strip()][-1]
        self.assertEqual(got, ",".join(sorted(os.path.basename(p) for p in keep)))


@unittest.skipUnless(PWSH, "нет pwsh")
class ExplicitSaveAndHeldModelsTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="bvs_hold_")
        self.addCleanup(shutil.rmtree, self.tmp, True)

    def test_old_unknown_rows_stay_out_of_the_table_until_explicit_save(self):
        path = os.path.join(self.tmp, "распределение_OLD.txt")
        out = run(self.tmp, r"""
$p = '%s'
$t = "`t"
$lines = @(
'[project]','stage=П','template=','models_dir=','summary_model=','google_link=','summary_done=0','insert_mode=','asg_scope=',
'[sections]', ('OV' + $t + 'ОВ'), ('UNKNOWN' + $t + 'Неизв.'),
'[models]', ('A_IOS_OV_1' + $t + 'OV'), ('RG_ORG_BF_K1' + $t + 'UNKNOWN'), ('RG_ORG_GU_K1' + $t + 'BF'), ($t + 'UNKNOWN'),
'[specs]', '1_А',
'[cells]',
('A_IOS_OV_1' + $t + '1_А' + $t + '1' + $t + 'm' + $t + '3'),
('RG_ORG_BF_K1' + $t + '1_А' + $t + '1' + $t + 'm' + $t + '9'),
('RG_ORG_GU_K1' + $t + '1_А' + $t + '0' + $t + 'a' + $t)
)
[IO.File]::WriteAllLines($p, [string[]]$lines, (New-Object System.Text.UTF8Encoding $true))
$c = Read-ProjectConfig $p
'models: ' + (($c.Models | ForEach-Object { $_.Name + '=' + $_.Section }) -join ',')
'secs: ' + (($c.Sections | ForEach-Object { $_.Code }) -join ',')
'held: ' + (($c.HeldModels | ForEach-Object { $_.Name + '=' + $_.Section }) -join ',')
'cell ov: ' + ($(if (Get-ProjectCell $c 'A_IOS_OV_1' '1_А') { 'yes' } else { 'no' }))
'cell bf: ' + ($(if (Get-ProjectCell $c 'RG_ORG_BF_K1' '1_А') { 'yes' } else { 'no' }))
'paths: ' + (@(Get-DistProjectModelPaths).Paths.Count)
$script:Project = $c
$script:DistModelPaths = @{ 'A_IOS_OV_1' = 'C:\m\A_IOS_OV_1.rvt'; 'RG_ORG_BF_K1' = 'C:\m\RG_ORG_BF_K1.rvt' }
$mp = Get-DistProjectModelPaths
'run: ' + ($mp.Paths -join ',')
Write-ProjectConfig $c ($p + '.keep')
$keep = [IO.File]::ReadAllText($p + '.keep')
'kept bf: ' + $keep.Contains("RG_ORG_BF_K1`tUNKNOWN")
'kept gu: ' + $keep.Contains("RG_ORG_GU_K1`tBF")
'kept cell: ' + $keep.Contains("RG_ORG_BF_K1`t1_А`t1`tm`t9")
'kept ov: ' + $keep.Contains("A_IOS_OV_1`tOV")
Clear-ProjectHeld $c
Write-ProjectConfig $c ($p + '.save')
$saved = [IO.File]::ReadAllText($p + '.save')
'dropped bf: ' + (-not $saved.Contains('RG_ORG_BF_K1'))
'dropped gu: ' + (-not $saved.Contains('RG_ORG_GU_K1'))
'still ov: ' + $saved.Contains("A_IOS_OV_1`t1_А`t1`tm`t3")
'sel: ' + (@(Select-DistModelPaths @('C:\m\A_IOS_OV_1.rvt','C:\m\RG_ORG_BF_K1.rvt','C:\m\RG_ORG_GU_K1.rvt') '' '') -join ';')
""" % path)
        self.assertEqual(out[0], "models: A_IOS_OV_1=OV")
        self.assertEqual(out[1], "secs: OV")
        self.assertEqual(out[2], "held: RG_ORG_BF_K1=UNKNOWN,RG_ORG_GU_K1=BF")
        self.assertEqual(out[3:6], ["cell ov: yes", "cell bf: no", "paths: 0"])
        self.assertEqual(out[6], "run: C:\\m\\A_IOS_OV_1.rvt")
        self.assertEqual(out[7:11], ["kept bf: True", "kept gu: True", "kept cell: True", "kept ov: True"])
        self.assertEqual(out[11:14], ["dropped bf: True", "dropped gu: True", "still ov: True"])
        self.assertEqual(out[14], "sel: C:\\m\\A_IOS_OV_1.rvt")

    def test_summary_done_does_not_write_unconfirmed_edits(self):
        path = os.path.join(self.tmp, "распределение_S.txt")
        out = run(self.tmp, r"""
function Show-DistGrid { }
$ui = @{
    DistStage = [pscustomobject]@{ Text = 'П' }
    DistTemplate = [pscustomobject]@{ Text = '' }
    DistModelsDir = [pscustomobject]@{ Text = '' }
    Summary = [pscustomobject]@{ Text = '' }
    Status = [pscustomobject]@{ Text = '' }
}
$script:DistRsnModels = @()
$c = New-ProjectConfig 'П'
Update-ProjectModels $c @('A_IOS_OV_1')
Update-ProjectSpecs $c @('1_А')
$c.InProject = @{}; $c.InProject[(Get-ProjectSpecKey '1_А')] = $true
Set-ProjectCell $c 'A_IOS_OV_1' '1_А' 1 'm' 4
$c.SummaryDone = $false
$p = '%s'
Write-ProjectConfig $c $p
$script:Project = Read-ProjectConfig $p
$script:ProjectPath = $p
$script:SummaryDone = $false
Enter-DistDirty
Set-ProjectCell $script:Project 'A_IOS_OV_1' '1_А' 0 'm' 4
Publish-SummaryDone $true
$txt = [IO.File]::ReadAllText($p)
'done: ' + $txt.Contains('summary_done=1')
'file still on: ' + $txt.Contains("A_IOS_OV_1`t1_А`t1`tm`t4")
'file not off: ' + (-not $txt.Contains("A_IOS_OV_1`t1_А`t0`tm`t4"))
'mem off: ' + ((Get-ProjectCell $script:Project 'A_IOS_OV_1' '1_А').State)
'snap on: ' + ((Get-ProjectCell $script:DistDirtySnapshot 'A_IOS_OV_1' '1_А').State)
'snap done: ' + $script:DistDirtySnapshot.SummaryDone
Restore-DistDirty
'after revert on: ' + ((Get-ProjectCell $script:Project 'A_IOS_OV_1' '1_А').State)
'after revert done: ' + $script:Project.SummaryDone
""" % path)
        self.assertEqual(out, [
            "done: True", "file still on: True", "file not off: True", "mem off: 0", "snap on: 1", "snap done: True",
            "after revert on: 1", "after revert done: True",
        ])
