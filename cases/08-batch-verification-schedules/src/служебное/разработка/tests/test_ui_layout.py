"""Раскладка окна «Проверочные спецификации» (ui_schedules.ps1) — статическая проверка без запуска WinForms.

Окно запускается только на Windows: здесь проверяется, что шаги «Файлы» и «Задание» построены на TableLayoutPanel
(тянутся вслед за окном), списки делят ширину пополам и занимают высоту, длинные подписи переносятся, а все вызываемые
функции раскладки определены, и скрипт разбирается PowerShell без ошибок.
"""
import os
import re
import shutil
import subprocess
import tempfile
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(os.path.dirname(HERE)))
TOOL = os.path.join(ROOT, "служебное", "скрипты")      # папка с .ps1/.py окна
UI = os.path.join(TOOL, "ui_schedules.ps1")
PWSH = os.environ.get("PWSH") or shutil.which("pwsh")


def src():
    with open(UI, encoding="utf-8-sig") as f:
        return f.read()


def block(text, start, end):
    a = text.index(start)
    return text[a:text.index(end, a)]


class UiLayoutTests(unittest.TestCase):
    def test_specs_tab_is_a_stretching_table_top_to_bottom(self):
        d = open(os.path.join(TOOL, "project_dist_ui.ps1"), encoding="utf-8-sig").read()
        page = d[d.index("function New-DistPage"):]
        self.assertIn("$t = New-LayoutTable", page)
        self.assertIn('Add-LayoutColumn $t "percent" 100', page)
        self.assertNotIn("Location", page)                                   # никаких абсолютных координат
        # порядок строк: пути, стадия, управление таблицей, таблица (забирает остаток высоты), кнопки шагов
        marks = ["Add-LayoutControl $t $paths 0 $r", "Add-LayoutControl $t $flow 0 $r", "Add-LayoutControl $t $flow3 0 $r", "Add-LayoutControl $t $g 0 $r",
                 "Add-LayoutControl $t $bar 0 $r"]
        idx = [page.index(m) for m in marks]
        self.assertEqual(idx, sorted(idx))
        self.assertRegex(page, r"\$r = Add-LayoutRow \$t \$true\s+\$g = New-Object System\.Windows\.Forms\.DataGridView")
        fn = d[d.index("function Add-PathRow"):d.index("function New-DistButton")]
        self.assertIn('$Box.Anchor = "Left, Right"', fn)                       # поле пути забирает всю ширину
        self.assertIn("[System.Windows.Forms.SizeType]::Percent, 100", page)
        # две вкладки, общая строка статуса и «Открыть отчёт» внизу окна
        t = src()
        self.assertEqual(t.count("New-Object System.Windows.Forms.TabPage"), 2)
        self.assertIn('$ui.Tabs.Dock = "Fill"', t)
        self.assertIn('$bottom.Dock = "Bottom"', t)

    def test_assign_page_columns_lists_and_wrapping(self):
        t = src()
        page = block(t, "# ---- раскладка вкладки «Задание»", "$ui.Tabs = New-Object")
        self.assertEqual(len(re.findall(r'Add-LayoutColumn \$asg "percent" 50', page)), 2)       # списки делят ширину пополам
        # «Всё» над списком моделей, «Выборочно» над списком спецификаций — один ряд, колонки 0 и 1
        m = re.search(r"Add-LayoutControl \$asg \$ui\.AsgAll 0 (\$r)\s+Add-LayoutControl \$asg \$ui\.AsgPart 1 (\$r)", page)
        self.assertTrue(m, "радиокнопки не в одном ряду")
        # списки в одной строке, которая забирает остаток высоты, и растягиваются целиком
        self.assertRegex(page, r"\$r = Add-LayoutRow \$asg \$true\s+Add-LayoutControl \$asg \$ui\.AsgTree 0 \$r\s+Add-LayoutControl \$asg \$ui\.AsgSpecs 1 \$r")
        self.assertIn('$ui.AsgTree.Dock = "Fill"', t)
        self.assertIn('$ui.AsgSpecs.Dock = "Fill"', t)
        # пояснения под списками — последняя строка (внизу), все длинные тексты переносятся
        self.assertLess(page.index("$ui.AsgSpecs 1 $r"), page.index("Register-WrapControl $asg $ui.AsgHint"))
        for ctl in ("$ui.CollapseDay", "$ui.AsgHint"):
            self.assertIn("Register-WrapControl $asg " + ctl, page)
        # поле ссылки тянется, кнопки «все»/«ничего» — в правых столбцах заголовка своего списка
        self.assertIn('$ui.Google.Anchor = "Left, Right"', page)
        self.assertIn("$head.Controls.Add($hd.None, 2, 0)", page)
        # маленькое окно — прокрутка, а не сжатые в ноль списки
        self.assertIn("AutoScrollMinSize", page)

    def test_helpers_defined_and_no_absolute_coordinates_in_new_layout(self):
        t = src()
        for name in ("New-LayoutTable", "Add-LayoutColumn", "Add-LayoutRow", "Add-LayoutControl", "Register-WrapControl"):
            self.assertIn("function " + name, t)
        layout = block(t, "# ---- раскладка вкладки «Задание»", "$ui.Tabs = New-Object")
        self.assertNotIn(".Location", layout)

    @unittest.skipUnless(PWSH, "нет pwsh")
    def test_script_parses_and_layout_functions_are_defined(self):
        tmp = tempfile.mkdtemp(prefix="bvs_ui_")
        self.addCleanup(shutil.rmtree, tmp, True)
        ps = os.path.join(tmp, "ast.ps1")
        with open(ps, "w", encoding="utf-8-sig") as f:
            f.write(r"""
$errs = $null; $tok = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile('%s', [ref]$tok, [ref]$errs)
"errors=" + @($errs).Count
$defined = @{}
foreach ($fn in $ast.FindAll({ $args[0] -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)) { $defined[$fn.Name] = $true }
$missing = @()
foreach ($c in $ast.FindAll({ $args[0] -is [System.Management.Automation.Language.CommandAst] }, $true)) {
  $n = $c.GetCommandName()
  if ($n -match '^(Add|New)-Layout|^Register-Wrap|^Add-PathRow|^New-InsertModeToggle' -and -not $defined.ContainsKey($n)) { $missing += $n }
}
"missing=" + (($missing | Sort-Object -Unique) -join ',')
""" % UI)
        env = dict(os.environ, DOTNET_SYSTEM_GLOBALIZATION_INVARIANT="1")
        r = subprocess.run([PWSH, "-NoProfile", "-File", ps], capture_output=True, text=True, encoding="utf-8", env=env)
        self.assertEqual(r.returncode, 0, r.stdout + r.stderr)
        self.assertIn("errors=0", r.stdout)
        self.assertRegex(r.stdout, r"missing=\s*$")


class LinkedListsWiringTests(unittest.TestCase):
    """«Выборочно»: справа спецификации ВЫДЕЛЕННОГО узла, отметки — по парам «модель × спека» (логика — tests/test_partial.py)."""

    def test_wiring(self):
        t = src()
        # клик по модели/разделу -> справа спецификации узла
        self.assertIn("Show-AsgSpecList", block(t, "$ui.AsgTree.Add_AfterSelect", "})\n"))
        self.assertIn("$node.Nodes.Count -gt 0", block(t, "function Get-AsgSelectedModels", "function Get-AsgSelectedTitle"))
        # галочки моделей слева («выгружать модель»): при отметке спеки модели отмечаются все; список и итог пересчитываются
        tree = block(t, "$ui.AsgTree.Add_AfterCheck", "$ui.AsgModelsAll.Add_Click")
        self.assertLess(tree.index("finally"), tree.index("Sync-AsgSpecsWithModels"))
        for btn in ("$ui.AsgModelsAll.Add_Click", "$ui.AsgModelsNone.Add_Click"):
            self.assertIn("Sync-AsgSpecsWithModels", block(t, btn, "})\n"))
        self.assertIn("Update-PartialPicksOnModelCheck", block(t, "function Sync-AsgSpecsWithModels", "# «все»/«ничего» справа"))
        # галочка спеки справа: у модели — её отметка, у раздела — у всех моделей раздела; из «частично» — отметить
        item = block(t, "$ui.AsgSpecs.Add_ItemCheck", "# «все»/«ничего» справа — только для выбранного узла")
        self.assertIn("Set-PartialSpecPick $script:AsgPicks @(Get-AsgSelectedModels)", item)
        self.assertIn("CheckState]::Indeterminate", item)
        self.assertIn("Update-AsgCaptions", item)
        show = block(t, "function Show-AsgSpecList", "function Sync-AsgSpecsWithModels")
        self.assertIn("Get-PartialSpecState", show)
        self.assertIn("[частично]", show)
        self.assertIn("SetItemCheckState", show)
        # «все»/«ничего» справа — только выбранный узел
        self.assertIn("Set-AsgNodeSpecs $true", t)
        self.assertIn("Set-AsgNodeSpecs $false", t)
        self.assertIn("Set-PartialNodeAll", block(t, "function Set-AsgNodeSpecs", "# Для выгрузки"))
        self.assertNotIn("Set-AsgVisibleSpecs", t)
        self.assertNotIn("Update-PartialSpecChecks", t)                   # автоотметка прежней версии удалена
        self.assertNotIn("Get-PartialVisibleSpecs", t)
        # подписи
        caps = block(t, "function Update-AsgCaptions", "# Правый список = спецификации")
        self.assertIn("Выберите модель или раздел", caps)
        self.assertIn("Format-PartialNodeCaption", caps)
        self.assertIn("Format-PartialTotalCaption (Get-AsgPairs)", caps)
        self.assertIn("Caption = $ui.AsgTotal", t)
        self.assertIn("Caption = $ui.AsgSpecCount", t)
        # выгрузка: пары в файл в служебной папке прогона + прежние assign_models/assign_specs
        batch = block(t, "$partialPairs = @(Get-AsgPairs)", "function Start-Batch" if False else "$worksets = \"close_all\"")
        self.assertIn("Write-PartialPairsFile $pairsFile $partialPairs", batch)
        self.assertIn("ConvertTo-PartialCfgLines $pairSummary.Specs $pairSummary.Models $pairsFile", batch)
        self.assertIn('Join-Path $work "assign_pairs.txt"', batch)
        # сохранение: пары; прежний asgmodels=/asgspecs= читается
        self.assertIn("'^(?i)asgpair=(.+)$'", t)
        self.assertIn("'^(?i)asgspecs=(.*)$'", t)
        self.assertIn("'^(?i)asgpairsfmt='", t)                                              # старый ui_settings.txt читается (перенос в проект)
        d = open(os.path.join(TOOL, "project_config.ps1"), encoding="utf-8-sig").read()
        self.assertIn('$L.Add("[asg_pairs]")', d)                                            # а сохраняются пары в настройке проекта
        self.assertIn("$script:AsgSel.Pairs, $script:AsgSel.Specs".replace(", ", " "), block(t, "$script:AsgPicks = New-PartialPicks", "Show-AsgSpecList").replace(", ", " "))
        # причина «57 из 57»: наборы берутся из распределения проекта; диагностика, если набор не определён
        refresh = block(t, "function Refresh-PartialLists", "function Set-AsgScopeEnabled")
        self.assertIn("Get-PartialModelSpecInfo $script:AsgAllModels $script:Project", refresh)
        self.assertIn("Write-AsgSetsDiagnostics", refresh)
        self.assertIn("набор спецификаций не определён", refresh)
        self.assertIn("partial_sets.log", t)
        # длинное пояснение «Выборочно» перенесено в ИНСТРУКЦИЯ.txt, в окне — только диагностика
        self.assertIn("$ui.AsgHint.Text = $diag.Trim()", refresh)
        instr = open(os.path.join(ROOT, "ИНСТРУКЦИЯ.txt"), encoding="utf-8-sig").read()
        for word in ("для каждой модели отдельно", "[частично]", "Всего к выгрузке"):
            self.assertIn(word, instr)

    def test_no_long_help_texts_on_screens(self):
        ui = open(os.path.join(TOOL, "ui_schedules.ps1"), encoding="utf-8-sig").read()
        dist = open(os.path.join(TOOL, "project_dist_ui.ps1"), encoding="utf-8-sig").read()
        for gone in ("Ниже список из блока [сводная]", "«Собрать задание» — выгрузка через RBP", "Ссылка для чата. Чтобы таблица",
                     "Слева галочка модели или раздела", "Настройте распределение спецификаций", "$hintAsg", "$hintUrl", "$hintSum"):
            self.assertNotIn(gone, ui)
        self.assertNotIn("Порядок: общий файл", dist)
        self.assertNotIn("$lbl3", dist)
        # пояснения — во всплывающих подсказках
        self.assertIn('Set-UiTip $ui.DistStage "В фильтрах {СТАДИЯ}', dist)
        self.assertIn('Set-UiTip $box "Совпадение по номеру спецификации', ui)
        self.assertIn("System.Windows.Forms.ToolTip", ui)
        self.assertNotIn("FolderBrowserDialog", ui + dist)
        # строки пояснений на экране — не длиннее одной короткой строки
        for m in re.finditer(r'(?:hint|Info|Hint)\.Text = "([^"]*)"', ui + dist):
            self.assertLessEqual(len(m.group(1)), 120, m.group(1))


if __name__ == "__main__":
    unittest.main()
