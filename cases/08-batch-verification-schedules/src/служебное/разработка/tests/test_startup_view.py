"""Открытие окна: таблица из кэша / пустая таблица с подсказкой, ширины столбцов, Excel после вставки (cleanup, CSV вместо Excel), режим вставки и подтверждение «Заменить все»."""
import os
import shutil
import sys
import tempfile
import time
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from test_chain import PWSH  # noqa: E402
from test_project_config import run as run_cfg, P, TOOL  # noqa: E402
import test_two_tabs as tt  # noqa: E402

DIST = os.path.join(TOOL, "project_dist_ui.ps1")
UI = os.path.join(TOOL, "ui_schedules.ps1")
EXCEL = os.path.join(TOOL, "excel_isolated.ps1")


def read(p):
    with open(p, encoding="utf-8-sig") as f:
        return f.read()


@unittest.skipUnless(PWSH, "нет pwsh")
class CacheAndWidthsTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="bvs_sv_")
        self.addCleanup(shutil.rmtree, self.tmp, True)

    def cfg_file(self, name, template, specs, age):
        d = os.path.join(self.tmp, "проект")
        os.makedirs(d, exist_ok=True)
        p = os.path.join(d, name)
        text = "[project]\nstage=П\ntemplate=%s\n[specs]\n%s\n" % (template, "\n".join(specs))
        with open(p, "w", encoding="utf-8-sig", newline="") as f:
            f.write(text)
        t = time.time() - age
        os.utime(p, (t, t))
        return p

    def test_cache_lookup_saved_path_then_template_match(self):
        newest = self.cfg_file("распределение_AAA.txt", "C:\\общий\\Спеки.rvt", ["3.1_Имя_Все"], 10)
        self.cfg_file("распределение_OLD.txt", "C:\\общий\\Спеки.rvt", ["3.1_Имя_Все"], 1000)
        self.cfg_file("распределение_OTHER.txt", "C:\\другой.rvt", ["3.1_Имя_Все"], 1)
        empty = self.cfg_file("распределение_EMPTY.txt", "C:\\общий\\Спеки.rvt", [], 0)
        out = run_cfg(r"""
$rd = '%s'
$tpl = 'C:\общий\Спеки.rvt'
'by template: ' + (Split-Path -Leaf (Find-ProjectCachePath $rd '' $tpl))
'saved wins: ' + (Split-Path -Leaf (Find-ProjectCachePath $rd '%s' ''))
'saved but empty -> template: ' + (Split-Path -Leaf (Find-ProjectCachePath $rd '%s' $tpl))
'case: ' + (Split-Path -Leaf (Find-ProjectCachePath $rd '' 'c:\ОБЩИЙ\спеки.RVT'))
'other template: [' + (Find-ProjectCachePath $rd '' 'C:\нет.rvt') + ']'
'no template, no saved: [' + (Find-ProjectCachePath $rd '' '') + ']'
'no dir: [' + (Find-ProjectCachePath 'C:\нет\такой' '' $tpl) + ']'
""" % (self.tmp.replace("'", "''"), newest.replace("'", "''"), empty.replace("'", "''")))
        self.assertEqual(out[0], "by template: распределение_AAA.txt")                  # самый свежий с тем же общим файлом
        self.assertEqual(out[1], "saved wins: распределение_AAA.txt")
        self.assertEqual(out[2], "saved but empty -> template: распределение_AAA.txt")
        self.assertEqual(out[3], "case: распределение_AAA.txt")
        self.assertEqual(out[4], "other template: []")
        self.assertEqual(out[5], "no template, no saved: []")
        self.assertEqual(out[6], "no dir: []")

    def test_column_widths_fixed_then_equal_share_with_minimum(self):
        out = run_cfg(r"""
. '%s'
$k = @('name', 'inproject', 'summary', 'section', 'section', 'section')
'wide: ' + ((Get-DistColumnWidths 1360 $k) -join ',')
'narrow: ' + ((Get-DistColumnWidths 600 $k) -join ',')
'models: ' + ((Get-DistColumnWidths 1000 @('name', 'inproject', 'summary', 'section', 'model', 'model', 'section')) -join ',')
'no flex: ' + ((Get-DistColumnWidths 800 @('name', 'inproject')) -join ',')
""" % DIST.replace("'", "''"))
        wide = [int(x) for x in out[0].split(": ")[1].split(",")]
        self.assertEqual(wide[:2], [360, 78])                                           # «Группа» и «В проекте» — фиксированные
        self.assertEqual(len(set(wide[2:])), 1)                                         # остальные поровну
        self.assertEqual(sum(wide), 1360 - (1360 - 360 - 78) % 4)
        narrow = [int(x) for x in out[1].split(": ")[1].split(",")]
        self.assertEqual(narrow, [360, 78, 70, 84, 84, 84])                             # узкое окно: минимумы, дальше — прокрутка
        models = [int(x) for x in out[2].split(": ")[1].split(",")]
        self.assertEqual(len(set(models[2:])), 1)
        self.assertEqual(out[3], "no flex: 360,78")


class StaticTests(unittest.TestCase):
    def setUp(self):
        self.ui, self.dist, self.excel = read(UI), read(DIST), read(EXCEL)

    def test_placeholder_table_and_hint(self):
        ph = self.dist[self.dist.index("function Show-DistPlaceholder"):self.dist.index("# Вид таблицы из настройки")]
        for part in ('"Группа / спецификация"', '"В проекте"', '"Сводная"', "$script:PcSections", "$i -lt 40", "[void]$g.Rows.Add()", "FromArgb(243, 243, 243)", "Update-DistHintPosition"):
            self.assertIn(part, ph)
        self.assertIn("Укажите общий файл и нажмите ", self.dist)
        self.assertIn("[char]0x201E", self.dist)                                         # „Загрузить спецификации и модели“
        self.assertIn("Загрузить спецификации и модели", self.dist[self.dist.index("$hint.Text"):self.dist.index("$ui.DistHint = $hint")])
        show = self.dist[self.dist.index("function Show-DistGrid"):self.dist.index("function Save-DistProject")]
        self.assertIn("$null -eq $script:Project -or $script:Project.Specs.Count -eq 0", show)      # нечего показать — заглушка
        self.assertIn("Show-DistPlaceholder", show)
        self.assertIn("$g.BackgroundColor = [System.Drawing.Color]::White", show)
        # центрирование и привязка к краям
        self.assertIn('$g.Dock = "Fill"', self.dist)
        self.assertIn("$g.Add_Resize({ Update-DistColumnWidths; Update-DistHintPosition })", self.dist)
        self.assertIn('$bottom.Dock = "Bottom"', self.ui)                                   # кнопки/статус у нижнего края

    def test_cache_used_at_startup_and_min_window_size(self):
        self.assertIn("Load-DistProject $(if ($open) { Get-ProjectConfigPath $ReportDir $open } else { \"\" })", self.ui)          # таблица — из кэша открытого проекта, Revit не нужен
        self.assertIn("$form.MinimumSize = New-Object System.Drawing.Size(960, 640)", self.ui)
        w = self.dist[self.dist.index("function Update-DistColumnWidths"):self.dist.index("# Подсказка по центру")]
        self.assertIn("MinimumWidth", w)
        self.assertIn("VScrollBar", w)                                                       # вертикальная полоса не съедает ширину
        self.assertIn("ScrollBars]::Both", self.dist)                                        # узкое окно — прокрутка, не сжатие

    def test_excel_is_hidden_and_cleaned_up_on_any_error(self):
        e = self.excel
        self.assertIn("-WindowStyle Hidden", e)
        self.assertIn("$excel.Visible = $false", e)
        fn = e[e.index("function New-IsolatedExcelApplication"):e.index("function Stop-IsolatedExcelApplication")]
        fin = fn[fn.index("finally {"):]
        self.assertIn("Remove-Item -LiteralPath $seed", fin)                                 # своя seed-книга удаляется при любой ошибке
        self.assertIn("Stop-OwnExcelProcess $ownPid", fn)                                    # закрывается только свой процесс (подробно — test_excel_own_process.py)
        self.assertNotIn("Stop-NewExcelProcesses", e)
        self.assertIn("for ($attempt = 1; $attempt -le 2; $attempt++)", fn)
        self.assertIn("TimeoutSec 60", fn)
        self.assertNotIn("Stop-IsolatedExcelApplication", self.ui)                            # отчёт прогона Excel больше не запускает (test_report_xlsx.py); очистка нужна запасным путям

    def test_no_excel_popup_csv_link_instead(self):
        ui = self.ui
        tick = ui[ui.index("$xlsx = Export-ReportXlsx"):ui.index("$form.Controls.Add($ui.Timer)")]
        self.assertIn('Set-CsvLink $script:Run.Summary', tick)
        self.assertNotIn("Excel не собран", tick)
        link = ui[ui.index("function Open-RunCsv"):ui.index("# «Заменить все»: что будет удалено")]
        self.assertIn("explorer.exe", link)
        self.assertNotIn("excel", link.lower())                                              # CSV показывается в проводнике, Excel не запускается
        self.assertIn('$ui.OpenCsv.Text = "Открыть CSV"', ui)
        self.assertIn("Set-CsvLink \"\"", ui[ui.index("function Start-Batch"):])             # новый прогон — ссылка скрыта
        self.assertIn('"Модель;Раздел;Удалено;Вставлено;Пропущено;Ошибка;Режим', ui)       # колонка «Режим» (первичная / добавление / замена)


@unittest.skipUnless(PWSH, "нет pwsh")
class InsertModeBoxTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="bvs_im_")
        self.addCleanup(shutil.rmtree, self.tmp, True)

    def test_default_is_add_missing_and_old_values_map(self):
        out = tt.run(self.tmp, r"""
$script:InsertModeValues = @("add_diff", "replace", "initial")
$script:InsertModeLabels = @("Добавить недостающие", "Заменить все")
$box = [pscustomobject]@{ SelectedIndex = 0; Items = @('a', 'b') }
$ui = @{ InsertModeBox = $box }
'default: ' + (Get-InsertModeValue)
Set-InsertModeBox $box 'initial'; 'initial -> ' + $box.SelectedIndex + ' ' + (Get-InsertModeValue)
Set-InsertModeBox $box 'replace'; 'replace -> ' + $box.SelectedIndex + ' ' + (Get-InsertModeValue)
Set-InsertModeBox $box ''; 'empty -> ' + $box.SelectedIndex
'saved none: ' + (Get-CombinedInsertMode '' '')
""", ["Normalize-InsertModeValue", "Get-InsertModeValue", "Set-InsertModeBox", "Get-CombinedInsertMode"])
        self.assertEqual(out, ["default: add_diff", "initial -> 0 add_diff", "replace -> 1 replace", "empty -> 0", "saved none: add_diff"])

    def test_replace_confirmation_has_models_and_specs_counts(self):
        out = tt.run(self.tmp, r"""
$script:Project = New-ProjectConfig 'П'
Update-ProjectSpecs $script:Project @('П        3.1_Имя_Все', 'П        3.1_Имя_Пустые', 'П        7.8_Диаметр_Все')
Update-ProjectModels $script:Project @('X_IOS_OV_T_1', 'X_IOS_VK_T_1')
$script:Project.InProject = $null; $script:Project.Summary = $null
Set-ProjectCell $script:Project 'X_IOS_OV_T_1' 'П        3.1_Имя_Все' 1 'a'; Set-ProjectCell $script:Project 'X_IOS_OV_T_1' 'П        3.1_Имя_Пустые' 1 'a'
Set-ProjectCell $script:Project 'X_IOS_VK_T_1' 'П        7.8_Диаметр_Все' 1 'a'
$p = Get-ReplacePlan 'sections' @('C:\m\X_IOS_OV_T_1.rvt', 'C:\m\X_IOS_VK_T_1.rvt')
'sections: ' + $p.Models + ' ' + $p.Specs
$s = Get-ReplacePlan 'summary' @('C:\m\SVOD.rvt')
'summary: ' + $s.Models + ' ' + $s.Specs
$t = Get-ReplaceConfirmText 'sections' @('C:\m\X_IOS_OV_T_1.rvt', 'C:\m\X_IOS_VK_T_1.rvt')
'text has: ' + $t.Contains('Моделей: 2') + ' ' + $t.Contains('Спецификаций списка: 3') + ' ' + $t.Contains('Продолжить?') + ' ' + $t.Contains('первичную загрузку')
""", ["Get-ReplacePlan", "Get-ReplaceConfirmText"])
        self.assertEqual(out, ["sections: 2 3", "summary: 1 3", "text has: True True True True"])

    def test_start_batch_asks_before_replace_and_default_no(self):
        ui = read(UI)
        start = ui[ui.index("function Start-Batch"):]
        ask = start[start.index('(Get-InsertModeValue) -eq "replace"'):start.index("Set-CsvLink")]
        self.assertIn("YesNo", ask)
        self.assertIn("MessageBoxDefaultButton]::Button2", ask)                              # по умолчанию «Нет»
        self.assertIn("return", ask)
        self.assertLess(start.index("Get-ReplaceConfirmText"), start.index("Find-BatchRvt"))   # до запуска Revit


if __name__ == "__main__":
    unittest.main()
