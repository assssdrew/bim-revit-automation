"""«Распределить» по категориям: справочник «раздел -> номера/категории», распределение, подсветка, кэш категорий; «Распределить» не открывает модели."""
import os
import shutil
import sys
import tempfile
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from test_chain import PWSH  # noqa: E402
from test_project_config import run, P, TOOL  # noqa: E402
from test_dist_groups import read_names, TSV  # noqa: E402

DIST = os.path.join(TOOL, "project_dist_ui.ps1")

# небольшой проект: спецификации с категориями и четыре модели
SMALL = r"""
function New-Small {
    $c = New-ProjectConfig 'П'
    $specs = @(
        @('0_Ошибка типоразмера', ''), @('1_Классификатор_Все', ''), @('9_Наименование_Все', ''),
        @('7.1_Диаметр_Воздуховоды_Все', 'Воздуховоды'), @('7.1_Диаметр_Детали возд-в_Все', 'Соединительные детали воздуховодов'),
        @('7.2_Диаметр_Трубы_Все', 'Трубы'), @('13_Материал_Трубы_Все', 'Трубы'), @('99_Толщина стенки воздуховодов', 'Воздуховоды'),
        @('15.1_Мощность тепловая_Все', 'Оборудование'), @('16.1_Корпус_Оборудование_Все', 'Оборудование'),
        @('7.8_Диаметр_Короба_Все', 'Короба'), @('6.6_Высота_Воздухор-ли_Все', 'Воздухораспределители'))
    Update-ProjectSpecs $c @($specs | ForEach-Object { $P + $_[0] })
    foreach ($s in $specs) { if ($s[1]) { $c.Categories[(Get-ProjectSpecKey ($P + $s[0]))] = $s[1] } }
    $c.InProject = $null
    return $c
}
function New-TestCatalog([string]$text) {
    $f = Join-Path ([IO.Path]::GetTempPath()) ('bvs_cat_' + [guid]::NewGuid().ToString('N') + '.txt')
    [IO.File]::WriteAllText($f, $text, (New-Object System.Text.UTF8Encoding $true))
    $r = Read-SectionCatalog $f
    Remove-Item $f
    return $r
}
function Ticked($c, [string]$model) { (@(Get-ProjectEnabledSpecs $c $model) | ForEach-Object { (Remove-SpecStagePrefix $_) }) -join ' | ' }
"""


@unittest.skipUnless(PWSH, "нет pwsh")
class SectionCatalogTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="bvs_cat_")
        self.addCleanup(shutil.rmtree, self.tmp, True)

    def test_default_file_is_created_once_and_has_all_sections(self):
        out = run(r"""
$f = Join-Path ([IO.Path]::GetTempPath()) ('bvs_catdef_' + [guid]::NewGuid().ToString('N') + '.txt')
Initialize-SectionCatalogFile $f
$r = Read-SectionCatalog $f
'ok: ' + $r.Ok + ' sections: ' + (($script:PcSections | ForEach-Object { $_.Code + '=' + $r.Sections.ContainsKey($_.Code) }) -join ',')
'ov cats: ' + ($r.Sections['OV'].Categories -join ';')
'tray cats: ' + ($r.Sections['EOM'].Categories -join ';')
'except: ' + @($r.Sections['VK'].Except).Count + ' ' + @($r.Sections['OT'].Except).Count + ' ' + @($r.Sections['PD'].Except).Count
[IO.File]::WriteAllText($f, 'мой файл', (New-Object System.Text.UTF8Encoding $true))
Initialize-SectionCatalogFile $f
'kept: ' + [IO.File]::ReadAllText($f, [Text.Encoding]::UTF8).Contains('мой файл')
'path: ' + (Get-SectionCatalogPath 'отчеты')
Remove-Item $f
""")
        self.assertEqual(out[0], "ok: True sections: OV=True,PD=True,VK=True,PT=True,OT=True,HS=True,ITP=True,EOM=True,SS=True,AVT=True,HC=True")
        self.assertEqual(out[1], "ov cats: Воздуховоды;Соединительные детали воздуховодов;Оборудование")
        self.assertEqual(out[2], "tray cats: Кабельные лотки;Соединительные детали кабельных лотков;Короба")
        self.assertEqual(out[3], "except: 1 0 0")
        self.assertEqual(out[4], "kept: True")                                        # существующий файл не перезаписывается
        self.assertEqual(out[5], "path: " + os.path.join("отчеты", "проект", "справочник_разделов.txt"))        # рядом с настройками проектов

    def test_distribution_by_section_category_and_number(self):
        out = run(SMALL + r"""
$c = New-Small
Update-ProjectModels $c @('X_IOS_OV_T_1', 'X_IOS_VK_T_1', 'X_IOS_OT_T_1', 'X_IOS_EOM_T_1', 'X_IOS_PD_T_1')
$cat = New-TestCatalog (Get-DefaultSectionCatalogText)
$r = Set-ProjectDistribution $c $cat
'OV: ' + (Ticked $c 'X_IOS_OV_T_1')
'PD: ' + (Ticked $c 'X_IOS_PD_T_1')
'VK: ' + (Ticked $c 'X_IOS_VK_T_1')
'OT: ' + (Ticked $c 'X_IOS_OT_T_1')
'EOM: ' + (Ticked $c 'X_IOS_EOM_T_1')
'nocat: ' + (($r.NoCategory | ForEach-Object { Remove-SpecStagePrefix $_ }) -join ' | ')
'nosection: ' + (($r.NoSection | ForEach-Object { Remove-SpecStagePrefix $_ }) -join ' | ')
'unknown: ' + @($r.UnknownSections).Count + ' models: ' + $r.Models
""")
        general = "0_Ошибка типоразмера | 1_Классификатор_Все"
        self.assertEqual(out[0], "OV: " + general + " | 7.1_Диаметр_Воздуховоды_Все | 7.1_Диаметр_Детали возд-в_Все | 99_Толщина стенки воздуховодов | 16.1_Корпус_Оборудование_Все")   # 15.1 в ОВ исключён
        self.assertEqual(out[1], "PD: " + general + " | 7.1_Диаметр_Воздуховоды_Все | 7.1_Диаметр_Детали возд-в_Все | 99_Толщина стенки воздуховодов")      # без оборудования и 16.1
        self.assertEqual(out[2], "VK: " + general + " | 7.2_Диаметр_Трубы_Все | 13_Материал_Трубы_Все | 16.1_Корпус_Оборудование_Все")             # без 15.1
        self.assertEqual(out[3], "OT: " + general + " | 7.2_Диаметр_Трубы_Все | 13_Материал_Трубы_Все | 15.1_Мощность тепловая_Все | 16.1_Корпус_Оборудование_Все")   # ОТ = ВК + 15.1
        self.assertEqual(out[4], "EOM: " + general + " | 16.1_Корпус_Оборудование_Все | 7.8_Диаметр_Короба_Все")                                  # 16.1 — по номеру, короба — по категории
        self.assertEqual(out[5], "nocat: 9_Наименование_Все")                                              # нет категории и номера нет в справочнике
        self.assertEqual(out[6], "nosection: 6.6_Высота_Воздухор-ли_Все")                                  # категория ни в одном разделе
        self.assertEqual(out[7], "unknown: 0 models: 5")

    def test_number_zero_is_ticked_in_every_section(self):
        out = run(SMALL + r"""
$c = New-Small
$models = @($script:PcSections | ForEach-Object { 'X_IOS_' + $_.Code + '_T_1' })
Update-ProjectModels $c $models
[void](Set-ProjectDistribution $c (New-TestCatalog (Get-DefaultSectionCatalogText)))
foreach ($s in $script:PcSections) {
    $m = 'X_IOS_' + $s.Code + '_T_1'
    $names = @(Get-ProjectEnabledSpecs $c $m | ForEach-Object { Remove-SpecStagePrefix $_ })
    $s.Code + ': ' + ($names -contains '0_Ошибка типоразмера')
}
""")
        codes = ["OV", "PD", "VK", "PT", "OT", "HS", "ITP", "EOM", "SS", "AVT", "HC"]
        self.assertEqual(out, [c + ": True" for c in codes])

    def test_suffixless_check_numbers_are_ticked_in_every_section(self):
        out = run(SMALL + r"""
$c = New-Small
$extra = @('1_Классификатор_999', '16.0_Корпус_Кириллица', '17.0_Секция_Кириллица')
Update-ProjectSpecs $c @($c.Specs + @($extra | ForEach-Object { $P + $_ }))
$models = @($script:PcSections | ForEach-Object { 'X_IOS_' + $_.Code + '_T_1' })
Update-ProjectModels $c $models
[void](Set-ProjectDistribution $c (New-TestCatalog (Get-DefaultSectionCatalogText)))
$need = @('0_Ошибка типоразмера', '1_Классификатор_999', '16.0_Корпус_Кириллица', '17.0_Секция_Кириллица')
foreach ($s in $script:PcSections) {
    $m = 'X_IOS_' + $s.Code + '_T_1'
    $names = @(Get-ProjectEnabledSpecs $c $m | ForEach-Object { Remove-SpecStagePrefix $_ })
    $miss = @($need | Where-Object { $names -notcontains $_ })
    $s.Code + ': ' + ($(if ($miss.Count -eq 0) { 'ok' } else { $miss -join ',' }))
}
""")
        codes = ["OV", "PD", "VK", "PT", "OT", "HS", "ITP", "EOM", "SS", "AVT", "HC"]
        self.assertEqual(out, [c + ": ok" for c in codes])

    def test_empty_model_gets_section_specs_and_nothing_is_opened(self):
        out = run(SMALL + r"""
$c = New-Small
Update-ProjectModels $c @('X_IOS_VK_T_1')
'cells before: ' + $c.Cells.Count
[void](Set-ProjectDistribution $c (New-TestCatalog (Get-DefaultSectionCatalogText)))
'cells after: ' + $c.Cells.Count + ' ticked: ' + @(Get-ProjectEnabledSpecs $c 'X_IOS_VK_T_1').Count
'distributed: ' + (Test-ProjectDistributed $c)
""")
        self.assertEqual(out[0], "cells before: 0")
        self.assertEqual(out[1], "cells after: 12 ticked: 5")
        self.assertEqual(out[2], "distributed: True")

    def test_not_in_project_manual_and_counts_are_kept(self):
        out = run(SMALL + r"""
$c = New-Small
$c.InProject = @{}
foreach ($n in '1_Классификатор_Все', '7.2_Диаметр_Трубы_Все', '13_Материал_Трубы_Все') { $c.InProject[(Get-ProjectSpecKey ($P + $n))] = $true }
Update-ProjectModels $c @('X_IOS_VK_T_1')
Set-ProjectCellsFromCounts $c @(@{ Model = 'X_IOS_VK_T_1'; Spec = ($P + '7.2_Диаметр_Трубы_Все'); Count = 41 })
Set-ProjectCell $c 'X_IOS_VK_T_1' ($P + '13_Материал_Трубы_Все') 0 'm' 5            # ручная правка: снята
[void](Set-ProjectDistribution $c (New-TestCatalog (Get-DefaultSectionCatalogText)))
'VK: ' + (Ticked $c 'X_IOS_VK_T_1')
$k = Get-ProjectCell $c 'X_IOS_VK_T_1' ($P + '7.2_Диаметр_Трубы_Все'); 'count: ' + $k.State + $k.Origin + $k.Count
$m = Get-ProjectCell $c 'X_IOS_VK_T_1' ($P + '13_Материал_Трубы_Все'); 'manual: ' + $m.State + $m.Origin + $m.Count
""")
        self.assertEqual(out[0], "VK: 1_Классификатор_Все | 7.2_Диаметр_Трубы_Все")            # воздуховоды не в проекте; ручная правка сохранена
        self.assertEqual(out[1], "count: 1a41")
        self.assertEqual(out[2], "manual: 0m5")

    def test_unknown_section_and_highlight(self):
        out = run(SMALL + r"""
$c = New-Small
Update-ProjectModels $c @('X_IOS_VK_T_1', 'X_IOS_ITP_T_1', 'X_IOS_SVOD_T_1', 'Чужая_модель')
$cat = New-TestCatalog "[ВК]`nномера = 0, 1`nкатегории = Трубы`n"                    # раздела ИТП в справочнике нет
$script:PcCatalog = $cat
$r = Set-ProjectDistribution $c $cat
'unknown: ' + ($r.UnknownSections -join ',') + ' models ticked: ' + $r.Models
'ITP: [' + (Ticked $c 'X_IOS_ITP_T_1') + '] cells: ' + (@($c.Specs | Where-Object { Get-ProjectCell $c 'X_IOS_ITP_T_1' $_ }).Count)
$e = @{ Sections = @{}; Groups = @{}; Hide = $false }
$m = Get-ProjectGridModel $c $e
'headers: ' + (($m.Columns | Where-Object { $_.Kind -eq 'section' } | ForEach-Object { $_.Title + '=' + $_.Issue }) -join ';')
$rows = @($m.Rows)
'issues: ' + (($rows | Where-Object { $_.Issue } | ForEach-Object { $_.Title + '=' + $_.Issue }) -join '; ')
$script:PcCatalog = $null
Update-ProjectGridRowState $c $rows[0]
'no catalog: [' + $rows[0].Issue + ']'
""")
        self.assertEqual(out[0], "unknown: ITP models ticked: 1")
        self.assertEqual(out[1], "ITP: [] cells: 0")                                  # не отмечается молча — ячеек нет
        self.assertEqual(out[2], "headers: ВК=;ИТП=unknown")                              # заголовок раздела подсвечивается
        self.assertIn("9_Наименование=nocat", out[3])
        self.assertIn("6.6_Высота_Воздухор-ли=nosection", out[3])
        self.assertEqual(out[4], "no catalog: []")

    def test_numbers_ranges_and_exceptions(self):
        out = run(SMALL + r"""
$cat = New-TestCatalog "[ОВ]`nномера = 16.1-18.1, 0`nкатегории = Воздуховоды`nкроме = 17.0`n"
$t = { param($n) Test-SectionTakesSpec $cat 'OV' $n '' }
'16.1: ' + (& $t '16.1') + ' 17.5: ' + (& $t '17.5') + ' 17.0: ' + (& $t '17.0') + ' 18.1: ' + (& $t '18.1') + ' 18.2: ' + (& $t '18.2') + ' 0: ' + (& $t '0') + ' 1: ' + (& $t '1') + ' 16: ' + (& $t '16')
'cat: ' + (Test-SectionTakesSpec $cat 'OV' '99' 'воздуховоды') + ' except wins: ' + (Test-SectionTakesSpec $cat 'OV' '17.0' 'Воздуховоды')
'other section: ' + (Test-SectionTakesSpec $cat 'VK' '0' '')
""")
        self.assertEqual(out[0], "16.1: True 17.5: True 17.0: False 18.1: True 18.2: False 0: True 1: False 16: False")
        self.assertEqual(out[1], "cat: True except wins: False")
        self.assertEqual(out[2], "other section: False")

    def test_catalog_errors_are_messages_not_crashes(self):
        out = run(SMALL + r"""
$cases = [ordered]@{
  'unknown key' = "[ОВ]`nномера = 0`nцвет = красный`n"
  'unknown section' = "[КС]`nномера = 0`n"
  'bad number' = "[ОВ]`nномера = 0, abc`n"
  'bad range' = "[ОВ]`nномера = 1-2-3`n"
  'no equals' = "[ОВ]`nномера 0`n"
  'outside block' = "номера = 0`n"
  'duplicate' = "[ОВ]`nномера = 0`n[OV]`nномера = 1`n"
  'empty' = "# только комментарий`n"
}
foreach ($k in $cases.Keys) { $r = New-TestCatalog $cases[$k]; $k + ' => ' + $r.Ok + ' | ' + $r.Error }
$r = Read-SectionCatalog 'C:\нет\такого.txt'
'missing => ' + $r.Ok + ' | ' + $r.Error
$bad = New-TestCatalog "[ОВ]`nномера = zzz`n"
try { [void](Set-ProjectDistribution (New-Small) $bad); 'no throw' } catch { 'throws: ' + $_.Exception.Message }
""")
        self.assertIn("unknown key => False | Справочник разделов, строка 3: неизвестный ключ «цвет»", out[0])
        self.assertIn("unknown section => False | Справочник разделов, строка 1: раздел «КС» не распознан", out[1])
        self.assertIn("bad number => False | Справочник разделов, строка 2: «abc» — не номер спецификации", out[2])
        self.assertIn("bad range => False | Справочник разделов, строка 2: «1-2-3» — не номер", out[3])
        self.assertIn("no equals => False | Справочник разделов, строка 2: нет знака «=»", out[4])
        self.assertIn("outside block => False | Справочник разделов, строка 1: строка вне блока", out[5])
        self.assertIn("duplicate => False | Справочник разделов, строка 3: раздел «OV» указан дважды", out[6])
        self.assertIn("empty => False | Справочник разделов пуст", out[7])
        self.assertIn("missing => False | Нет файла справочника разделов", out[8])
        self.assertEqual(out[9], "throws: Справочник разделов недоступен")                     # прямой вызов без проверки Ok — понятная ошибка, окно проверяет Ok заранее

    def test_categories_roundtrip_in_project_file_and_template_reader(self):
        out = run(SMALL + r"""
$c = New-Small
$f = Join-Path ([IO.Path]::GetTempPath()) ('bvs_catrt_' + [guid]::NewGuid().ToString('N') + '.txt')
Write-ProjectConfig $c $f
$d = Read-ProjectConfig $f
'file has section: ' + ((Get-Content $f -Raw) -match '\[categories\]')
'roundtrip: ' + $d.Categories.Count + ' ' + (Get-ProjectSpecCategory $d ($P + '7.2_Диаметр_Трубы_Все'))
Remove-Item $f
""")
        self.assertEqual(out[0], "file has section: True")
        self.assertEqual(out[1], "roundtrip: 9 Трубы")

    def test_samplesite_default_catalog(self):
        names = read_names()
        body = SMALL + "$names = @(" + ",".join("'" + n.replace("'", "''") + "'" for n in names) + ")\n" + r"""
$c = New-ProjectConfig 'РД'; $c.InProject = $null
Update-ProjectSpecs $c $names
$cats = [System.IO.File]::ReadAllLines('%s', (New-Object System.Text.UTF8Encoding $true)) | Where-Object { $_.Trim() }
foreach ($l in $cats) { $p = $l.Split("`t"); if ($p.Count -ge 2 -and $p[1].Trim()) { $c.Categories[(Get-ProjectSpecKey $p[0])] = $p[1].Trim() } }
Update-ProjectModels $c @('PROJ_IOS_OV_ORG_K01_R24', 'PROJ_IOS_PD_ORG_K01_R24', 'PROJ_IOS_VK_ORG_K01_R24', 'PROJ_IOS_OT_ORG_K01_R24', 'PROJ_IOS_EOM_ORG_K01_R24')
$cat = New-TestCatalog (Get-DefaultSectionCatalogText)
$r = Set-ProjectDistribution $c $cat
foreach ($m in $c.Models) { 'N|' + $m.Section + '|' + @(Get-ProjectEnabledSpecs $c $m.Name).Count }
'nocat|' + @($r.NoCategory).Count + '|nosection|' + @($r.NoSection).Count
function Has($m, $n) { (@(Get-ProjectEnabledSpecs $c $m) | Where-Object { (Remove-SpecStagePrefix $_) -like $n }).Count -gt 0 }
'OV 7.1 duct: ' + (Has 'PROJ_IOS_OV_ORG_K01_R24' '7.1_Диаметр_Воздуховоды_*') + ' pipe: ' + (Has 'PROJ_IOS_OV_ORG_K01_R24' '7.2_Диаметр_Трубы_*')
'VK 99 pipe: ' + (Has 'PROJ_IOS_VK_ORG_K01_R24' '99_Толщина стенки труб*') + ' 99 duct: ' + (Has 'PROJ_IOS_VK_ORG_K01_R24' '99_Толщина стенки воздуховодов*')
'15.1 VK: ' + (Has 'PROJ_IOS_VK_ORG_K01_R24' '15.1_*') + ' OT: ' + (Has 'PROJ_IOS_OT_ORG_K01_R24' '15.1_*') + ' OV: ' + (Has 'PROJ_IOS_OV_ORG_K01_R24' '15.1_*')
'EOM tray: ' + (Has 'PROJ_IOS_EOM_ORG_K01_R24' '*Лотки*') + ' duct: ' + (Has 'PROJ_IOS_EOM_ORG_K01_R24' '*Воздуховоды*') + ' 16.1: ' + (Has 'PROJ_IOS_EOM_ORG_K01_R24' '16.1_*')
""" % TSV.replace("'", "''")
        out = run(body)
        counts = {l.split("|")[1]: int(l.split("|")[2]) for l in out if l.startswith("N|")}
        self.assertEqual(set(counts), {"OV", "PD", "VK", "OT", "EOM"})
        self.assertTrue(all(v > 0 for v in counts.values()), counts)
        self.assertGreater(counts["OT"], counts["VK"])                                     # ОТ = ВК + 15.1
        self.assertEqual(counts["OT"] - counts["VK"], 4)                                   # 15.1: номинальная и тепловая × (Все, Без значения)
        line = [l for l in out if l.startswith("nocat|")][0].split("|")
        self.assertEqual((int(line[1]), int(line[3])), NOCAT_NOSECTION)
        self.assertIn("OV 7.1 duct: True pipe: False", out)
        self.assertIn("VK 99 pipe: True 99 duct: False", out)
        self.assertIn("15.1 VK: False OT: True OV: False", out)
        self.assertIn("EOM tray: True duct: False 16.1: True", out)


NOCAT_NOSECTION = (30, 56)      # SampleSite: 30 без категории (9–12, 19, 20) и 56 с категорией вне справочника (воздухораспределители 27, сантехприборы 15, осветительные 8, электрооборудование 6)


class DistAssignDoesNotOpenModelsTests(unittest.TestCase):
    def setUp(self):
        with open(DIST, encoding="utf-8-sig") as f:
            self.d = f.read()

    def body(self, name, nxt):
        return self.d[self.d.index("function " + name):self.d.index("function " + nxt)]

    def test_assign_is_pure_and_count_runs_revit(self):
        assign = self.body("Invoke-DistAssign", "Invoke-DistCount")
        for bad in ("Invoke-RbpAndWait", "Resolve-RevitRun", "Start-Process", "Find-BatchRvt", "Write-RunFiles", "Set-ActionsEnabled", "insert_schedules", "Get-DistProjectModelPaths"):
            self.assertNotIn(bad, assign, bad)
        self.assertIn("Set-ProjectDistribution", assign)
        self.assertIn("Get-SectionCatalogPath", assign)
        self.assertIn("$catalog.Ok", assign)                                                # ошибка справочника — сообщение
        count = self.body("Invoke-DistCount", "Invoke-DistCellClick")
        self.assertIn("Invoke-RbpAndWait", count)                                           # пробная вставка (prefill) — под «Посчитать элементы»
        self.assertIn("Merge-PrefillIntoProject", count)
        self.assertIn('"prefill"', count)
        merge = self.body("Merge-PrefillIntoProject", "Get-DistLogLines")
        self.assertIn("Set-ProjectCellsFromCounts", merge)

    def test_template_categories_are_cached_on_load(self):
        load = self.body("Invoke-DistLoad", "Invoke-DistAssign")
        self.assertIn("Read-TemplateSpecCategories $out", load)
        self.assertIn("$cfgObj.Categories", load)


if __name__ == "__main__":
    unittest.main()
