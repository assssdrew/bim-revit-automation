"""project_config.ps1: настройка проекта («Распределение») — единый источник «какие спеки в какие модели/разделы»."""
import os
import shutil
import sys
import tempfile
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from test_chain import pwsh, PWSH  # noqa: E402

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(HERE)))
TOOL = os.path.join(ROOT, "служебное", "скрипты")      # папка с .ps1/.py окна
LIB = os.path.join(TOOL, "project_config.ps1")
from project_fixture import project_text  # noqa: E402
P = "П        "


def run(body, tmp=None):
    tmp = tmp or tempfile.mkdtemp(prefix="bvs_pc_")
    script = os.path.join(tmp, "t.ps1")
    with open(script, "w", encoding="utf-8-sig") as f:
        f.write(". '%s'\n$P = '%s'\n%s\n" % (LIB, P, body))
    r = pwsh("-File", script)
    assert r.returncode == 0, r.stdout + r.stderr
    return [l for l in r.stdout.splitlines() if l.strip()]


@unittest.skipUnless(PWSH, "нет pwsh")
class ProjectConfigTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="bvs_pc_")
        self.addCleanup(shutil.rmtree, self.tmp, True)

    def test_group_names_and_order(self):
        out = run(r"""
foreach ($n in @(($P + '7.8_Диаметр_Фитингов коробов_Без значения'), ($P + '7.8_Диаметр_Фитингов коробов_Все'), '13_Материал_Трубы_Пустые', ($P + '2.1_Клас-я системы_Воздуховоды_Без значния'),
                 ($P + '99_Толщина стенки труб'), ($P + '0_Ошибка типоразмера'), 'Пров   Ошибка типоразмера', ($P + '1_Классификатор_Уникальные_Без повторов'))) { (Get-SpecGroupName $n) }
$c = New-ProjectConfig 'П'
Update-ProjectSpecs $c @(($P + '13_Материал_Трубы_Пустые'), ($P + '5.1_Ширина_Воздуховоды_Без значения'), ($P + '5.1_Ширина_Воздуховоды_Пустые'), ($P + '5.1_Ширина_Воздуховоды_Все'), ($P + '7.8_Диаметр_Фитингов коробов_Все'), ($P + '2.1_Клас-я_Все'))
'order: ' + ((Get-ProjectGroups $c | ForEach-Object { $_.Name }) -join ' | ')
'tails: ' + (((Get-ProjectGroups $c | Where-Object { $_.Name -like '5.1*' }).Specs | ForEach-Object { $_ -replace '^.*_', '' }) -join ',')
Set-ProjectSpecGroup $c ($P + '2.1_Клас-я_Все') 'Особая группа'
'override: ' + ((Get-ProjectGroups $c | ForEach-Object { $_.Name }) -join ' | ')
Set-ProjectSpecGroup $c ($P + '2.1_Клас-я_Все') ''
'cleared: ' + $c.GroupOverride.Count
""")
        self.assertEqual(out[:8], ["7.8_Диаметр_Фитингов коробов", "7.8_Диаметр_Фитингов коробов", "13_Материал_Трубы", "2.1_Клас-я системы_Воздуховоды",   # группа = имя без хвоста и без префикса стадии
                                   "99_Толщина стенки труб", "0_Ошибка типоразмера", "Пров Ошибка типоразмера", "1_Классификатор"])
        self.assertEqual(out[8], "order: 2.1_Клас-я | 5.1_Ширина_Воздуховоды | 7.8_Диаметр_Фитингов коробов | 13_Материал_Трубы")    # 13 после 7.8: версия, не строка
        self.assertEqual(out[9], "tails: Все,Пустые,Без значения")
        self.assertEqual(out[10], "override: 5.1_Ширина_Воздуховоды | 7.8_Диаметр_Фитингов коробов | 13_Материал_Трубы | Особая группа")
        self.assertEqual(out[11], "cleared: 0")

    def test_sections_from_file_names_only(self):
        out = run(r"""
$c = New-ProjectConfig 'РД'
Update-ProjectModels $c @('PROJ_IOS_OV_ORG_K01_R24.rvt', '\\srv\проект\PROJ_IOS_ВК_ORG_K01_R24', 'PROJ_IOS_SVOD_ORG_R24', 'Чужое_имя', 'PROJ_IOS_OV_ORG_K01_R24', 'PROJ_IOS_TASK_K01', 'PROJ_ПД')
'models: ' + (($c.Models | ForEach-Object { $_.Name + '=' + $_.Section }) -join ';')
'sections: ' + (($c.Sections | ForEach-Object { $_.Code + '=' + $_.Title }) -join ';')
'sec: ' + (Get-ProjectModelSection $c 'Чужое_имя') + '|' + (@(Get-ProjectSectionModels $c 'OV') -join ',')
'view: ' + ((Get-ProjectViewSections $c | ForEach-Object { $_.Title }) -join ',')
'manual api gone: ' + ($null -eq (Get-Command Set-ProjectModelSection -ErrorAction SilentlyContinue))
""")
        self.assertEqual(out[0], "models: PROJ_IOS_OV_ORG_K01_R24=OV;PROJ_IOS_ВК_ORG_K01_R24=VK")    # без кода (сводная, TASK, «Чужое_имя», «_ПД_» без второго «_») — не принимаются
        self.assertEqual(out[1], "sections: OV=ОВ;VK=ВК")
        self.assertEqual(out[2], "sec: |PROJ_IOS_OV_ORG_K01_R24")                  # модели без кода раздела присвоить нельзя
        self.assertEqual(out[3], "view: ОВ,ВК")
        self.assertEqual(out[4], "manual api gone: True")

    def test_section_code_filter_cyrillic_latin_case_and_no_false_hits(self):
        names = [("PROJ_IOS_OV_ORG_K01_R24", "OV"), ("PROJ_IOS_ОВ_ORG_K01", "OV"), ("a_ios_ов_x", "OV"), ("A_pd_X", "PD"), ("A_ПД_X", "PD"),
                 ("A_vk_X", "VK"), ("A_ВК_X", "VK"), ("A_pt_X", "PT"), ("A_ПТ_X", "PT"), ("A_ot_X", "OT"), ("A_От_X", "OT"), ("A_hs_X", "HS"), ("A_хс_X", "HS"),
                 ("A_itp_X", "ITP"), ("A_ИТП_X", "ITP"), ("A_eom_X", "EOM"), ("A_эом_X", "EOM"), ("A_ss_X", "SS"), ("A_СС_X", "SS"),
                 ("\\\\srv\\Проект\\PROJ_IOS_VK_ORG_K01_R24.rvt", "VK"),
                 # ложные срабатывания: IOS/ИОС/ВИС/TASK, граничные условия, сводная, код не между «_»
                 ("PROJ_IOS_SVOD_ORG_R24", ""), ("PROJ_IOS_TASK_K01_R24", ""), ("PROJ_ИОС_ORG", ""), ("PROJ_ВИС_ORG_K01", ""), ("PROJ_IOS_BC_ORG", ""),
                 ("PROJ_IOS_ORG_K01_R24", ""), ("PROJ_OV.rvt", ""), ("OV_PROJ_X", ""), ("PROJ_OVX_Y", ""), ("PROJ_XOV_Y", ""), ("Чужое_имя", ""),
                 ("PROJ_IOS_СВОД_OV_R24", ""), ("PROJ_сводная_OV_R24", "")]
        lines = "\n".join("'%s|' + (Get-ModelSectionFromName '%s') + '|' + (Test-ProjectEngineeringModel '%s')" % (n.replace("'", "''"), n.replace("'", "''"), n.replace("'", "''")) for n, _ in names)
        out = run(lines)
        self.assertEqual(len(out), len(names))
        for line, (n, want) in zip(out, names):
            code = line.split("|")[-2]
            self.assertEqual(code, want, n)
            self.assertEqual(line.split("|")[-1], str(bool(want)), n)

    def test_cells_aggregate_click_and_colors(self):
        out = run(r"""
$c = New-ProjectConfig 'П'; $c.InProject = $null
Update-ProjectSpecs $c @(($P + '3.1_Имя_Все'), ($P + '3.1_Имя_Пустые'))
Update-ProjectModels $c @('X_IOS_OV_T_1', 'X_IOS_OV_T_2', 'X_IOS_ВК_T_1')
# галочки — распределение (a), числа — «Посчитать элементы» (только информация)
Set-ProjectCell $c 'X_IOS_OV_T_1' ($P + '3.1_Имя_Все') 1 'a'; Set-ProjectCell $c 'X_IOS_OV_T_1' ($P + '3.1_Имя_Пустые') 1 'a'
Set-ProjectCell $c 'X_IOS_OV_T_2' ($P + '3.1_Имя_Все') 0 'a'; Set-ProjectCell $c 'X_IOS_OV_T_2' ($P + '3.1_Имя_Пустые') 0 'x'
Set-ProjectCellsFromCounts $c @(@{ Model = 'X_IOS_OV_T_1'; Spec = ($P + '3.1_Имя_Все'); Count = 412 }, @{ Model = 'X_IOS_OV_T_1'; Spec = ($P + '3.1_Имя_Пустые'); Count = 136 },
    @{ Model = 'X_IOS_OV_T_2'; Spec = ($P + '3.1_Имя_Все'); Count = 0 }, @{ Model = 'X_IOS_OV_T_2'; Spec = ($P + '3.1_Имя_Пустые'); Count = -1 },
    @{ Model = 'X_IOS_ВК_T_1'; Spec = ($P + '3.1_Имя_Все'); Count = 7 })
$ct = Get-ProjectCell $c 'X_IOS_ВК_T_1' ($P + '3.1_Имя_Все'); 'count only: ' + $ct.State + $ct.Origin + $ct.Count
$c1 = Get-ProjectCell $c 'X_IOS_OV_T_1' ($P + '3.1_Имя_Все'); 'ticks kept: ' + $c1.State + $c1.Origin + $c1.Count
[void]$c.Cells.Remove((Get-ProjectCellKey 'X_IOS_ВК_T_1' ($P + '3.1_Имя_Все')))
$g = (Get-ProjectGroups $c)[0]
$ov = @('X_IOS_OV_T_1', 'X_IOS_OV_T_2')
function Show($m, $s) { $a = Get-ProjectAggregate $c $m $s; $a.State + '/' + (Get-ProjectAggregateColor $a) + '/' + (Format-ProjectAggregate $a) }
'cell1: ' + (Show @('X_IOS_OV_T_1') @($g.Specs))
'cell2: ' + (Show @('X_IOS_OV_T_2') @($g.Specs))
'section: ' + (Show $ov @($g.Specs))
'spec row: ' + (Show @('X_IOS_OV_T_1') @($P + '3.1_Имя_Все'))
# ручная правка: снять у модели 1 -> жёлтый, ручная
$row = [pscustomobject]@{ Kind = 'group'; Specs = @($g.Specs) }
Invoke-ProjectGridClick $c $row ([pscustomobject]@{ Kind = 'model'; Models = @('X_IOS_OV_T_1') })
'after click: ' + (Show @('X_IOS_OV_T_1') @($g.Specs)) + ' origin=' + (Get-ProjectCell $c 'X_IOS_OV_T_1' ($P + '3.1_Имя_Все')).Origin + ' count=' + (Get-ProjectCell $c 'X_IOS_OV_T_1' ($P + '3.1_Имя_Все')).Count
# по разделу: частично -> отметить у всех
Invoke-ProjectGridClick $c $row ([pscustomobject]@{ Kind = 'section'; Models = $ov })
'section after: ' + (Show $ov @($g.Specs))
Invoke-ProjectGridClick $c $row ([pscustomobject]@{ Kind = 'section'; Models = $ov })
'section off: ' + (Show $ov @($g.Specs))
# автозаполнение не затирает ручное
Set-ProjectCellsFromCounts $c @(@{ Model = 'X_IOS_OV_T_1'; Spec = ($P + '3.1_Имя_Все'); Count = 999 })
$k = Get-ProjectCell $c 'X_IOS_OV_T_1' ($P + '3.1_Имя_Все'); 'manual kept: ' + $k.State + $k.Origin + $k.Count
# модель без ячеек
'empty: ' + (Show @('X_IOS_ВК_T_1') @($g.Specs))
""")
        self.assertEqual(out[0], "count only: 0c7")                                   # число без галочки: признак «c»
        self.assertEqual(out[1], "ticks kept: 1a412")                                 # число обновилось, галочка и признак не изменились
        self.assertEqual(out[2], "cell1: Checked/green/\u2611 548")
        self.assertEqual(out[3], "cell2: Unchecked/red/\u2610 0")                       # «не определено» перекрывает: у одной из двух нет данных
        self.assertEqual(out[4], "section: Indeterminate/red/\u25ea 548")
        self.assertEqual(out[5], "spec row: Checked/green/\u2611 412")
        self.assertTrue(out[6].startswith("after click: Unchecked/yellow/\u2610 548 origin=m count=412"), out[6])
        self.assertTrue(out[7].startswith("section after: Checked/"), out[7])
        self.assertTrue(out[8].startswith("section off: Unchecked/"), out[8])
        self.assertEqual(out[9], "manual kept: 0m999")
        self.assertEqual(out[10], "empty: Unchecked/none/\u2610")

    def test_grid_view_tree_glyphs_single_spec_groups_sections_and_hide(self):
        out = run(r"""
$c = New-ProjectConfig 'П'
Update-ProjectSpecs $c @(($P + '3.1_Имя_Все'), ($P + '3.1_Имя_Пустые'), ($P + '7.8_Диаметр_Все'), ($P + '1_Классификатор_999'))
Update-ProjectModels $c @('PROJ_IOS_OV_ORG_K01_R24', 'PROJ_IOS_OV_ORG_K02_R24', 'PROJ_IOS_PD_ORG_K01_R24', 'PROJ_IOS_SVOD_ORG_R24')
$e = @{ Sections = @{}; Groups = @{}; Hide = $false }
$m = Get-ProjectGridModel $c $e
'collapsed cols: ' + (($m.Columns | ForEach-Object { $_.Kind + ':' + $_.Title }) -join ' | ')
'collapsed rows: ' + (($m.Rows | ForEach-Object { $_.Title + '/' + $_.Expandable }) -join ' | ')
$e.Sections['OV'] = $true; $e.Groups['3.1_Имя'] = $true
$m = Get-ProjectGridModel $c $e
'open cols: ' + (($m.Columns | Where-Object { $_.Kind -ne 'sep' -and $_.Kind -ne 'name' -and $_.Kind -ne 'inproject' -and $_.Kind -ne 'summary' } | ForEach-Object { $_.Kind + ':' + $_.Title + ':' + $_.Tint + ':' + $_.Open }) -join ' | ')
'open rows: ' + (($m.Rows | ForEach-Object { $_.Kind + ':' + $_.Title + ':' + $_.Open }) -join ' | ')
'seps: ' + @($m.Columns | Where-Object { $_.Kind -eq 'sep' }).Count
Invoke-ProjectGroupInClick $c @(($P + '3.1_Имя_Все'), ($P + '3.1_Имя_Пустые'))
$e.Hide = $true
$m = Get-ProjectGridModel $c $e
'hidden rows: ' + (($m.Rows | ForEach-Object { $_.Title }) -join ' | ')
$e.Hide = $false; $m = Get-ProjectGridModel $c $e
$r0 = $m.Rows | Where-Object { $_.Group -eq '3.1_Имя' -and $_.Kind -eq 'group' }
'in-project cell: ' + (Get-ProjectGridCell $c $r0 $m.Columns[1]).Text
$r1 = $m.Rows | Where-Object { $_.Group -eq '7.8_Диаметр' -and $_.Kind -eq 'group' }
'outside: [' + (Get-ProjectGridCell $c $r1 $m.Columns[4]).Text + '] ' + (Get-ProjectGridCell $c $r1 $m.Columns[1]).Text
""")
        # значков раскрытия в тексте нет (рисует окно); Open — признак «раскрыто»
        self.assertEqual(out[0], "collapsed cols: name:Группа / спецификация | inproject:В проекте | summary:Сводная | sep: | section:ОВ | sep: | section:ПД")   # сводная — не раздел
        self.assertEqual(out[1], "collapsed rows: 1_Классификатор/False | 3.1_Имя/True | 7.8_Диаметр/False")               # у группы из одной спеки нет значка
        self.assertEqual(out[2], "open cols: section:ОВ:0:True | model:K01_R24:0:False | model:K02_R24:0:False | section:ПД:1:False")
        self.assertEqual(out[3], "open rows: group:1_Классификатор:False | group:3.1_Имя:True | spec:Все:False | spec:Пустые:False | group:7.8_Диаметр:False")
        self.assertEqual(out[4], "seps: 2")
        self.assertEqual(out[5], "hidden rows: 3.1_Имя | Все | Пустые")
        self.assertEqual(out[6], "in-project cell: \u2611")
        self.assertEqual(out[7], "outside: [] \u2610")

    def test_expanded_group_has_no_blank_rows_and_group_99_has_exactly_two_specs(self):
        """Спецификация, совпадающая с названием группы (99_Толщина стенки труб), раньше давала строку «…» без подписи."""
        sets = os.path.join(HERE, "fixtures", "schedule_sets.txt")
        out = run(r"""
$names = [System.IO.File]::ReadAllLines('%s', [Text.Encoding]::UTF8) | Where-Object { $_ -and $_ -notmatch '^\[' } | ForEach-Object { $_.Trim() } | Select-Object -Unique
$c = New-ProjectConfig 'П'
Update-ProjectSpecs $c @($names)
$e = @{ Sections = @{}; Groups = @{}; Hide = $false }
foreach ($g in @(Get-ProjectGroups $c)) { $e.Groups[$g.Name] = $true }
$m = Get-ProjectGridModel $c $e
$g99 = @($m.Rows | Where-Object { $_.Group -eq '99_Толщина стенки труб' })
'g99: ' + (($g99 | ForEach-Object { $_.Kind + ':' + $_.Title }) -join ' | ')
'blank: ' + @($m.Rows | Where-Object { -not $_.Title.Trim() -or $_.Title.Trim() -eq [string][char]0x2026 }).Count
'spec rows: ' + @($m.Rows | Where-Object { $_.Kind -eq 'spec' }).Count + ' specs in groups>1: ' + (@(Get-ProjectGroups $c | Where-Object { @($_.Specs).Count -gt 1 } | ForEach-Object { @($_.Specs).Count } | Measure-Object -Sum).Sum)
'names clean: ' + @($c.Specs | Where-Object { -not $_.Trim() }).Count
""" % sets.replace("'", "''"))
        self.assertEqual(out[0], "g99: group:99_Толщина стенки труб | spec:(основная) | spec:Без значения")      # группа + ровно две спецификации
        self.assertEqual(out[1], "blank: 0")                                                       # ни в одной группе нет пустых строк
        spec_rows, total = [int(x) for x in out[2].replace("spec rows: ", "").split(" specs in groups>1: ")]
        self.assertEqual(spec_rows, total)
        self.assertEqual(out[3], "names clean: 0")

    def test_bulk_in_project_and_key_cache(self):
        out = run(r"""
$c = New-ProjectConfig 'П'
Update-ProjectSpecs $c @(($P + '3.1_Имя_Все'), ($P + '3.1_Имя_Пустые'), ($P + '7.8_Диаметр_Все'))
Set-ProjectInProjectBulk $c $c.Specs $true
'all: ' + (@($c.Specs | Where-Object { Test-ProjectSpecIn $c $_ }).Count)
Set-ProjectInProjectBulk $c @($P + '3.1_Имя_Все') $false
'one off: ' + (@($c.Specs | Where-Object { Test-ProjectSpecIn $c $_ }).Count)
Set-ProjectInProjectBulk $c $c.Specs $false
'none: ' + (@($c.Specs | Where-Object { Test-ProjectSpecIn $c $_ }).Count)
'cache: ' + ((Get-ProjectSpecKey ($P + '3.1_Имя_Все')) -eq (Get-ProjectSpecKey '3.1_имя_все')) + ' ' + $script:PcSpecKeyCache.ContainsKey($P + '3.1_Имя_Все')
[void](Get-ProjectGroups $c)
'groups cached: ' + ($null -ne $c.GroupsCache)
Update-ProjectSpecs $c @(($P + '3.1_Имя_Все'))
'groups after change: ' + (@(Get-ProjectGroups $c).Count)
""")
        self.assertEqual(out, ["all: 3", "one off: 2", "none: 0", "cache: True True", "groups cached: True", "groups after change: 1"])

    def test_in_project_group_click_and_consumers(self):
        out = run(r"""
$c = New-ProjectConfig 'П'
Update-ProjectSpecs $c @(($P + '3.1_Имя_Все'), ($P + '3.1_Имя_Пустые'), ($P + '7.8_Диаметр_Все'))
Update-ProjectModels $c @('PROJ_IOS_OV_ORG_K01_R24')
foreach ($s in $c.Specs) { Set-ProjectCell $c 'PROJ_IOS_OV_ORG_K01_R24' $s 1 'a' 5 }
'none in project: ' + @(Get-ProjectEnabledSpecs $c 'PROJ_IOS_OV_ORG_K01_R24').Count + ' summary ' + @(Get-ProjectSummarySpecs $c).Count
$g = (Get-ProjectGroups $c)[0]
$row = [pscustomobject]@{ Kind = 'group'; Specs = @($g.Specs); In = 'Unchecked' }
Invoke-ProjectGridClick $c $row ([pscustomobject]@{ Kind = 'inproject'; Models = @() })
'after in-click: ' + (Get-ProjectGroupInState $c $g.Specs) + ' enabled: ' + ((Get-ProjectEnabledSpecs $c 'PROJ_IOS_OV_ORG_K01_R24') -join ',') + ' summary: ' + ((Get-ProjectSummarySpecs $c) -join ',')
'section specs: ' + (@(Get-ProjectSectionSpecs $c 'OV').Count)
# отметка снимается обратным щелчком; частично -> отметить все
Set-ProjectCell $c 'PROJ_IOS_OV_ORG_K01_R24' ($P + '3.1_Имя_Пустые') 0 'm' 5
$c.InProject.Remove((Get-ProjectSpecKey ($P + '3.1_Имя_Пустые')))
'partial: ' + (Get-ProjectGroupInState $c $g.Specs)
Invoke-ProjectGroupInClick $c $g.Specs
'partial->all: ' + (Get-ProjectGroupInState $c $g.Specs)
Invoke-ProjectGroupInClick $c $g.Specs
'all->none: ' + (Get-ProjectGroupInState $c $g.Specs)
# новая спека в отмеченной группе отмечается сама при перечитывании шаблона
Invoke-ProjectGroupInClick $c $g.Specs
Update-ProjectSpecs $c @(($P + '3.1_Имя_Все'), ($P + '3.1_Имя_Пустые'), ($P + '3.1_Имя_Без значения'), ($P + '7.8_Диаметр_Все'))
'new spec in marked group: ' + (Test-ProjectSpecIn $c ($P + '3.1_Имя_Без значения')) + ' other: ' + (Test-ProjectSpecIn $c ($P + '7.8_Диаметр_Все'))
# клик по ячейке распределения группы вне проекта ничего не меняет
$row2 = [pscustomobject]@{ Kind = 'group'; Specs = @($P + '7.8_Диаметр_Все'); In = 'Unchecked' }
Invoke-ProjectGridClick $c $row2 ([pscustomobject]@{ Kind = 'model'; Models = @('PROJ_IOS_OV_ORG_K01_R24') })
'outside click: ' + (Get-ProjectCell $c 'PROJ_IOS_OV_ORG_K01_R24' ($P + '7.8_Диаметр_Все')).Origin
""")
        self.assertEqual(out[0], "none in project: 0 summary 0")
        self.assertEqual(out[1], "after in-click: Checked enabled: 3.1_Имя_Все,3.1_Имя_Пустые summary: 3.1_Имя_Все,3.1_Имя_Пустые")
        self.assertEqual(out[2], "section specs: 2")
        self.assertEqual(out[3], "partial: Indeterminate")
        self.assertEqual(out[4], "partial->all: Checked")
        self.assertEqual(out[5], "all->none: Unchecked")
        self.assertEqual(out[6], "new spec in marked group: True other: False")
        self.assertEqual(out[7], "outside click: a")

    def test_project_file_name_prefix_models_dir_and_run_blocker(self):
        out = run(r"""
'prefix: ' + (Get-ProjectPrefix @('PROJ_IOS_OV_ORG_K01_R24', '\\srv\PROJ_IOS_VK_ORG_K01_R24.rvt', 'XYZ_IOS_PD_X'))
'prefix none: [' + (Get-ProjectPrefix @()) + ']'
'path: ' + (Split-Path -Leaf (Get-ProjectConfigPath 'rep' 'PROJ')) + ' | ' + (Split-Path -Leaf (Get-ProjectConfigPath 'rep' ''))
'dir ok: [' + (Get-ModelsDirProblem '\\srv\проект\RVT') + ']'
'dir empty: ' + (Get-ModelsDirProblem '')
'dir file: ' + (Get-ModelsDirProblem 'C:\проект\PROJ_IOS_OV_ORG_K01_R24.RVT')
'dir file2: ' + (Get-ModelsDirProblem '\\srv\проект\PROJ_IOS_OV_ORG_K01_R24.rvt\')
'blocker empty: ' + (Get-ProjectRunBlocker '')
'blocker spaces: ' + (Get-ProjectRunBlocker '   ')
'blocker missing: ' + (Get-ProjectRunBlocker 'C:\нет\файла.txt')
'blocker ok: [' + (Get-ProjectRunBlocker $PSCommandPath) + ']'
""")
        self.assertEqual(out[0], "prefix: PROJ")
        self.assertEqual(out[1], "prefix none: []")
        self.assertEqual(out[2], "path: распределение_PROJ.txt | распределение_PROJ.txt")
        self.assertEqual(out[3], "dir ok: []")
        self.assertIn("Укажите папку", out[4])
        self.assertIn("указан файл модели", out[5])
        self.assertIn("указан файл модели", out[6])
        self.assertIn("Не задан путь к файлу настройки проекта", out[7])
        self.assertIn("Не задан путь", out[8])
        self.assertIn("не найден", out[9])
        self.assertEqual(out[10], "blocker ok: []")

    def test_file_roundtrip_with_cyrillic_path(self):
        path = os.path.join(self.tmp, "отчеты", "проект", "распределение_PROJ.txt")
        out = run(r"""
$c = New-ProjectConfig 'РД'; $c.Template = 'C:\шаблон\Спеки.rvt'; $c.ModelsDir = '\\srv\проект\RVT'
Update-ProjectSpecs $c @(($P + '3.1_Имя_Все'), ($P + '7.8_Диаметр_Все'))
'stripped: ' + ($c.Specs -join '|')
Update-ProjectModels $c @('PROJ_IOS_OV_ORG_K01_R24', 'PROJ_IOS_ЭОМ_ORG_K01_R24', 'PROJ_IOS_SVOD_ORG_R24')
Set-ProjectSpecGroup $c ($P + '7.8_Диаметр_Все') 'Моя группа'
Invoke-ProjectGroupInClick $c @($P + '3.1_Имя_Все')
Set-ProjectCell $c 'PROJ_IOS_OV_ORG_K01_R24' ($P + '3.1_Имя_Все') 1 'a' 412
Set-ProjectCell $c 'PROJ_IOS_OV_ORG_K01_R24' ($P + '7.8_Диаметр_Все') 0 'm' -1
Set-ProjectCell $c 'PROJ_IOS_ЭОМ_ORG_K01_R24' ($P + '3.1_Имя_Все') 0 'x' -1
Write-ProjectConfig $c '%s'
$b = Read-ProjectConfig '%s'
'proj: ' + $b.Stage + ' ' + $b.Template + ' ' + $b.ModelsDir
'counts: ' + $b.Sections.Count + ' ' + $b.Models.Count + ' ' + $b.Specs.Count + ' ' + $b.Cells.Count + ' ' + $b.GroupOverride.Count + ' ' + $b.InProject.Count
$k = Get-ProjectCell $b 'PROJ_IOS_OV_ORG_K01_R24' ($P + '3.1_Имя_Все'); 'cell: ' + $k.State + $k.Origin + $k.Count
$k = Get-ProjectCell $b 'PROJ_IOS_ЭОМ_ORG_K01_R24' ($P + '3.1_Имя_Все'); 'x: ' + $k.State + $k.Origin + $k.Count
'group: ' + (Get-SpecGroup $b ($P + '7.8_Диаметр_Все'))
'in project: ' + (Test-ProjectSpecIn $b ($P + '3.1_Имя_Все')) + ' ' + (Test-ProjectSpecIn $b ($P + '7.8_Диаметр_Все'))
'missing: ' + ($null -eq (Read-ProjectConfig '%s'))
'bom: ' + ([System.IO.File]::ReadAllBytes('%s')[0] -eq 0xEF)
'no code line: ' + (-not ([System.IO.File]::ReadAllText('%s').Contains('code=')))
""" % (path, path, path + ".нет", path, path))
        self.assertEqual(out[0], "stripped: 3.1_Имя_Все|7.8_Диаметр_Все")
        out = out[1:]
        self.assertEqual(out[0], "proj: РД C:\\шаблон\\Спеки.rvt \\\\srv\\проект\\RVT")
        self.assertEqual(out[1], "counts: 2 2 2 3 1 1")
        self.assertEqual(out[2], "cell: 1a412")
        self.assertEqual(out[3], "x: 0x-1")
        self.assertEqual(out[4], "group: Моя группа")
        self.assertEqual(out[5], "in project: True False")
        self.assertEqual(out[6], "missing: True")
        self.assertEqual(out[7], "bom: True")
        self.assertEqual(out[8], "no code line: True")
        self.assertFalse(os.path.exists(path + ".tmp"))

    def test_old_file_without_inproject_section_keeps_everything_in_project(self):
        path = os.path.join(self.tmp, "старый.txt")
        with open(path, "w", encoding="utf-8-sig", newline="") as f:
            f.write(project_text())
        out = run(r"""
$c = Read-ProjectConfig '%s'
'legacy: ' + ($null -eq $c.InProject) + ' ' + (Test-ProjectSpecIn $c '1_Классификатор_Все')
""" % path)
        self.assertEqual(out[0], "legacy: True True")

    def test_per_model_sets_parity_from_fixture_file(self):
        """Файл настройки (данные теста по прежним пакетам) даёт ту же выборку по моделям, что и прежние зашитые таблицы раздел -> пакеты."""
        path = os.path.join(self.tmp, "fixture.txt")
        models = ["PROJ_IOS_%s_ORG_K01_R24" % c for c in ("OV", "PD", "VK", "OT", "EOM")] + ["PROJ_IOS_SS_ORG_K02_R24"]
        with open(path, "w", encoding="utf-8-sig", newline="") as f:
            f.write(project_text(models + ["PROJ_IOS_SVOD_ORG_R24"]))
        out = run(r"""
$c = Read-ProjectConfig '%s'
foreach ($m in $c.Models) { $r = Get-ProjectModelSpecs $c $m.Name; if ($null -eq $r.Specs) { $m.Name + ': -- ' + $r.Why } else { $m.Name + ': ' + @($r.Specs).Count } }
'models: ' + $c.Models.Count
$vk = (Get-ProjectModelSpecs $c 'PROJ_IOS_VK_ORG_K01_R24').Specs
'vk has pipes: ' + ($vk -contains '7.2_Диаметр_Трубы_Все') + ' ducts: ' + ($vk -contains '5.1_Ширина_Воздуховоды_Все') + ' trays: ' + ($vk -contains '5.7_Ширина_Лотки_Все')
$u = Get-ProjectModelSpecs $c '\\srv\PROJ_IOS_VK_ORG_K05_R24.rvt'
'new vk model falls back to section union: ' + (@($u.Specs).Count -eq @($vk).Count)
$x = Get-ProjectModelSpecs $c 'Чужая_модель'; 'unknown: ' + ($null -eq $x.Specs) + ' | ' + $x.Why
'display: ' + (Get-ProjectSpecDisplayName $c 'П        5.1_Ширина') + '|' + (Get-ProjectSpecDisplayName $c '5.1_Ширина')
""" % path)
        counts = {l.split(": ")[0]: l.split(": ", 1)[1] for l in out[:6]}
        self.assertEqual(counts["PROJ_IOS_PD_ORG_K01_R24"], "38")
        self.assertEqual(counts["PROJ_IOS_OV_ORG_K01_R24"], "55")
        self.assertEqual(counts["PROJ_IOS_VK_ORG_K01_R24"], "53")
        self.assertEqual(counts["PROJ_IOS_OT_ORG_K01_R24"], "55")
        self.assertEqual(counts["PROJ_IOS_EOM_ORG_K01_R24"], "35")
        self.assertEqual(out[6], "models: 6")                                      # сводная в настройку не попала
        self.assertEqual(out[7], "vk has pipes: True ducts: False trays: False")
        self.assertEqual(out[8], "new vk model falls back to section union: True")
        self.assertIn("раздел не определён по имени", out[9])
        self.assertEqual(out[10], "display: " + P + "5.1_Ширина|" + P + "5.1_Ширина")

    def test_profile_substitution_and_unplaced_model_reasons(self):
        out = run(r"""
$c = New-ProjectConfig 'РД'; $c.InProject = $null
'f: ' + (Expand-ProjectProfile $c '{СТАДИЯ}_{ШифР}_{stage}-{CODE}_x_{ЧТО-ТО}')
Update-ProjectSpecs $c @('1_А_Все')
Update-ProjectModels $c @('PROJ_IOS_OV_ORG_K01_R24', 'PROJ_IOS_VK_ORG_K01_R24')
Set-ProjectCell $c 'PROJ_IOS_OV_ORG_K01_R24' '1_А_Все' 1 'a' 5
$r = Get-ProjectModelSpecs $c 'PROJ_IOS_VK_ORG_K01_R24'; 'unmarked: ' + ($null -eq $r.Specs) + ' | ' + $r.Why
$r = Get-ProjectModelSpecs $c 'PROJ_IOS_PD_ORG_K09_R24'; 'no section models: ' + ($null -eq $r.Specs) + ' | ' + $r.Why
""")
        self.assertEqual(out[0], "f: РД_{ШифР}_РД-{CODE}_x_{ЧТО-ТО}")                # {ШИФР} больше не подставляется
        self.assertIn("спецификации для неё не размечены", out[1])
        self.assertIn("нет размеченных моделей", out[2])


if __name__ == "__main__":
    unittest.main()
