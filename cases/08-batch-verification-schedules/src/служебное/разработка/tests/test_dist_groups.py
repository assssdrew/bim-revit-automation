"""Таблица распределения: группы = имя без хвоста, «В проекте» по спецификациям, колонка «Сводная», RSN://."""
import os
import sys
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from test_chain import PWSH  # noqa: E402
from test_project_config import run, P, TOOL  # noqa: E402

# Short stand-in for a template catalog. The client template list is not published.


def synthetic_names():
    return [
        "01_Alpha_Все",
        "01_Alpha_Пустые",
        "01_Alpha_Без значения",
        "99_Beta",
        "99_Beta_Без значения",
    ]


@unittest.skipUnless(PWSH, "нет pwsh")
class GroupSplitTests(unittest.TestCase):
    def test_tails_fall_off_the_group_name(self):
        names = synthetic_names()
        body = "$names = @(" + ",".join("'" + n + "'" for n in names) + ")\n" + r"""
$c = New-ProjectConfig 'РД'
Update-ProjectSpecs $c $names
'specs=' + $c.Specs.Count
foreach ($g in (Get-ProjectGroups $c)) { 'G|' + $g.Name + '|' + @($g.Specs).Count }
"""
        out = run(body)
        self.assertEqual(out[0], "specs=5")
        groups = dict(l.split("|")[1:3] for l in out if l.startswith("G|"))
        self.assertEqual(groups.get("01_Alpha"), "3")
        self.assertEqual(groups.get("99_Beta"), "2")

    def test_latin_c_and_cyrillic_c_share_a_group_key(self):
        out = run(r"""
$lat = '2.5_Class_' + [string][char]0x43 + 'антехприборы_Все'
$cyr = '2.5_Class_' + [string][char]0x421 + 'антехприборы_Все'
$c = New-ProjectConfig 'РД'
Update-ProjectSpecs $c @($lat)
$before = @(Get-ProjectGroups $c | ForEach-Object { $_.Name }) -join ';'
Update-ProjectSpecs $c @($cyr)
$after = @(Get-ProjectGroups $c | ForEach-Object { $_.Name }) -join ';'
'same: ' + ($before -eq $after)
'latin left: ' + @($c.Specs | Where-Object { $_.Contains([string][char]0x43 + 'антехприборы') }).Count
'keys: ' + ((Get-ProjectSpecKey $lat) -eq (Get-ProjectSpecKey $cyr))
""")
        self.assertEqual(out[0], "same: True")
        self.assertEqual(out[1], "latin left: 0")
        self.assertEqual(out[2], "keys: True")



class SpecInProjectAndSummaryTests(unittest.TestCase):
    def test_per_spec_in_project_partial_group_and_summary_column(self):
        out = run(r"""
$c = New-ProjectConfig 'П'
Update-ProjectSpecs $c @(($P + '3.1_Имя_Все'), ($P + '3.1_Имя_Пустые'), ($P + '7.8_Диаметр_Все'))
Update-ProjectModels $c @('PROJ_IOS_OV_ORG_K01_R24')
$g = (Get-ProjectGroups $c)[0]
Invoke-ProjectSpecInClick $c ($P + '3.1_Имя_Все')
'one spec: ' + (Get-ProjectGroupInState $c $g.Specs) + ' ' + (Test-ProjectSpecIn $c ($P + '3.1_Имя_Все')) + ' ' + (Test-ProjectSpecIn $c ($P + '3.1_Имя_Пустые'))
'summary after in: ' + ((Get-ProjectSummarySpecs $c) -join ',')
Invoke-ProjectSummaryClick $c @($P + '3.1_Имя_Все')
'summary off: ' + @(Get-ProjectSummarySpecs $c).Count + ' still in project: ' + (Test-ProjectSpecIn $c ($P + '3.1_Имя_Все'))
Invoke-ProjectSpecInClick $c ($P + '3.1_Имя_Все')
Invoke-ProjectSpecInClick $c ($P + '3.1_Имя_Все')
'summary state after re-add: ' + (Test-ProjectSpecInSummary $c ($P + '3.1_Имя_Все'))
Invoke-ProjectSpecInClick $c ($P + '3.1_Имя_Пустые')
Invoke-ProjectSummaryClick $c @($P + '3.1_Имя_Пустые')
'group summary partial: ' + (Get-ProjectSummaryState $c $g.Specs)
# сохранение и чтение
$f = Join-Path ([IO.Path]::GetTempPath()) 'bvs_sum_test.txt'
Write-ProjectConfig $c $f
$d = Read-ProjectConfig $f
'roundtrip summary: ' + ((Get-ProjectSummarySpecs $d) -join ',')
'file has section: ' + ((Get-Content $f -Raw) -match '\[summary\]')
Remove-Item $f
""")
        self.assertEqual(out[0], "one spec: Indeterminate True False")
        self.assertEqual(out[1], "summary after in: 3.1_Имя_Все")
        self.assertEqual(out[2], "summary off: 0 still in project: True")
        self.assertEqual(out[3], "summary state after re-add: True")
        self.assertEqual(out[4], "group summary partial: Indeterminate")
        self.assertEqual(out[5], "roundtrip summary: 3.1_Имя_Все")
        self.assertEqual(out[6], "file has section: True")

    def test_old_file_without_summary_section_sends_all_project_specs_to_summary(self):
        out = run(r"""
$c = New-ProjectConfig 'П'
Update-ProjectSpecs $c @(($P + '3.1_Имя_Все'), ($P + '7.8_Диаметр_Все'))
Set-ProjectInProjectBulk $c $c.Specs $true
$c.Summary = $null
'all: ' + @(Get-ProjectSummarySpecs $c).Count
""")
        self.assertEqual(out, ["all: 2"])

    def test_hide_unchecked_hides_specs_outside_project_and_rsn_models_get_section(self):
        out = run(r"""
$c = New-ProjectConfig 'П'
Update-ProjectSpecs $c @(($P + '3.1_Имя_Все'), ($P + '3.1_Имя_Пустые'))
Update-ProjectModels $c @('RSN://srv/proj/PROJ_IOS_VK_ORG_K01_R24.rvt', 'RSN://srv/proj/PROJ_IOS_SVOD_ORG_R24.rvt')
'models: ' + (($c.Models | ForEach-Object { $_.Name + '=' + $_.Section }) -join ',')
'rsn source: ' + (Test-RsnModelsSource 'RSN://srv/proj')
Invoke-ProjectSpecInClick $c ($P + '3.1_Имя_Все')
$e = @{ Sections = @{}; Groups = @{}; Hide = $true }
$g = (Get-ProjectGroups $c)[0]; $e.Groups[$g.Name] = $true
$m = Get-ProjectGridModel $c $e
'spec rows: ' + @($m.Rows | Where-Object { $_.Kind -eq 'spec' }).Count
""")
        self.assertTrue(out[0].startswith("models: PROJ_IOS_VK_ORG_K01_R24=VK"), out[0])
        self.assertEqual(out[1], "rsn source: True")
        self.assertEqual(out[2], "spec rows: 1")


if __name__ == "__main__":
    unittest.main()
