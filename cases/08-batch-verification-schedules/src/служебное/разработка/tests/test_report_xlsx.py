"""Отчёт прогона вставки (Сводка, Вставлено, Параметры, Пропуски) собирается в xlsx без Excel: zip + XML (report_xlsx.ps1)."""
import glob
import os
import re
import shutil
import sys
import tempfile
import unittest
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from test_chain import PWSH  # noqa: E402
import test_two_tabs as tt  # noqa: E402

TOOL = tt.TOOL
UI = tt.UI
XLSX_LIB = os.path.join(TOOL, "report_xlsx.ps1")

FUNCS = ["ConvertTo-PlainArray", "Read-ReportCsv", "Get-ModelFileName", "Convert-ReportNumber", "Build-SummaryRows", "Build-GroupedRows", "Export-ReportXlsx"]

SUMMARY = ["Модель;Раздел;Удалено;Вставлено;Пропущено;Ошибка;Режим",
           "C:/m/PROJ_IOS_OV_ORG_K01_R24.rvt;ОВ;0;3;1;;первичная (автоматически)",
           "C:/m/PROJ_IOS_VK_ORG_K01_R24.rvt;ВК;2;2;0;Модель не сохранена;замена"]
INSERTED = ["Модель;Раздел;Спецификация",
            "C:/m/PROJ_IOS_OV_ORG_K01_R24.rvt;ОВ;РД        1_Классификатор_Все",
            "C:/m/PROJ_IOS_OV_ORG_K01_R24.rvt;ОВ;РД        5.1_Ширина_Воздуховоды_Все",
            "C:/m/PROJ_IOS_VK_ORG_K01_R24.rvt;ВК;РД        7.2_Диаметр_Трубы_Все"]
PARAMS = ["Модель;Спецификация;Параметр;Замечание",
          "C:/m/PROJ_IOS_OV_ORG_K01_R24.rvt;РД        5.1_Ширина_Воздуховоды_Все;ADSK_Размер;нет параметра & <угловые> скобки"]
SKIPS = ["Модель;Спецификация;Причина",
         "C:/m/PROJ_IOS_OV_ORG_K01_R24.rvt;РД        20_Комплект_Все;уже в пакете модели",
         "C:/m/PROJ_IOS_VK_ORG_K01_R24.rvt;РД        20_Комплект_Все;нет в файле со спецификациями"]


def write_csv(path, lines):
    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        f.write("\r\n".join(lines) + "\r\n")


def read(p):
    with open(p, encoding="utf-8-sig") as f:
        return f.read()


@unittest.skipUnless(PWSH, "нет pwsh")
class ReportXlsxTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="bvs_rx_")
        self.addCleanup(shutil.rmtree, self.tmp, True)
        self.files = {}
        for key, lines in (("sum", SUMMARY), ("ins", INSERTED), ("par", PARAMS), ("skp", SKIPS)):
            self.files[key] = os.path.join(self.tmp, key + ".csv")
            write_csv(self.files[key], lines)

    def export(self, xlsx):
        body = (". '%s'\n" % XLSX_LIB.replace("'", "''")) + r"""
$script:ReportXlsxError = ''
$before = @(Get-ChildItem ([IO.Path]::GetTempPath()) -Filter 'bvs_excel_seed_*' -ErrorAction SilentlyContinue).Count
$r = Export-ReportXlsx -Summary '%s' -Params '%s' -Skips '%s' -Inserted '%s' -XlsxPath '%s'
'result: ' + $(if ($r) { $r } else { 'NULL' })
'error: ' + $script:ReportXlsxError
'seed files created: ' + (@(Get-ChildItem ([IO.Path]::GetTempPath()) -Filter 'bvs_excel_seed_*' -ErrorAction SilentlyContinue).Count - $before)
""" % tuple(x.replace("'", "''") for x in (self.files["sum"], self.files["par"], self.files["skp"], self.files["ins"], xlsx))
        return tt.run(self.tmp, body, FUNCS)

    def test_xlsx_is_built_without_excel_and_matches_csv(self):
        import openpyxl
        xlsx = os.path.join(self.tmp, "out", "2026-11-02.xlsx")
        out = self.export(xlsx)
        self.assertEqual(out[0], "result: " + xlsx)
        self.assertEqual(out[1].strip(), "error:")
        self.assertEqual(out[2], "seed files created: 0")                                     # seed-книги нет, Excel не запускался
        self.assertTrue(zipfile.is_zipfile(xlsx))                                              # файл открывается как zip
        with zipfile.ZipFile(xlsx) as z:
            names = set(z.namelist())
            self.assertIn("[Content_Types].xml", names)
            self.assertEqual(len([n for n in names if n.startswith("xl/worksheets/sheet")]), 4)
            self.assertIsNone(z.testzip())
        self.assertFalse(os.path.exists(xlsx + ".tmp"))
        wb = openpyxl.load_workbook(xlsx)                                                      # и как книга Excel (структура валидна)
        self.assertEqual(wb.sheetnames, ["Сводка", "Вставлено", "Параметры", "Пропуски"])
        sh = wb["Сводка"]
        self.assertEqual([c.value for c in sh[1]], ["Модель", "Раздел", "Удалено", "Вставлено", "Пропущено", "Ошибка", "Режим"])
        self.assertEqual([c.value for c in sh[2]], ["PROJ_IOS_OV_ORG_K01_R24.rvt", "ОВ", 0, 3, 1, None, "первичная (автоматически)"])      # данные как в CSV, числа — числами
        self.assertEqual([c.value for c in sh[3]], ["PROJ_IOS_VK_ORG_K01_R24.rvt", "ВК", 2, 2, 0, "Модель не сохранена", "замена"])
        self.assertEqual(sh.freeze_panes, "A2")
        self.assertEqual(sh.auto_filter.ref, "A1:G3")
        ins = wb["Вставлено"]
        rows = [[c.value for c in r] for r in ins.iter_rows()]
        self.assertEqual(rows[0], ["Модель", "Раздел", "Число", "Спецификация"])
        self.assertEqual(rows[1], ["PROJ_IOS_OV_ORG_K01_R24.rvt", "ОВ", 2, None])           # строка модели: число спецификаций
        self.assertEqual(rows[2][3], "РД        1_Классификатор_Все")
        self.assertEqual(rows[3][3], "РД        5.1_Ширина_Воздуховоды_Все")
        self.assertEqual(rows[4], ["PROJ_IOS_VK_ORG_K01_R24.rvt", "ВК", 1, None])
        self.assertEqual(rows[5][3], "РД        7.2_Диаметр_Трубы_Все")
        self.assertFalse(ins.row_dimensions[2].hidden)                                          # строка модели видна
        self.assertTrue(ins.row_dimensions[3].hidden and ins.row_dimensions[3].outlineLevel == 1)   # её спецификации свёрнуты (как «+» в Excel)
        self.assertTrue(ins.row_dimensions[4].hidden and ins.row_dimensions[5].hidden is False)
        par = wb["Параметры"]
        prow = [[c.value for c in r] for r in par.iter_rows()]
        self.assertEqual(prow[0], ["Модель", "Число", "Спецификация", "Параметр", "Замечание"])
        self.assertEqual(prow[1][:2], ["PROJ_IOS_OV_ORG_K01_R24.rvt", 1])
        self.assertEqual(prow[2], [None, None, "РД        5.1_Ширина_Воздуховоды_Все", "ADSK_Размер", "нет параметра & <угловые> скобки"])   # XML-символы экранированы
        skp = wb["Пропуски"]
        srow = [[c.value for c in r] for r in skp.iter_rows()]
        self.assertEqual([r[0] for r in srow if r[0]], ["Модель", "PROJ_IOS_OV_ORG_K01_R24.rvt", "PROJ_IOS_VK_ORG_K01_R24.rvt"])
        self.assertEqual([r[3] for r in srow if r[3]], ["Причина", "уже в пакете модели", "нет в файле со спецификациями"])

    def test_existing_file_is_replaced_and_write_error_gives_reason_not_exception(self):
        xlsx = os.path.join(self.tmp, "r.xlsx")
        with open(xlsx, "w") as f:
            f.write("старый")
        out = self.export(xlsx)
        self.assertEqual(out[0], "result: " + xlsx)
        self.assertTrue(zipfile.is_zipfile(xlsx))
        blocked = os.path.join(self.tmp, "занято")
        open(blocked, "w").write("файл вместо папки")                                          # родитель — файл: записать нельзя
        out = self.export(os.path.join(blocked, "r.xlsx"))
        self.assertEqual(out[0], "result: NULL")
        self.assertNotEqual(out[1].strip(), "error:")                                          # причина записана в статус
        self.assertEqual(out[2], "seed files created: 0")


class NoExcelInReportCodeTests(unittest.TestCase):
    def test_report_code_has_no_excel_com(self):
        ui = read(UI)
        fn = ui[ui.index("function Export-ReportXlsx"):ui.index("function Get-NextReportPath")]
        lib = read(XLSX_LIB)
        for text in (fn, lib):
            for bad in ("New-Object -ComObject", "ComObject", "Start-Process", "New-IsolatedExcelApplication", "Excel.Application", "excel.exe", "EXCEL.EXE",
                        "seed", "Marshal]::ReleaseComObject", "Workbooks"):
                self.assertNotIn(bad, text, bad)
        self.assertIn("Write-ReportXlsx -Path $XlsxPath", fn)
        self.assertIn('. (Join-Path $ToolDir "report_xlsx.ps1")', ui)
        self.assertIn('"report_xlsx.ps1"', ui[:ui.index("SELFCHECK OK")])                       # SelfCheck знает новый файл
        self.assertNotIn("Один раз откройте Excel", fn + lib)                                   # COM-ветки отчёта и её сообщения нет
        # статус при ошибке: причина + ссылка «Открыть CSV»
        tick = ui[ui.index("$xlsx = Export-ReportXlsx"):ui.index("$form.Controls.Add($ui.Timer)")]
        self.assertIn("Set-CsvLink $script:Run.Summary", tick)
        self.assertIn("Причина: ", tick)

    def test_excel_com_remaining_places_are_only_assignment_fallback_and_google(self):
        users = []
        for p in glob.glob(os.path.join(TOOL, "*.ps1")):
            t = read(p)
            if "New-IsolatedExcelApplication" in t and os.path.basename(p) != "excel_isolated.ps1":
                users.append(os.path.basename(p))
        self.assertEqual(sorted(users), ["push_google.ps1"])             # чтение xlsx для Google, если нет payload; отчёт прогона и сборка задания — без COM


if __name__ == "__main__":
    unittest.main()
