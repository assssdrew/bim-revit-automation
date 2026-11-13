"""Выборочная выгрузка: сборка по буферу «последнее состояние по паре модель + спецификация» (build_assignment.ps1 -Partial).

Сценарии: полная -> выборочная по одной модели -> полная; незагруженная связь; удалённая спецификация и модель;
старые история и журнал без буфера (засев); защита от несовпадения режима выгрузки и сборки.
"""
import json
import os
import shutil
import subprocess
import sys
import tempfile
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(os.path.dirname(os.path.dirname(HERE)))
TOOL = os.path.join(ROOT, "служебное", "скрипты")      # папка с .ps1/.py окна
sys.path.insert(0, HERE)
from baseline_lib import LEGACY_HISTORY, LEGACY_LEDGER, dump_xlsx, read_state, run_pwsh  # noqa: E402
from project_fixture import project_text  # noqa: E402

PWSH = os.environ.get("PWSH") or shutil.which("pwsh")

E1 = "ORG_K1_PROJ_IOS_EOM_R23"
E2 = "ORG_K2_PROJ_IOS_EOM_R23"
SSM = "ORG_K1_PROJ_IOS_SS_R23"
S1 = "ЭОМ Проверка заполнения"
S2 = "СС Проверка марки"
S3 = "ОВ Проверка прочее"

HEAD = ['Проверить, чтобы была заполнена "Длина"', "Наименование", "ADSK_Размер_Длина", "Количество", "Связанный файл"]
FIELDS = "|".join(["Семейство и типоразмер", "Наименование", "ADSK_Размер_Длина", "Количество", "Связанный файл"])

ROWS = {
    S1: [("Светильник: тип А", "Светильник 1", "", E1), ("Светильник: тип Б", "Светильник 2", "007", E1),
         ("Розетка", "Розетка 1", "пробел", E2)],
    S2: [("Извещатель", "ИП-1", "", SSM)],
}


def write_export(folder, rows, partial_pairs=None, catalog_specs=None, catalog_models=None):
    """rows: {спецификация: [(семейство, наименование, значение, модель)]}; пустой список -> «*empty*».
    partial_pairs: None — полная выгрузка, иначе список (модель, спецификация) проверенных пар."""
    os.makedirs(folder, exist_ok=True)
    man = []
    n = 0
    for spec, items in rows.items():
        if not items:
            man.append("*empty*\t%s" % spec)
            continue
        n += 1
        fn = "%04d.tsv" % n
        with open(os.path.join(folder, fn), "w", encoding="utf-8-sig", newline="") as f:
            f.write("\t".join(HEAD) + "\n")
            for fam, item, val, model in items:
                f.write("\t".join([fam, item, val, "1", model + ".rvt"]) + "\n")
        man.append("%s\t%s\t%s\tADSK_Размер_Длина" % (fn, spec, FIELDS))
    if partial_pairs is not None:
        man.append("*scope*\tpartial")
        for m, sp in partial_pairs:
            man.append("*pair*\t%s\t%s" % (m, sp))
    with open(os.path.join(folder, "manifest.txt"), "w", encoding="utf-8-sig", newline="") as f:
        f.write("\n".join(man) + "\n")
    if catalog_specs is not None or catalog_models is not None:
        cat = ["spec\t" + s for s in (catalog_specs or [])] + ["model\t%s\t%s" % (m, st) for m, st in (catalog_models or [])]
        with open(os.path.join(folder, "catalog.txt"), "w", encoding="utf-8-sig", newline="") as f:
            f.write("\n".join(cat) + "\n")


def read_ledger(path):
    out = {}
    with open(path, encoding="utf-8-sig") as f:
        for line in f.read().splitlines():
            p = line.split("\t")
            out[(p[0], p[2], p[3], p[4])] = {"first": p[7], "last": p[8], "status": p[9], "fixed": p[10], "reopened": p[11]}
    return out


def read_history(path):
    with open(path, encoding="utf-8-sig") as f:
        return [l.split("\t") for l in f.read().splitlines() if l]


ALL_CAT_SPECS = [S1, S2]
ALL_CAT_MODELS = [(E1, "loaded"), (E2, "loaded"), (SSM, "loaded")]


@unittest.skipUnless(PWSH, "нет pwsh")
class PartialTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="bvs_part_")
        self.addCleanup(shutil.rmtree, self.tmp, True)
        self.out = os.path.join(self.tmp, "out")
        os.makedirs(self.out)
        self.n = 0
        self.led = os.path.join(self.out, "задание_журнал.txt")
        self.hist = os.path.join(self.out, "задание_история.txt")

    def build(self, now, rows, partial_pairs=None, flag=None, specs=None, models=None, expect_ok=True):
        self.n += 1
        exp = os.path.join(self.tmp, "exp%d" % self.n)
        write_export(exp, rows, partial_pairs, ALL_CAT_SPECS if specs is None else specs, ALL_CAT_MODELS if models is None else models)
        args = ["-File", os.path.join(TOOL, "build_assignment.ps1"), "-ExportDir", exp, "-OutXlsx", os.path.join(self.out, "task.xlsx")]
        if flag is None:
            flag = partial_pairs is not None
        if flag:
            args.append("-Partial")
        r = run_pwsh(PWSH, args, {"BVS_FAKE_NOW": now})
        if expect_ok:
            self.assertEqual(r.returncode, 0, r.stdout + r.stderr)
        return r

    def payload(self):
        with open(os.path.join(self.out, "google_push_payload.json"), encoding="utf-8") as f:
            return json.load(f)

    def sheet_texts(self, name):
        x = dump_xlsx(os.path.join(self.out, "task.xlsx"))
        for sh in x["sheets"]:
            if sh["name"] == name:
                return {str(c["v"]) for c in sh["cells"] if c["v"] is not None}
        return set()

    # ---- основной сценарий: полная -> выборочная по одной модели -> полная ----
    def test_full_partial_full(self):
        self.build("2026-02-01 10:00:00", ROWS)
        led1 = read_ledger(self.led)
        self.assertEqual(len([v for v in led1.values() if v["status"] == "open"]), 4)
        # состояние — один файл <история>_состояние.txt (до фазы 7 — три файла _буфер, _проверки, _каталог)
        self.assertTrue(os.path.exists(os.path.join(self.out, "задание_история_состояние.txt")))
        for f in ("задание_история_буфер.txt", "задание_история_проверки.txt", "задание_история_каталог.txt"):
            self.assertFalse(os.path.exists(os.path.join(self.out, f)), f)
        st = read_state(self.out)
        self.assertEqual((len(st["buffer"]), len(st["checks"]), len(st["catalog"])), (4, 2, 5))

        # выборочно: только E2 × S1, замечание «Розетка» исправлено
        rows = {S1: []}
        r = self.build("2026-02-02 10:00:00", rows, partial_pairs=[(E2, S1)])
        self.assertIn("FIXED=1", r.stdout)
        self.assertIn("NEW=0", r.stdout)
        self.assertIn("PARTIAL=1 PAIRS=1", r.stdout)
        led2 = read_ledger(self.led)
        rose = led2[(E2, S1, "Розетка", "Розетка 1")]
        self.assertEqual((rose["status"], rose["fixed"]), ("fixed", "2026-02-02 10:00:00"))
        for k, v in led2.items():
            if k[0] != E2:
                # непроверенные пары не тронуты: открыты, LastSeen прежний
                self.assertEqual((v["status"], v["last"]), ("open", "2026-02-01 10:00:00"), k)
        # видимый отчёт собран из буфера целиком: замечания непроверенных моделей на месте
        eom = self.sheet_texts("ЭОМ")
        self.assertIn("Светильник 1", eom)
        self.assertIn("Светильник 2", eom)
        self.assertIn("Розетка 1", eom)                # исправленное — в списке со статусом «исправлено»
        self.assertIn(S2, self.sheet_texts("Сводка"))
        self.assertIn("Извещатель", self.sheet_texts("СС"))
        # история: срез полный (E1/SS перенесены без изменений), у проверенной пары — ноль
        h = read_history(self.hist)
        s2 = [x for x in h if x[7] == "2026-02-02 10:00:00"]
        by = {(x[2], x[3]): int(x[5]) for x in s2}
        self.assertEqual(by[(E1, S1)], 2)
        self.assertEqual(by[(SSM, S2)], 1)
        self.assertEqual(by[(E2, S1)], 0)

        # полная: Розетка не вернулась, остальные как были; «Новое» нет
        r = self.build("2026-02-03 10:00:00", {S1: ROWS[S1][:2], S2: ROWS[S2]})
        self.assertIn("FIXED=0", r.stdout)
        self.assertIn("NEW=0", r.stdout)
        led3 = read_ledger(self.led)
        self.assertEqual(led3[(E2, S1, "Розетка", "Розетка 1")]["status"], "fixed")
        self.assertEqual(led3[(E1, S1, "Светильник: тип А", "Светильник 1")]["last"], "2026-02-03 10:00:00")

    def test_partial_new_and_reopened_only_in_checked_pairs(self):
        self.build("2026-02-01 10:00:00", ROWS)
        self.build("2026-02-02 10:00:00", {S1: []}, partial_pairs=[(E2, S1)])
        # E2 × S1: розетка вернулась + новая; E1 не проверялась
        rows = {S1: [("Розетка", "Розетка 1", "", E2), ("Розетка", "Розетка 9", "", E2)]}
        r = self.build("2026-02-03 10:00:00", rows, partial_pairs=[(E2, S1)])
        self.assertIn("NEW=1", r.stdout)
        led = read_ledger(self.led)
        self.assertEqual(led[(E2, S1, "Розетка", "Розетка 1")]["reopened"], "1")
        self.assertEqual(led[(E2, S1, "Розетка", "Розетка 9")]["first"], "2026-02-03 10:00:00")
        self.assertEqual(led[(E1, S1, "Светильник: тип А", "Светильник 1")]["last"], "2026-02-01 10:00:00")

    # ---- незагруженная связь ----
    def test_unloaded_link_does_not_close_issues(self):
        self.build("2026-02-01 10:00:00", ROWS)
        # E1 выбрана, но связь не загружена: пар нет (выгрузчик их не пишет), строк нет
        r = self.build("2026-02-02 10:00:00", {S1: []}, partial_pairs=[],
                       models=[(E1, "unloaded"), (E2, "loaded"), (SSM, "loaded")])
        self.assertIn("FIXED=0", r.stdout)
        led = read_ledger(self.led)
        self.assertTrue(all(v["status"] == "open" for v in led.values()))
        self.assertIn("Светильник 1", self.sheet_texts("ЭОМ"))

    # ---- удалённая спецификация и модель ----
    def test_removed_spec_and_model_leave_report_without_false_fixed(self):
        self.build("2026-02-01 10:00:00", ROWS)
        r = self.build("2026-02-02 10:00:00", {S1: ROWS[S1][:2]}, partial_pairs=[(E1, S1)],
                       specs=[S1], models=[(E1, "loaded"), (E2, "loaded")])     # S2 и модель SS исчезли
        self.assertIn("FIXED=0", r.stdout)
        sm = self.sheet_texts("Сводка")
        self.assertNotIn(S2, sm)
        self.assertNotIn(SSM, sm)
        buf = "\n".join(read_state(self.out)["buffer"])
        self.assertNotIn("Извещатель", buf)
        self.assertIn("Розетка 1", buf)
        led = read_ledger(self.led)
        self.assertEqual(led[(SSM, S2, "Извещатель", "ИП-1")]["status"], "open")     # как в «Всё»: журнал не закрывает

    # ---- старые файлы без буфера ----
    def test_seed_from_legacy_history_and_ledger(self):
        with open(self.hist, "w", encoding="utf-8-sig", newline="") as f:
            f.write(LEGACY_HISTORY)
        with open(self.led, "w", encoding="utf-8-sig", newline="") as f:
            f.write(LEGACY_LEDGER)
        # проверяем E1 × S1: те же два замечания, что в старом журнале (одно «Светильник 1», второе «Светильник 2»)
        r = self.build("2026-02-01 10:00:00", {S1: ROWS[S1][:2]}, partial_pairs=[(E1, S1)])
        self.assertIn("NEW=0", r.stdout)
        self.assertIn("FIXED=0", r.stdout)
        led = read_ledger(self.led)
        self.assertEqual(len(led), 4)
        self.assertTrue(all(v["status"] == "open" for v in led.values()))
        self.assertIn("Извещатель", self.sheet_texts("СС"))           # пары, которых не проверяли, показаны из журнала

    # ---- свежесть и Google ----
    def summary_rows(self):
        x = dump_xlsx(os.path.join(self.out, "task.xlsx"))
        sh = [s for s in x["sheets"] if s["name"] == "Сводка"][0]
        grid = {}
        for c in sh["cells"]:
            grid[c["ref"]] = c
        return grid, sh

    def test_summary_freshness_only_in_partial(self):
        self.build("2026-02-01 10:00:00", ROWS)
        grid, sh = self.summary_rows()
        self.assertFalse(any(str(c["v"]) == "Проверено" for c in grid.values()))
        self.assertNotIn("dim", json.dumps(self.payload(), ensure_ascii=False))
        self.assertNotIn("partial", self.payload())          # «Всё»: payload прежний

        # через 10 дней проверяем только E2 × S1 (E1 и SS — «давно не проверялись»)
        self.build("2026-02-11 10:00:00", {S1: ROWS[S1][2:], S2: []}, partial_pairs=[(E2, S1), (E2, S2)])
        grid, sh = self.summary_rows()
        last = max(c["ref"] for c in grid.values() if c["ref"].endswith("2") and str(c["v"]) == "Проверено")
        col = "".join(ch for ch in last if ch.isalpha())
        by_model = {str(grid["A%d" % r]["v"]): grid["%s%d" % (col, r)] for r in range(3, sh["dims"][0] + 1) if "A%d" % r in grid}
        self.assertEqual(by_model[E2]["v"], "11.02.2026")
        self.assertEqual(by_model[E1]["v"], "01.02.2026")
        self.assertEqual(by_model[SSM]["v"], "01.02.2026")
        self.assertNotEqual(by_model[E1]["fill"], by_model[E2]["fill"])      # старше 7 дней подсвечено
        self.assertEqual(by_model[E1]["fill"], by_model[SSM]["fill"])

    def test_partial_payload_has_summary_and_touched_sections_and_dim(self):
        self.build("2026-02-01 10:00:00", ROWS)
        # проверяем E2 × S1 и СС не трогаем: в Google — Сводка и ЭОМ; СС не отправляется и не очищается
        self.build("2026-02-02 10:00:00", {S1: ROWS[S1][2:]}, partial_pairs=[(E2, S1)])
        pl = self.payload()
        self.assertTrue(pl["partial"])
        self.assertEqual([s["name"] for s in pl["sheets"] if s["name"] != "Динамика"], ["Сводка", "ЭОМ"])   # «Динамика» не отправляется (push_google берёт только Сводку и разделы)
        self.assertEqual(pl["clear"], [])
        summ = pl["sheets"][0]
        dims = [f for f in summ["formats"] if f.get("dim")]
        self.assertTrue(dims, "непроверенные пары должны быть серыми")
        # проверенная пара (E2 × S1) серой не бывает: ровно её ячейка без dim
        rows = summ["rows"]
        r_e2 = [i for i, r in enumerate(rows) if r[0] == E2][0] + 1
        c_s1 = rows[1].index(S1) + 1
        self.assertFalse(any(f["r"] == r_e2 and f["c"] == c_s1 for f in dims))
        # заметка листа раздела: что не обновлялось
        eom = pl["sheets"][1]
        self.assertIn("Выборочная выгрузка", eom["rows"][0][0])
        self.assertIn(E1, eom["rows"][0][0])

    def test_partial_clears_touched_section_without_open_issues(self):
        self.build("2026-02-01 10:00:00", ROWS)
        # E1 и E2 проверены и чисты: ЭОМ без открытых замечаний -> очистка этого листа; СС не затрагивается
        r = self.build("2026-02-02 10:00:00", {S1: []}, partial_pairs=[(E1, S1), (E2, S1)])
        pl = self.payload()
        self.assertEqual([s["name"] for s in pl["sheets"] if s["name"] != "Динамика"], ["Сводка"])
        self.assertEqual(pl["clear"], ["ЭОМ"])

    def test_partial_without_any_base_fails_with_clear_message(self):
        r = self.build("2026-02-01 10:00:00", {S1: ROWS[S1]}, partial_pairs=[(E1, S1)], expect_ok=False)
        self.assertNotEqual(r.returncode, 0)
        self.assertIn("Сначала сделайте полную выгрузку", r.stdout + r.stderr)

    # ---- режим выгрузки и режим сборки должны совпадать ----
    def test_mode_mismatch_is_rejected(self):
        self.build("2026-02-01 10:00:00", ROWS)
        r = self.build("2026-02-02 10:00:00", {S1: []}, partial_pairs=[(E2, S1)], flag=False, expect_ok=False)
        self.assertNotEqual(r.returncode, 0)
        self.assertIn("режиме «Всё»", " ".join((r.stdout + r.stderr).replace("|", " ").split()))  # pwsh переносит длинное сообщение по ширине консоли
        r = self.build("2026-02-02 10:00:00", ROWS, flag=True, expect_ok=False)
        self.assertNotEqual(r.returncode, 0)
        self.assertIn("выгрузка полная", r.stdout + r.stderr)


@unittest.skipUnless(PWSH, "нет pwsh")
class PartialUiLogicTests(unittest.TestCase):
    """Логика окна без самого окна: списки моделей/спецификаций, строки run.cfg, разбор ui_schedules.ps1."""

    def run_ps(self, tmp, body):
        script = os.path.join(tmp, "t.ps1")
        with open(script, "w", encoding="utf-8-sig") as f:
            f.write(". '%s'\n. '%s'\n. '%s'\n. '%s'\n" % tuple(os.path.join(TOOL, n) for n in (
                "project_config.ps1", "assignment_export_stream.ps1", "assignment_partial.ps1", "assignment_partial_ui.ps1")) + body)
        r = run_pwsh(PWSH, ["-File", script])
        self.assertEqual(r.returncode, 0, r.stdout + r.stderr)
        return r.stdout

    def test_choices_sorted_grouped_and_filtered_by_catalog(self):
        tmp = tempfile.mkdtemp(prefix="bvs_pui_")
        self.addCleanup(shutil.rmtree, tmp, True)
        hist = os.path.join(tmp, "история_X.txt")
        with open(os.path.join(tmp, "история_X_каталог.txt"), "w", encoding="utf-8-sig") as f:
            f.write("spec\t10 Б Пустые\nspec\t2 А Пустые\nspec\t4.4 В Пустые\n"
                    "model\t%s\tloaded\nmodel\t%s\tunloaded\nmodel\t%s.rvt\tloaded\n" % (E2, E1, SSM))
        with open(os.path.join(tmp, "история_X_буфер.txt"), "w", encoding="utf-8-sig") as f:
            # в буфере есть модель и спецификация, которых уже нет в сводной — предлагать их нельзя
            f.write("EOM\tORG_K9_PROJ_IOS_EOM_R23\t2 А Пустые\tp\tp\tф\tи\tз\n"
                    "EOM\t%s\t99 Удалённая\tp\tp\tф\tи\tз\n" % E1)
        out = self.run_ps(tmp, "$c = Get-PartialChoices '%s' '' ; "
                          "($c.Models | ForEach-Object { $_.SectionTitle + '|' + $_.Name + '|' + $_.Loaded }) -join ';'; '#'; $c.Specs -join ';'" % hist)
        models, specs = out.strip().split("#")
        self.assertEqual(models.strip().rstrip(";").split(";"), [
            "ЭОМ|%s|False" % E1, "ЭОМ|%s|True" % E2, "СС|%s|True" % SSM])
        self.assertEqual(specs.strip().lstrip(";").split(";"), ["2 А Пустые", "4.4 В Пустые", "10 Б Пустые"])

    def test_type_size_error_is_offered_for_every_section(self):
        """Номер 0 в справочнике: «0_Ошибка типоразмера» есть у модели каждого раздела и первая в списке «Выборочно»."""
        tmp = tempfile.mkdtemp(prefix="bvs_pui_")
        self.addCleanup(shutil.rmtree, tmp, True)
        hist = os.path.join(tmp, "история_X.txt")
        spec0 = "П        0_Ошибка типоразмера"
        spec1 = "П        1_Классификатор_Пустые"
        spec999 = "П        1_Классификатор_999"
        spec16 = "П        16.0_Корпус_Кириллица"
        spec17 = "П        17.0_Секция_Кириллица"
        with open(os.path.join(tmp, "история_X_каталог.txt"), "w", encoding="utf-8-sig") as f:
            f.write("spec\t%s\nspec\t%s\nspec\t%s\nspec\t%s\nspec\t%s\nmodel\tX_IOS_OV_T_1\tloaded\n" % (spec1, spec999, spec17, spec0, spec16))
        out = self.run_ps(tmp, r"""
$spec0 = 'П        0_Ошибка типоразмера'
$spec1 = 'П        1_Классификатор_Пустые'
$spec999 = 'П        1_Классификатор_999'
$spec16 = 'П        16.0_Корпус_Кириллица'
$spec17 = 'П        17.0_Секция_Кириллица'
$c = New-ProjectConfig 'П'
Update-ProjectSpecs $c @($spec0, $spec1, $spec999, $spec16, $spec17, 'П        9_Наименование_Все')
$c.InProject = $null
$models = @($script:PcSections | ForEach-Object { 'X_IOS_' + $_.Code + '_T_1' })
Update-ProjectModels $c $models
$f = Join-Path ([IO.Path]::GetTempPath()) ('bvs_cat_' + [guid]::NewGuid().ToString('N') + '.txt')
[IO.File]::WriteAllText($f, (Get-DefaultSectionCatalogText), (New-Object System.Text.UTF8Encoding $true))
$cat = Read-SectionCatalog $f
Remove-Item $f
if (-not $cat.Ok) { throw $cat.Error }
[void](Set-ProjectDistribution $c $cat)
$info = Get-PartialModelSpecInfo $models $c
$all = @($spec0, $spec1, $spec999, $spec16, $spec17, 'П        9_Наименование_Все')
foreach ($s in $script:PcSections) {
    $m = 'X_IOS_' + $s.Code + '_T_1'
    $list = @(Get-PartialModelSpecList $m $all $info.Sets | ForEach-Object { $_ -replace '^П\s+', '' })
    $s.Code + ': ' + ($list -join '|')
}
$ch = Get-PartialChoices '%s' ''
'choices: ' + ($ch.Specs -join '|')
""" % hist)
        lines = [l for l in out.splitlines() if l.strip()]
        codes = ["OV", "PD", "VK", "PT", "OT", "HS", "ITP", "EOM", "SS", "AVT", "HC"]
        want = "0_Ошибка типоразмера|1_Классификатор_Пустые|1_Классификатор_999|16.0_Корпус_Кириллица|17.0_Секция_Кириллица"
        self.assertEqual(lines[:len(codes)], [c + ": " + want for c in codes])
        self.assertEqual(lines[-1], "choices: П        " + want.replace("|", "|П        "))

    def test_base_ready_and_cfg_lines(self):
        tmp = tempfile.mkdtemp(prefix="bvs_pui_")
        self.addCleanup(shutil.rmtree, tmp, True)
        hist = os.path.join(tmp, "история_X.txt")
        out = self.run_ps(tmp, "Test-PartialBaseReady '%s' '%s'; ConvertTo-PartialCfgLines @('С1','С2') @('М1')" % (hist, os.path.join(tmp, "j.txt")))
        lines = out.strip().splitlines()
        self.assertEqual(lines, ["False", "assign_scope=partial", "assign_specs=С1|С2", "assign_models=М1"])
        open(hist, "w").close()
        out = self.run_ps(tmp, "Test-PartialBaseReady '%s' ''" % hist)
        self.assertEqual(out.strip(), "True")

    SETS = os.path.join(HERE, "fixtures", "schedule_sets.txt")

    def logic(self, body):
        tmp = tempfile.mkdtemp(prefix="bvs_link_")
        self.addCleanup(shutil.rmtree, tmp, True)
        cfg = os.path.join(tmp, "проект.txt")
        with open(cfg, "w", encoding="utf-8-sig", newline="") as f:
            f.write(project_text(["PROJ_IOS_%s_ORG_K0%d_R24" % x for x in (("OV", 1), ("PD", 1), ("VK", 1), ("OT", 2), ("EOM", 1))] + ["PROJ_IOS_SVOD_ORG_R24"]))
        pre = r"""
$cfgFile = '%s'
$P = 'П        '
$ov = 'PROJ_IOS_OV_ORG_K01_R24'; $pd = 'PROJ_IOS_PD_ORG_K01_R24'; $vk = 'PROJ_IOS_VK_ORG_K01_R24'; $ot = 'PROJ_IOS_OT_ORG_K02_R24'
$eom = 'PROJ_IOS_EOM_ORG_K01_R24'; $svod = 'PROJ_IOS_SVOD_ORG_R24'
$models = @($ov, $pd, $vk, $ot, $eom)
$duct = $P + '5.1_Ширина_Воздуховоды_Все'; $tray = $P + '5.7_Ширина_Лотки_Все'; $pipe = $P + '7.2_Диаметр_Трубы_Все'; $box = $P + '7.8_Диаметр_Фитингов коробов_Все'
$ltray = $P + '14.7_Длина_Лотки_Все'; $common = $P + '1_Классификатор_Все'; $geo = $P + '16.1_Корпус_Оборудование_Все'; $heat = $P + '15.1_Мощность тепловая_Все'; $other = 'Проверка_вне_наборов'
$all = @($common, $duct, $tray, $ltray, $box, $pipe, $geo, $heat, $other)
$proj = Read-ProjectConfig $cfgFile                        # файл настройки проекта (данные теста: прежние пакеты как отметки)
$info = Get-PartialModelSpecInfo ($models + @($svod)) $proj
$ms = $info.Sets
function Join-S($x) { return ((@($x) | Sort-Object) -join '|') }
function Has($m, $sp) { return (Test-PartialModelHasSpec $ms $m $sp) }
""" % cfg.replace("'", "''")
        out = self.run_ps(tmp, pre + body)
        return [l for l in out.splitlines() if l.strip() != ""]

    D, T, PI, B, LT, C, G, H, O = ("П        5.1_Ширина_Воздуховоды_Все", "П        5.7_Ширина_Лотки_Все", "П        7.2_Диаметр_Трубы_Все",
                                   "П        7.8_Диаметр_Фитингов коробов_Все", "П        14.7_Длина_Лотки_Все", "П        1_Классификатор_Все",
                                   "П        16.1_Корпус_Оборудование_Все", "П        15.1_Мощность тепловая_Все", "Проверка_вне_наборов")

    def test_sets_come_from_project_file_by_section_code_in_model_name(self):
        out = self.logic(r"""
foreach ($m in @($ov, $pd, $vk, $ot, $eom)) { $m + ': ' + ((@($duct, $tray, $ltray, $box, $pipe, $heat, $common, $geo) | ForEach-Object { if (Has $m $_) { '1' } else { '0' } }) -join '') }
'svod: ' + ($null -eq $ms[$svod]); 'svod why: ' + $info.Why[$svod]
""")
        # порядок: duct, tray, ltray, box, pipe, heat, common, geo
        self.assertEqual(out[0], "PROJ_IOS_OV_ORG_K01_R24: 10000011")
        self.assertEqual(out[1], "PROJ_IOS_PD_ORG_K01_R24: 10000010")      # ПД: воздуховоды и общие, без оборудования/«гео»
        self.assertEqual(out[2], "PROJ_IOS_VK_ORG_K01_R24: 00001011")
        self.assertEqual(out[3], "PROJ_IOS_OT_ORG_K02_R24: 00001111")      # ОТ: трубы + тепловая мощность
        self.assertEqual(out[4], "PROJ_IOS_EOM_ORG_K01_R24: 01110011")     # ЭОМ: лотки и короба
        self.assertEqual(out[5], "svod: True")
        self.assertIn("раздел не определён по имени", out[6])             # сводная — не раздел и в настройку не попадает, набора у неё нет

    def test_no_project_unknown_models_and_union_without_trays(self):
        """Причина «57 из 57» — набор не находился ни у одной модели (искались зашитые блоки в неверном файле). Теперь наборы — из настройки
        проекта, а если набор определить нельзя, причина видна."""
        out = self.logic(r"""
$none = Get-PartialModelSpecInfo @($vk) $null
'no project: ' + ($null -eq $none.Sets[$vk]) + ' | ' + $none.Why[$vk]
$unk = Get-PartialModelSpecInfo @('Чужая_модель', 'PROJ_IOS_HS_ORG_K01_R24') $proj
'unknown: ' + $unk.Why['Чужая_модель']
'hs (нет в распределении, у раздела нет размеченных моделей): ' + $unk.Why['PROJ_IOS_HS_ORG_K01_R24']
# модель из распределения, но без разметки
$p2 = New-ProjectConfig 'X' 'П'; Update-ProjectSpecs $p2 @('1_А_Все'); Update-ProjectModels $p2 @('X_IOS_VK_T_1')
$unm = Get-PartialModelSpecInfo @('X_IOS_VK_T_1') $p2
'unmarked: ' + $unm.Why['X_IOS_VK_T_1']
# объединение по моделям без ЭОМ: ни коробов, ни лотков
$u = @(Get-PartialNodeSpecs @($vk, $ot, $pd, $ov) $all $ms)
'union: ' + (Join-S $u)
# имена из истории с префиксом стадии сопоставляются с именами настройки без префикса
'prefixed history name matches: ' + (Has $vk ($P + '7.2_Диаметр_Трубы_Все')) + ' ' + (Has $vk '7.2_диаметр_трубы_ВСЕ')
""")
        self.assertTrue(out[0].startswith("no project: True | нет настройки проекта"), out[0])
        self.assertIn("раздел не определён по имени", out[1])
        self.assertIn("нет размеченных моделей", out[2])
        self.assertIn("спецификации для неё не размечены", out[3])
        union = out[4].split(": ", 1)[1].split("|")
        self.assertEqual(sorted(union), sorted([self.C, self.D, self.PI, self.G, self.H]))
        self.assertEqual(out[5], "prefixed history name matches: True True")

    def test_node_specs_states_and_per_model_picks(self):
        out = self.logic(r"""
$picks = New-PartialPicks $models $all $ms
'default vk: ' + (Join-S ($picks[$vk].Keys | Where-Object { $picks[$vk][$_] }))
'default ov has tray: ' + $picks[$ov].ContainsKey($tray)
$sec = @($vk, $ot)                                                          # раздел из двух моделей
'sec specs: ' + (Join-S (Get-PartialNodeSpecs $sec $all $ms))
'pipe state: ' + (Get-PartialSpecState $picks $sec $ms $pipe)
# снимаем трубы только у ВК (выделена модель)
Set-PartialSpecPick $picks @($vk) $ms $pipe $false
'vk pipe: ' + (Get-PartialSpecState $picks @($vk) $ms $pipe) + '; ot pipe: ' + (Get-PartialSpecState $picks @($ot) $ms $pipe) + '; section: ' + (Get-PartialSpecState $picks $sec $ms $pipe)
# спека есть только у ОТ (тепловая мощность): состояние раздела = состояние ОТ
'heat sec: ' + (Get-PartialSpecState $picks $sec $ms $heat)
# выделен раздел: галочка ставит у всех моделей, где есть
Set-PartialSpecPick $picks $sec $ms $pipe $true
'after set: ' + (Get-PartialSpecState $picks $sec $ms $pipe) + ' ' + $picks[$vk][$pipe] + ' ' + $picks[$ot][$pipe]
Set-PartialSpecPick $picks $sec $ms $heat $false
'heat off: ' + $picks[$ot][$heat] + '; vk has heat key: ' + $picks[$vk].ContainsKey($heat)
# спека, которой нет у модели, ей не добавляется
Set-PartialSpecPick $picks @($ov) $ms $tray $true
'ov tray key: ' + $picks[$ov].ContainsKey($tray)
# узел без моделей
'empty: ' + @(Get-PartialNodeSpecs @() $all $ms).Count
""")
        self.assertEqual(out[0].split(": ", 1)[1].split("|") and sorted(out[0].split(": ", 1)[1].split("|")), sorted([self.C, self.PI, self.G]))
        self.assertEqual(out[1], "default ov has tray: False")
        self.assertEqual(sorted(out[2].split(": ", 1)[1].split("|")), sorted([self.C, self.PI, self.G, self.H]))
        self.assertEqual(out[3], "pipe state: Checked")
        self.assertEqual(out[4], "vk pipe: Unchecked; ot pipe: Checked; section: Indeterminate")        # частичное состояние раздела
        self.assertEqual(out[5], "heat sec: Checked")
        self.assertEqual(out[6], "after set: Checked True True")
        self.assertEqual(out[7], "heat off: False; vk has heat key: False")
        self.assertEqual(out[8], "ov tray key: False")
        self.assertEqual(out[9], "empty: 0")

    def test_model_check_marks_all_its_specs_uncheck_keeps_picks(self):
        out = self.logic(r"""
$picks = New-PartialPicks $models $all $ms
Set-PartialNodeAll $picks @($vk) $all $ms $false                           # у ВК всё снято
Set-PartialSpecPick $picks @($ot) $ms $pipe $false
'vk none: ' + @(Get-PartialPairs $picks @($vk) $all $ms).Count
# модель отметили слева: все её спеки отмечены (ВК была не отмечена)
Update-PartialPicksOnModelCheck $picks $all $ms @($ot) @($ot, $vk)
'vk on: ' + @(Get-PartialPairs $picks @($vk) $all $ms).Count
# ОТ не менялась (была отмечена раньше): ручное снятие трубы сохранилось
'ot pipe kept: ' + $picks[$ot][$pipe]
# снятие галочки модели отметки спек не трогает
$before = @(Get-PartialPairs $picks @($ot) $all $ms).Count
Update-PartialPicksOnModelCheck $picks $all $ms @($ot, $vk) @($vk)
'after uncheck: ' + $picks[$ot][$pipe] + ' ' + $before + ' ' + @(Get-PartialPairs $picks @($ot) $all $ms).Count
# повторная отметка ОТ — снова все
Update-PartialPicksOnModelCheck $picks $all $ms @($vk) @($vk, $ot)
'ot rechecked: ' + $picks[$ot][$pipe]
""")
        self.assertEqual(out[0], "vk none: 0")
        self.assertEqual(out[1], "vk on: 3")                  # ВК: общие, трубы, «гео»
        self.assertEqual(out[2], "ot pipe kept: False")
        self.assertEqual(out[3], "after uncheck: False 3 3")
        self.assertEqual(out[4], "ot rechecked: True")

    def test_pairs_are_checked_models_times_their_picks_and_captions(self):
        out = self.logic(r"""
$picks = New-PartialPicks $models $all $ms
Set-PartialNodeAll $picks @($ov) $all $ms $false
Set-PartialSpecPick $picks @($ov) $ms $duct $true                         # у ОВ только воздуховоды
Set-PartialSpecPick $picks @($vk) $ms $pipe $false                        # у ВК без труб
$pairs = @(Get-PartialPairs $picks @($ov, $vk) $all $ms)
'pairs: ' + (($pairs | ForEach-Object { ($_.Model -replace 'PROJ_IOS_|_ORG.*', '') + ':' + ($_.Spec -replace '^П\s+', '').Substring(0, 3) }) -join ',')
Format-PartialTotalCaption $pairs
'unchecked model gives nothing: ' + @(Get-PartialPairs $picks @($ot) $all $ms).Count + ' vs ' + @(Get-PartialPairs $picks @() $all $ms).Count
Format-PartialNodeCaption $vk $picks @($vk) $all $ms
Format-PartialNodeCaption 'ВК' $picks @($vk, $ot) $all $ms
Format-PartialTotalCaption @()
""")
        self.assertEqual(out[0], "pairs: OV:5.1,VK:1_К,VK:16.")
        self.assertEqual(out[1], "Всего к выгрузке: 3 пар, 2 моделей")
        self.assertEqual(out[2], "unchecked model gives nothing: 4 vs 0")
        self.assertEqual(out[3], "PROJ_IOS_VK_ORG_K01_R24: выбрано 2 из 3")
        self.assertEqual(out[4], "ВК: выбрано 3 из 4 (частично: 1)")       # объединение ВК+ОТ; трубы отмечены только у ОТ
        self.assertEqual(out[5], "Всего к выгрузке: 0 пар, 0 моделей")

    def test_all_none_buttons_only_affect_selected_node(self):
        out = self.logic(r"""
$picks = New-PartialPicks $models $all $ms
Set-PartialNodeAll $picks @($vk) $all $ms $false
'vk none, ot untouched: ' + @(Get-PartialPairs $picks @($vk) $all $ms).Count + ' ' + @(Get-PartialPairs $picks @($ot) $all $ms).Count
Set-PartialNodeAll $picks @($vk, $ot) $all $ms $true                        # раздел
'section all: ' + @(Get-PartialPairs $picks @($vk) $all $ms).Count + ' ' + @(Get-PartialPairs $picks @($ot) $all $ms).Count + ' ov: ' + @(Get-PartialPairs $picks @($ov) $all $ms).Count
Set-PartialNodeAll $picks @($ot) $all $ms $false
'ot none: ' + @(Get-PartialPairs $picks @($ot) $all $ms).Count + ' vk: ' + @(Get-PartialPairs $picks @($vk) $all $ms).Count
""")
        self.assertEqual(out, ["vk none, ot untouched: 0 4", "section all: 3 4 ov: 3", "ot none: 0 vk: 3"])

    def test_saved_pairs_roundtrip_new_models_and_legacy_formats(self):
        out = self.logic(r"""
$picks = New-PartialPicks @($vk, $ot) $all $ms
Set-PartialSpecPick $picks @($vk) $ms $pipe $false
Set-PartialNodeAll $picks @($ot) $all $ms $false
$lines = @(ConvertTo-PartialPairLines $picks)
'lines: ' + $lines.Count
# как их пишет окно: + «модель|*» для известных моделей
$saved = @{}
foreach ($l in $lines) { $saved[($l -replace '^asgpair=', '')] = $true }
foreach ($m in @($vk, $ot)) { $saved[$m + '|*'] = $true }
$back = New-PartialPicks @($vk, $ot, $ov) $all $ms $saved
'vk same: ' + ((Join-S ($picks[$vk].Keys | Where-Object { $picks[$vk][$_] })) -eq (Join-S ($back[$vk].Keys | Where-Object { $back[$vk][$_] })))
'ot none: ' + @(Get-PartialPairs $back @($ot) $all $ms).Count + '; new model ov all: ' + @(Get-PartialPairs $back @($ov) $all $ms).Count
# прежний формат: asgspecs (одни отметки на все модели)
$legacy = @{}; $legacy[$duct] = $true; $legacy[$common] = $true
$old = New-PartialPicks @($ov, $vk) $all $ms $null $legacy
'legacy: ' + @(Get-PartialPairs $old @($ov, $vk) $all $ms).Count
# файл пар и строки run.cfg
$f = Join-Path ([System.IO.Path]::GetTempPath()) ('pairs_' + [guid]::NewGuid().ToString('N') + '.txt')
Write-PartialPairsFile $f @(Get-PartialPairs $picks @($vk) $all $ms)
$bytes = [System.IO.File]::ReadAllBytes($f); 'bom: ' + ($bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB)
$txt = [System.IO.File]::ReadAllText($f); $first = ($txt -split "`r?`n")[0]
'first: ' + ($first -replace "`t", ' || ')
Remove-Item $f
(ConvertTo-PartialCfgLines @('С1') @('М1') 'C:\отчеты\служебные\x\assign_pairs.txt') -join ' ;; '
(ConvertTo-PartialCfgLines @('С1') @('М1')) -join ' ;; '
""")
        self.assertEqual(out[0], "lines: 2")                                        # у ВК две пары (общая, «гео»), у ОТ ни одной
        self.assertEqual(out[1], "vk same: True")
        self.assertEqual(out[2], "ot none: 0; new model ov all: 3")
        self.assertEqual(out[3], "legacy: 3")                                # ОВ: воздуховоды + общая, ВК: общая
        self.assertEqual(out[4], "bom: True")
        self.assertEqual(out[5], "first: PROJ_IOS_VK_ORG_K01_R24 || П        1_Классификатор_Все")
        self.assertEqual(out[6], "assign_scope=partial ;; assign_specs=С1 ;; assign_models=М1 ;; assign_pairs_file=C:\\отчеты\\служебные\\x\\assign_pairs.txt")
        self.assertEqual(out[7], "assign_scope=partial ;; assign_specs=С1 ;; assign_models=М1")

    def test_ui_script_parses_and_partial_functions_exist(self):
        script = os.path.join(HERE, "..", "..", "..", "..")  # не используется
        ps = os.path.join(tempfile.mkdtemp(prefix="bvs_pui_"), "ast.ps1")
        self.addCleanup(shutil.rmtree, os.path.dirname(ps), True)
        with open(ps, "w", encoding="utf-8-sig") as f:
            f.write(r"""
$errs = $null; $tok = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile('%s', [ref]$tok, [ref]$errs)
"errors=" + @($errs).Count
$defined = @{}
foreach ($f in @('ui_schedules.ps1','assignment_partial.ps1','assignment_partial_ui.ps1','assignment_export_stream.ps1','bvs_assignment_guard.ps1')) {
  $a = [System.Management.Automation.Language.Parser]::ParseFile((Join-Path '%s' $f), [ref]$tok, [ref]$errs)
  foreach ($fn in $a.FindAll({ $args[0] -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)) { $defined[$fn.Name] = $true }
}
$missing = @()
foreach ($c in $ast.FindAll({ $args[0] -is [System.Management.Automation.Language.CommandAst] }, $true)) {
  $n = $c.GetCommandName()
  if ($n -match '(Asg|Partial|PartialChoices|AsgChecked)' -and -not $defined.ContainsKey($n)) { $missing += $n }
}
"missing=" + (($missing | Sort-Object -Unique) -join ',')
""" % (os.path.join(TOOL, "ui_schedules.ps1"), TOOL))
        r = run_pwsh(PWSH, ["-File", ps])
        self.assertEqual(r.returncode, 0, r.stdout + r.stderr)
        self.assertIn("errors=0", r.stdout)
        self.assertIn("missing=\n", r.stdout + "\n")
        self.assertNotRegex(r.stdout, r"missing=\S")


if __name__ == "__main__":
    unittest.main()
