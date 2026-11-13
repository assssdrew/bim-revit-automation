"""Фаза 7: состояние «Выборочно» одним файлом <история>_состояние.txt вместо <история>_буфер.txt + _проверки.txt + _каталог.txt.

* Сборки «Всё», «Выборочно» (незагруженная связь, удалённые спецификация и модель, выгрузка без каталога) и «Собрать динамику»
  дают ровно эталон tests/baseline_partial, снятый на коде фазы 6: xlsx, payload, история, журнал, итоговые строки, содержимое состояния.
* Посреди цепочки состояние в прежнем формате (три файла) — читается так же; после переноса — тоже.
* Перенос при открытии окна: резервная копия до записи, прежние файлы -> .migrated_<дата> (не удаляются), повторный запуск ничего не делает;
  после отката (прежние файлы новее) берутся прежние файлы.
* Пустой раздел [inproject] / [summary] — «ни одна не отмечена», раздела нет — «все»."""
import json
import os
import shutil
import sys
import tempfile
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import baseline_lib as bl  # noqa: E402
from test_baseline import first_diff  # noqa: E402
from test_chain import pwsh, PWSH  # noqa: E402

TOOL = bl.TOOL
STEPS = [s[0] for s in bl.partial_steps()]


def q(s):
    return str(s).replace("'", "''")


def split_state(out, stem="задание_история", touch_newer=False):
    """<stem>_состояние.txt -> прежние три файла (как их писала фаза 6), сам файл состояния удаляется (тестовая папка)."""
    merged = os.path.join(out, stem + bl.STATE_NEW)
    with open(merged, encoding="utf-8-sig") as f:
        lines = f.read().splitlines()
    parts, cur = {}, None
    for line in lines:
        if line in ("[buffer]", "[checks]", "[catalog]"):
            cur = line[1:-1]
            parts[cur] = []
        elif cur and line.strip():
            parts[cur].append(line)
    os.remove(merged)
    for part, body in parts.items():
        with open(os.path.join(out, stem + bl.STATE_OLD[part]), "w", encoding="utf-8-sig", newline="") as f:
            f.write("".join(l + "\r\n" for l in body))
    return parts


def jsonish(x):
    return json.loads(json.dumps(x))


@unittest.skipUnless(PWSH, "нет pwsh")
class SameAsPhase6Tests(unittest.TestCase):
    def assert_steps(self, res, steps):
        for step in steps:
            got, exp = jsonish(res[step]), bl.load_partial_baseline(step)
            d = first_diff(got, exp, step)
            self.assertIsNone(d, d)

    def test_new_format_equals_phase6_baseline(self):
        res = bl.partial_scenario(PWSH)
        self.assert_steps(res, STEPS)

    def test_old_format_state_in_the_middle_is_read_the_same(self):
        """Перед каждой сборкой состояние переписывается в прежние три файла: результат тот же, после сборки — снова один файл."""
        seen = []

        def before(step, out):
            if os.path.exists(os.path.join(out, "задание_история" + bl.STATE_NEW)):
                split_state(out)
                seen.append(step)
        res = bl.partial_scenario(PWSH, before_step=before)
        self.assert_steps(res, STEPS)
        self.assertEqual(seen, STEPS[1:])

    def test_migration_in_the_middle_gives_the_same_result(self):
        def before(step, out):
            if step != "P4":
                return
            split_state(out)
            script = os.path.join(out, "..", "m.ps1")
            with open(script, "w", encoding="utf-8-sig") as f:
                f.write(". '%s'\n. '%s'\n$bk = Join-Path '%s' 'bk'\n$r = Invoke-PartialStateMigration (Join-Path '%s' 'задание_история.txt') { New-Item -ItemType Directory -Force -Path $bk | Out-Null; $bk } ([datetime]'2026-11-10')\n$r.Action"
                        % (q(os.path.join(TOOL, "assignment_export_stream.ps1")), q(os.path.join(TOOL, "assignment_partial.ps1")), q(out), q(out)))
            r = pwsh("-File", script)
            assert r.returncode == 0 and "merged" in r.stdout, r.stdout + r.stderr
            names = sorted(os.listdir(out))
            assert "задание_история_буфер.migrated_2026-11-10.txt" in names, names
            assert "задание_история_буфер.txt" not in names, names
        res = bl.partial_scenario(PWSH, before_step=before)
        self.assert_steps(res, STEPS)


@unittest.skipUnless(PWSH, "нет pwsh")
class MigrationTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="bvs_pst_")
        self.addCleanup(shutil.rmtree, self.tmp, True)
        self.rep = os.path.join(self.tmp, "отчеты")
        self.hist = os.path.join(self.rep, "история")
        os.makedirs(self.hist)
        os.makedirs(os.path.join(self.rep, "проект"))
        with open(os.path.join(self.rep, "проект", "распределение_PROJ.txt"), "w", encoding="utf-8-sig") as f:
            f.write("[project]\nstage=П\n")
        self.files = {
            "история_PROJ.txt": "2026-02-01\tEOM\tM\tS\tP\t1\tP\t2026-02-01 10:00:00\r\n",
            "задание_журнал_PROJ.txt": "M\tEOM\tS\tF\tI\tP\t\t2026-02-01 10:00:00\t2026-02-01 10:00:00\topen\t\t0\r\n",
            "история_PROJ_буфер.txt": "EOM\tM\tS\tP\tP\tF\tI\tV\r\nEOM\tM2\tS\tP\tP\tF\tI2\tV\r\n",
            "история_PROJ_проверки.txt": "*\tS\t2026-02-01 10:00:00\r\nM\tS\t2026-02-02 10:00:00\r\n",
            "история_PROJ_каталог.txt": "spec\tS\r\nmodel\tM\tloaded\r\n",
            "история_SITE_буфер.txt": "",                          # пустой буфер (была полная сборка без замечаний), без каталога
            "история_SITE_проверки.txt": "*\tS\t2026-02-03 10:00:00\r\n",
        }
        for n, t in self.files.items():
            with open(os.path.join(self.hist, n), "w", encoding="utf-8-sig", newline="") as f:
                f.write(t)

    def ps(self, body):
        script = os.path.join(self.tmp, "t.ps1")
        with open(script, "w", encoding="utf-8-sig") as f:
            f.write("".join(". '%s'\n" % q(os.path.join(TOOL, n)) for n in ("project_config.ps1", "assignment_export_stream.ps1", "assignment_partial.ps1", "assignment_partial_ui.ps1")) + body)
        r = pwsh("-File", script)
        self.assertEqual(r.returncode, 0, r.stdout + r.stderr)
        return [l for l in r.stdout.splitlines() if l.strip()]

    def state_of(self, stem):
        script = """
$s = Read-PartialState (Get-PartialStatePaths (Join-Path '%s' '%s.txt'))
'buf=' + $(if ($null -eq $s.Buffer) { 0 } else { @($s.Buffer).Count }) + ' has=' + $s.HasBuffer + ' null=' + ($null -eq $s.Buffer) + ' chk=' + (@($s.Checks.Keys | Sort-Object | ForEach-Object { $_ + '=' + $s.Checks[$_] }) -join ',') + ' cat=' + $s.Catalog.Has + ':' + (@($s.Catalog.Specs.Keys) -join ',')
'ready=' + (Test-PartialBaseReady (Join-Path '%s' '%s.txt') '')
""" % (q(self.hist), stem, q(self.hist), stem)
        return script

    def test_first_start_backup_rename_merge_and_second_start_does_nothing(self):
        before = self.ps(self.state_of("история_PROJ") + self.state_of("история_SITE"))
        out = self.ps(r"""
$r = Invoke-ProjectStateMigration '%s' ([datetime]'2026-11-10 09:00')
'backup=' + [bool]$r.Backup + ' notes=' + (@($r.Notes) -join '|') + ' errors=' + @($r.Errors).Count
$r2 = Invoke-ProjectStateMigration '%s' ([datetime]'2026-11-10 09:05')
'again backup=' + [bool]$r2.Backup + ' notes=' + @($r2.Notes).Count
""" % (q(self.rep), q(self.rep)) + self.state_of("история_PROJ") + self.state_of("история_SITE"))
        self.assertEqual(out[0], "backup=True notes=Файлы «Выборочно» объединены: буфер, проверки и каталог теперь в одном файле <история>_состояние.txt (2 шт.; технический шаг, один раз). errors=0")
        self.assertEqual(out[1], "again backup=False notes=0")
        self.assertEqual(out[2:], before)                          # то же состояние, что читалось из прежних файлов
        self.assertIn("buf=2 has=True null=False chk=*\x1fS=2026-02-01 10:00:00,M\x1fS=2026-02-02 10:00:00 cat=True:S", before[0])
        self.assertIn("buf=0 has=True null=True", before[2])        # пустой буфер: как прежде — засев, но база «готова»
        self.assertIn("cat=False", before[2])
        names = sorted(os.listdir(self.hist))
        self.assertEqual(names, sorted([
            "история_PROJ.txt", "задание_журнал_PROJ.txt", "история_PROJ_состояние.txt", "история_SITE_состояние.txt",
            "история_PROJ_буфер.migrated_2026-11-10.txt", "история_PROJ_проверки.migrated_2026-11-10.txt", "история_PROJ_каталог.migrated_2026-11-10.txt",
            "история_SITE_буфер.migrated_2026-11-10.txt", "история_SITE_проверки.migrated_2026-11-10.txt"]))
        # прежние файлы целы (только переименованы), история и журнал не тронуты
        for n, t in self.files.items():
            stem, ext = os.path.splitext(n)
            p = os.path.join(self.hist, n if n in names else stem + ".migrated_2026-11-10" + ext)
            with open(p, encoding="utf-8-sig", newline="") as f:
                self.assertEqual(f.read(), t, n)
        # резервная копия — до записи, в проект\_backup_* (вне автоочистки), с прежними файлами состояния
        backups = [d for d in os.listdir(os.path.join(self.rep, "проект")) if d.startswith("_backup_")]
        self.assertEqual(len(backups), 1)
        bdir = os.path.join(self.rep, "проект", backups[0])
        self.assertEqual(sorted(os.listdir(os.path.join(bdir, "история"))), sorted(self.files))
        with open(os.path.join(bdir, "перенос_состояния.txt"), encoding="utf-8-sig") as f:
            log = f.read()
        self.assertIn("история_PROJ: 3 файла(ов) объединены в история_PROJ_состояние.txt", log)
        self.assertIn("история_PROJ_буфер.txt -> история_PROJ_буфер.migrated_2026-11-10.txt", log)
        self.assertIn("история_SITE: 2 файла(ов) объединены в история_SITE_состояние.txt (строк: буфер 0, проверки 1, каталог нет)", log)

    def test_no_backup_no_changes(self):
        out = self.ps(r"""
try { [void](Invoke-PartialStateMigration (Join-Path '%s' 'история_PROJ.txt') { return '' } ([datetime]'2026-11-10')); 'no error' } catch { 'error: ' + $_.Exception.Message }
""" % q(self.hist))
        self.assertEqual(out, ["error: нет резервной копии — перенос не выполнен"])
        self.assertEqual(sorted(os.listdir(self.hist)), sorted(self.files))

    def test_old_files_newer_than_state_win_after_rollback(self):
        """Откат на прежнюю версию и снова обновление: прежние файлы новее файла состояния — читаются они, перенос берёт их."""
        self.ps("$r = Invoke-ProjectStateMigration '%s' ([datetime]'2026-11-10')" % q(self.rep))
        import time
        time.sleep(0.05)
        with open(os.path.join(self.hist, "история_PROJ_буфер.txt"), "w", encoding="utf-8-sig", newline="") as f:
            f.write("EOM\tM3\tS\tP\tP\tF\tI3\tV\r\n")                     # прежняя версия записала новый буфер
        out = self.ps(self.state_of("история_PROJ") + r"""
$r = Invoke-ProjectStateMigration '%s' ([datetime]'2026-11-11')
'notes=' + @($r.Notes).Count
""" % q(self.rep) + self.state_of("история_PROJ"))
        self.assertIn("buf=1 ", out[0])
        self.assertIn("chk=*\x1fS=2026-02-01 10:00:00,M\x1fS=2026-02-02 10:00:00 cat=True:S", out[0])   # остальные части — из файла состояния
        self.assertEqual(out[2], "notes=1")
        self.assertEqual(out[3:], out[:2])
        self.assertIn("история_PROJ_буфер.migrated_2026-11-11.txt", os.listdir(self.hist))

    def test_choices_same_from_old_and_new_format(self):
        body = "$c = Get-PartialChoices (Join-Path '%s' 'история_PROJ.txt') ''\n($c.Models | ForEach-Object { $_.Name + '|' + $_.Loaded }) -join ';'\n$c.Specs -join ';'\n" % q(self.hist)
        old = self.ps(body)
        self.ps("$r = Invoke-ProjectStateMigration '%s' ([datetime]'2026-11-10')" % q(self.rep))
        self.assertEqual(self.ps(body), old)
        self.assertTrue(old)


@unittest.skipUnless(PWSH, "нет pwsh")
class EmptyInProjectTests(unittest.TestCase):
    def test_empty_section_means_none_missing_means_all(self):
        tmp = tempfile.mkdtemp(prefix="bvs_inp_")
        self.addCleanup(shutil.rmtree, tmp, True)
        base = "[project]\nstage=П\n[models]\nA_IOS_OV_1\tOV\n[specs]\n1_А\n2_Б\n"
        files = {"missing": base, "header_only": base + "[inproject]\n[summary]\n", "marked": base + "[inproject]\n2_Б\n"}
        for k, t in files.items():
            with open(os.path.join(tmp, k + ".txt"), "w", encoding="utf-8-sig") as f:
                f.write(t)
        script = os.path.join(tmp, "t.ps1")
        with open(script, "w", encoding="utf-8-sig") as f:
            f.write(". '%s'\n" % q(os.path.join(TOOL, "project_config.ps1")) + r"""
function Show($c) { 'in=' + ($null -ne $c.InProject) + ':' + (@($c.Specs | Where-Object { Test-ProjectSpecIn $c $_ }) -join ',') + ' sum=' + ($null -ne $c.Summary) + ':' + (@($c.Specs | Where-Object { Test-ProjectSpecInSummary $c $_ }) -join ',') }
foreach ($k in @('missing', 'header_only', 'marked')) { $k + ' ' + (Show (Read-ProjectConfig (Join-Path '%s' ($k + '.txt')))) }
# пользователь снял все «В проекте» и нажал «Сохранить» -> при следующем открытии не отмечено ничего
$c = Read-ProjectConfig (Join-Path '%s' 'missing.txt')
Set-ProjectInProjectBulk $c @($c.Specs) $false
Write-ProjectConfig $c (Join-Path '%s' 'saved.txt')
'saved ' + (Show (Read-ProjectConfig (Join-Path '%s' 'saved.txt')))
$n = Get-ProjectInProjectNames (Read-ProjectConfig (Join-Path '%s' 'saved.txt'))
'names=' + $n.Count + ' isnull=' + ($null -eq $n) + ' null=' + ($null -eq (Get-ProjectInProjectNames (Read-ProjectConfig (Join-Path '%s' 'missing.txt'))))
# снова отметить одну -> пометка «пусто» уходит
$c = Read-ProjectConfig (Join-Path '%s' 'saved.txt'); Set-ProjectInProjectBulk $c @('1_А') $true; Write-ProjectConfig $c (Join-Path '%s' 'saved2.txt')
'saved2 ' + (Show (Read-ProjectConfig (Join-Path '%s' 'saved2.txt')))
""".replace("%s", q(tmp)))
        r = pwsh("-File", script)
        self.assertEqual(r.returncode, 0, r.stdout + r.stderr)
        out = [l for l in r.stdout.splitlines() if l.strip()]
        self.assertEqual(out[0], "missing in=False:1_А,2_Б sum=False:1_А,2_Б")        # прежние файлы: все
        self.assertEqual(out[1], "header_only in=True: sum=True:")                   # пустой раздел: ни одной (как у insert_schedules.py)
        self.assertEqual(out[2], "marked in=True:2_Б sum=False:1_А,2_Б")
        self.assertEqual(out[3], "saved in=True: sum=False:1_А,2_Б")
        self.assertEqual(out[4], "names=0 isnull=False null=True")        # пустой список (не «раздела нет») -> «Выбор восстановлен: 0 спек»
        self.assertEqual(out[5], "saved2 in=True:1_А sum=False:1_А,2_Б")
        with open(os.path.join(tmp, "saved.txt"), encoding="utf-8-sig") as f:
            saved = f.read()
        self.assertIn("[inproject]\n# пусто: ни одна спецификация не отмечена «В проекте»\n[groups]", saved)
        with open(os.path.join(tmp, "saved2.txt"), encoding="utf-8-sig") as f:
            self.assertNotIn("# пусто", f.read())

    def test_insert_sees_the_same_as_window(self):
        """insert_schedules.py (запасной разбор) и план окна: файл с пометкой «пусто» — в сводную ничего, как в окне."""
        import test_insert_project as tip
        proj = tip.PROJECT.replace("[cells]", "[inproject]\n# пусто: ни одна спецификация не отмечена «В проекте»\n[cells]")
        r = tip.Run("C:\\m\\PROJ_IOS_SVOD_ORG_R24.rvt", mode="summary", project=proj).go()
        self.addCleanup(r.cleanup)
        self.assertIn("В распределении нет отмеченных спецификаций", r.read("r1_сводка.csv") + str(r.error))
        self.assertEqual(r.names(), [])


if __name__ == "__main__":
    unittest.main()
