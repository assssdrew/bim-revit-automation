"""excel_isolated.ps1: закрывается ТОЛЬКО свой процесс Excel (PID из запуска / по hwnd своего COM-объекта); чужой EXCEL.EXE, появившийся во время прогона,
не трогается; PID свой определить не удалось — ничего не убивается, удаляется только своя seed-книга. «Excel» в тесте — копия sleep под именем EXCEL."""
import os
import shutil
import stat
import sys
import tempfile
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from test_chain import pwsh, PWSH  # noqa: E402

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(HERE)))
TOOL = os.path.join(ROOT, "служебное", "скрипты")
EXCEL = os.path.join(TOOL, "excel_isolated.ps1")
SLEEP = shutil.which("sleep")


def read(p):
    with open(p, encoding="utf-8-sig") as f:
        return f.read()


PRELUDE = r"""
. '%s'
$fake = '%s'
function Start-Fake { (Start-Process -FilePath $fake -ArgumentList '120' -PassThru).Id }
function Alive([int]$id) { [bool](Get-Process -Id $id -ErrorAction SilentlyContinue) }
function Seeds { @(Get-ChildItem ([IO.Path]::GetTempPath()) -Filter 'bvs_excel_seed_*.xlsx' -ErrorAction SilentlyContinue).Count }
function Get-ExcelExePath { $fake }
$script:Foreign = 0
"""


@unittest.skipUnless(PWSH and SLEEP and sys.platform != "win32", "нужны pwsh и Linux (sleep под именем EXCEL)")
class OwnExcelOnlyTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="bvs_xl_")
        self.fake = os.path.join(self.tmp, "EXCEL")
        shutil.copy(SLEEP, self.fake)
        os.chmod(self.fake, os.stat(self.fake).st_mode | stat.S_IXUSR)
        self.addCleanup(self.kill_fakes)
        self.addCleanup(shutil.rmtree, self.tmp, True)

    def kill_fakes(self):
        os.system("pkill -x EXCEL 2>/dev/null")

    def run_ps(self, body):
        script = os.path.join(self.tmp, "t.ps1")
        with open(script, "w", encoding="utf-8-sig") as f:
            f.write(PRELUDE % (EXCEL.replace("'", "''"), self.fake.replace("'", "''")) + body)
        r = pwsh("-File", script)
        self.assertEqual(r.returncode, 0, r.stdout + r.stderr)
        return [l for l in r.stdout.splitlines() if l.strip()]

    def test_stop_closes_only_own_process_and_removes_seed(self):
        out = self.run_ps(r"""
$own = Start-Fake
$seed = New-BvsExcelSeedFile
$foreign = Start-Fake                      # «чужой» Excel появился после старта нашего
Start-Sleep -Milliseconds 400
Stop-IsolatedExcelApplication ([pscustomobject]@{ Application = $null; ProcessId = $own; SeedPath = $seed })
Start-Sleep -Milliseconds 400
'own alive: ' + (Alive $own)
'foreign alive: ' + (Alive $foreign)
'seed exists: ' + (Test-Path -LiteralPath $seed)
Stop-Process -Id $foreign -Force
""")
        self.assertEqual(out, ["own alive: False", "foreign alive: True", "seed exists: False"])

    def test_failed_start_kills_own_pid_only_and_removes_seed(self):
        out = self.run_ps(r"""
$script:ownPid = 0; $script:foreignPid = 0
function Start-IsolatedExcelProcess { param($Exe, $SeedPath)
    $script:ownPid = Start-Fake
    if ($script:foreignPid -eq 0) { $script:foreignPid = Start-Fake }       # чужой Excel запускается во время прогона
    Start-Sleep -Milliseconds 400
    return $script:ownPid }
function Wait-ExcelComForProcessId { param($ProcessId, $TimeoutSec) return $null }    # COM не подключился
$before = Seeds
$msg = ''
try { [void](New-IsolatedExcelApplication) } catch { $msg = $_.Exception.Message }
Start-Sleep -Milliseconds 400
'own alive: ' + (Alive $script:ownPid)
'foreign alive: ' + (Alive $script:foreignPid)
'seeds left: ' + ((Seeds) - $before)
'pid unknown flag: ' + $script:ExcelPidUnknown
'message: ' + $msg.Substring(0, 22)
Stop-Process -Id $script:foreignPid -Force
""")
        self.assertEqual(out[0], "own alive: False")                                    # свой закрыт
        self.assertEqual(out[1], "foreign alive: True")                                 # чужой — нет
        self.assertEqual(out[2], "seeds left: 0")                                       # seed удалён
        self.assertEqual(out[3], "pid unknown flag: False")
        self.assertTrue(out[4].startswith("message: Excel запущен (PID"), out[4])

    def test_unknown_pid_kills_nothing_and_reports(self):
        out = self.run_ps(r"""
$script:foreignPid = Start-Fake
function Start-IsolatedExcelProcess { param($Exe, $SeedPath) return 0 }              # свой PID определить не удалось
$before = Seeds
$msg = ''
try { [void](New-IsolatedExcelApplication) } catch { $msg = $_.Exception.Message }
'foreign alive: ' + (Alive $script:foreignPid)
'seeds left: ' + ((Seeds) - $before)
'pid unknown flag: ' + $script:ExcelPidUnknown
'message: ' + $msg.Substring(0, 36)
Stop-Process -Id $script:foreignPid -Force
""")
        self.assertEqual(out[0], "foreign alive: True")
        self.assertEqual(out[1], "seeds left: 0")
        self.assertEqual(out[2], "pid unknown flag: True")
        self.assertTrue(out[3].startswith("message: Excel не закрыт, PID неизвестен"), out[3])

    def test_launched_process_exits_at_once_means_unknown_pid_nothing_killed(self):
        out = self.run_ps(r"""
$script:foreignPid = Start-Fake
function Start-IsolatedExcelProcess { param($Exe, $SeedPath)
    $p = Start-Process -FilePath $fake -ArgumentList '0' -PassThru; $p.WaitForExit(); return $p.Id }     # запущенный процесс завершился сразу
function Wait-ExcelComForProcessId { param($ProcessId, $TimeoutSec) return $null }
$before = Seeds
try { [void](New-IsolatedExcelApplication) } catch { }
'foreign alive: ' + (Alive $script:foreignPid)
'seeds left: ' + ((Seeds) - $before)
'pid unknown flag: ' + $script:ExcelPidUnknown
Stop-Process -Id $script:foreignPid -Force
""")
        self.assertEqual(out, ["foreign alive: True", "seeds left: 0", "pid unknown flag: True"])


class StaticTests(unittest.TestCase):
    def test_no_new_process_heuristic_left(self):
        t = read(EXCEL)
        for gone in ("Stop-NewExcelProcesses", "$initial", "BeforePids", "-not $BeforePids"):
            self.assertNotIn(gone, t, gone)
        # единственное место, где убивается процесс Excel — Stop-OwnExcelProcess с проверкой имени и PID
        self.assertEqual(t.count("Stop-Process"), 1)
        self.assertIn("function Stop-OwnExcelProcess", t)
        self.assertNotIn("Get-Process -Name EXCEL", t)                                    # «все EXCEL.EXE» не перебираются вовсе
        self.assertIn("[BvsExcelRot]::GetExcelProcessId($excel)", t)                      # PID по hwnd своего COM-объекта
        self.assertIn("-PassThru", t)                                                       # PID запущенного нами процесса

    def test_ui_reports_unknown_pid_with_csv_link(self):
        ui = read(os.path.join(TOOL, "ui_schedules.ps1"))
        self.assertIn("Excel не закрыт, PID неизвестен", ui)
        tick = ui[ui.index("$xlsx = Export-ReportXlsx"):ui.index("$form.Controls.Add($ui.Timer)")]
        self.assertIn("$script:ExcelPidUnknown", tick)
        self.assertIn("Set-CsvLink $script:Run.Summary", tick)


if __name__ == "__main__":
    unittest.main()
