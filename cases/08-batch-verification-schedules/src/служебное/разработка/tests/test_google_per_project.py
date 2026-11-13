"""Google: адрес веб-приложения и секрет на проект (google_<проект>.cfg), общий google_push.cfg — только после подтверждения; секрет не в логах и не в git."""
import os
import shutil
import subprocess
import sys
import tempfile
import unittest

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from test_chain import pwsh, PWSH  # noqa: E402

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(HERE)))     # tools/batch_verification_schedules
TOOL = os.path.join(ROOT, "служебное", "скрипты")
CFGLIB = os.path.join(TOOL, "google_push_cfg.ps1")
LAYOUT = os.path.join(TOOL, "report_layout.ps1")
PUSH = os.path.join(TOOL, "push_google.ps1")
UI = os.path.join(TOOL, "ui_schedules.ps1")

URL_A = "https://script.google.com/macros/s/AKfycbAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA1111/exec"
URL_B = "https://script.google.com/macros/s/AKfycbBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBBB2222/exec"
URL_SHARED = "https://script.google.com/macros/s/AKfycbSHARED0000000000000000000000000000000000000000000000009999/exec"
SECRET_A, SECRET_B, SECRET_SH = "SECRET-A-do-not-log-111", "SECRET-B-do-not-log-222", "SECRET-SHARED-do-not-log-999"


def read(p):
    with open(p, encoding="utf-8-sig") as f:
        return f.read()


PRELUDE = r"""
. '%s'
. '%s'
$errs = $null; $tok = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile('%s', [ref]$tok, [ref]$errs)
$texts = foreach ($n in @('Get-GooglePushDecision', 'Add-GooglePushLog')) {
    ($ast.FindAll({ param($x) $x -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $x.Name -eq $n }, $true) | Select-Object -First 1).Extent.Text
}
$tmpFile = Join-Path ([IO.Path]::GetTempPath()) ('bvs_fn_' + [guid]::NewGuid().ToString('N') + '.ps1')
[IO.File]::WriteAllText($tmpFile, ($texts -join "`n"), (New-Object System.Text.UTF8Encoding $true))
. $tmpFile
Remove-Item $tmpFile
""" % (CFGLIB.replace("'", "''"), LAYOUT.replace("'", "''"), UI.replace("'", "''"))


@unittest.skipUnless(PWSH, "нет pwsh")
class GooglePerProjectTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="bvs_gp_")
        self.addCleanup(shutil.rmtree, self.tmp, True)
        self.proj = os.path.join(self.tmp, "отчеты", "проект")
        os.makedirs(self.proj)
        self.svc = os.path.join(self.tmp, "служебное")              # служебное\google_push\google_push.cfg — общий запасной
        self.toolsdir = os.path.join(self.svc, "скрипты")
        os.makedirs(self.toolsdir)
        os.makedirs(os.path.join(self.svc, "google_push"))

    def cfg_a(self):
        return os.path.join(self.proj, "распределение_PROJ.txt")

    def cfg_b(self):
        return os.path.join(self.proj, "распределение_OTHER.txt")

    def write_project_google(self, name, url, secret):
        with open(os.path.join(self.proj, "google_%s.cfg" % name), "w", encoding="utf-8-sig") as f:
            f.write("webapp_url=%s\nsecret=%s\n" % (url, secret))

    def write_shared(self, url=URL_SHARED, secret=SECRET_SH):
        with open(os.path.join(self.svc, "google_push", "google_push.cfg"), "w", encoding="utf-8-sig") as f:
            f.write("webapp_url=%s\nsecret=%s\n" % (url, secret))

    def run_ps(self, body):
        script = os.path.join(self.tmp, "t.ps1")
        with open(script, "w", encoding="utf-8-sig") as f:
            f.write(PRELUDE + "$cfgA = '%s'; $cfgB = '%s'; $root = '%s'; $AppDir = '%s'; $script:Layout = Get-ReportLayout '%s'\n" % (self.cfg_a(), self.cfg_b(), self.toolsdir, self.tmp, os.path.join(self.tmp, "отчеты").replace("'", "''")) + body)
        r = pwsh("-File", script)
        self.assertEqual(r.returncode, 0, r.stdout + r.stderr)
        return [l for l in r.stdout.splitlines() if l.strip()]

    def test_different_projects_have_different_addresses(self):
        self.write_project_google("PROJ", URL_A, SECRET_A)
        self.write_project_google("OTHER", URL_B, SECRET_B)
        self.write_shared()
        out = self.run_ps(r"""
$a = Resolve-GooglePushTarget $cfgA $root $false
$b = Resolve-GooglePushTarget $cfgB $root $false
'A: ' + $a.Kind + ' ' + $a.ProjectName + ' ' + ($a.Cfg.WebAppUrl.EndsWith('1111/exec')) + ' ' + $a.Cfg.Secret
'B: ' + $b.Kind + ' ' + $b.ProjectName + ' ' + ($b.Cfg.WebAppUrl.EndsWith('2222/exec')) + ' ' + $b.Cfg.Secret
'same: ' + ($a.Cfg.WebAppUrl -eq $b.Cfg.WebAppUrl)
'path: ' + (Split-Path -Leaf (Get-ProjectGoogleConfigPath $cfgA))
'labelA: ' + $a.Label
""")
        self.assertEqual(out[0], "A: project PROJ True " + SECRET_A)
        self.assertEqual(out[1], "B: project OTHER True " + SECRET_B)
        self.assertEqual(out[2], "same: False")
        self.assertEqual(out[3], "path: google_PROJ.cfg")
        self.assertIn("проект PROJ", out[4])
        self.assertTrue(out[4].endswith("1111/exec"), out[4])                         # проект + конец адреса
        self.assertNotIn(SECRET_A, out[4])

    def test_no_project_address_asks_and_default_is_refusal(self):
        self.write_shared()
        out = self.run_ps(r"""
$script:asked = ''
$no = Get-GooglePushDecision $cfgA $root { param($q) $script:asked = $q; $false }
'asked: ' + $script:asked
'no -> go: ' + $no.Go + ' shared: ' + $no.AllowShared
'text: ' + $no.Text
$yes = Get-GooglePushDecision $cfgA $root { param($q) $true }
'yes -> go: ' + $yes.Go + ' shared: ' + $yes.AllowShared + ' kind: ' + $yes.Target.Kind
'label: ' + $yes.Target.Label
""")
        self.assertEqual(out[0], "asked: У проекта нет своего адреса Google. Отправить в общую таблицу из google_push.cfg?")
        self.assertEqual(out[1], "no -> go: False shared: False")
        self.assertIn("в общую таблицу не отправлено", out[2])
        self.assertEqual(out[3], "yes -> go: True shared: True kind: shared")
        self.assertIn("ОБЩАЯ таблица", out[4])
        self.assertNotIn(SECRET_SH, "\n".join(out))

    def test_project_with_address_does_not_ask_and_no_address_anywhere_does_not_send(self):
        self.write_project_google("PROJ", URL_A, SECRET_A)
        out = self.run_ps(r"""
$script:asked = 0
$d = Get-GooglePushDecision $cfgA $root { param($q) $script:asked++; $true }
'project: go ' + $d.Go + ' asked ' + $script:asked + ' shared ' + $d.AllowShared
$n = Get-GooglePushDecision $cfgB $root { param($q) $script:asked++; $true }      # у OTHER адреса нет, общего тоже нет
'none: go ' + $n.Go + ' kind ' + $n.Target.Kind + ' asked ' + $script:asked
""")
        self.assertEqual(out, ["project: go True asked 0 shared False", "none: go False kind none asked 1".replace("asked 1", "asked 0")])

    def test_push_script_refuses_shared_without_confirmation_and_uses_project_cfg(self):
        payload = os.path.join(self.tmp, "google_push_payload.json")
        with open(payload, "w", encoding="utf-8") as f:
            f.write('{"sheets":[]}')
        self.write_shared()

        def push(*extra):
            r = pwsh("-File", PUSH, "-PayloadOnly", "-PayloadPath", payload, "-ToolDir", self.toolsdir, "-ProjectConfigPath", self.cfg_a(), *extra)
            return r.returncode, r.stdout + r.stderr
        code, text = push()
        self.assertNotEqual(code, 0)
        self.assertIn("Общая таблица не используется без подтверждения", text)               # нет своего адреса: общий не берётся молча
        self.assertNotIn(SECRET_SH, text)
        # у проекта адрес есть, но секрет не задан — сообщение про настройку Google проекта (общий cfg не используется)
        self.write_project_google("PROJ", URL_A, "")
        code, text = push()
        self.assertNotEqual(code, 0)
        self.assertIn("Настройка Google:", text)
        self.assertIn("Не задан secret", text)
        self.assertIn("google_PROJ.cfg", text)
        self.assertNotIn(SECRET_SH, text)
        # подтверждённый общий адрес (в нём незаменённый секрет — дальше проверки формата не идём, сети нет)
        os.remove(os.path.join(self.proj, "google_PROJ.cfg"))
        self.write_shared(secret="REPLACE_ME")
        code, text = push("-AllowShared")
        self.assertNotEqual(code, 0)
        self.assertIn("secret не заменён", text)

    def test_shared_fallback_is_read_with_a_warning_and_never_written_or_deleted(self):
        self.write_shared()
        shared = os.path.join(self.svc, "google_push", "google_push.cfg")
        with open(shared, "rb") as f:
            before = f.read()
        out = self.run_ps(r"""
$t = Resolve-GooglePushTarget $cfgA $root $false
'kind: ' + $t.Kind
'warn: ' + $t.Warning
$d = Get-GooglePushDecision $cfgA $root { param($q) $false }
'go: ' + $d.Go
'log: ' + ([IO.File]::ReadAllText((Join-Path $script:Layout.Logs 'google_push.log')).Trim() -replace '^\S+ \S+\s+', '')
""")
        self.assertEqual(out[0], "kind: needs-confirm")                               # общий читается как запасной…
        self.assertIn("ВНИМАНИЕ: у проекта PROJ нет своего google_PROJ.cfg, найден общий запасной google_push.cfg", out[1])
        self.assertEqual(out[2], "go: False")                                         # …но без подтверждения не используется
        self.assertIn("нет своего google_PROJ.cfg", out[3])                           # предупреждение в журнале
        self.assertNotIn(SECRET_SH, "\n".join(out))
        with open(shared, "rb") as f:
            self.assertEqual(f.read(), before)                                       # общий файл не переписан
        self.assertFalse(os.path.exists(os.path.join(self.proj, "google_PROJ.cfg")))  # и файл проекта сам не создан
        # свой файл проекта есть — он главный, предупреждения нет
        self.write_project_google("PROJ", URL_A, SECRET_A)
        out = self.run_ps(r"""
$t = Resolve-GooglePushTarget $cfgA $root $false
'kind: ' + $t.Kind + ' warn: [' + $t.Warning + ']'
""")
        self.assertEqual(out, ["kind: project warn: []"])
        with open(shared, "rb") as f:
            self.assertEqual(f.read(), before)

    def test_no_code_creates_or_deletes_the_shared_cfg(self):
        for name in os.listdir(TOOL):
            if not name.endswith(".ps1"):
                continue
            t = read(os.path.join(TOOL, name))
            for line in t.splitlines():
                if "google_push.cfg" in line and not line.lstrip().startswith("#"):
                    for bad in ("WriteAll", "Set-Content", "Out-File", "Remove-Item", "::Delete", "::Move", "Copy-Item"):
                        self.assertNotIn(bad, line, name + ": " + line)

    def test_secret_never_reaches_log_or_status_text(self):
        self.write_project_google("PROJ", URL_A, SECRET_A)
        self.write_shared()
        out = self.run_ps(r"""
$t = Resolve-GooglePushTarget $cfgA $root $false
Add-GooglePushLog ('отправка: ' + $t.Label)
$s = Resolve-GooglePushTarget $cfgB $root $true
Add-GooglePushLog ('отправка: ' + $s.Label)
'log: ' + (Get-Content -Raw -LiteralPath (Join-Path $script:Layout.Logs 'google_push.log'))
'tail: ' + (Get-GoogleUrlTail 'https://x/exec') + ' | ' + (Get-GoogleUrlTail '%s')
$bad = Test-GooglePushConnection @{ WebAppUrl = 'https://example.com/lib'; Secret = '%s'; ConfigPath = '' }
'shape lines: ' + ($bad.Lines -join ' / ')
""" % (URL_A, SECRET_A))
        text = "\n".join(out)
        for secret in (SECRET_A, SECRET_SH, SECRET_B):
            self.assertNotIn(secret, text)
        self.assertIn("проект PROJ", text)
        self.assertIn("1111/exec", text)
        self.assertIn("ОБЩАЯ таблица", text)
        self.assertIn("не должен быть библиотекой".replace("не должен быть библиотекой", "webapp_url должен быть веб-приложением"), text)

    def test_write_read_roundtrip_keeps_other_field(self):
        out = self.run_ps(r"""
$p = Get-ProjectGoogleConfigPath $cfgA
Write-ProjectGoogleConfig $p '' ''
'empty creates file: ' + (Test-Path -LiteralPath $p)
Write-ProjectGoogleConfig $p 'https://script.google.com/macros/s/X/exec' $null
Write-ProjectGoogleConfig $p $null 'sec'
$c = Read-ProjectGoogleConfig $p
'url: ' + $c.WebAppUrl + ' secret: ' + $c.Secret
Write-ProjectGoogleConfig $p 'https://script.google.com/macros/s/Y/exec' $null
$c = Read-ProjectGoogleConfig $p
'url changed, secret kept: ' + $c.WebAppUrl + ' ' + $c.Secret
'other project untouched: ' + (Test-Path -LiteralPath (Get-ProjectGoogleConfigPath $cfgB))
""")
        self.assertEqual(out[0], "empty creates file: False")
        self.assertEqual(out[1], "url: https://script.google.com/macros/s/X/exec secret: sec")
        self.assertEqual(out[2], "url changed, secret kept: https://script.google.com/macros/s/Y/exec sec")
        self.assertEqual(out[3], "other project untouched: False")


class StaticTests(unittest.TestCase):
    def setUp(self):
        self.ui, self.push, self.cfg = read(UI), read(PUSH), read(CFGLIB)

    def test_ui_has_no_url_secret_fields_but_status_check_and_dialog(self):
        # полей адреса и секрета в окне нет: читается google_<проект>.cfg, показывается строка статуса, кнопки «Проверить» и «Подключить/Изменить»
        for gone in ("$ui.GoogleUrl", "$ui.GoogleSecret", "Save-ProjectGoogle", "Адрес веб-приложения (…/exec):"):
            self.assertNotIn(gone, self.ui)
        for part in ('$ui.GoogleCheck.Text = "Проверить"', "$ui.GoogleCheck.Add_Click({ Invoke-GoogleCheck })", "$ui.GoogleConnect.Add_Click({ Show-GoogleConnectDialog })",
                     "$ui.GoogleInfo = New-Object", '"Изменить подключение…"', '"Подключить Google…"', "$tSec.UseSystemPasswordChar = $true",
                     "Google: проект {0}, адрес {1}"):
            self.assertIn(part, self.ui)
        check = self.ui[self.ui.index("function Invoke-GoogleCheck"):self.ui.index("function Get-GooglePushDecision")]
        self.assertIn("Test-GooglePushConnection $cfg", check)                               # ping + пробный POST, формат — Test-BvsGooglePushConfigShape
        self.assertNotIn(".Secret", check)                                                    # секрет не попадает в статус и журнал
        self.assertIn("Test-BvsGooglePushConfigShape $Cfg", self.cfg[self.cfg.index("function Test-GooglePushConnection"):])
        fin = self.ui[self.ui.index("Get-GooglePushDecision $script:ProjectPath $ToolDir"):self.ui.index("Refresh-PartialLists\n        Show-Finished")]
        self.assertIn("MessageBoxDefaultButton]::Button2", fin)                              # по умолчанию «Нет»
        self.assertIn("Отправляю в Google: {0}", fin)                                        # куда уходит — в статус
        self.assertIn("Add-GooglePushLog", fin)
        self.assertNotIn("$hasPushCfg", self.ui)                                              # прежняя «отправка в общий cfg по умолчанию» убрана
        self.assertIn("$pushParams.ProjectConfigPath = $script:ProjectPath", self.ui)

    def test_push_script_and_logs_do_not_print_secret(self):
        for m in __import__("re").finditer(r"^.*Secret.*$", self.push, __import__("re").M):
            line = m.group(0)
            self.assertFalse(any(w in line for w in ("Write-Host", "Write-Output", "Write-Verbose", "throw", "Add-GooglePushLog")), line)
        for m in __import__("re").finditer(r"^.*(Add-GooglePushLog|\$ui\.Status\.Text).*$", self.ui, __import__("re").M):
            self.assertNotIn(".Secret", m.group(0))
            self.assertNotIn("tSec.Text", m.group(0))

    @unittest.skipUnless(shutil.which("git"), "нет git")
    def test_secret_files_are_git_ignored(self):
        def ignored(rel):
            return subprocess.run(["git", "check-ignore", "-q", rel], cwd=ROOT, capture_output=True).returncode == 0
        self.assertTrue(ignored("отчеты/проект/google_PROJ.cfg"))
        self.assertTrue(ignored("служебное/google_push/google_push.cfg"))
        self.assertFalse(ignored("служебное/google_push/google_push.cfg.example"))
        tracked = subprocess.run(["git", "ls-files", "--", "*google_*.cfg", "*/проект/*"], cwd=ROOT, capture_output=True, text=True).stdout
        self.assertEqual([l for l in tracked.splitlines() if l.endswith(".cfg")], [])         # ни один cfg с секретом не в индексе


if __name__ == "__main__":
    unittest.main()
