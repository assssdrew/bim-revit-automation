# UTF-8 WITH BOM required (PowerShell 5.1).
# Чтение настроек Google и POST в Apps Script. Главный файл — у каждого проекта свой: отчеты\проект\google_<проект>.cfg.
# Общий google_push.cfg (служебное\google_push) — только запасной вариант для проекта без своего файла: читается, но код его
# не создаёт, не переписывает и не удаляет. google_<проект>.cfg пишется только из диалога «Подключить Google…» после «Сохранить».

function Read-GooglePushConfig([string]$Root) {
    $map = @{ WebAppUrl = ""; Secret = ""; ConfigPath = "" }
    $service = Split-Path -Parent $Root          # $Root — служебное\скрипты; google_push лежит в служебное\google_push
    foreach ($path in @(
        (Join-Path $service "google_push\google_push.cfg"),
        (Join-Path $Root "google_push.cfg")
    )) {
        if (-not (Test-Path -LiteralPath $path)) { continue }
        $map.ConfigPath = $path
        foreach ($line in @(Get-Content -LiteralPath $path -Encoding UTF8)) {
            $t = $line.Trim()
            if (-not $t -or $t.StartsWith("#")) { continue }
            if ($t -match '^(?i)webapp_url=(.*)$') { $map.WebAppUrl = $Matches[1].Trim() }
            elseif ($t -match '^(?i)secret=(.*)$') { $map.Secret = $Matches[1].Trim() }
        }
        if ($map.WebAppUrl) { break }
    }
    return $map
}

# ---------- адрес Google на проект ----------
# Адрес веб-приложения и секрет — для каждого проекта свои: отчеты\проект\google_<проект>.cfg (рядом с распределение_<проект>.txt, строки webapp_url= и secret=).
# Папка отчётов в git не попадает (.gitignore), в логи секрет не пишется. Общий google_push.cfg — запасной вариант: используется только после подтверждения пользователя.
function Get-ProjectGoogleConfigPath([string]$ProjectConfigPath) {
    if (-not $ProjectConfigPath) { return "" }
    $dir = Split-Path -Parent $ProjectConfigPath
    $leaf = Split-Path -Leaf $ProjectConfigPath
    if ($leaf -match '^распределение_(.+)\.txt$') { $leaf = "google_" + $Matches[1] + ".cfg" } else { $leaf = "google_" + [System.IO.Path]::GetFileNameWithoutExtension($leaf) + ".cfg" }
    return (Join-Path $dir $leaf)
}

function Get-ProjectNameFromConfigPath([string]$ProjectConfigPath) {
    $leaf = Split-Path -Leaf $ProjectConfigPath
    if ($leaf -match '^распределение_(.+)\.txt$') { return $Matches[1] }
    return [System.IO.Path]::GetFileNameWithoutExtension($leaf)
}

function Read-ProjectGoogleConfig([string]$Path) {
    $map = @{ WebAppUrl = ""; Secret = ""; ConfigPath = $Path }
    if (-not $Path -or -not (Test-Path -LiteralPath $Path)) { return $map }
    foreach ($line in @(Get-Content -LiteralPath $Path -Encoding UTF8)) {
        $t = $line.Trim()
        if (-not $t -or $t.StartsWith("#")) { continue }
        if ($t -match '^(?i)webapp_url=(.*)$') { $map.WebAppUrl = $Matches[1].Trim() }
        elseif ($t -match '^(?i)secret=(.*)$') { $map.Secret = $Matches[1].Trim() }
    }
    return $map
}

# Запись; Url/Secret = $null — оставить прежнее значение из файла. Оба пустые — файл не создаётся (и существующий не трогается)
function Write-ProjectGoogleConfig([string]$Path, $Url, $Secret) {
    if (-not $Path) { return }
    $cur = Read-ProjectGoogleConfig $Path
    $u = if ($null -eq $Url) { $cur.WebAppUrl } else { ([string]$Url).Trim() }
    $sec = if ($null -eq $Secret) { $cur.Secret } else { ([string]$Secret).Trim() }
    if (-not (Test-Path -LiteralPath $Path) -and -not $u -and -not $sec) { return }
    if ($u -eq $cur.WebAppUrl -and $sec -eq $cur.Secret -and (Test-Path -LiteralPath $Path)) { return }
    $dir = Split-Path -Parent $Path
    if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $tmp = $Path + ".tmp"
    $lines = @("# Google-таблица этого проекта. Файл с секретом: в git не хранится (папка отчётов в .gitignore).", ("webapp_url={0}" -f $u), ("secret={0}" -f $sec))
    [System.IO.File]::WriteAllLines($tmp, [string[]]$lines, (New-Object System.Text.UTF8Encoding $true))
    if (Test-Path -LiteralPath $Path) { [System.IO.File]::Delete($Path) }
    [System.IO.File]::Move($tmp, $Path)
}

# Конец адреса для показа («…ABCDEF/exec»): сам адрес целиком и секрет в статус не выводятся
function Get-GoogleUrlTail([string]$Url) {
    $t = ([string]$Url).Trim()
    if ($t.Length -le 24) { return $t }
    return ([string][char]0x2026 + $t.Substring($t.Length - 24))
}

# Куда уйдёт отправка. Kind: project — у проекта свой адрес; needs-confirm — своего адреса нет, есть общий google_push.cfg (нужен вопрос пользователю);
# shared — общий адрес подтверждён (AllowShared); none — адреса нет нигде. Label — текст для статуса и лога (проект + конец адреса, без секрета)
function Resolve-GooglePushTarget([string]$ProjectConfigPath, [string]$Root, [bool]$AllowShared = $false) {
    $name = Get-ProjectNameFromConfigPath $ProjectConfigPath
    $proj = Read-ProjectGoogleConfig (Get-ProjectGoogleConfigPath $ProjectConfigPath)
    if ($proj.WebAppUrl) {
        return [pscustomobject]@{ Kind = "project"; Cfg = $proj; ProjectName = $name; Label = ("проект {0}, таблица {1}" -f $name, (Get-GoogleUrlTail $proj.WebAppUrl)) }
    }
    $shared = Read-GooglePushConfig $Root
    if (-not $shared.WebAppUrl) {
        return [pscustomobject]@{ Kind = "none"; Cfg = $null; ProjectName = $name; Label = "" }
    }
    $label = ("проект {0}, ОБЩАЯ таблица из google_push.cfg {1}" -f $name, (Get-GoogleUrlTail $shared.WebAppUrl))
    # запасной вариант: предупреждение для журнала (без секрета и без полного адреса)
    $warn = ("ВНИМАНИЕ: у проекта {0} нет своего {1}, найден общий запасной google_push.cfg ({2})." -f $name, (Split-Path -Leaf (Get-ProjectGoogleConfigPath $ProjectConfigPath)), (Get-GoogleUrlTail $shared.WebAppUrl))
    if ($AllowShared) { return [pscustomobject]@{ Kind = "shared"; Cfg = $shared; ProjectName = $name; Label = $label; Warning = $warn } }
    return [pscustomobject]@{ Kind = "needs-confirm"; Cfg = $shared; ProjectName = $name; Label = $label; Warning = $warn }
}

# Проверка связи (ping + пустой POST), как проверить_google.ps1. Возвращает @{ Ok; Lines } — строки без секрета
function Test-GooglePushConnection($Cfg) {
    $lines = New-Object System.Collections.Generic.List[string]
    $issues = @(Test-BvsGooglePushConfigShape $Cfg)
    if ($issues.Count -gt 0) { foreach ($i in $issues) { $lines.Add("ОШИБКА: " + $i) }; return @{ Ok = $false; Lines = $lines.ToArray() } }
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        $ping = (Invoke-WebRequest -Uri ($Cfg.WebAppUrl.TrimEnd('/') + "?ping=1") -UseBasicParsing -TimeoutSec 60).Content.Trim()
        if ($ping -match '(?i)<\s*(!doctype|html)') { $ping = "html" }
        if ($ping -eq "ok") { $lines.Add("ping ok") }
        elseif ($ping -match "doGet") { $lines.Add("ВНИМАНИЕ: в развернутом скрипте нет doGet — вставьте Code.gs и сделайте «новая версия» развертывания.") }
        else { $lines.Add("ВНИМАНИЕ: ping не ok (не блокирует POST, если секрет верный).") }
    }
    catch { $lines.Add("GET ping не удался (не критично): " + $_.Exception.Message) }
    try {
        $body = @{ secret = $Cfg.Secret; hide = @(); sheets = @() } | ConvertTo-Json -Compress
        $text = Invoke-GoogleAppsScriptPost -Url $Cfg.WebAppUrl -Body ((New-Object System.Text.UTF8Encoding $false).GetBytes($body)) -Attempts 1
        $res = Get-GooglePushResult $text
        if ($res.Ok) { $lines.Add($(if ($res.Legacy) { "POST ok (старая версия Code.gs: обновите развёртывание)" } else { "POST ok: секрет принят, веб-приложение отвечает" })); return @{ Ok = $true; Lines = $lines.ToArray() } }
        if ($res.Message -match 'forbidden') { $lines.Add("ОШИБКА: forbidden — secret не совпадает с PUSH_SECRET в Code.gs.") } else { $lines.Add("ОШИБКА: " + $res.Message) }
    }
    catch { $lines.Add("ОШИБКА: " + $_.Exception.Message) }
    return @{ Ok = $false; Lines = $lines.ToArray() }
}

function Test-BvsGooglePushConfigShape {
    param($Cfg)
    $issues = New-Object System.Collections.Generic.List[string]
    if ([string]::IsNullOrWhiteSpace($Cfg.WebAppUrl)) {
        [void]$issues.Add("Не задан webapp_url.")
    }
    elseif ($Cfg.WebAppUrl -notmatch '(?i)/macros/s/') {
        [void]$issues.Add("webapp_url должен быть веб-приложением (…/macros/s/…/exec), не библиотекой.")
    }
    if ([string]::IsNullOrWhiteSpace($Cfg.Secret)) {
        [void]$issues.Add("Не задан secret.")
    }
    elseif ($Cfg.Secret -match '(?i)^(REPLACE|CHANGE)[_-]?ME') {
        [void]$issues.Add("secret не заменён: впишите свой PUSH_SECRET (тот же, что в Code.gs).")
    }
    elseif ($Cfg.Secret -match '^AKfycb') {
        [void]$issues.Add("secret похож на идентификатор развертывания. Укажите PUSH_SECRET из Code.gs.")
    }
    return $issues
}

# Разбор ответа Apps Script. Успех — только JSON {"ok":true,...} (новый Code.gs) или ровно "ok" (старый Code.gs).
# Всё остальное (forbidden, error:..., HTML-страница входа/ошибки, пустой ответ) — отказ.
function Get-GooglePushResult {
    param([string]$ResponseText)
    $t = ""
    if ($null -ne $ResponseText) { $t = $ResponseText.Trim() }
    if ($t -eq "") {
        return [pscustomobject]@{ Ok = $false; Legacy = $false; Message = "пустой ответ от Apps Script" }
    }
    if ($t.StartsWith("{")) {
        $o = $null
        try { $o = $t | ConvertFrom-Json } catch { $o = $null }
        if ($null -eq $o) {
            return [pscustomobject]@{ Ok = $false; Legacy = $false; Message = ("ответ не разобран как JSON: " + $t.Substring(0, [Math]::Min(300, $t.Length))) }
        }
        if ($o.ok -eq $true) {
            return [pscustomobject]@{ Ok = $true; Legacy = $false; Message = $t }
        }
        $err = [string]$o.error
        if (-not $err) { $err = $t }
        return [pscustomobject]@{ Ok = $false; Legacy = $false; Message = $err }
    }
    if ($t -match '(?i)^ok$') {
        return [pscustomobject]@{ Ok = $true; Legacy = $true; Message = "ok" }
    }
    if ($t -match '(?i)<\s*(!doctype|html)') {
        $plain = ($t -replace '(?s)<(script|style).*?</\1>', ' ' -replace '<[^>]+>', ' ' -replace '\s+', ' ').Trim()
        if ($plain.Length -gt 300) { $plain = $plain.Substring(0, 300) }
        return [pscustomobject]@{ Ok = $false; Legacy = $false; Message = ("вместо ответа скрипта пришла HTML-страница (нет доступа к развёртыванию, неверный URL или ошибка Apps Script): " + $plain) }
    }
    return [pscustomobject]@{ Ok = $false; Legacy = $false; Message = $t.Substring(0, [Math]::Min(300, $t.Length)) }
}

function Test-GooglePushSucceeded {
    param([string]$ResponseText)
    return [bool](Get-GooglePushResult $ResponseText).Ok
}

# POST JSON в Apps Script. Таймаут 5 минут (у WebClient он жёстко 100 с, а Apps Script с разметкой работает дольше).
# Повтор безопасен: Code.gs каждый раз полностью пересоздаёт листы (идемпотентно) и держит блокировку.
function Invoke-GoogleAppsScriptPost {
    param(
        [string]$Url,
        [byte[]]$Body,
        [int]$TimeoutSec = 300,
        [int]$Attempts = 3
    )
    $lastError = $null
    for ($i = 1; $i -le $Attempts; $i++) {
        $resp = $null
        try {
            $req = [System.Net.HttpWebRequest]::Create($Url)
            $req.Method = "POST"
            $req.ContentType = "application/json; charset=utf-8"
            $req.ContentLength = $Body.Length
            $req.Timeout = $TimeoutSec * 1000
            $req.ReadWriteTimeout = $TimeoutSec * 1000
            $req.AllowAutoRedirect = $true
            $rs = $req.GetRequestStream()
            try { $rs.Write($Body, 0, $Body.Length) } finally { $rs.Dispose() }
            $resp = $req.GetResponse()
            $sr = New-Object System.IO.StreamReader($resp.GetResponseStream(), (New-Object System.Text.UTF8Encoding $false))
            try { return $sr.ReadToEnd() } finally { $sr.Dispose() }
        }
        catch [System.Net.WebException] {
            $we = $_.Exception
            $code = 0
            $text = ""
            if ($we.Response) {
                try { $code = [int]$we.Response.StatusCode } catch { }
                try {
                    $er = New-Object System.IO.StreamReader($we.Response.GetResponseStream(), (New-Object System.Text.UTF8Encoding $false))
                    $text = $er.ReadToEnd()
                    $er.Dispose()
                }
                catch { }
            }
            $lastError = ("HTTP {0}: {1} {2}" -f $code, $we.Message, $text.Substring(0, [Math]::Min(200, $text.Length)))
            # 4xx (кроме 408/429) повторять бессмысленно
            if ($code -ge 400 -and $code -lt 500 -and $code -ne 408 -and $code -ne 429) { throw $lastError }
        }
        finally {
            if ($resp) { $resp.Close() }
        }
        if ($i -lt $Attempts) { Start-Sleep -Seconds (5 * $i * $i) }
    }
    throw ("нет ответа от Apps Script после {0} попыток. {1}" -f $Attempts, $lastError)
}
