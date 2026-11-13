# Юнит-проверки PowerShell-функций цепочки xlsx -> Google. Запуск (pwsh или Windows PowerShell):
#   powershell -NoProfile -File служебное\разработка\tests\ps_unit.ps1
$ErrorActionPreference = "Stop"
$tool = Join-Path (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)) "скрипты"     # служебное\разработка\tests -> служебное\скрипты
$script:BvsAssignmentDir = $tool
. (Join-Path $tool "google_push_cfg.ps1")
. (Join-Path $tool "google_payload_lib.ps1")
. (Join-Path $tool "assignment_xlsx.ps1")
. (Join-Path $tool "assignment_export_stream.ps1")

$script:failed = 0
function Check([string]$Name, [bool]$Cond, [string]$Info = "") {
    if ($Cond) { Write-Output ("ok   - " + $Name) }
    else { $script:failed++; Write-Output ("FAIL - " + $Name + " " + $Info) }
}

# --- ответ Apps Script ---
$r = Get-GooglePushResult '{"ok":true,"version":3,"sheets":[]}'
Check "JSON ok:true -> успех" ($r.Ok -eq $true -and -not $r.Legacy)
$r = Get-GooglePushResult 'ok'
Check "старый ответ 'ok' -> успех (legacy)" ($r.Ok -eq $true -and $r.Legacy)
$r = Get-GooglePushResult '{"ok":false,"error":"row group depth cannot exceed 8"}'
Check "JSON ok:false -> отказ с текстом ошибки" ($r.Ok -eq $false -and $r.Message -match 'depth')
$r = Get-GooglePushResult 'forbidden'
Check "forbidden -> отказ" ($r.Ok -eq $false)
$r = Get-GooglePushResult 'error: Exception: boom'
Check "error: ... -> отказ" ($r.Ok -eq $false)
$r = Get-GooglePushResult '<!DOCTYPE html><html><body><h1>Authorization needed</h1><script>x()</script></body></html>'
Check "HTML-страница -> отказ" ($r.Ok -eq $false -and $r.Message -match 'HTML' -and $r.Message -match 'Authorization')
$r = Get-GooglePushResult "OK but HTML <html>"
Check "'OK ...' не принимается как успех" ($r.Ok -eq $false)
$r = Get-GooglePushResult ''
Check "пустой ответ -> отказ" ($r.Ok -eq $false)
Check "Test-GooglePushSucceeded обёртка" ((Test-GooglePushSucceeded 'ok') -and -not (Test-GooglePushSucceeded '<html>'))

# --- конфиг ---
$bad = @(Test-BvsGooglePushConfigShape @{ WebAppUrl = "https://script.google.com/macros/s/AKfycbX/exec"; Secret = "REPLACE_ME_WITH_LONG_RANDOM_SECRET" })
Check "secret-заглушка отклоняется" ($bad.Count -ge 1)
$good = @(Test-BvsGooglePushConfigShape @{ WebAppUrl = "https://script.google.com/macros/s/AKfycbX/exec"; Secret = "long-random-1" })
Check "нормальный конфиг принимается" ($good.Count -eq 0)

# --- цвета Excel (BGR-число) ---
Check "Convert-BgrToRgbHex красный" ((Convert-BgrToRgbHex 0x0000FF) -eq "FF0000")
Check "Convert-BgrToRgbHex жёлтый" ((Convert-BgrToRgbHex 0x00FFFF) -eq "FFFF00")
Check "Convert-BgrToRgbHex оранжевый" ((Convert-BgrToRgbHex 0x80C0FF) -eq "FFC080")

# --- строки payload: одна строка не разворачивается ---
$one = [pscustomobject]@{ name = "Сводка"; rows = @(, @("Только заметка")); formats = @(); columnWidths = @(48); meta = $null }
$n = Normalize-PayloadSheet $one
$json = $n | ConvertTo-Json -Depth 6 -Compress
Check "одна строка остаётся массивом массивов" ($json -match '"rows":\[\["Только заметка"\]\]') $json
$parsed = '{"name":"Сводка","rows":[["a","b"],["c"]]}' | ConvertFrom-Json
$n = Normalize-PayloadSheet $parsed
Check "строки выравниваются до прямоугольника" (@($n.rows).Count -eq 2 -and @($n.rows[1]).Count -eq 2)

# --- значения ячеек xlsx ---
Ensure-OpenXmlLoaded
function Cell-For([string]$Text) {
    $c = New-Object DocumentFormat.OpenXml.Spreadsheet.Cell
    Set-XlsxCellValue -Cell $c -Text $Text
    return ,$c
}
$c = Cell-For "007"
Check "'007' остаётся текстом" ($c.DataType.Value -eq [DocumentFormat.OpenXml.Spreadsheet.CellValues]::InlineString)
$c = Cell-For "42"
Check "'42' — число" ($c.DataType.Value -eq [DocumentFormat.OpenXml.Spreadsheet.CellValues]::Number -and $c.CellValue.Text -eq "42")
$c = Cell-For ("a" + [char]1 + "b")
Check "управляющий символ вырезается" ($c.InlineString.Text.Text -eq "ab")

# --- порядок спецификаций (номер в начале имени как версия) ---
$names = @("П        99_Толщина стенки_Без значения", "П        14.7_Длина_Лотки_Без значения", "П        13_Материал_Фитинги_Без значния",
    "П        4.4_Имя системы_Оборудование_Без значения", "П        4.4_Имя системы_Оборудование_Пустые", "П        5.7_Ширина_Без значения",
    "П        1_Классификатор_Пустые", "П        1_Классификатор_999", "П        17.1_Секция_Оборудование_Пустые", "П        17.0_Секция_Кириллица",
    "П        16.1_Корпус_Оборудование_Пустые", "П        16.0_Корпус_Кириллица", "Без номера", "П        0_Ошибка типоразмера")
$sorted = @(Sort-CheckNames $names)
$expect = @("П        0_Ошибка типоразмера", "П        1_Классификатор_Пустые", "П        1_Классификатор_999",
    "П        4.4_Имя системы_Оборудование_Пустые", "П        4.4_Имя системы_Оборудование_Без значения", "П        5.7_Ширина_Без значения",
    "П        13_Материал_Фитинги_Без значния", "П        14.7_Длина_Лотки_Без значения", "П        16.0_Корпус_Кириллица",
    "П        16.1_Корпус_Оборудование_Пустые", "П        17.0_Секция_Кириллица", "П        17.1_Секция_Оборудование_Пустые",
    "П        99_Толщина стенки_Без значения", "Без номера")
Check "числовой порядок: 0, затем Пустые и сразу 999, 16.0 раньше 16.1, 99 в конце" (($sorted -join "|") -eq ($expect -join "|")) ($sorted -join "|")
Check "Sort-CheckNames одного имени -> массив" (@(Sort-CheckNames @("П 1_А_Пустые")).Count -eq 1)
$p1 = Get-CheckNameParts "П        4.4_Имя системы_Оборудование_Пустые"
$p2 = Get-CheckNameParts "П        4.4_Имя системы_Оборудование_Без значния"
Check "пара Пустые/Без значения: одинаковая база" ($p1.Base -eq $p2.Base -and $p1.Kind -eq "empty" -and $p2.Kind -eq "novalue")

if ($script:failed -gt 0) { Write-Output ("FAILED: " + $script:failed); exit 1 }
Write-Output "ALL OK"
