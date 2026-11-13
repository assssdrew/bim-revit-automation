# UTF-8 WITH BOM required (PowerShell 5.1).
# Сравнение двух выгрузок фильтров спецификаций: «было» и «стало». Спеки сопоставляются по номеру и имени без префикса «П»;
# на выходе таблица (CSV для Excel): спека, что изменилось (добавлено / удалено / изменено), было, стало, и отдельно спеки,
# которые есть только в одном файле. Порядок фильтров внутри спеки не сравнивается.
#
# Запуск: перетащите на «Сравнить фильтры спек.bat» два файла (сначала «было», потом «стало»), либо
#   powershell -File compare_filter_dumps.ps1 было.txt стало.txt [-OutCsv результат.csv] [-Swap]

param(
    [Parameter(Position = 0)][string]$Old = "",
    [Parameter(Position = 1)][string]$New = "",
    [string]$OutCsv = "",
    [switch]$Swap,
    [switch]$NoPause
)

$ErrorActionPreference = "Stop"
. (Join-Path $PSScriptRoot "compare_filter_lib.ps1")
$ServiceDir = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot "..\.."))  # служебное (скрипт лежит в служебное\разработка\инструменты)
$ToolDir = Join-Path $ServiceDir "скрипты"          # служебное\скрипты: общие скрипты окна
. (Join-Path $ToolDir "report_layout.ps1")

function Stop-Here([string]$Message, [int]$Code = 1) {
    Write-Host $Message -ForegroundColor Red
    if (-not $NoPause) { [void](Read-Host "Нажмите Enter") }
    exit $Code
}

if (-not $Old) { $Old = (Read-Host "Файл «было» (выгрузка фильтров)").Trim('"') }
if (-not $New) { $New = (Read-Host "Файл «стало» (выгрузка фильтров)").Trim('"') }
if ($Swap) { $t = $Old; $Old = $New; $New = $t }
foreach ($p in @($Old, $New)) { if (-not (Test-Path -LiteralPath $p)) { Stop-Here ("Нет файла: " + $p) } }
if (-not $OutCsv) { $OutCsv = Join-Path (Initialize-ReportLayout (Get-ReportRootDir $ToolDir)).Layout.Ops ("сравнение_фильтров_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".csv") }

Write-Host ""
Write-Host ("Было : {0}" -f $Old)
Write-Host ("Стало: {0}" -f $New)
$oldSpecs = @(Read-FilterDump $Old)
$newSpecs = @(Read-FilterDump $New)
$res = Compare-FilterDumps $oldSpecs $newSpecs
Write-CompareCsv $OutCsv $res.Rows

Write-Host ""
Write-Host ("Спецификаций: было {0}, стало {1}; сопоставлено {2}, из них без изменений {3}, с изменениями {4}" -f $oldSpecs.Count, $newSpecs.Count, $res.Matched, $res.Unchanged, ($res.Matched - $res.Unchanged)) -ForegroundColor Cyan
Write-Host ("Только в «было»: {0}; только в «стало»: {1}" -f $res.OnlyOld, $res.OnlyNew) -ForegroundColor Cyan
$show = @($res.Rows | Select-Object -First 60)
foreach ($r in $show) {
    $color = "Gray"
    if ($r.Change -like "спека только*") { $color = "Yellow" } elseif ($r.Change -eq "добавлено") { $color = "Green" } elseif ($r.Change -eq "удалено") { $color = "Red" }
    Write-Host ("{0,-8} {1}" -f $r.Number, $r.Spec) -ForegroundColor $color
    Write-Host ("         {0}: {1}  ->  {2}" -f $r.Change, $r.Was, $r.Became) -ForegroundColor $color
}
if ($res.Rows.Count -gt $show.Count) { Write-Host ("... и ещё {0} строк (все — в CSV)" -f ($res.Rows.Count - $show.Count)) }
if ($res.Rows.Count -eq 0) { Write-Host "Различий нет." -ForegroundColor Green }
Write-Host ""
Write-Host ("Таблица: {0}" -f $OutCsv) -ForegroundColor Green
if (-not $NoPause) { [void](Read-Host "Нажмите Enter") }
