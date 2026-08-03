@echo off
chcp 65001 >nul
setlocal EnableExtensions

REM Убрать повторяющееся окно "Надстройка без подписи" (BatchRvtAddin).
REM Запустить ОДИН РАЗ на ПК Revit под тем же пользователем, что и RBP/Revit.

echo.
echo ============================================
echo   Разрешить надстройку BatchRvt (Revit)
echo ============================================
echo.
echo Если окно "Всегда загружать" не помогает - этот скрипт
echo добавляет доверие в реестр Windows (CodeSigning).
echo.

set /p REVIT_YEAR=Версия Revit (2021/2022/2023/2024) [Enter = 2024]: 
if "%REVIT_YEAR%"=="" set "REVIT_YEAR=2024"

echo.
echo Ищем BatchRvt .addin в Addins\%REVIT_YEAR% ...
dir /s /b "%APPDATA%\Autodesk\Revit\Addins\%REVIT_YEAR%\*BatchRvt*.addin" 2>nul
echo.

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$year = '%REVIT_YEAR%';" ^
  "$addinsRoot = Join-Path $env:APPDATA ('Autodesk\Revit\Addins\' + $year);" ^
  "$files = Get-ChildItem -LiteralPath $addinsRoot -Recurse -Filter *.addin -ErrorAction SilentlyContinue | Where-Object { $_.Name -match 'BatchRvt|BatchRvtAddin' };" ^
  "if (-not $files) { Write-Host ('ОШИБКА: BatchRvt .addin не найден в ' + $addinsRoot); exit 1 };" ^
  "$guid = $null;" ^
  "foreach ($f in $files) {" ^
  "  $txt = Get-Content -LiteralPath $f.FullName -Raw;" ^
  "  if ($txt -match '<AddInId>\s*\{?([0-9A-Fa-f\-]{36})\}?\s*</AddInId>') { $guid = $Matches[1].ToUpper(); Write-Host ('Найден AddInId: ' + $guid); Write-Host ('Файл: ' + $f.FullName); break }" ^
  "};" ^
  "if (-not $guid) { Write-Host 'ОШИБКА: AddInId не найден в .addin'; exit 1 };" ^
  "$key = 'HKCU:\SOFTWARE\Autodesk\Revit\Autodesk Revit ' + $year + '\CodeSigning';" ^
  "if (-not (Test-Path $key)) { New-Item -Path $key -Force | Out-Null };" ^
  "New-ItemProperty -Path $key -Name $guid -PropertyType DWord -Value 1 -Force | Out-Null;" ^
  "Write-Host '';" ^
  "Write-Host ('Готово. Реестр: ' + $key);" ^
  "Write-Host ('DWORD ' + $guid + ' = 1');" ^
  "Write-Host 'Полностью закройте Revit и запустите снова.'"

echo.
pause
