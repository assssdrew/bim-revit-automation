REM Name: разрешить_надстройку_BatchRvt.cmd
REM Version: 1.0
REM What it does: Trust BatchRvt add-in for installed Revit years (registry + optional UI wait).
REM Inputs: None.
REM Outputs: HKCU CodeSigning entries; operator may confirm Revit load dialog.
REM How to run: Double-click.
REM Notes: Calls разрешить_надстройку_BatchRvt.ps1 from the same folder.
@echo off
chcp 65001 >nul
setlocal EnableExtensions
set "HERE=%~dp0"

echo.
echo ============================================
echo   Разрешить надстройку BatchRvt (Revit)
echo ============================================
echo.
echo 1) Пишем доверие в реестр (все годы Revit на этом ПК).
echo 2) 3 минуты ждём окно «Всегда загружать» и нажимаем сами.
echo    RBP на русской Revit это окно не нажимает.
echo.
echo Закройте Revit полностью.
echo После «реестр готов» СРАЗУ запустите Revit Batch Processor.
echo.

set "REVIT_YEAR="
set /p REVIT_YEAR=Версия Revit [Enter = все]: 

echo.
echo --- Реестр ---
if "%REVIT_YEAR%"=="" (
  powershell -NoProfile -ExecutionPolicy Bypass -File "%HERE%разрешить_надстройку_BatchRvt.ps1"
) else (
  powershell -NoProfile -ExecutionPolicy Bypass -File "%HERE%разрешить_надстройку_BatchRvt.ps1" -Year %REVIT_YEAR%
)
if errorlevel 1 (
  echo Реестр: ничего не записано. Проверьте, что RBP установлен.
  pause
  exit /b 1
)

echo.
echo --- Автоклик 3 минуты ---
echo СЕЙЧАС откройте Revit Batch Processor и нажмите Start.
echo Это окно не закрывайте.
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%HERE%разрешить_надстройку_BatchRvt.ps1" -Watch -Seconds 180

echo.
pause
exit /b %ERRORLEVEL%
