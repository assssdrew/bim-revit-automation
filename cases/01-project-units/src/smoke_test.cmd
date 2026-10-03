REM Name: smoke_test.cmd
REM Version: 1.0
REM What it does: Verify case 01 toolkit files and default accuracy.cfg without opening Revit.
REM Inputs: Toolkit folder layout under src\.
REM Outputs: Console PASS/FAIL and optional default accuracy.cfg creation.
REM How to run: Double-click before first RBP run on a new copy.
REM Notes: Does not exercise Revit API or RBP.
@echo off
setlocal EnableExtensions
pushd "%~dp0" 2>nul
if errorlevel 1 (
  echo ERROR: Cannot open script folder.
  pause
  exit /b 1
)

echo.
echo ============================================
echo   Batch Set Project Units - Smoke Test
echo ============================================
echo.
echo This test checks only helper scripts/config, not Revit API changes.
echo.

set "FAIL=0"

if not exist "%CD%\choose_models.ps1" (
  echo [FAIL] choose_models.ps1 is missing
  set "FAIL=1"
) else (
  echo [ OK ] choose_models.ps1
)

if not exist "%CD%\set_length_accuracy.py" (
  echo [FAIL] set_length_accuracy.py is missing
  set "FAIL=1"
) else (
  echo [ OK ] set_length_accuracy.py
)

if not exist "%CD%\rbp_rsn_patch\apply_rsn_support.ps1" (
  echo [FAIL] rbp_rsn_patch\apply_rsn_support.ps1 is missing
  set "FAIL=1"
) else (
  echo [ OK ] rbp_rsn_patch\apply_rsn_support.ps1
)

if not exist "%CD%\servers.cfg" (
  echo [FAIL] servers.cfg is missing
  set "FAIL=1"
) else (
  echo [ OK ] servers.cfg
)

if not exist "%CD%\accuracy.cfg" (
  echo [WARN] accuracy.cfg missing - creating default 0.1
  > "%CD%\accuracy.cfg" echo # Length rounding accuracy in millimeters
  >> "%CD%\accuracy.cfg" echo # Set by choose_accuracy.cmd
  >> "%CD%\accuracy.cfg" echo 0.1
)

if exist "%CD%\accuracy.cfg" (
  echo [ OK ] accuracy.cfg
  echo.
  echo accuracy.cfg:
  type "%CD%\accuracy.cfg"
)

echo.
if "%FAIL%"=="0" (
  echo RESULT: PASS
  echo Next: run choose_models_path.cmd and process 1 test model in RBP.
) else (
  echo RESULT: FAIL
)
echo.
popd 2>nul
pause
exit /b %FAIL%
