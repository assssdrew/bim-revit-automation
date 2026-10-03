REM Name: choose_accuracy.cmd
REM Version: 1.0
REM What it does: Prompt for Length rounding accuracy (mm) and write accuracy.cfg.
REM Inputs: Interactive numeric value from the operator.
REM Outputs: accuracy.cfg in the toolkit folder.
REM How to run: Double-click; enter value at the prompt.
REM Notes: UNC-safe pushd; cancel leaves cfg unchanged if input empty.
@echo off
setlocal EnableExtensions
REM IMPORTANT: UNC cannot be current directory for cmd.exe, use pushd.
pushd "%~dp0" 2>nul
if errorlevel 1 (
  echo ERROR: Cannot open script folder.
  echo Path: %~dp0
  pause
  exit /b 1
)

echo.
echo ============================================
echo   Revit Length Accuracy (mm)
echo ============================================
echo.
echo   Examples: 0.1   1.0   0.01   0.000001
echo.

set "VAL="
set /p VAL=Enter value: 
if not defined VAL (
  echo Cancelled.
  popd 2>nul
  pause
  exit /b 0
)

> "%CD%\accuracy.cfg" echo # Length rounding accuracy in millimeters
>> "%CD%\accuracy.cfg" echo # Set by choose_accuracy.cmd
>> "%CD%\accuracy.cfg" echo %VAL%

echo.
echo Saved to: "%CD%\accuracy.cfg"
echo Value: %VAL%
echo.
popd 2>nul
pause
exit /b 0
