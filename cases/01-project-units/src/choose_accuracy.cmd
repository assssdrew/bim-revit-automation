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
