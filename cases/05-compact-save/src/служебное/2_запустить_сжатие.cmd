REM Name: 2_запустить_сжатие.cmd
REM Version: 1.0
REM What it does: Legacy launcher: open compact UI to run compact (step 2).
REM Inputs: None.
REM Outputs: ui_compact.ps1.
REM How to run: Double-click.
REM Notes: Operator flow documented in case README.
@echo off
chcp 65001 >nul
setlocal EnableExtensions
set "HERE=%~dp0..\"
echo %HERE%| findstr /b /c:"\\" >nul
if errorlevel 1 (
  pushd "%HERE%" >nul 2>&1
) else (
  rem UNC share - keep system cwd
)
powershell -NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File "%HERE%ui_compact.ps1"
set "ERR=%ERRORLEVEL%"
popd >nul 2>&1
exit /b %ERR%
