REM Name: choose_models_path.cmd
REM Version: 1.0
REM What it does: Open compact UI from the служебное subfolder (parent src as toolkit root).
REM Inputs: None.
REM Outputs: ui_compact.ps1.
REM How to run: Double-click.
REM Notes: Handles UNC parent path via pushd when possible.
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
