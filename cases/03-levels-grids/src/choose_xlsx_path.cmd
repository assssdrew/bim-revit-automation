REM Name: choose_xlsx_path.cmd
REM Version: 1.0
REM What it does: Launch Excel folder picker.
REM Inputs: None.
REM Outputs: choose_xlsx_path.ps1.
REM How to run: Double-click.
REM Notes: UNC-safe pushd.
@echo off
chcp 65001 >nul
pushd "%~dp0" 2>nul
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0choose_xlsx_path.ps1"
set "ERR=%ERRORLEVEL%"
popd 2>nul
pause
exit /b %ERR%
