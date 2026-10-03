REM Name: choose_exemplar_link.cmd
REM Version: 1.0
REM What it does: Launch exemplar model picker.
REM Inputs: None.
REM Outputs: choose_exemplar_link.ps1.
REM How to run: Double-click.
REM Notes: UNC-safe pushd.
@echo off
chcp 65001 >nul
pushd "%~dp0" 2>nul
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0choose_exemplar_link.ps1"
set "ERR=%ERRORLEVEL%"
popd 2>nul
pause
exit /b %ERR%
