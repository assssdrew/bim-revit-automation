REM Name: presets.cmd
REM Version: 1.0
REM What it does: Console launcher for presets.ps1 (STA).
REM Inputs: None.
REM Outputs: Model ops UI.
REM How to run: run_cascade menu option 1 or direct.
REM Notes: UNC-safe pushd.
@echo off
chcp 65001 >nul
pushd "%~dp0" 2>nul
powershell -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0presets.ps1"
set "ERR=%ERRORLEVEL%"
popd 2>nul
exit /b %ERR%
