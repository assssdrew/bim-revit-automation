REM Name: choose_models_path.cmd
REM Version: 1.0
REM What it does: Launcher for the case 01 model picker (UNC-safe pushd).
REM Inputs: None.
REM Outputs: Runs choose_models.ps1; exit code from PowerShell.
REM How to run: Double-click or run from cmd in src\.
REM Notes: Uses pushd so UNC script paths work.
@echo off
chcp 65001 >nul
REM pushd works with UNC (maps a temp drive); cd /d does not
pushd "%~dp0" 2>nul
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0choose_models.ps1"
set "ERR=%ERRORLEVEL%"
popd 2>nul
pause
exit /b %ERR%
