REM Name: choose_models_path.cmd
REM Version: 1.0
REM What it does: Launch the case 02 model picker.
REM Inputs: None.
REM Outputs: Invokes choose_models.ps1.
REM How to run: Double-click.
REM Notes: UNC-safe pushd.
@echo off
chcp 65001 >nul
REM pushd works with UNC (maps a temp drive); cd /d does not
pushd "%~dp0" 2>nul
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0choose_models.ps1"
set "ERR=%ERRORLEVEL%"
popd 2>nul
pause
exit /b %ERR%
