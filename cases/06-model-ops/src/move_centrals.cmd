REM Name: move_centrals.cmd
REM Version: 1.0
REM What it does: Wrapper for move_centrals.ps1.
REM Inputs: None.
REM Outputs: Moves per job_paths.csv.
REM How to run: Double-click.
REM Notes: UNC-safe pushd.
@echo off
chcp 65001 >nul
pushd "%~dp0" 2>nul
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0move_centrals.ps1"
set "ERR=%ERRORLEVEL%"
popd 2>nul
pause
exit /b %ERR%
