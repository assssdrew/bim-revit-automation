REM Name: установить_поддержку_RSN_в_RBP.cmd
REM Version: 1.0
REM What it does: Wrapper to apply RSN patch from compact-save toolkit.
REM Inputs: None.
REM Outputs: apply_rsn_support.ps1.
REM How to run: Double-click.
REM Notes: Run on the PC where Revit Batch Processor is installed.
@echo off
chcp 65001 >nul
setlocal EnableExtensions
set "HERE=%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%HERE%apply_rsn_support.ps1" %*
set "ERR=%ERRORLEVEL%"
echo.
pause
exit /b %ERR%
