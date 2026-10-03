REM Name: patch_rbp_open_failures.cmd
REM Version: 1.0
REM What it does: Wrapper to patch RBP open-failure handling.
REM Inputs: None.
REM Outputs: patch_rbp_open_failures.ps1.
REM How to run: Pause then PowerShell.
REM Notes: Close RBP first.
@echo off
chcp 65001 >nul
echo.
echo Patch RBP so Open is not canceled on Cyrillic Revit warnings.
echo Close Revit Batch Processor before continuing.
echo.
pause
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0patch_rbp_open_failures.ps1"
echo.
pause
exit /b %ERRORLEVEL%
