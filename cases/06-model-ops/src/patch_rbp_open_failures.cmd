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
