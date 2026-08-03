@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

echo.
echo ============================================================
echo  Patch Revit Batch Processor for RSN:// (Revit Server)
echo  Must run on the PC where RBP is installed (e.g. potapov)
echo ============================================================
echo.

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0rbp_rsn_patch\apply_rsn_support.ps1" %*
set ERR=%ERRORLEVEL%
echo.
pause
exit /b %ERR%
