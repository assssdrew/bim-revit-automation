REM Name: install_RSN_support_in_RBP.cmd
REM Version: 1.0
REM What it does: Install RSN:// (Revit Server) support patch into the local RBP scripts folder.
REM Inputs: Installed Revit Batch Processor; optional script path args forwarded to apply_rsn_support.ps1.
REM Outputs: Patched RBP helper scripts on disk.
REM How to run: Double-click from case src\; close RBP before patching if it is open.
REM Notes: Delegates to rbp_rsn_patch\apply_rsn_support.ps1.
@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

echo.
echo ============================================================
echo  Patch Revit Batch Processor for RSN:// (Revit Server)
echo  Must run on the PC where RBP is installed (the RBP workstation)
echo ============================================================
echo.

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0rbp_rsn_patch\apply_rsn_support.ps1" %*
set ERR=%ERRORLEVEL%
echo.
pause
exit /b %ERR%
