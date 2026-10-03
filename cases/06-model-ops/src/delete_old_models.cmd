REM Name: delete_old_models.cmd
REM Version: 1.0
REM What it does: Wrapper for delete_old_models.ps1 with safety echo.
REM Inputs: None.
REM Outputs: Old model cleanup.
REM How to run: After Pass 2 success.
REM Notes: Explains RSN manual step.
@echo off
chcp 65001 >nul
pushd "%~dp0" 2>nul
echo After Pass2 success only.
echo Local/UNC old files will be deleted if new twin exists.
echo RSN paths go to to_delete_rsn.txt for manual Admin delete.
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0delete_old_models.ps1"
set "ERR=%ERRORLEVEL%"
popd 2>nul
pause
exit /b %ERR%
