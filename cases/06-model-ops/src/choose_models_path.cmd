@echo off
chcp 65001 >nul
REM pushd works with UNC (maps a temp drive); cd /d does not
pushd "%~dp0" 2>nul
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0choose_models.ps1"
set "ERR=%ERRORLEVEL%"
popd 2>nul
pause
exit /b %ERR%
