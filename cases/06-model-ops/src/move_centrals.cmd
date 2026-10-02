@echo off
chcp 65001 >nul
pushd "%~dp0" 2>nul
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0move_centrals.ps1"
set "ERR=%ERRORLEVEL%"
popd 2>nul
pause
exit /b %ERR%
