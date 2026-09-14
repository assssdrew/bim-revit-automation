@echo off
chcp 65001 >nul
pushd "%~dp0" 2>nul
powershell -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0presets.ps1"
set "ERR=%ERRORLEVEL%"
popd 2>nul
exit /b %ERR%
