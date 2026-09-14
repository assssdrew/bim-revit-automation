@echo off
chcp 65001 >nul
setlocal EnableExtensions
set "HERE=%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%HERE%apply_rsn_support.ps1" %*
set "ERR=%ERRORLEVEL%"
echo.
pause
exit /b %ERR%
