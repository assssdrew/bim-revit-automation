@echo off
REM Compact save without the window. Parameters: CLI.txt in this folder.
REM Example: run_list.cmd -Mode deep -ListPath C:\lists\models.txt
set "HERE=%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%HERE%..\start_compact.ps1" -NonInteractive %*
exit /b %ERRORLEVEL%
