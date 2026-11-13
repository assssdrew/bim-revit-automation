@echo off
rem Do not CD onto a UNC path: CMD warns and drops to Windows. On a share, pushd maps a drive letter; on a local disk, do not change directory.
rem Dump schedule filters. Drop .rvt files, a folder, or a .txt list (one path per line). See dump_schedule_filters.ps1
if "%~d0"=="\\" goto unc
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0dump_schedule_filters.ps1" %*
exit /b %ERRORLEVEL%
:unc
pushd "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0dump_schedule_filters.ps1" %*
set RC=%ERRORLEVEL%
popd
exit /b %RC%
