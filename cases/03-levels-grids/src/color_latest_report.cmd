@echo off
chcp 65001 >nul
pushd "%~dp0" 2>nul

set "RBP_ROOT=%USERPROFILE%\Documents\doc\script\rbp"
set "REPORTS=%RBP_ROOT%\reports"
if not exist "%REPORTS%" mkdir "%REPORTS%" 2>nul
> "xlsx_out_path.cfg" echo %REPORTS%

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0csv_to_colored_xlsx.ps1"
set "ERR=%ERRORLEVEL%"
popd 2>nul
if "%ERR%"=="0" (
  echo XLSX rebuilt from latest levels_grids_audit_*.csv
  echo Folder: %REPORTS%
) else (
  echo XLSX rebuild failed. See levels_grids_audit_xlsx_export.log in %REPORTS%
)
pause
exit /b %ERR%
