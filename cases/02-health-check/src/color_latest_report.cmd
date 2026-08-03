@echo off
chcp 65001 >nul
pushd "%~dp0" 2>nul

if not exist "xlsx_out_path.cfg" (
  echo XLSX folder is not set. Choose local folder...
  powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0choose_xlsx_path.ps1"
  if errorlevel 1 (
    echo Cancelled.
    popd 2>nul
    pause
    exit /b 1
  )
)

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0csv_to_colored_xlsx.ps1"
set "ERR=%ERRORLEVEL%"
popd 2>nul
if "%ERR%"=="0" (
  echo XLSX rebuilt from latest health_*.csv / *_links.csv
) else (
  echo XLSX rebuild failed. See xlsx_export.log in folder from xlsx_out_path.cfg
)
pause
exit /b %ERR%
