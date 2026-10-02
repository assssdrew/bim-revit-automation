@echo off
chcp 65001 >nul
pushd "%~dp0" 2>nul

if not exist "reports" mkdir "reports"
if not exist "reports\cvc" mkdir "reports\cvc"
if exist "reports\cvc\_active_report.txt" del /f /q "reports\cvc\_active_report.txt"
if exist "reports\cvc\_active_links.txt" del /f /q "reports\cvc\_active_links.txt"

echo.
echo === New Health Check session ===
echo 1) Choose LOCAL folder for Excel report...
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0choose_xlsx_path.ps1"
if errorlevel 1 (
  echo.
  echo ERROR: XLSX folder was not selected. Session not started.
  popd 2>nul
  pause
  exit /b 1
)

echo.
echo Session ready.
echo - CSV will be written to: reports\cvc\health_YYYY-MM-DD_HHmm.csv
echo - Links detail CSV:       reports\cvc\health_YYYY-MM-DD_HHmm_links.csv
echo - XLSX will be saved to the folder from xlsx_out_path.cfg
echo.
type "%~dp0xlsx_out_path.cfg" 2>nul
echo.
popd 2>nul
pause
exit /b 0
