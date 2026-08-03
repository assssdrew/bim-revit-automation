@echo off
chcp 65001 >nul
pushd "%~dp0" 2>nul

if not exist "reports" mkdir "reports"
if not exist "reports\cvc" mkdir "reports\cvc"
if exist "reports\cvc\_active_report.txt" del /f /q "reports\cvc\_active_report.txt"
if exist "reports\cvc\_active_details.txt" del /f /q "reports\cvc\_active_details.txt"

echo.
echo === New Levels and Grids session ===
echo 1) Choose LOCAL folder for colored XLSX...
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0choose_xlsx_path.ps1"
if errorlevel 1 (
  echo.
  echo ERROR: XLSX folder was not selected. Session not started.
  popd 2>nul
  pause
  exit /b 1
)

echo.
echo 2) Choose EXEMPLAR model (same picker as models: file or Revit Server)...
echo    BF vs AR  -^> pick AR
echo    Sections vs BF -^> pick BF
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0choose_exemplar_link.ps1"
if errorlevel 1 (
  echo.
  echo WARN: exemplar not selected. Script will use first loaded RVT link.
)

echo.
echo Session ready.
echo - Host models: choose_models_path.cmd -^> rvt_list.txt
echo - Exemplar:    exemplar_model.cfg
echo - CSV: reports\cvc\lg_*.csv
echo - XLSX folder:
type "%~dp0xlsx_out_path.cfg" 2>nul
echo.
echo Exemplar:
type "%~dp0exemplar_model.cfg" 2>nul
echo.
popd 2>nul
pause
exit /b 0
