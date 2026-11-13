@echo off
chcp 65001 >nul
pushd "%~dp0" 2>nul

set "RBP_ROOT=%USERPROFILE%\Documents\doc\script\rbp"
set "REPORTS=%RBP_ROOT%\reports"
if not exist "%REPORTS%" mkdir "%REPORTS%" 2>nul
if exist "%REPORTS%\_active_levels_grids_audit.txt" del /f /q "%REPORTS%\_active_levels_grids_audit.txt"
if exist "%REPORTS%\_active_levels_grids_audit_details.txt" del /f /q "%REPORTS%\_active_levels_grids_audit_details.txt"
> "xlsx_out_path.cfg" echo %REPORTS%

echo.
echo === New Levels and Grids session ===
echo 2) Exemplar mode = AUTO
echo    Script picks AR/BF link inside each host by building code (K1, K2, ...).
> "%~dp0exemplar_model.cfg" echo AUTO
if exist "%~dp0exemplar_link.cfg" del /f /q "%~dp0exemplar_link.cfg"

echo.
echo Session ready.
echo - Host models: run_cascade.cmd or choose_models.ps1 -^> rvt_list.txt
echo - Exemplar:    AUTO (exemplar_model.cfg)
echo - Reports:     %REPORTS%\levels_grids_audit_YYYY-MM-DD_HHmm.csv / .xlsx
echo.
echo Exemplar:
type "%~dp0exemplar_model.cfg" 2>nul
echo.
popd 2>nul
pause
exit /b 0
