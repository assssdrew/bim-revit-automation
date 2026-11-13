@echo off
chcp 65001 >nul
pushd "%~dp0" 2>nul

REM Optional. Each new RBP run already starts a new health_check_*.csv on the first model.
REM This cmd creates the shared reports folder and can clear stuck markers.

set "RBP_ROOT=%USERPROFILE%\Documents\doc\script\rbp"
set "REPORTS=%RBP_ROOT%\reports"
if not exist "%REPORTS%" mkdir "%REPORTS%" 2>nul
if exist "%REPORTS%\_active_health_check.txt" del /f /q "%REPORTS%\_active_health_check.txt"
if exist "%REPORTS%\_active_health_check_links.txt" del /f /q "%REPORTS%\_active_health_check_links.txt"
> "xlsx_out_path.cfg" echo %REPORTS%

echo.
echo === Health Check folders ready ===
echo New RBP run = new report file automatically (old reports are kept).
echo Folder: %REPORTS%
echo Files:  health_check_YYYY-MM-DD_HHmm.csv / .xlsx
echo.
popd 2>nul
pause
exit /b 0
