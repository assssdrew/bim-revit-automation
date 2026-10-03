REM Name: reset_report_session.cmd
REM Version: 1.0
REM What it does: Clear active RBP report marker files under the default RBP reports folder.
REM Inputs: RBP_ROOT path in script.
REM Outputs: Deleted _active_* txt files.
REM How to run: run_cascade menu 6 or direct.
REM Notes: Uses %USERPROFILE%\Documents\doc\script\rbp\reports.
@echo off
chcp 65001 >nul
pushd "%~dp0" 2>nul

set "RBP_ROOT=%USERPROFILE%\Documents\doc\script\rbp"
set "REPORTS=%RBP_ROOT%\reports"
if not exist "%REPORTS%" mkdir "%REPORTS%" 2>nul
del /q "%REPORTS%\_active_update_rvt_links.txt" 2>nul
del /q "%REPORTS%\_active_update_rvt_links_details.txt" 2>nul
del /q "%REPORTS%\_active_saveas_new_name.txt" 2>nul
del /q "%REPORTS%\_active_saveas_central.txt" 2>nul
del /q "%REPORTS%\_active_detach_to_buffer.txt" 2>nul
echo Report session reset.
echo Folder: %REPORTS%
popd 2>nul
pause
