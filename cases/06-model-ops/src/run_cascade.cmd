@echo off
chcp 65001 >nul
pushd "%~dp0" 2>nul

:MENU
cls
echo ============================================
echo  batch_model_ops  (rename / move / upgrade / relink)
echo ============================================
echo.
if exist "rbp_checklist.txt" (
  echo --- rbp_checklist.txt ---
  type "rbp_checklist.txt"
  echo -------------------------
) else (
  echo No saved job yet. Open presets ^(1^) and Save job.
)
echo.
echo 1 - Presets window ^(names, location, version, preview, mapping.csv^)
echo 2 - Choose models ^(rvt_list.txt^)
echo 3 - Windows Move UNC from job_paths.csv
echo 4 - Delete old UNC + buffer / list RSN
echo 5 - Open reports folder
echo 6 - Reset report session
echo 7 - Re-apply RBP open patches
echo 0 - Exit
echo.
set /p CHOICE=Number:

if "%CHOICE%"=="0" goto END
if "%CHOICE%"=="1" goto PRESETS
if "%CHOICE%"=="2" goto PICK
if "%CHOICE%"=="3" goto MOVE
if "%CHOICE%"=="4" goto DEL
if "%CHOICE%"=="5" goto REPORTS
if "%CHOICE%"=="6" goto RESET
if "%CHOICE%"=="7" goto PATCH
echo Invalid choice.
pause
goto MENU

:PRESETS
powershell -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0presets.ps1"
pause
goto MENU

:PICK
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0choose_models.ps1" -OutFile "rvt_list.txt"
pause
goto MENU

:MOVE
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0move_centrals.ps1"
pause
goto MENU

:DEL
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0delete_old_models.ps1"
pause
goto MENU

:REPORTS
set "RBP_ROOT=%USERPROFILE%\Documents\doc\script\rbp"
set "AUTO_CSV=%RBP_ROOT%\reports"
if not exist "%AUTO_CSV%" mkdir "%AUTO_CSV%" 2>nul
explorer "%AUTO_CSV%"
goto MENU

:RESET
call "%~dp0reset_report_session.cmd"
goto MENU

:PATCH
echo Close Revit Batch Processor first.
pause
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0patch_rbp_open_failures.ps1"
pause
goto MENU

:END
popd 2>nul
exit /b 0
