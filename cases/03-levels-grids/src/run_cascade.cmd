@echo off
chcp 65001 >nul
REM Weekly wizard: multi-BF auto AR from each host link
pushd "%~dp0" 2>nul

set "RBP_ROOT=%USERPROFILE%\Documents\doc\script\rbp"
set "AUTO_XLSX=%RBP_ROOT%\reports"

:MENU
cls
echo.
echo ============================================================
echo   Levels and Grids
echo ============================================================
echo.
echo   Weekly work (do in order):
echo.
echo   1 - Check BF vs AR          (pick BF only; AR auto from links)
echo   2 - Fix BF from report      (prepare Apply; AR auto)
echo   3 - Check sections vs BF    (pick sections; BF auto from links)
echo   4 - Fix sections from report (prepare Apply; BF auto)
echo   5 - Open reports folder
echo   0 - Exit
echo.
echo   Reports XLSX: %AUTO_XLSX%
echo.
set /p CHOICE=Number: 

if "%CHOICE%"=="0" goto END
if "%CHOICE%"=="1" goto STEP1
if "%CHOICE%"=="2" goto STEP2
if "%CHOICE%"=="3" goto STEP3
if "%CHOICE%"=="4" goto STEP4
if "%CHOICE%"=="5" goto OPEN_REPORTS
echo Invalid choice.
pause
goto MENU

:RESET_SESSION
if not exist "%AUTO_XLSX%" mkdir "%AUTO_XLSX%" 2>nul
if exist "%AUTO_XLSX%\_active_levels_grids_audit.txt" del /f /q "%AUTO_XLSX%\_active_levels_grids_audit.txt"
if exist "%AUTO_XLSX%\_active_levels_grids_audit_details.txt" del /f /q "%AUTO_XLSX%\_active_levels_grids_audit_details.txt"
> "xlsx_out_path.cfg" echo %AUTO_XLSX%
exit /b 0

:SET_AUTO_EXEMPLAR
REM AUTO = script picks AR/BF link inside each host by building code (K1/K2/...)
> "exemplar_model.cfg" echo AUTO
if exist "exemplar_link.cfg" del /f /q "exemplar_link.cfg"
exit /b 0

:PICK_HOST
echo.
echo Select HOST models (local file/folder OR Revit Server).
echo In picker: 1=folder  2=files  3=Revit Server
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0choose_models.ps1"
if errorlevel 1 exit /b 1
exit /b 0

:SHOW_AUDIT
echo.
echo ------------------------------------------------------------
echo NEXT in Revit Batch Processor:
echo   Task Script = levels_grids_audit.py
echo   File List   = rvt_list.txt
echo   Create New Local = ON
echo   Detach           = OFF
echo   Worksets         = Open All   (IMPORTANT - not Close All)
echo.
echo After run:
echo   Reports = %AUTO_XLSX%\levels_grids_audit_YYYY-MM-DD_HHmm.csv / .xlsx
echo ------------------------------------------------------------
exit /b 0

:SHOW_APPLY
echo.
echo ------------------------------------------------------------
echo NEXT in Revit Batch Processor:
echo   Task Script = levels_grids_apply.py
if /I "%APPLY_HOST%"=="SECTIONS" (
  echo   File List   = rvt_list.txt   (sections from fix plan)
  echo   Exemplar         = AUTO (BF link from each section)
) else (
  echo   File List   = rvt_list.txt   (BF from fix plan)
  echo   Exemplar         = AUTO (AR link from each BF)
)
echo   Create New Local = ON
echo   Detach           = OFF
echo   Worksets         = Open All
echo   Plan             = fix_plan.csv
echo ------------------------------------------------------------
exit /b 0

:SHOW_PICKED
echo.
echo Selected now:
echo   Host list:
if exist "rvt_list.txt" (
  powershell -NoProfile -Command "Get-Content -LiteralPath 'rvt_list.txt' -EA SilentlyContinue | Where-Object { $_ -and -not $_.StartsWith('#') } | Select-Object -First 8 | ForEach-Object { Write-Host ('    ' + $_) }"
) else (
  echo     (empty)
)
echo   Exemplar mode:
if exist "exemplar_model.cfg" (
  powershell -NoProfile -Command "Get-Content -LiteralPath 'exemplar_model.cfg' -EA SilentlyContinue | Where-Object { $_ -and -not $_.StartsWith('#') } | Select-Object -First 1 | ForEach-Object { Write-Host ('    ' + $_) }"
) else (
  echo     (empty)
)
echo.
exit /b 0

:STRIP_DUMMY_RSN
REM Drop RSN://x/... placeholders (Revit Server dummy - RBP open error)
powershell -NoProfile -Command "$p='rvt_list.txt'; if (Test-Path -LiteralPath $p) { $k=@(); Get-Content -LiteralPath $p | ForEach-Object { if ($_ -and ($_ -notmatch '(?i)^RSN://x(/|$)')) { $k += $_ } }; $k | Set-Content -LiteralPath $p -Encoding UTF8 }"
exit /b 0

:STEP1
cls
echo.
echo === 1) Check BF vs AR ===
echo.
echo Pick any number of BF models.
echo Audit = Coordination Review + Project Base Point.
echo Script will take AR link from EACH BF (same building: K1-^>AR_K1, ...).
echo Do NOT put AR models into the file list.
echo.
echo Step A: pick BF models
call :PICK_HOST
if errorlevel 1 (pause & goto MENU)
call :STRIP_DUMMY_RSN

call :SET_AUTO_EXEMPLAR
call :RESET_SESSION
call :SHOW_PICKED
call :SHOW_AUDIT
echo.
echo Then open report. If BF needs fix - run menu 2.
pause
goto MENU

:STEP2
cls
echo.
echo === 2) Fix BF from report ===
echo.
echo Builds fix_plan.csv via checkbox window:
echo   BF models, Revit rules, per-issue actions, bulk buttons
echo.
call :SET_AUTO_EXEMPLAR
set "APPLY_HOST=BF"
powershell -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0choose_fix_plan.ps1" -HostKind BF
if errorlevel 1 (pause & goto MENU)
call :SHOW_PICKED
call :SHOW_APPLY
echo.
echo Apply reads fix_plan.csv (not all diffs blindly).
pause
goto MENU

:STEP3
cls
echo.
echo === 3) Check sections vs BF ===
echo.
echo Same as menu 1: Coordination Review + Project Base Point.
echo Pick any number of section models.
echo Script will take BF link from EACH section (same building: K1-^>BF_K1, ...).
echo Do NOT put BF models into the file list.
echo.
echo Step A: pick section models
call :PICK_HOST
if errorlevel 1 (pause & goto MENU)
call :STRIP_DUMMY_RSN

call :SET_AUTO_EXEMPLAR
call :RESET_SESSION
call :SHOW_PICKED
call :SHOW_AUDIT
echo.
echo Then open report. If sections need fix - run menu 4.
pause
goto MENU

:STEP4
cls
echo.
echo === 4) Fix sections from report ===
echo.
echo Builds fix_plan.csv via checkbox window:
echo   section models, PBP Z, per-issue actions
echo   Uses the last Audit from menu 3.
echo.
call :SET_AUTO_EXEMPLAR
set "APPLY_HOST=SECTIONS"
powershell -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0choose_fix_plan.ps1" -HostKind SECTIONS
if errorlevel 1 (pause & goto MENU)
call :SHOW_PICKED
call :SHOW_APPLY
echo.
echo Apply reads fix_plan.csv (not all diffs blindly).
pause
goto MENU

:OPEN_REPORTS
if not exist "%AUTO_XLSX%" mkdir "%AUTO_XLSX%" 2>nul
echo.
echo Opening: %AUTO_XLSX%
start "" "%AUTO_XLSX%"
echo.
pause
goto MENU

:END
popd 2>nul
exit /b 0
