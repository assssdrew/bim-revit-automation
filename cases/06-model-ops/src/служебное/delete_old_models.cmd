@echo off
chcp 65001 >nul
pushd "%~dp0.." 2>nul
echo After Pass 1/2 success only.
echo Step 1: DRY-RUN - shows what would happen, changes nothing.
echo Old files are MOVED to _to_delete next to them (not deleted), and only
echo when the Pass 1/2 reports say OK and the new file exists and is not empty.
echo RSN paths go to to_delete_rsn.txt for manual Admin delete.
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0delete_old_models.ps1"
set "ERR=%ERRORLEVEL%"
if not "%ERR%"=="0" goto DONE
echo.
set "ANS="
set /p ANS=Type YES to apply (anything else = exit): 
if not "%ANS%"=="YES" goto DONE
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0delete_old_models.ps1" -Execute -Yes
set "ERR=%ERRORLEVEL%"
:DONE
popd 2>nul
pause
exit /b %ERR%
