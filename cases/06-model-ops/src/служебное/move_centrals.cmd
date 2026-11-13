@echo off
chcp 65001 >nul
pushd "%~dp0.." 2>nul
echo Step 1: DRY-RUN - shows what would be moved, changes nothing.
echo Existing destination files are never overwritten (use move_centrals.ps1 -Overwrite by hand).
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0move_centrals.ps1"
set "ERR=%ERRORLEVEL%"
if not "%ERR%"=="0" goto DONE
echo.
set "ANS="
set /p ANS=Type YES to apply (anything else = exit): 
if not "%ANS%"=="YES" goto DONE
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0move_centrals.ps1" -Execute -Yes
set "ERR=%ERRORLEVEL%"
:DONE
popd 2>nul
pause
exit /b %ERR%
