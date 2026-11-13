@echo off
rem UNC: cmd cannot use a network folder as current dir; pushd maps it to a temporary drive letter. Script paths come from the location of this .bat
pushd "%~dp0" 2>nul
rem Compare two filter dumps: first = was, second = became. See compare_filter_dumps.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0compare_filter_dumps.ps1" %*
popd 2>nul
