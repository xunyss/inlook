@echo off
rem Shrink the DB: drop HTML bodies that can be read back from the .msg files, then VACUUM.
rem Close the inlook app first. Usage: compact.bat [archive folder] [--dry-run]
setlocal
cd /d "%~dp0"
chcp 65001 >nul
set PYTHONUTF8=1
echo Close the inlook app before running this.
echo.
".venv\Scripts\python.exe" -m inlook compact %*
echo.
pause
