@echo off
rem Re-read attachment lists (without inline images) and sender SMTP addresses from the .msg files.
rem Close the inlook app first. Usage: repair.bat [archive folder] [--dry-run]
setlocal
cd /d "%~dp0"
chcp 65001 >nul
set PYTHONUTF8=1
echo Close the inlook app before running this.
echo.
".venv\Scripts\python.exe" -m inlook repair %*
echo.
pause
