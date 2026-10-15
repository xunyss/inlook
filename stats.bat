@echo off
rem DB size report (sizes and counts only, no mail content). Usage: stats.bat [archive folder]
setlocal
cd /d "%~dp0"
chcp 65001 >nul
set PYTHONUTF8=1
".venv\Scripts\python.exe" -m inlook stats %*
echo.
pause
