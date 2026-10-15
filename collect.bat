@echo off
rem Sync: save new mail as .msg + DB. Usage: collect.bat [archive folder] [--dry-run] [--full]
rem Default archive folder: D:\mailbox
setlocal
cd /d "%~dp0"
chcp 65001 >nul
set PYTHONUTF8=1

if not exist ".venv\Scripts\python.exe" (
    echo [collect] .venv not found. Run setup.bat first.
    pause
    exit /b 1
)
echo To stop, press Ctrl+C and wait. Do NOT close this window while it is running.
echo Saved mail is kept. Run again to continue.
echo.
".venv\Scripts\python.exe" -m inlook collect %*
echo.
pause
