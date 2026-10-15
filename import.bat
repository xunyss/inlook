@echo off
rem Import .msg files saved earlier from Outlook ("Save As").
rem Usage: import.bat <folder> [--dry-run] [--to D:\mailbox]
rem   files directly in <folder>  -> label "Inbox"
rem   files in <folder>\X         -> label "Inbox/X"
setlocal
cd /d "%~dp0"
chcp 65001 >nul
set PYTHONUTF8=1

if "%~1"=="" (
    echo Usage: import.bat ^<folder^> [--dry-run]
    pause
    exit /b 1
)
if not exist ".venv\Scripts\python.exe" (
    echo [import] .venv not found. Run setup.bat first.
    pause
    exit /b 1
)
echo Close the inlook app before importing.
echo.
".venv\Scripts\python.exe" -m inlook import %*
echo.
pause
