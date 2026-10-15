@echo off
rem Self-diagnosis. Usage: probe.bat [archive folder]   e.g. probe.bat H:\mail
setlocal
cd /d "%~dp0"
chcp 65001 >nul
set PYTHONUTF8=1

if not exist ".venv\Scripts\python.exe" (
    echo [probe] .venv not found. Run setup.bat first.
    pause
    exit /b 1
)
echo To stop, press Ctrl+C and wait. Do NOT close this window while it is running.
echo.
".venv\Scripts\python.exe" -m inlook probe %*
echo.
pause
