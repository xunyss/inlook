@echo off
rem Viewer app (search, labels, sync). Usage: app.bat [archive folder]   default D:\mailbox
rem No console window. If the app does not open, run in cmd to see the error:
rem   .venv\Scripts\python.exe -m mailer app
setlocal
cd /d "%~dp0"
set PYTHONUTF8=1

if not exist ".venv\Scripts\pythonw.exe" (
    echo [app] .venv not found. Run setup.bat first.
    pause
    exit /b 1
)
start "" ".venv\Scripts\pythonw.exe" -m mailer app %*
