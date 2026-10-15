@echo off
rem Create .venv and install dependencies. Run once, and again after git pull if pyproject.toml changed.
rem Close app.bat / collect.bat before running this.
setlocal
cd /d "%~dp0"

where uv >nul 2>nul
if not %errorlevel%==0 goto :pip

echo [setup] using uv
rem Security software on the VDI scans new files while uv writes them and can
rem briefly lock them ("Access is denied", os error 5). Install one package at
rem a time and retry; already downloaded packages are reused.
set UV_CONCURRENT_INSTALLS=1
set /a tries=0
:retry
set /a tries+=1
uv sync --no-dev && goto :ok
if %tries% GEQ 5 goto :fail
echo.
echo [setup] install failed (try %tries%/5). Retrying in 10 seconds...
timeout /t 10 /nobreak >nul
goto :retry

:pip
echo [setup] uv not found, using python -m venv + pip
if not exist ".venv\Scripts\python.exe" python -m venv .venv || goto :fail
".venv\Scripts\python.exe" -m pip install "pywin32>=311" "pyside6-essentials>=6.8" "olefile>=0.47" || goto :fail

:ok
echo.
echo [setup] OK
pause
exit /b 0

:fail
echo.
echo [setup] FAILED - capture this screen and send it.
pause
exit /b 1
