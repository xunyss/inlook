@echo off
rem Create .venv and install dependencies. Run once, and again after git pull if pyproject.toml changed.
setlocal
cd /d "%~dp0"

where uv >nul 2>nul
if %errorlevel%==0 (
    echo [setup] using uv
    uv sync --no-dev || goto :fail
) else (
    echo [setup] uv not found, using python -m venv + pip
    if not exist ".venv\Scripts\python.exe" python -m venv .venv || goto :fail
    ".venv\Scripts\python.exe" -m pip install "pywin32>=311" || goto :fail
)
echo.
echo [setup] OK
pause
exit /b 0

:fail
echo.
echo [setup] FAILED - capture this screen and send it.
pause
exit /b 1
