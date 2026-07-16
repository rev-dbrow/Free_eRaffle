@echo off
setlocal
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
    echo The project is not set up yet. Run setup.bat first.
    pause
    exit /b 1
)
".venv\Scripts\python.exe" raffle.py prepare --config raffle_config.json
echo.
pause
