@echo off
setlocal
cd /d "%~dp0"

echo Setting up the raffle project...
where py >nul 2>nul
if %errorlevel%==0 (
    py -3 -m venv .venv
) else (
    python -m venv .venv
)
if errorlevel 1 goto :error

call ".venv\Scripts\activate.bat"
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
if errorlevel 1 goto :error

echo.
echo Setup complete. You can now use run_prepare.bat, run_validate.bat, run_draw.bat, and run_website.bat.
pause
exit /b 0

:error
echo.
echo Setup failed. Confirm that Python 3 is installed and that you can install packages.
pause
exit /b 1
