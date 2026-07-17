@echo off
setlocal
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
    echo The project is not set up yet. Run setup.bat first.
    pause
    exit /b 1
)
if not exist "draw.html" (
    echo draw.html is missing from this folder. Re-download the project.
    pause
    exit /b 1
)
if not exist "draw_data.js" (
    echo draw_data.js does not exist. Run run_draw.bat first.
    pause
    exit /b 1
)
echo Starting a private local website at http://127.0.0.1:8765/draw.html
echo Keep this window open during the raffle. Press Ctrl+C here when the event is finished.
start "" powershell -NoProfile -WindowStyle Hidden -Command "Start-Sleep -Seconds 1; Start-Process 'http://127.0.0.1:8765/draw.html'"
".venv\Scripts\python.exe" -m http.server 8765 --bind 127.0.0.1
