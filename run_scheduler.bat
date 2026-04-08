@echo off
REM Simple runner for the service scheduler

REM Change to the folder where this script lives
cd /d "%~dp0"

REM Activate virtual environment
call .venv\Scripts\activate.bat

REM Run the scheduler
python src\scheduler.py

REM Pause so the window doesn't close immediately
echo.
echo Finished. Press any key to close this window...
pause >nul
