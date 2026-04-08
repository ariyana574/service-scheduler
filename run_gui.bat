@echo off
cd /d "%~dp0"

echo Starting Service Scheduler GUI...
echo.

call .venv\Scripts\activate.bat

streamlit run gui_app.py

pause
