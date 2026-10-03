@echo off
REM Double-click to start the consolidation app. Keep this window open while using it.
cd /d "%~dp0"
python consolidation_app.py
if errorlevel 1 (
    echo.
    echo Could not start the app. Check that Python, pandas and openpyxl are installed.
    pause
)
