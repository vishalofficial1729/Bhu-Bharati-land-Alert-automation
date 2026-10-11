@echo off
REM One-time setup on Windows. Double-click this file.
cd /d "%~dp0"
py -3.12 -m venv .venv 2>nul || python -m venv .venv
if not exist .venv\Scripts\python.exe (echo Python not found. Install Python 3.12 from python.org and tick "Add python.exe to PATH". & pause & exit /b 1)
.venv\Scripts\python -m pip install --upgrade pip --quiet
.venv\Scripts\pip install -r requirements.txt || (pause & exit /b 1)
.venv\Scripts\playwright install chromium || (pause & exit /b 1)
echo.
echo Setup finished. Double-click start.bat to open the app.
pause
