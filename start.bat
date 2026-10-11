@echo off
REM Start the app on Windows. Double-click this file; close the window to stop.
cd /d "%~dp0"
if not exist .venv\Scripts\python.exe (echo Run setup.bat first. & pause & exit /b 1)
.venv\Scripts\python app.py
pause
