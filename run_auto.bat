@echo off
title Graphic Automation Server
color 0A
cd /d "%~dp0"

:: Do not silently fall back to a different/broken system Python.
if not exist "venv\Scripts\python.exe" (
    echo [ERROR] Python environment is missing.
    echo Please run Setup_System.bat first.
    pause & exit /b 1
)
"venv\Scripts\python.exe" --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python environment is broken or was copied from another PC.
    echo Please run Setup_System.bat to rebuild it.
    pause & exit /b 1
)

echo Starting Smart Automation Server...
echo Checking for updates from GitHub first...

:: הרצת המעדכן (הוא כבר יפעיל את node server.js בסוף)
"venv\Scripts\python.exe" run_me.py

pause
