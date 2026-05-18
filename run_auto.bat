@echo off
title Graphic Automation Server
color 0A
cd /d "%~dp0"

:: הפעלת הסביבה הווירטואלית
call venv\Scripts\activate

echo Starting Smart Automation Server...
echo Checking for updates from GitHub first...

:: הרצת המעדכן (הוא כבר יפעיל את node server.js בסוף)
python run_me.py

pause