@echo off
title DEVELOPER MODE - No Updates
color 0B
cd /d "%~dp0"

:: הפעלת הסביבה הווירטואלית (למקרה שתצטרכי להריץ פייתון בחלון זה)
call venv\Scripts\activate

echo DEVELOPER MODE ACTIVATED (Safe for Coding)
echo Running YOUR local code directly!

node server.js
pause