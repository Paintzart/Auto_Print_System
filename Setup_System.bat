@echo off
title Full System Setup - Pre-Press Automation
echo ==========================================
echo   Step 1: Checking Environments
echo ==========================================

:: בדיקת פייתון
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python is not installed!
    pause & exit
)

:: בדיקת Node.js
node -v >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Node.js is not installed!
    pause & exit
)

echo ==========================================
echo   Step 2: Python Environment & Libraries
echo ==========================================
:: יצירת סביבה וירטואלית
if not exist venv (
    echo Creating venv...
    python -m venv venv
)

:: הפעלת הסביבה והתקנת ספריות
call venv\Scripts\activate
echo Installing Python Libraries (pywin32, requests, pymupdf)...
pip install pywin32 requests pymupdf streamlit axios

:: שלב קריטי: רישום ה-COM עבור אילוסטרייטור ופוטושופ
echo Registering COM objects...
python venv\Scripts\pywin32_postinstall.py -install

echo ==========================================
echo   Step 3: Node.js Server Setup
echo ==========================================
:: התקנת ספריות שרת (npm install יחפש את package.json בתיקייה הנוכחית)
if exist package.json (
    echo Installing Node.js dependencies...
    call npm install
) else (
    echo [WARNING] package.json not found in this folder!
)

echo.
echo ==========================================
echo        SETUP COMPLETE! 🚀
echo ==========================================
echo You can now use the Run scripts.
pause