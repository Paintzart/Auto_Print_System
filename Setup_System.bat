@echo off
title Full System Setup - Pre-Press Automation
cd /d "%~dp0"
echo ==========================================
echo   Step 1: Checking Environments
echo ==========================================

:: Find a real Python installation.  On Windows, "python" can be only the
:: Microsoft Store alias, so prefer the Python launcher when it is available.
set "PYTHON_CMD="
py -3 --version >nul 2>&1
if %errorlevel% equ 0 set "PYTHON_CMD=py -3"
if not defined PYTHON_CMD (
    python --version >nul 2>&1
    if %errorlevel% equ 0 set "PYTHON_CMD=python"
)
if not defined PYTHON_CMD (
    echo [ERROR] Python 3 is not installed or is not available in PATH.
    echo Install Python 3 from python.org and enable "Add Python to PATH".
    pause & exit /b 1
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
:: A copied/old venv contains an absolute link to Python on the other PC.
:: Validate it instead of assuming that an existing folder is usable.
if exist "venv\Scripts\python.exe" (
    "venv\Scripts\python.exe" --version >nul 2>&1
    if errorlevel 1 (
        echo Existing Python environment is broken or belongs to another PC.
        echo Rebuilding venv...
        rmdir /s /q "venv"
    )
) else if exist "venv" (
    echo Existing Python environment is incomplete. Rebuilding venv...
    rmdir /s /q "venv"
)

if not exist "venv\Scripts\python.exe" (
    echo Creating venv...
    %PYTHON_CMD% -m venv venv
    if errorlevel 1 (
        echo [ERROR] Could not create the Python environment.
        pause & exit /b 1
    )
)

:: Always use the venv interpreter directly; activation is shell-dependent.
echo Installing Python Libraries from requirements.txt...
"venv\Scripts\python.exe" -m pip install --upgrade pip
"venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 (
    echo [ERROR] Python library installation failed.
    pause & exit /b 1
)

:: שלב קריטי: רישום ה-COM עבור אילוסטרייטור ופוטושופ
echo Registering COM objects...
if exist "venv\Scripts\pywin32_postinstall.py" (
    "venv\Scripts\python.exe" "venv\Scripts\pywin32_postinstall.py" -install
)

echo ==========================================
echo   Step 3: Node.js Server Setup
echo ==========================================
:: התקנת ספריות שרת (npm install יחפש את package.json בתיקייה הנוכחית)
if exist package.json (
    echo Installing Node.js dependencies...
    call npm install
    if errorlevel 1 (
        echo [ERROR] Node.js dependency installation failed.
        pause & exit /b 1
    )
) else (
    echo [WARNING] package.json not found in this folder!
)

echo.
echo ==========================================
echo        SETUP COMPLETE! 🚀
echo ==========================================
echo You can now use the Run scripts.
pause
