@echo off
chcp 65001 >nul
title Fund Plan App - first time setup
cd /d "%~dp0"

echo.
echo  ============================================
echo    First time setup (run once)
echo  ============================================
echo.

python --version >nul 2>&1
if errorlevel 1 (
    echo  [ERROR] Python is not installed.
    echo.
    echo   1. Download Python from https://www.python.org
    echo   2. Check "Add Python to PATH" during install
    echo   3. Run this file again
    echo.
    pause
    exit /b
)

echo  Python found:
python --version
echo.
echo  Installing required packages. This may take a few minutes...
echo.

python -m pip install -r requirements.txt

echo.
if errorlevel 1 (
    echo  [ERROR] Installation failed. Check the messages above.
) else (
    echo  ============================================
    echo    Setup complete!  Now double-click START.bat
    echo  ============================================
    echo.
    echo  Reminder: put the original Excel workbook into
    echo            templates\ as  fund_plan_template.xlsm
)
echo.
pause
