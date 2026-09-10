@echo off
chcp 65001 >nul
title Fund Plan App
cd /d "%~dp0"

echo.
echo  ============================================
echo    Fund Plan App - starting...
echo  ============================================
echo.
echo  A browser will open shortly.
echo  If not, open:  http://localhost:8501
echo.
echo  To quit: close this window or press Ctrl+C
echo.

python -m streamlit run app/Home.py

echo.
echo  App stopped.
pause
