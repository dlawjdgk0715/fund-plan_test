@echo off
chcp 65001 >nul
title Fund Plan App (shared)
cd /d "%~dp0"

echo.
echo  ============================================
echo    Fund Plan App - team shared mode
echo  ============================================
echo.
echo  Teammates on the same office network can connect.
echo  Share the "Network URL" shown below.
echo.
echo  Note: closing this window disconnects everyone.
echo.

python -m streamlit run app/Home.py --server.address 0.0.0.0

echo.
pause
