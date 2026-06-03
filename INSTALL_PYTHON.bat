@echo off
title Install Python for POS Book Report

echo ==========================================
echo Install Python for POS Book Report
echo ==========================================
echo.

echo Checking if Python is already installed...
py --version >nul 2>&1

if %errorlevel%==0 (
    echo Python is already installed.
    py --version
    echo.
    pause
    exit /b 0
)

echo Python not found.
echo Installing Python 3 using winget...
echo.

winget install Python.Python.3.12 --accept-package-agreements --accept-source-agreements

echo.
echo Installation command completed.
echo Please close and reopen CMD/VS Code Terminal, then run:
echo py --version
echo.
pause