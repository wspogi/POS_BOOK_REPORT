@echo off
title POS Book Report Automation

echo ==========================================
echo POS Book Report Automation
echo ==========================================
echo.

cd /d "%~dp0"

echo Running POS Book Report...
echo.

py main.py

if errorlevel 1 (
    echo.
    echo POS Book Report failed.
    echo Please check the error above or check the logs folder.
    echo.
    pause
    exit /b 1
)

echo.
echo POS Book Report completed successfully.
echo.
pause