@echo off
title POS Book Report - Test SQL Connection

echo ==========================================
echo POS Book Report - Test SQL Connection
echo ==========================================
echo.

cd /d "%~dp0"

py test_connection.py

if errorlevel 1 (
    echo.
    echo SQL connection test failed.
    echo Please check your .env database credentials or network connection.
    echo.
    pause
    exit /b 1
)

echo.
echo SQL connection test completed successfully.
echo.
pause