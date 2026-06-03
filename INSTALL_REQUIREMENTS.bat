@echo off
title Install POS Book Report Python Requirements

echo ==========================================
echo Install POS Book Report Requirements
echo ==========================================
echo.

cd /d "%~dp0"

echo Checking Python...
py --version

if errorlevel 1 (
    echo.
    echo Python was not found.
    echo Please run INSTALL_PYTHON.bat first.
    echo.
    pause
    exit /b 1
)

echo.
echo Upgrading pip...
py -m pip install --upgrade pip

echo.
echo Installing requirements from requirements.txt...
py -m pip install -r requirements.txt

if errorlevel 1 (
    echo.
    echo Failed to install requirements.
    echo Please check your internet connection or Python installation.
    echo.
    pause
    exit /b 1
)

echo.
echo Requirements installed successfully.
echo.
pause