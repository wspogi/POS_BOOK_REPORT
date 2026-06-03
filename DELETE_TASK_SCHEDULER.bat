@echo off
title Delete Task Scheduler - POS Book Report

echo ==========================================
echo Delete Task Scheduler - POS Book Report
echo ==========================================
echo.

set TASK_NAME=POS Book Report Automation

echo Deleting scheduled task:
echo %TASK_NAME%
echo.

schtasks /Delete /TN "%TASK_NAME%" /F

if errorlevel 1 (
    echo.
    echo Failed to delete scheduled task or task does not exist.
    echo.
    pause
    exit /b 1
)

echo.
echo Scheduled task deleted successfully.
echo.
pause