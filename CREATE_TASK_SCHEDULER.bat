@echo off
title Create Task Scheduler - POS Book Report

echo ==========================================
echo Create Task Scheduler - POS Book Report
echo ==========================================
echo.

cd /d "%~dp0"

set TASK_NAME=POS Book Report Automation
set PROJECT_DIR=%~dp0
set RUN_BAT=%PROJECT_DIR%RUN_POS_BOOK_REPORT.bat

echo Task Name:
echo %TASK_NAME%
echo.
echo Project Directory:
echo %PROJECT_DIR%
echo.
echo Run File:
echo %RUN_BAT%
echo.

echo Creating scheduled task...
echo Schedule: Monthly, every 1st day, 08:00 AM
echo.

schtasks /Create ^
 /TN "%TASK_NAME%" ^
 /TR "\"%RUN_BAT%\"" ^
 /SC MONTHLY ^
 /D 1 ^
 /ST 08:00 ^
 /F

if errorlevel 1 (
    echo.
    echo Failed to create scheduled task.
    echo Try running this BAT file as Administrator.
    echo.
    pause
    exit /b 1
)

echo.
echo Scheduled task created successfully.
echo.

echo Testing task information...
schtasks /Query /TN "%TASK_NAME%"

echo.
echo Done.
echo.
pause