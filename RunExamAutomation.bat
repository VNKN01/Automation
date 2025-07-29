@echo off
setlocal
REM Batch file to run the portable PowerShell script with Bypass ExecutionPolicy

REM Get the folder where this BAT file is located
set SCRIPT_DIR=%~dp0

echo Launching Exam Folder Automation Tool...
echo.

REM Run PowerShell with Bypass ExecutionPolicy only for this run
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%CreateExamFolders_AnyPC.ps1"

echo.
echo Script finished.
pause
