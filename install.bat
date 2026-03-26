@echo off
:: ============================================================
::  install.bat
::  Self-elevates to Administrator then runs the PowerShell
::  installer once. Both files must be in the same folder.
:: ============================================================

:: Check for admin rights - if not present, re-launch elevated
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo Requesting Administrator privileges...
    PowerShell.exe -NoProfile -Command "Start-Process -FilePath cmd.exe -ArgumentList '/c \"%~f0\"' -Verb RunAs"
    exit /b
)

:: Already running as Administrator - execute the PowerShell script
set "SCRIPT=%~dp0install.ps1"

if not exist "%SCRIPT%" (
    echo ERROR: install.ps1 not found in the same folder as this batch file.
    echo Expected location: %SCRIPT%
    pause
    exit /b 1
)

PowerShell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%"

if %errorLevel% neq 0 (
    echo.
    echo [ERROR] Installation did not complete successfully. Review the output above.
    pause
    exit /b %errorLevel%
)

pause
