@echo off
:: ============================================================
::  2-Install-HRS-PaymentPro.bat
::  Self-elevates to Administrator then runs the HRS
::  PaymentPro.Net PowerShell installer.
::  Both files must be in the same folder.
:: ============================================================

net session >nul 2>&1
if %errorLevel% neq 0 (
    echo Requesting Administrator privileges...
    PowerShell.exe -NoProfile -Command "Start-Process -FilePath cmd.exe -ArgumentList '/c \"%~f0\"' -Verb RunAs"
    exit /b
)

set "SCRIPT=%~dp02-Install-HRS-PaymentPro.ps1"

if not exist "%SCRIPT%" (
    echo ERROR: 2-Install-HRS-PaymentPro.ps1 not found in the same folder as this batch file.
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
