@echo off
setlocal EnableExtensions
set "SELF=%~f0"

:: Require elevation for DISM, IIS admin tasks, service control, and scheduled task creation
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo Requesting administrative privileges...
    powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "Start-Process -FilePath $env:ComSpec -ArgumentList ('/c ""' + $env:SELF + '""') -Verb RunAs"
    exit /b 0
)

echo Installing IIS, .NET and ASP.NET...

:: .NET Framework 3.5 (includes 2.0 and 3.0)
dism /online /enable-feature /featurename:NetFx3 /all /norestart || exit /b %errorlevel%

:: IIS Base (required before IIS sub-features)
dism /online /enable-feature /featurename:IIS-WebServerRole /all /norestart || exit /b %errorlevel%
dism /online /enable-feature /featurename:IIS-WebServer /all /norestart || exit /b %errorlevel%

:: .NET Extensibility 3.5
dism /online /enable-feature /featurename:IIS-NetFxExtensibility /all /norestart || exit /b %errorlevel%

:: .NET Extensibility 4.8
dism /online /enable-feature /featurename:IIS-NetFxExtensibility45 /all /norestart || exit /b %errorlevel%

:: ASP.NET 3.5
dism /online /enable-feature /featurename:IIS-ASPNET /all /norestart || exit /b %errorlevel%

:: ASP.NET 4.8
dism /online /enable-feature /featurename:IIS-ASPNET45 /all /norestart || exit /b %errorlevel%

:: ISAPI Extensions
dism /online /enable-feature /featurename:IIS-ISAPIExtensions /all /norestart || exit /b %errorlevel%

:: ISAPI Filters
dism /online /enable-feature /featurename:IIS-ISAPIFilter /all /norestart || exit /b %errorlevel%

echo.
echo IIS and related features installed.
echo Running post-install PowerShell steps...
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -Command ^
    "$content = Get-Content -LiteralPath $env:SELF -Raw; " ^
    "$marker = ':__POWERSHELL__'; " ^
    "$idx = $content.IndexOf($marker); " ^
    "if ($idx -lt 0) { throw 'Embedded PowerShell marker not found.' }; " ^
    "$ps = $content.Substring($idx + $marker.Length); " ^
    "Invoke-Expression $ps"

set "EC=%ERRORLEVEL%"
if not "%EC%"=="0" echo Post-install PowerShell failed with exit code %EC%.
endlocal & exit /b %EC%

goto :eof

:__POWERSHELL__
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

function Write-Step {
    param([Parameter(Mandatory)][string]$Message)
    Write-Host "`n=== $Message ==="
}

function Test-PathOrThrow {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Label
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "$Label not found: $Path"
    }
}

function Start-And-VerifyService {
    param(
        [Parameter(Mandatory)][string]$ServiceName,
        [int]$WaitSeconds = 5
    )

    $svc = Get-Service -Name $ServiceName -ErrorAction Stop

    if ($svc.Status -ne 'Running') {
        Write-Host "Starting service '$ServiceName'..."
        Start-Service -Name $ServiceName -ErrorAction Stop
    }
    else {
        Write-Host "Service '$ServiceName' is already running."
    }

    Start-Sleep -Seconds $WaitSeconds
    $svc = Get-Service -Name $ServiceName -ErrorAction Stop

    if ($svc.Status -eq 'Running') {
        Write-Host "Service '$ServiceName' is running after initial check."
        return
    }

    Write-Warning "Service '$ServiceName' is not running after $WaitSeconds seconds. Attempting one restart..."

    try {
        if ($svc.Status -eq 'Running') {
            Restart-Service -Name $ServiceName -Force -ErrorAction Stop
        }
        else {
            Start-Service -Name $ServiceName -ErrorAction Stop
        }
    }
    catch {
        Write-Warning "Restart/Start attempt reported: $($_.Exception.Message)"
        Start-Service -Name $ServiceName -ErrorAction Stop
    }

    Start-Sleep -Seconds $WaitSeconds
    $svc = Get-Service -Name $ServiceName -ErrorAction Stop

    if ($svc.Status -ne 'Running') {
        throw "Service '$ServiceName' keeps stopping. Review the application logs and Windows Event Viewer."
    }

    Write-Host "Service '$ServiceName' is running after the retry."
}

try {
    $installerPath       = 'C:\Program Files\KCX PaymentPro.Net\KCX.PPN.ServiceInstaller.exe'
    $serviceName         = 'KCXPPNService'
    $expectedServiceExe  = 'C:\Program Files\KCX PaymentPro.Net\KCXPPNService.exe'
    $taskName            = 'PPN'
    $taskExe             = 'C:\Program Files\KCX PaymentPro.Net\KCX.PPN.AutoSettlement.exe'

    Write-Step 'Removing Default IIS website'
    Import-Module WebAdministration -ErrorAction Stop

    $defaultSite = Get-Website -Name 'Default Web Site' -ErrorAction SilentlyContinue
    if ($null -ne $defaultSite) {
        Remove-Website -Name 'Default Web Site'
        Write-Host "Removed IIS site: Default Web Site"
    }
    else {
        Write-Host "IIS site 'Default Web Site' was not found. Skipping removal."
    }

    Write-Step 'Running KCX service installer'
    Test-PathOrThrow -Path $installerPath -Label 'Installer'
    $installerProc = Start-Process -FilePath $installerPath -Wait -PassThru

    if ($installerProc.ExitCode -ne 0) {
        throw "Installer exited with code $($installerProc.ExitCode)."
    }

    Write-Host 'Installer completed successfully.'

    Write-Step 'Validating service registration'
    $svcCim = Get-CimInstance Win32_Service -Filter "Name='$serviceName'" -ErrorAction SilentlyContinue
    if (-not $svcCim) {
        throw "Service '$serviceName' was not found after running the installer."
    }

    $actualPath = ($svcCim.PathName -replace '"', '')
    if ($actualPath -notlike "*$expectedServiceExe*") {
        Write-Warning "Service path does not match the expected executable."
        Write-Warning "Expected: $expectedServiceExe"
        Write-Warning "Found:    $($svcCim.PathName)"
    }
    else {
        Write-Host "Service path matches expected executable."
    }

    Write-Step 'Starting and verifying service'
    Start-And-VerifyService -ServiceName $serviceName -WaitSeconds 5

    Write-Step 'Creating scheduled task'
    Test-PathOrThrow -Path $taskExe -Label 'Scheduled task executable'
    Import-Module ScheduledTasks -ErrorAction Stop

    if (Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue) {
        Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
        Write-Host "Existing scheduled task '$taskName' removed."
    }

    $action   = New-ScheduledTaskAction -Execute $taskExe
    $trigger  = New-ScheduledTaskTrigger -Daily -At 8:00PM
    $settings = New-ScheduledTaskSettingsSet `
        -DisallowDemandStart:$false `
        -ExecutionTimeLimit (New-TimeSpan -Days 3) `
        -DisallowHardTerminate:$false

    $currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
    $principal   = New-ScheduledTaskPrincipal -UserId $currentUser -LogonType S4U -RunLevel Highest

    Register-ScheduledTask `
        -TaskName $taskName `
        -Action $action `
        -Trigger $trigger `
        -Settings $settings `
        -Principal $principal `
        -Description 'PPN AutoSettlement' `
        -Force | Out-Null

    Write-Host "Scheduled task '$taskName' created successfully."
    Write-Host "Run as user: $currentUser"
    Write-Host "Trigger: Daily at 8:00 PM"

    Write-Step 'Completed'
    Write-Host 'All steps completed successfully.'
    exit 0
}
catch {
    Write-Error $_.Exception.Message
    exit 1
}
