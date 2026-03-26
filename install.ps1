#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Installs IIS / .NET features and configures KCX PaymentPro.Net.
.DESCRIPTION
    1. Installs IIS, .NET 3.5, ASP.NET 3.5/4.8 and related features via DISM
    2. Removes the Default IIS Website
    3. Runs the KCX PaymentPro.Net Service Installer
    4. Starts KCXPPNService and verifies it stays running
    5. Creates the PPN Scheduled Task in Task Scheduler
.NOTES
    Must be run as Administrator.
    Run with: PowerShell.exe -ExecutionPolicy Bypass -File install.ps1
#>

# ============================================================
#  Helper: Write a section header
# ============================================================
function Write-Header {
    param([string]$Title)
    Write-Host ""
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host "  $Title" -ForegroundColor Cyan
    Write-Host "============================================================" -ForegroundColor Cyan
}

# ============================================================
#  DISM Features
# ============================================================
Write-Header "Installing IIS, .NET and ASP.NET Features"

$dismFeatures = @(
    'NetFx3',                   # .NET Framework 3.5 (includes 2.0 and 3.0)
    'IIS-WebServerRole',        # IIS Base
    'IIS-WebServer',            # IIS Web Server
    'IIS-NetFxExtensibility',   # .NET Extensibility 3.5
    'IIS-NetFxExtensibility45', # .NET Extensibility 4.8
    'IIS-ASPNET',               # ASP.NET 3.5
    'IIS-ASPNET45',             # ASP.NET 4.8
    'IIS-ISAPIExtensions',      # ISAPI Extensions
    'IIS-ISAPIFilter'           # ISAPI Filters
)

foreach ($feature in $dismFeatures) {
    Write-Host "Enabling: $feature" -ForegroundColor Gray
    $result = dism /online /enable-feature /featurename:$feature /all /norestart
    if ($LASTEXITCODE -eq 0 -or $LASTEXITCODE -eq 3010) {
        Write-Host "[OK] $feature" -ForegroundColor Green
    } else {
        Write-Host "[WARN] $feature exited with code $LASTEXITCODE" -ForegroundColor Yellow
    }
}

Write-Host ""
Write-Host "[OK] IIS and related features installed." -ForegroundColor Green

# ============================================================
#  STEP 1 - Remove the Default IIS Website
# ============================================================
Write-Header "STEP 1: Removing Default IIS Website"

Import-Module WebAdministration -ErrorAction SilentlyContinue

if (Test-Path 'IIS:\Sites\Default Web Site') {
    Remove-WebSite -Name 'Default Web Site'
    Write-Host "[OK] Default Web Site removed." -ForegroundColor Green
} else {
    Write-Host "[INFO] Default Web Site not found - skipping." -ForegroundColor Yellow
}

# ============================================================
#  STEP 2 - Install KCXPPNService via InstallUtil
# ============================================================
Write-Header "STEP 2: Installing KCXPPNService via InstallUtil"

$serviceExe  = 'C:\Program Files\KCX PaymentPro.Net\KCXPPNService.exe'
$installUtil = "$env:SystemRoot\Microsoft.NET\Framework\v4.0.30319\InstallUtil.exe"

# Fall back to 64-bit only if 32-bit is somehow not present
if (-not (Test-Path $installUtil)) {
    $installUtil = "$env:SystemRoot\Microsoft.NET\Framework64\v4.0.30319\InstallUtil.exe"
}

if (-not (Test-Path $serviceExe)) {
    Write-Host "[ERROR] Service executable not found at: $serviceExe" -ForegroundColor Red
    exit 1
}

if (-not (Test-Path $installUtil)) {
    Write-Host "[ERROR] InstallUtil.exe not found. Ensure .NET Framework 4.x is installed." -ForegroundColor Red
    exit 1
}

Write-Host "Using InstallUtil: $installUtil"
Write-Host "Installing service from: $serviceExe"
Write-Host ""

& $installUtil /LogFile= /LogToConsole=true "$serviceExe"

if ($LASTEXITCODE -eq 0) {
    Write-Host "[OK] Service installed successfully." -ForegroundColor Green
} else {
    Write-Host "[ERROR] InstallUtil failed with exit code: $LASTEXITCODE" -ForegroundColor Red
    Write-Host "        Review the output above for the exact error message."
    exit 1
}

# ============================================================
#  STEP 3 - Start KCXPPNService
# ============================================================
Write-Header "STEP 3: Starting KCXPPNService"

$serviceName  = 'KCXPPNService'
$retrySeconds = 60
$interval     = 5
$elapsed      = 0
$svc          = $null

Write-Host "Waiting for '$serviceName' to appear (up to $retrySeconds seconds)..."

while ($elapsed -lt $retrySeconds) {
    $svc = Get-Service -Name $serviceName -ErrorAction SilentlyContinue
    if ($null -ne $svc) {
        Write-Host "[OK] Service '$serviceName' found after $elapsed seconds." -ForegroundColor Green
        break
    }
    Start-Sleep -Seconds $interval
    $elapsed += $interval
    Write-Host "  Still waiting... ($elapsed / $retrySeconds seconds)"
}

if ($null -eq $svc) {
    Write-Host ""
    Write-Host "[ERROR] Service '$serviceName' was not found after $retrySeconds seconds." -ForegroundColor Red
    Write-Host ""
    Write-Host "--- Services currently registered that contain 'KCX' or 'PPN' ---" -ForegroundColor Yellow
    $matches = Get-Service | Where-Object { $_.Name -match 'KCX|PPN' -or $_.DisplayName -match 'KCX|PPN' }
    if ($matches) {
        $matches | Format-Table Name, DisplayName, Status -AutoSize
    } else {
        Write-Host "  None found. The installer may have failed silently." -ForegroundColor Yellow
    }
    Write-Host ""
    Write-Host "--- All services registered in the last 5 minutes ---" -ForegroundColor Yellow
    Get-Service | Sort-Object Name | Format-Table Name, DisplayName, Status -AutoSize
    exit 1
}

if ($svc.Status -ne 'Running') {
    Start-Service -Name $serviceName
    Write-Host "[OK] Start command sent to '$serviceName'." -ForegroundColor Green
} else {
    Write-Host "[INFO] '$serviceName' is already running." -ForegroundColor Yellow
}

# ============================================================
#  STEP 4 - Verify KCXPPNService stays running
# ============================================================
Write-Header "STEP 4: Verifying KCXPPNService Remains Running"

Write-Host "Waiting 5 seconds for stability check 1 of 2..."
Start-Sleep -Seconds 5
$svc.Refresh()

if ($svc.Status -ne 'Running') {
    Write-Host "[WARN] Service stopped after first check. Attempting restart..." -ForegroundColor Yellow
    Start-Service -Name $serviceName

    Write-Host "Waiting 5 seconds for stability check 2 of 2..."
    Start-Sleep -Seconds 5
    $svc.Refresh()

    if ($svc.Status -ne 'Running') {
        Write-Host ""
        Write-Host "************************************************************" -ForegroundColor Red
        Write-Host "*  ERROR: KCXPPNService keeps stopping after two attempts. *" -ForegroundColor Red
        Write-Host "*  Please check the service logs and Windows Event Viewer  *" -ForegroundColor Red
        Write-Host "*  for details before attempting to run this script again. *" -ForegroundColor Red
        Write-Host "************************************************************" -ForegroundColor Red
        exit 1
    } else {
        Write-Host "[OK] Service recovered and is now running." -ForegroundColor Green
    }
} else {
    Write-Host "[OK] KCXPPNService is running and stable." -ForegroundColor Green
}

# ============================================================
#  STEP 5 - Create the PPN Scheduled Task
# ============================================================
Write-Header "STEP 5: Creating PPN Scheduled Task"

$taskName    = 'PPN'
$taskExe     = 'C:\Program Files\KCX PaymentPro.Net\KCX.PPN.AutoSettlement.exe'
$triggerTime = '20:00'

# Remove any existing task with the same name for a clean install
if (Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue) {
    Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
    Write-Host "[INFO] Existing '$taskName' task removed; will be recreated." -ForegroundColor Yellow
}

# Action
$action = New-ScheduledTaskAction -Execute $taskExe

# Trigger - Daily at 8:00 PM
$trigger = New-ScheduledTaskTrigger -Daily -At $triggerTime

# Principal - run whether logged in or not, highest privileges
$principal = New-ScheduledTaskPrincipal `
    -UserId    (whoami) `
    -LogonType S4U `
    -RunLevel  Highest

# Settings
$settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -StartWhenAvailable `
    -ExecutionTimeLimit (New-TimeSpan -Days 3) `
    -MultipleInstances  IgnoreNew

# Register the task
Register-ScheduledTask `
    -TaskName  $taskName `
    -Action    $action `
    -Trigger   $trigger `
    -Principal $principal `
    -Settings  $settings `
    -Force | Out-Null

Write-Host "[OK] Scheduled task '$taskName' created successfully." -ForegroundColor Green
Write-Host ""
Write-Host "  Name          : PPN"
Write-Host "  Run as        : Logged in or not (S4U)"
Write-Host "  Trigger       : Daily at 8:00 PM"
Write-Host "  Action        : $taskExe"
Write-Host "  Run on demand : Yes"
Write-Host "  Max duration  : 3 days (force-stop if unresponsive)"

# ============================================================
#  Done
# ============================================================
Write-Host ""
Write-Host "============================================================" -ForegroundColor Green
Write-Host "  All steps completed successfully." -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green
