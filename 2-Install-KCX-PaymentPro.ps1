#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Step 2 of 2 - Configures KCX PaymentPro.Net.
.DESCRIPTION
    1. Removes the Default IIS Website
    2. Installs KCXPPNService via InstallUtil
    3. Starts KCXPPNService and verifies it stays running
    4. Creates the PPN Scheduled Task in Task Scheduler
.NOTES
    Must be run as Administrator.
    Run with: PowerShell.exe -ExecutionPolicy Bypass -File 2-Install-KCX-PaymentPro.ps1
    Requires 1-Install-IIS-Features.ps1 to have been run first.
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
#  STEP 1 - Install KCXPPNService via InstallUtil
# ============================================================
Write-Header "STEP 1: Installing KCXPPNService via InstallUtil"

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
#  STEP 2 - Start KCXPPNService
# ============================================================
Write-Header "STEP 2: Starting KCXPPNService"

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
    Write-Host "--- Services containing 'KCX' or 'PPN' ---" -ForegroundColor Yellow
    $matchedSvcs = Get-Service | Where-Object { $_.Name -match 'KCX|PPN' -or $_.DisplayName -match 'KCX|PPN' }
    if ($matchedSvcs) {
        $matchedSvcs | Format-Table Name, DisplayName, Status -AutoSize
    } else {
        Write-Host "  None found. The installer may have failed silently." -ForegroundColor Yellow
    }
    exit 1
}

if ($svc.Status -ne 'Running') {
    Start-Service -Name $serviceName
    Write-Host "[OK] Start command sent to '$serviceName'." -ForegroundColor Green
} else {
    Write-Host "[INFO] '$serviceName' is already running." -ForegroundColor Yellow
}

# ============================================================
#  STEP 3 - Verify KCXPPNService stays running
# ============================================================
Write-Header "STEP 3: Verifying KCXPPNService Remains Running"

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
#  STEP 4 - Create the PPN Scheduled Task
# ============================================================
Write-Header "STEP 4: Creating PPN Scheduled Task"

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

# Trigger - Daily at $triggerTime
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
