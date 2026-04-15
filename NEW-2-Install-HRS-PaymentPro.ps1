#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Step 2 of 2 - Configures KCX PaymentPro.Net.
.DESCRIPTION
    1. Checks for and removes if existing KCXPPNService
    2. Installs KCXPPNService via InstallUtil
    3. Starts KCXPPNService and verifies it stays running
    4. Creates a dedicated local service account (svc_ppn)
    5. Grants svc_ppn "Log on as a batch job" rights
    6. Grants svc_ppn read/execute on the PPN folder
    7. Creates the PPN Autosettlement Scheduled Task in Task Scheduler
.NOTES
    Must be run as Administrator.
    Run with: PowerShell.exe -ExecutionPolicy Bypass -File 2-Install-HRS-PaymentPro.ps1
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
#  Helper: Grant "Log on as a batch job" via secedit
# ============================================================
function Grant-LogOnAsBatchJob {
    param([string]$Username)

    $tempInf = "$env:TEMP\svc_ppn_batch.inf"
    $tempDb  = "$env:TEMP\svc_ppn_batch.sdb"
    $tempLog = "$env:TEMP\svc_ppn_batch.log"

    # Export current policy
    secedit /export /cfg $tempInf /quiet

    $infContent = Get-Content $tempInf

    # Check if SeBatchLogonRight already exists in the policy
    $batchLine = $infContent | Where-Object { $_ -match 'SeBatchLogonRight' }

    if ($batchLine) {
        # Right exists - append the user if not already there
        if ($batchLine -notmatch [regex]::Escape($Username)) {
            $infContent = $infContent -replace `
                '(SeBatchLogonRight\s*=\s*.*)', `
                "`$1,*$Username"
        } else {
            Write-Host "[INFO] '$Username' already has Log on as a batch job rights." -ForegroundColor Gray
            Remove-Item $tempInf, $tempDb, $tempLog -ErrorAction SilentlyContinue
            return
        }
    } else {
        # Right does not exist yet - add it under [Privilege Rights]
        $infContent = $infContent -replace `
            '\[Privilege Rights\]', `
            "[Privilege Rights]`r`nSeBatchLogonRight = *$Username"
    }

    $infContent | Set-Content $tempInf

    # Apply the updated policy
    secedit /configure /db $tempDb /cfg $tempInf /log $tempLog /quiet

    Remove-Item $tempInf, $tempDb, $tempLog -ErrorAction SilentlyContinue
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

# Check if the service is already installed and remove it first
$existingSvc = Get-Service -Name 'KCXPPNService' -ErrorAction SilentlyContinue
if ($null -ne $existingSvc) {
    Write-Host "[INFO] Existing KCXPPNService detected - removing before reinstall..." -ForegroundColor Yellow

    # Stop the service if it is running
    if ($existingSvc.Status -eq 'Running') {
        Write-Host "  Stopping service..."
        Stop-Service -Name 'KCXPPNService' -Force
        Start-Sleep -Seconds 3
    }

    # Uninstall via InstallUtil /u
    Write-Host "  Uninstalling existing service..."
    & $installUtil /u /LogFile= /LogToConsole=true "$serviceExe"

    if ($LASTEXITCODE -eq 0) {
        Write-Host "[OK] Existing service removed successfully." -ForegroundColor Green
    } else {
        Write-Host "[WARN] Uninstall exited with code $LASTEXITCODE - attempting to continue..." -ForegroundColor Yellow
    }

    # Brief pause to let the SCM release the service entry
    Start-Sleep -Seconds 3
} else {
    Write-Host "[INFO] No existing KCXPPNService found - proceeding with fresh install." -ForegroundColor Gray
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
#  STEP 4 - Create Dedicated Service Account
# ============================================================
Write-Header "STEP 4: Creating Dedicated Service Account (svc_ppn)"

$svcAccountName = 'svc_ppn'
$ppnFolder      = 'C:\Program Files\KCX PaymentPro.Net'

# Prompt for password - never stored as plain text
$svcPassword = Read-Host "Enter password for '$svcAccountName'" -AsSecureString
$svcPasswordConfirm = Read-Host "Confirm password for '$svcAccountName'" -AsSecureString

# Compare the two passwords
$pwd1Plain = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
    [Runtime.InteropServices.Marshal]::SecureStringToBSTR($svcPassword))
$pwd2Plain = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
    [Runtime.InteropServices.Marshal]::SecureStringToBSTR($svcPasswordConfirm))

if ($pwd1Plain -ne $pwd2Plain) {
    Write-Host "[ERROR] Passwords do not match. Exiting." -ForegroundColor Red
    exit 1
}

# Clear the plain text comparison variable immediately
$pwd1Plain = $null
$pwd2Plain = $null

# Remove existing account if present for a clean install
$existingAccount = Get-LocalUser -Name $svcAccountName -ErrorAction SilentlyContinue
if ($null -ne $existingAccount) {
    Write-Host "[INFO] Existing '$svcAccountName' account found - removing before recreating..." -ForegroundColor Yellow
    Remove-LocalUser -Name $svcAccountName
    Write-Host "[OK] Existing account removed." -ForegroundColor Green
}

# Create the account
New-LocalUser `
    -Name                  $svcAccountName `
    -Password              $svcPassword `
    -PasswordNeverExpires  $true `
    -UserMayNotChangePassword $true `
    -Description           "PPN AutoSettlement Service Account - do not add to any groups"

Write-Host "[OK] Local account '$svcAccountName' created successfully." -ForegroundColor Green

# ============================================================
#  STEP 5 - Grant "Log on as a batch job" Rights
# ============================================================
Write-Header "STEP 5: Granting 'Log on as a batch job' Rights"

Grant-LogOnAsBatchJob -Username $svcAccountName
Write-Host "[OK] 'Log on as a batch job' rights granted to '$svcAccountName'." -ForegroundColor Green

# ============================================================
#  STEP 6 - Grant Read/Execute on the PPN Folder
# ============================================================
Write-Header "STEP 6: Granting Folder Permissions to '$svcAccountName'"

if (-not (Test-Path $ppnFolder)) {
    Write-Host "[ERROR] PPN folder not found at: $ppnFolder" -ForegroundColor Red
    exit 1
}

$acl  = Get-Acl $ppnFolder
$rule = New-Object System.Security.AccessControl.FileSystemAccessRule(
    $svcAccountName,
    "ReadAndExecute",
    "ContainerInherit,ObjectInherit",
    "None",
    "Allow"
)
$acl.AddAccessRule($rule)
Set-Acl $ppnFolder $acl

Write-Host "[OK] Read/Execute permission granted on '$ppnFolder'." -ForegroundColor Green

# ============================================================
#  STEP 7 - Create the PPN Scheduled Task
# ============================================================
Write-Header "STEP 7: Creating PPN Scheduled Task"

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

# Principal - dedicated service account, no admin rights needed
$principal = New-ScheduledTaskPrincipal `
    -UserId    $svcAccountName `
    -LogonType Password `
    -RunLevel  Normal

# Settings
$settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -StartWhenAvailable `
    -ExecutionTimeLimit (New-TimeSpan -Days 3) `
    -MultipleInstances  IgnoreNew

# Convert SecureString password to plain text only for task registration
# then immediately clear it
$taskPassword = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
    [Runtime.InteropServices.Marshal]::SecureStringToBSTR($svcPassword))

try {
    Register-ScheduledTask `
        -TaskName  $taskName `
        -Action    $action `
        -Trigger   $trigger `
        -Principal $principal `
        -Settings  $settings `
        -Password  $taskPassword `
        -Force | Out-Null

    Write-Host "[OK] Scheduled task '$taskName' created successfully." -ForegroundColor Green
} catch {
    Write-Host "[ERROR] Failed to register scheduled task: $_" -ForegroundColor Red
    exit 1
} finally {
    # Always clear the plain text password from memory
    $taskPassword = $null
    $svcPassword  = $null
}

Write-Host ""
Write-Host "  Name          : PPN"
Write-Host "  Run as        : $svcAccountName (dedicated service account)"
Write-Host "  Trigger       : Daily at $triggerTime"
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
