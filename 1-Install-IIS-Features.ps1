#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Step 1 of 2 - Installs IIS, .NET 3.5, ASP.NET 3.5/4.8 and related features via DISM.
.NOTES
    Must be run as Administrator.
    Run with: PowerShell.exe -ExecutionPolicy Bypass -File 1-Install-IIS-Features.ps1
    Run 2-Install-KCX-PaymentPro.ps1 after this script completes.
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
    dism /online /enable-feature /featurename:$feature /all /norestart
    if ($LASTEXITCODE -eq 0 -or $LASTEXITCODE -eq 3010) {
        Write-Host "[OK] $feature" -ForegroundColor Green
    } else {
        Write-Host "[WARN] $feature exited with code $LASTEXITCODE" -ForegroundColor Yellow
    }
}

Write-Host ""
Write-Host "[OK] IIS and related features installed." -ForegroundColor Green

# ============================================================
#  Remove the Default IIS Website
# ============================================================
Write-Header "Removing Default IIS Website"

Import-Module WebAdministration -ErrorAction SilentlyContinue

if (Test-Path 'IIS:\Sites\Default Web Site') {
    Remove-WebSite -Name 'Default Web Site'
    Write-Host "[OK] Default Web Site removed." -ForegroundColor Green
} else {
    Write-Host "[INFO] Default Web Site not found - skipping." -ForegroundColor Yellow
}

Write-Host ""
Write-Host "============================================================" -ForegroundColor Green
Write-Host "  IIS features installed and configured successfully." -ForegroundColor Green
Write-Host "  You may now install the PPN software." -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green
