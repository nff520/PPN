# PPN Installation Instructions

Follow the steps below to install PPN correctly.

## Before You Begin

Make sure the following are installed on the SQL Server or PPN server before starting:

- Microsoft SQL Server
- SQL Server Management Studio (SSMS)

## Required Files

All of the following files must remain in the **same folder**:

- `1-Install-IIS-Features.bat`
- `1-Install-IIS-Features.ps1`
- `2-Install-HRS-PaymentPro.bat`
- `2-Install-HRS-PaymentPro.ps1`

## Installation Steps

### Step 1: Install PPN Prerequisites
Right-click `1-Install-IIS-Features.bat` and select **Run as Administrator**.

This will install the IIS features and other prerequisites required for PPN.

### Step 2: Install PPN
Run `PPN_Setup.exe` and complete the installation.

### Step 3: Finalize the Installation
Right-click `2-Install-HRS-PaymentPro.bat` and select **Run as Administrator**.

This will:
- finalize the PPN installation
- schedule the Autosettlement using Task Scheduler

## Important Notes
- Adjust the Autosettlement task name on line 146 & trigger time on line 148.
