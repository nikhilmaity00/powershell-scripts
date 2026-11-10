# PowerShell Scripts

A collection of handy PowerShell automation and utility scripts for Microsoft 365, Exchange Online, SaaS platforms, and general IT administration.

---

## Available Scripts

### 1. `Export-TenantDistributionGroups.ps1`
Exports all Distribution Lists, Mail-Enabled Security Groups, Microsoft 365 Groups, and Dynamic Distribution Groups into a clean, multi-tab Excel workbook (`.xlsx`).

* **Features:**
  * Interactive Device Code login (`https://microsoft.com/devicelogin`) for effortless multi-tenant switching.
  * Detects and reuses active sessions; prompts to keep or close session on exit.
  * Recursively resolves nested child groups.
  * Formats clean management tabs: `Executive_Summary`, `Group_Directory`, `Distribution_Lists`, `Microsoft_365_Groups`, and `Dynamic_Groups`.
* **Usage:**
  ```powershell
  # Full export
  .\Export-TenantDistributionGroups.ps1

  # Dry-run test (preview on screen)
  .\Export-TenantDistributionGroups.ps1 -DryRun
  ```

---

### 2. `Sync-LicenseInventory.ps1`
Synchronizes SaaS software licenses into an Excel tracker from multiple cloud vendor APIs (Microsoft Graph, Adobe Admin API, Zscaler, Barracuda, IAM Cloud, CloudPBX).

* **Usage:**
  ```powershell
  .\Sync-LicenseInventory.ps1 -ExcelPath "C:\Path\To\Enterprise_Software_License_Tracker.xlsx"
  ```

---

## Requirements & Portability

* PowerShell 5.1 or PowerShell 7+ on Windows.
* Scripts are self-bootstrapping and automatically install any required modules (`ExchangeOnlineManagement`, `ImportExcel`) into the `CurrentUser` scope if missing.
