# Executive Exchange Online Group & Distribution List Report

A streamlined PowerShell solution designed for higher management, security audits, and IT governance.

It exports all:
1. **Standard Distribution Lists & Mail Security Groups**
2. **Microsoft 365 Groups (Teams & Unified Groups)**
3. **Dynamic Distribution Groups**

...into a **clean, decluttered, executive-ready Excel workbook (`.xlsx`)**.

---

## What Makes It Management-Ready?

* **No Technical Clutter:** Stripped out GUIDs, LDAP filter code, internal PowerShell class names (`MailUniversalDistributionGroup`, `UserMailbox`), and empty attribute columns.
* **Plain-English Business Terms:** Group and member types are translated to clear terminology:
  * `Distribution List`, `Mail Security Group`, `Microsoft 365 Group`
  * `Internal User`, `External Contact`, `Sub-Group`, `Guest User`
  * `Direct` vs `Nested (via Sub-Group)`
* **Master Group Directory:** One master inventory table containing every group, its member count, assigned owners, and security ingress status.

---

## Excel Workbook Structure

File: `M365_Group_and_DL_Report_<TenantName>_<Timestamp>.xlsx`

| Tab Order | Sheet Name | Target Audience | What It Shows |
| :---: | :--- | :--- | :--- |
| **1** | **`Executive_Summary`** | C-Level / Leadership | High-level KPI metrics card (Total Email Groups Managed, DLs, M365 Groups, Dynamic Groups) |
| **2** | **`Group_Directory`** | IT Management | Master catalog of all groups, types, member counts, assigned owners, and public email status |
| **3** | **`Distribution_Lists`** | Operations / Helpdesk | Clean member roster for standard DLs (Group Name, Member Name, Email, Direct/Nested) |
| **4** | **`Microsoft_365_Groups`** | Operations / Helpdesk | Clean member & owner roster for M365 Groups (Group Name, Privacy, Member, Role) |
| **5** | **`Dynamic_Groups`** | Operations / Helpdesk | Evaluated recipients for dynamic groups |

---

## How to Run

```powershell
cd C:\Users\NikhilMaity\projects\powershell-scripts
.\Export-TenantDistributionGroups.ps1
```

### Dry Run (Preview on Screen)
```powershell
.\Export-TenantDistributionGroups.ps1 -DryRun
```
* Shows the Master Group Directory preview directly in the PowerShell console without writing files.

### Optional Switches
* `-ExportCsv`: Also generates supplementary flat CSV files alongside the Excel workbook.
* `-ReuseExistingConnection`: Reuses your existing session without prompting at startup.
* `-KeepSessionConnected`: Leaves Exchange Online connected without prompting at the end.
* `-SkipM365Groups` or `-SkipDynamicGroups`: Excludes those groups if not needed.
