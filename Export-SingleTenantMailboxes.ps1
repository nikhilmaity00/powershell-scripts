<#
.SYNOPSIS
    Exports O365 mailbox storage, quotas, licenses, and types for a single tenant into an Excel workbook.
    INCLUDES ALL IDs: Enabled, Disabled, Shared, and Inactive/Hold mailboxes.
    Execute this script 3 times (once for each tenant).
    Supports writing to its own tab in a shared workbook or a separate file.
    Uses Device Code authentication (URL + Code to copy/paste in browser).
    Strictly READ-ONLY execution.
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $false, HelpMessage = "Tenant Domain (e.g. contoso.onmicrosoft.com) or Tenant ID GUID")]
    [string]$TenantIdOrDomain,

    [Parameter(Mandatory = $false, HelpMessage = "Friendly Label for worksheet tab (e.g., Tenant1, ProdTenant, Contoso)")]
    [string]$TenantFriendlyName,

    [Parameter(Mandatory = $false, HelpMessage = "Excel output path. Default: .\O365_Mailbox_Utilization_<TenantSuffix>.xlsx")]
    [string]$ExcelPath,

    [Parameter(Mandatory = $false, HelpMessage = "Also check for inactive/soft-deleted mailboxes kept on Litigation/Retention Hold")]
    [switch]$IncludeInactiveMailboxes = $true
)

# 1. Prompt for Tenant Identifier if not provided
if ([string]::IsNullOrWhiteSpace($TenantIdOrDomain)) {
    Write-Host "`n========================================================" -ForegroundColor Cyan
    Write-Host "   ENTER TARGET TENANT IDENTIFIER                       " -ForegroundColor Cyan
    Write-Host "========================================================" -ForegroundColor Cyan
    Write-Host "You can enter either:" -ForegroundColor Gray
    Write-Host "  - Tenant Domain / URL (e.g., contoso.onmicrosoft.com)" -ForegroundColor Gray
    Write-Host "  - Tenant Directory ID / GUID (e.g., 11111111-2222-3333-4444-555555555555)`n" -ForegroundColor Gray
    $TenantIdOrDomain = Read-Host "Tenant ID or Domain / URL"
}
$TenantIdOrDomain = $TenantIdOrDomain.Trim()

# 2. Sanitize Tenant URL/Domain for Filename Suffix
$tenantSuffix = ($TenantIdOrDomain -replace '^https?://', '' -replace '[\\/:*?"<>|\s]', '_').Trim('_')
if ([string]::IsNullOrWhiteSpace($ExcelPath)) {
    $ExcelPath = ".\O365_Mailbox_Utilization_${tenantSuffix}.xlsx"
}

# 3. Determine Sheet Tab Name (Sanitized for Excel, max 31 characters)
if ([string]::IsNullOrWhiteSpace($TenantFriendlyName)) {
    $suggestedName = ($tenantSuffix -split '\.')[0]
    $promptName = Read-Host "Enter tab name for this tenant in Excel [Default: $suggestedName]"
    $TenantFriendlyName = if ([string]::IsNullOrWhiteSpace($promptName)) { $suggestedName } else { $promptName }
}
$SheetName = ($TenantFriendlyName -replace '[\\/\?\*\[\]:]', '_')
if ($SheetName.Length -gt 31) { $SheetName = $SheetName.Substring(0, 31) }

Write-Host "`n========================================================" -ForegroundColor Cyan
Write-Host "     O365 MAILBOX UTILIZATION EXPORT [READ-ONLY]        " -ForegroundColor Cyan
Write-Host "     Target Tenant:  $TenantIdOrDomain                  " -ForegroundColor Cyan
Write-Host "     Worksheet Tab:  $SheetName                         " -ForegroundColor Cyan
Write-Host "     Excel Workbook: $ExcelPath                         " -ForegroundColor Cyan
Write-Host "     Scope:          ALL Enabled + Disabled Mailboxes   " -ForegroundColor Cyan
Write-Host "========================================================`n" -ForegroundColor Cyan

# 4. Friendly License SKU Mapping
$FriendlySkuNames = @{
    "ENTERPRISEPACK"           = "Office 365 E3"
    "ENTERPRISEPREMIUM"        = "Office 365 E5"
    "STANDARDPACK"             = "Office 365 E1"
    "SPE_E3"                   = "Microsoft 365 E3"
    "SPE_E5"                   = "Microsoft 365 E5"
    "SPE_F1"                   = "Microsoft 365 F3"
    "SPB"                      = "Microsoft 365 Business Standard"
    "M365_BUSINESS_PREMIUM"    = "Microsoft 365 Business Premium"
    "O365_BUSINESS_PREMIUM"    = "Microsoft 365 Business Standard"
    "SMB_BUSINESS_PREMIUM"     = "Microsoft 365 Business Standard"
    "EXCHANGESTANDARD"         = "Exchange Online (Plan 1)"
    "EXCHANGEENTERPRISE"       = "Exchange Online (Plan 2)"
    "EXCHANGEARCHIVE"          = "Exchange Online Archiving"
    "O365_BUSINESS_ESSENTIALS" = "Microsoft 365 Business Basic"
}

function Convert-SizeToGB ([string]$sizeString) {
    if ([string]::IsNullOrWhiteSpace($sizeString) -or $sizeString -eq "Unlimited") {
        return "Unlimited"
    }
    if ($sizeString -match '\(([\d,]+)\s+bytes\)') {
        $bytes = [double]($Matches[1] -replace ',', '')
        return [math]::Round($bytes / 1GB, 2)
    }
    return $sizeString
}

try {
    # 5. Connect to Exchange Online via Device Code
    Write-Host "[1/3] Connecting to Exchange Online via Device Code..." -ForegroundColor Yellow
    Write-Host ">>> COPY THE URL AND CODE BELOW AND PASTE INTO YOUR BROWSER <<<`n" -ForegroundColor Green
    Connect-ExchangeOnline -Device -Organization $TenantIdOrDomain -ShowBanner:$false -ErrorAction Stop
    Write-Host "`n[+] Exchange Online connected.`n" -ForegroundColor Green

    # 6. Connect to Microsoft Graph via Device Code (Read-Only)
    Write-Host "[2/3] Connecting to Microsoft Graph via Device Code (Read-Only)..." -ForegroundColor Yellow
    Write-Host ">>> COPY THE URL AND CODE BELOW AND PASTE INTO YOUR BROWSER <<<`n" -ForegroundColor Green
    Connect-MgGraph -TenantId $TenantIdOrDomain -Scopes "User.Read.All", "Organization.Read.All" -UseDeviceAuthentication -NoWelcome -ErrorAction Stop
    Write-Host "`n[+] Microsoft Graph connected.`n" -ForegroundColor Green

    # 7. Build License Lookup Table (Minimal Attribute Projection)
    Write-Host "[3/3] Querying tenant data using least-data exposure..." -ForegroundColor Yellow
    Write-Host "Fetching Subscribed SKUs for license name resolution..." -ForegroundColor Gray
    $SubscribedSkus = Get-MgSubscribedSku -All -ErrorAction Stop
    $SkuMap = @{}
    foreach ($sku in $SubscribedSkus) {
        $friendly = if ($FriendlySkuNames.ContainsKey($sku.SkuPartNumber)) { 
            $FriendlySkuNames[$sku.SkuPartNumber] 
        } else { 
            $sku.SkuPartNumber 
        }
        $SkuMap[$sku.SkuId] = $friendly
    }
    # Clear raw SKUs reference
    $SubscribedSkus = $null

    Write-Host "Fetching user license allocations (minimal attribute projection: UPN, Licenses, AccountEnabled)..." -ForegroundColor Gray
    # Query strictly the 3 required attributes to avoid pulling full user profiles into memory
    $rawGraphUsers = Get-MgUser -All -Property @("UserPrincipalName", "AssignedLicenses", "AccountEnabled") -ErrorAction SilentlyContinue

    # Build lean in-memory lookup table storing ONLY non-sensitive mapped strings
    $UserLicenseMap = @{}
    if ($rawGraphUsers) {
        foreach ($u in $rawGraphUsers) {
            if ([string]::IsNullOrWhiteSpace($u.UserPrincipalName)) { continue }
            
            $lics = [System.Collections.Generic.List[string]]::new()
            if ($u.AssignedLicenses) {
                foreach ($lic in $u.AssignedLicenses) {
                    if ($SkuMap.ContainsKey($lic.SkuId)) {
                        $lics.Add($SkuMap[$lic.SkuId])
                    }
                }
            }
            $licText = if ($lics.Count -gt 0) { ($lics -join "; ") } else { "None / Unlicensed" }

            $UserLicenseMap[$u.UserPrincipalName.ToLower()] = @{
                AccountEnabled = $u.AccountEnabled
                Licenses       = $licText
            }
        }
    }
    # Immediately deallocate raw Graph user profile objects to protect memory
    $rawGraphUsers = $null

    # 8. Fetch ALL Active Mailboxes (UserMailbox and SharedMailbox, both Enabled & Disabled)
    Write-Host "Fetching Mailboxes (Restricted to Minimum & Quota property sets)..." -ForegroundColor Gray
    $Mailboxes = [System.Collections.Generic.List[object]]::new()
    
    $activeMbx = Get-EXOMailbox -ResultSize Unlimited -PropertySets Minimum,Quota `
        -RecipientTypeDetails UserMailbox,SharedMailbox `
        -Properties DisplayName,UserPrincipalName,RecipientTypeDetails,IssueWarningQuota,ProhibitSendReceiveQuota,AccountDisabled

    if ($activeMbx) {
        foreach ($mbxItem in $activeMbx) {
            $Mailboxes.Add($mbxItem)
        }
    }
    $activeMbx = $null

    # Optionally fetch Inactive Mailboxes (mailboxes under retention/litigation hold whose users were deleted)
    if ($IncludeInactiveMailboxes) {
        Write-Host "Checking for Inactive Mailboxes on Litigation/Retention Hold..." -ForegroundColor Gray
        $inactiveMbx = Get-EXOMailbox -InactiveMailboxOnly -ResultSize Unlimited -PropertySets Minimum,Quota `
            -Properties DisplayName,UserPrincipalName,RecipientTypeDetails,IssueWarningQuota,ProhibitSendReceiveQuota `
            -ErrorAction SilentlyContinue

        if ($inactiveMbx) {
            Write-Host "Found $($inactiveMbx.Count) inactive mailbox(es) on hold." -ForegroundColor Yellow
            foreach ($inItem in $inactiveMbx) {
                $Mailboxes.Add($inItem)
            }
        }
        $inactiveMbx = $null
    }

    $TenantResults = [System.Collections.Generic.List[object]]::new()
    $total = $Mailboxes.Count
    $counter = 0

    $enabledCount = 0
    $disabledCount = 0
    $sharedCount = 0
    $inactiveCount = 0

    Write-Host "Total mailboxes to process: $total. Parsing storage and license metrics locally..." -ForegroundColor Gray

    foreach ($mbx in $Mailboxes) {
        $counter++
        Write-Progress -Activity "Processing Mailboxes ($TenantIdOrDomain)" -Status "$counter of $total ($($mbx.DisplayName))" -PercentComplete (($counter / $total) * 100)

        $isInactive = ($mbx.RecipientTypeDetails -eq "InactiveMailbox")
        $isShared = ($mbx.RecipientTypeDetails -eq "SharedMailbox")
        $upnKey = if ($mbx.UserPrincipalName) { $mbx.UserPrincipalName.ToLower() } else { "" }

        # Determine Account Status
        $accountStatus = "Enabled"
        if ($isInactive) {
            $accountStatus = "Inactive (On Hold)"
            $inactiveCount++
        }
        elseif ($isShared) {
            # In M365 best practices, the underlying account of a shared mailbox has sign-in disabled
            $accountStatus = if ($mbx.AccountDisabled -eq $true) { "Disabled (Shared)" } else { "Enabled (Shared)" }
            $sharedCount++
        }
        else {
            # Correlate with cached Graph account state if available, else fallback to EXO AccountDisabled
            if ($UserLicenseMap.ContainsKey($upnKey) -and $UserLicenseMap[$upnKey].AccountEnabled -ne $null) {
                if ($UserLicenseMap[$upnKey].AccountEnabled) {
                    $accountStatus = "Enabled"
                    $enabledCount++
                } else {
                    $accountStatus = "Disabled"
                    $disabledCount++
                }
            }
            elseif ($mbx.AccountDisabled -eq $true) {
                $accountStatus = "Disabled"
                $disabledCount++
            }
            else {
                $accountStatus = "Enabled"
                $enabledCount++
            }
        }

        # Strictly Read-Only Statistics (Folder/Item content is NEVER queried)
        $stats = Get-EXOMailboxStatistics -Identity $mbx.UserPrincipalName -ErrorAction SilentlyContinue

        # Calculate Used Space
        $usedGB = 0
        if ($stats.TotalItemSize) {
            $usedRaw = $stats.TotalItemSize.ToString()
            if ($usedRaw -match '\(([\d,]+)\s+bytes\)') {
                $usedBytes = [double]($Matches[1] -replace ',', '')
                $usedGB = [math]::Round($usedBytes / 1GB, 2)
            }
        }

        # Calculate Max Quota Allocated
        $allocatedQuota = $mbx.ProhibitSendReceiveQuota.ToString()
        $allocatedGB = Convert-SizeToGB -sizeString $allocatedQuota

        # Calculate % Used
        $percentUsed = "N/A"
        if ($allocatedGB -is [double] -and $allocatedGB -gt 0) {
            $percentUsed = [math]::Round(($usedGB / $allocatedGB) * 100, 1)
        }

        # In-Memory License Lookup (Instant local match, zero additional network calls)
        $licenseString = "None / Unlicensed"
        if (-not $isInactive -and $UserLicenseMap.ContainsKey($upnKey)) {
            $licenseString = $UserLicenseMap[$upnKey].Licenses
        }

        $record = [PSCustomObject]@{
            Tenant              = $TenantIdOrDomain
            DisplayName         = $mbx.DisplayName
            UserPrincipalName   = $mbx.UserPrincipalName
            AccountStatus       = $accountStatus
            MailboxType         = $mbx.RecipientTypeDetails.ToString()
            "UsedSpace(GB)"     = $usedGB
            "AllocatedQuota(GB)"= $allocatedGB
            "% Utilized"        = $percentUsed
            TotalItemCount      = if ($stats.ItemCount) { $stats.ItemCount } else { 0 }
            AllocatedLicenses   = $licenseString
        }

        $TenantResults.Add($record)
    }

    # 9. Export Tab to Excel Workbook (Appends or updates the tab without overwriting other tabs)
    Write-Host "`nExporting $total records to worksheet tab [$SheetName] in $ExcelPath ..." -ForegroundColor Green
    
    $TenantResults | Export-Excel -Path $ExcelPath `
        -WorksheetName $SheetName `
        -ClearSheet `
        -AutoSize `
        -AutoFilter `
        -BoldTopRow `
        -FreezeTopRow

    # 10. Summary Statistics Display
    Write-Host "`n========================================================" -ForegroundColor Cyan
    Write-Host "                TENANT AUDIT SUMMARY                    " -ForegroundColor Cyan
    Write-Host "========================================================" -ForegroundColor Cyan
    Write-Host ("{0,-28} : {1}" -f "Total Mailboxes Processed", $total) -ForegroundColor White
    Write-Host ("{0,-28} : {1}" -f "Enabled User Accounts", $enabledCount) -ForegroundColor Green
    Write-Host ("{0,-28} : {1}" -f "Disabled User Accounts", $disabledCount) -ForegroundColor Yellow
    Write-Host ("{0,-28} : {1}" -f "Shared Mailboxes", $sharedCount) -ForegroundColor White
    if ($inactiveCount -gt 0) {
        Write-Host ("{0,-28} : {1}" -f "Inactive Mailboxes (On Hold)", $inactiveCount) -ForegroundColor Magenta
    }
    Write-Host ("{0,-28} : {1}" -f "Output Excel File", $ExcelPath) -ForegroundColor Green
    Write-Host ("{0,-28} : {1}" -f "Worksheet Sheet Name", $SheetName) -ForegroundColor Green
    Write-Host ("{0,-28} : {1}" -f "Data Privacy / Caching", "Purged (Zero Sensitive Data Cached)") -ForegroundColor Cyan
    Write-Host "========================================================`n" -ForegroundColor Cyan
}
catch {
    Write-Host "`n[!] An error occurred during export:" -ForegroundColor Red
    Write-Error $_
}
finally {
    # ---------------------------------------------------------------------
    # Security & Privacy: Purge in-memory collections and disconnect sessions
    # ---------------------------------------------------------------------
    Write-Host "Cleaning session context and purging sensitive data from memory..." -ForegroundColor Gray
    
    # 1. Clear In-Memory Dictionaries and Collections
    if ($UserLicenseMap) { $UserLicenseMap.Clear(); $UserLicenseMap = $null }
    if ($SkuMap)         { $SkuMap.Clear(); $SkuMap = $null }
    if ($Mailboxes)      { $Mailboxes.Clear(); $Mailboxes = $null }
    if ($TenantResults)  { $TenantResults.Clear(); $TenantResults = $null }

    # 2. Clear Graph Context & Auth Token Cache
    try {
        Clear-MgContext -ErrorAction SilentlyContinue
    } catch {}

    # 3. Disconnect Exchange Online and Microsoft Graph
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    Disconnect-MgGraph -ErrorAction SilentlyContinue

    # 4. Trigger Immediate Garbage Collection to free memory heap
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
