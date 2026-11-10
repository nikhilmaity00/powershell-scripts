<#
.SYNOPSIS
    Executive-Ready Exchange Online Group & Distribution List Exporter.
    Exports all Distribution Lists, Mail-Enabled Security Groups, Microsoft 365 Groups,
    and Dynamic Distribution Groups into a clean, decluttered Excel workbook.

.DESCRIPTION
    Designed for higher management and IT governance audits:
    - Decluttered presentation: strips out raw Exchange technical details (GUIDs, LDAP filters, 
      PowerShell class names) and replaces them with clear, human-readable terminology.
    - Master Group Directory: one comprehensive view of all groups, owners, and external email status.
    - Clean Member Rosters: simplified columns (Group Name, Email, Member Name, Member Email, Role, Direct vs Nested).
    - Multi-tenant ready with Device Code authentication, session reuse, and non-intrusive disconnects.

.PARAMETER Organization
    Optional tenant domain (e.g. contoso.onmicrosoft.com) or Tenant GUID.

.PARAMETER OutputDir
    Directory path to save exported Excel workbook, CSVs, and logs. Defaults to current directory.

.PARAMETER DryRun
    Simulates execution without generating files. Lists discovered groups, counts, and security settings.

.PARAMETER SkipM365Groups
    Switch to skip querying Microsoft 365 unified groups.

.PARAMETER SkipDynamicGroups
    Switch to skip querying dynamic distribution groups.

.PARAMETER ExportCsv
    Switch to also generate separate CSV files alongside the Excel workbook.

.PARAMETER ReuseExistingConnection
    Automatically reuse an active Exchange Online connection without prompting.

.PARAMETER KeepSessionConnected
    Keep Exchange Online session connected after script completes without prompting.

.PARAMETER DisconnectSession
    Force disconnect Exchange Online session after script completes without prompting.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false, HelpMessage = "Optional Tenant domain (e.g. contoso.onmicrosoft.com) or Tenant GUID")]
    [string]$Organization,

    [Parameter(Mandatory = $false, HelpMessage = "Output directory for Excel workbook, CSVs, and logs")]
    [string]$OutputDir = (Get-Location).Path,

    [Parameter(Mandatory = $false, HelpMessage = "Test run mode: discover groups and counts without exporting files")]
    [switch]$DryRun,

    [Parameter(Mandatory = $false, HelpMessage = "Skip Microsoft 365 unified groups discovery")]
    [switch]$SkipM365Groups,

    [Parameter(Mandatory = $false, HelpMessage = "Skip dynamic distribution group discovery")]
    [switch]$SkipDynamicGroups,

    [Parameter(Mandatory = $false, HelpMessage = "Also export individual CSV files alongside the Excel workbook")]
    [switch]$ExportCsv,

    [Parameter(Mandatory = $false, HelpMessage = "Automatically reuse an existing active connection without prompting")]
    [switch]$ReuseExistingConnection,

    [Parameter(Mandatory = $false, HelpMessage = "Keep Exchange Online session connected after script completes")]
    [switch]$KeepSessionConnected,

    [Parameter(Mandatory = $false, HelpMessage = "Force disconnect Exchange Online session after script completes")]
    [switch]$DisconnectSession
)

# Set strict error behavior
$ErrorActionPreference = "Stop"

# Script-scoped log path
$script:LogFilePath = $null

# --- Function: Unified Console & File Logger ---
function Write-Log {
    param(
        [Parameter(Mandatory = $false)]
        [AllowEmptyString()]
        [string]$Message = "",

        [ValidateSet("INFO", "WARN", "ERROR", "SUCCESS", "DRYRUN")]
        [string]$Level = "INFO"
    )
    if ([string]::IsNullOrWhiteSpace($Message)) { return }

    $timeStr = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $formattedLog = "[$timeStr] [$Level] $Message"

    switch ($Level) {
        "INFO"    { Write-Host $formattedLog -ForegroundColor Cyan }
        "SUCCESS" { Write-Host $formattedLog -ForegroundColor Green }
        "WARN"    { Write-Host $formattedLog -ForegroundColor Yellow }
        "ERROR"   { Write-Host $formattedLog -ForegroundColor Red }
        "DRYRUN"  { Write-Host $formattedLog -ForegroundColor Magenta }
    }

    if ($script:LogFilePath) {
        try {
            Add-Content -Path $script:LogFilePath -Value $formattedLog -ErrorAction SilentlyContinue
        } catch { }
    }
}

# --- Management Decluttering Helpers ---
function Get-FriendlyRecipientType {
    param(
        [string]$RawType,
        [string]$Details
    )
    $combined = "$RawType $Details"
    if ($combined -match "MailContact|Contact") { return "External Contact" }
    if ($combined -match "Group|SecurityGroup|DistributionGroup") { return "Sub-Group" }
    if ($combined -match "Guest") { return "Guest User" }
    return "Internal User"
}

function Get-FriendlyGroupType {
    param([string]$RawType)
    switch -Regex ($RawType) {
        "MailUniversalSecurityGroup"     { return "Mail Security Group" }
        "MailUniversalDistributionGroup" { return "Distribution List" }
        "GroupMailbox|UnifiedGroup"      { return "Microsoft 365 Group" }
        "Dynamic"                        { return "Dynamic Distribution Group" }
        default                          { return "Distribution List" }
    }
}

function Get-CleanManagers {
    param($ManagedBy)
    if (-not $ManagedBy) { return "Unassigned" }
    $names = [System.Collections.Generic.List[string]]::new()
    foreach ($item in $ManagedBy) {
        $str = $item.ToString()
        if ($str -match '/([^/]+)$') {
            $names.Add($matches[1])
        } else {
            $names.Add($str)
        }
    }
    if ($names.Count -eq 0) { return "Unassigned" }
    return ($names -join ", ")
}

# --- Function: Self-Bootstrap Environment & Modules ---
function Assert-Prerequisites {
    Write-Log "=== [1/6] Checking Environment & Modules ===" "INFO"
    
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

    # 1. NuGet PackageProvider
    $nuget = Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue
    if (-not $nuget -or ($nuget.Version -lt [version]"2.8.5.201")) {
        Write-Log "Installing/updating NuGet package provider..." "WARN"
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser | Out-Null
    }

    # 2. ExchangeOnlineManagement
    $exoModule = Get-Module -ListAvailable -Name ExchangeOnlineManagement
    if (-not $exoModule) {
        Write-Log "ExchangeOnlineManagement module not found. Installing into CurrentUser scope..." "WARN"
        Set-PSRepository -Name "PSGallery" -InstallationPolicy Trusted -ErrorAction SilentlyContinue
        Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber -Repository PSGallery
        Write-Log "ExchangeOnlineManagement installed successfully." "SUCCESS"
    } else {
        Write-Log "ExchangeOnlineManagement module is available (Version: $($exoModule[0].Version))." "SUCCESS"
    }
    Import-Module ExchangeOnlineManagement -ErrorAction Stop

    # 3. ImportExcel
    $excelModule = Get-Module -ListAvailable -Name ImportExcel
    if (-not $excelModule) {
        Write-Log "ImportExcel module not found. Installing into CurrentUser scope..." "WARN"
        Set-PSRepository -Name "PSGallery" -InstallationPolicy Trusted -ErrorAction SilentlyContinue
        Install-Module -Name ImportExcel -Scope CurrentUser -Force -AllowClobber -Repository PSGallery
        Write-Log "ImportExcel installed successfully." "SUCCESS"
    } else {
        Write-Log "ImportExcel module is available (Version: $($excelModule[0].Version))." "SUCCESS"
    }
    Import-Module ImportExcel -ErrorAction Stop
}

# --- Function: Recursive Standard Group Member Resolver ---
function Get-RecursiveGroupMembers {
    param(
        [Parameter(Mandatory = $true)]
        [object]$RootGroup,

        [Parameter(Mandatory = $true)]
        [object]$CurrentGroup,

        [Parameter(Mandatory = $false)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.HashSet[string]]$VisitedGroupIds = $null,

        [string]$Path = ""
    )

    if ($null -eq $VisitedGroupIds) {
        $VisitedGroupIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    }

    $groupIdKey = if ($CurrentGroup.ExternalDirectoryObjectId) {
        $CurrentGroup.ExternalDirectoryObjectId.ToString().ToLower()
    } elseif ($CurrentGroup.Identity) {
        $CurrentGroup.Identity.ToString().ToLower()
    } else {
        $CurrentGroup.PrimarySmtpAddress.ToString().ToLower()
    }

    # Circular reference protection
    if (-not $VisitedGroupIds.Add($groupIdKey)) {
        Write-Log "Circular membership detected for '$($CurrentGroup.DisplayName)'. Skipping recursion." "WARN"
        return
    }

    $currentPath = if ($Path) { "$Path -> $($CurrentGroup.DisplayName)" } else { $CurrentGroup.DisplayName }

    try {
        $members = Get-DistributionGroupMember -Identity $CurrentGroup.Identity -ResultSize Unlimited -ErrorAction Stop
    }
    catch {
        Write-Log "Could not fetch members for group '$($CurrentGroup.DisplayName)': $_" "WARN"
        return
    }

    if (-not $members -or $members.Count -eq 0) {
        if ($CurrentGroup.Identity -eq $RootGroup.Identity) {
            # Top-level group has no members
            [PSCustomObject]@{
                "Group Name"   = $RootGroup.DisplayName
                "Group Email"  = $RootGroup.PrimarySmtpAddress
                "Group Type"   = (Get-FriendlyGroupType $RootGroup.RecipientTypeDetails)
                "Member Name"  = "<No Members>"
                "Member Email" = ""
                "Member Type"  = ""
                "Membership"   = "Direct"
            }
        }
        return
    }

    foreach ($member in $members) {
        $isSubGroup = ($member.RecipientType -match "Group") -or ($member.RecipientTypeDetails -match "Group")

        $friendlyMemberType = Get-FriendlyRecipientType -RawType $member.RecipientType -Details $member.RecipientTypeDetails
        $membershipDetail = if ($CurrentGroup.Identity -eq $RootGroup.Identity) {
            "Direct"
        } else {
            "Nested (via $($CurrentGroup.DisplayName))"
        }

        # Emit clean member record
        [PSCustomObject]@{
            "Group Name"   = $RootGroup.DisplayName
            "Group Email"  = $RootGroup.PrimarySmtpAddress
            "Group Type"   = (Get-FriendlyGroupType $RootGroup.RecipientTypeDetails)
            "Member Name"  = $member.DisplayName
            "Member Email" = $member.PrimarySmtpAddress
            "Member Type"  = $friendlyMemberType
            "Membership"   = $membershipDetail
        }

        # If child is a group, recurse
        if ($isSubGroup) {
            Get-RecursiveGroupMembers -RootGroup $RootGroup `
                                      -CurrentGroup $member `
                                      -VisitedGroupIds $VisitedGroupIds `
                                      -Path $currentPath
        }
    }
}

# --- Main Script Execution ---
try {
    if (-not (Test-Path -Path $OutputDir)) {
        New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
    }

    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $script:LogFilePath = Join-Path $OutputDir "DL_Export_Execution_${timestamp}.log"
    Write-Log "Execution started. Log file: $script:LogFilePath" "INFO"

    if ($DryRun) {
        Write-Log "=== RUNNING IN DRY-RUN MODE (Validation only) ===" "DRYRUN"
    }

    Assert-Prerequisites

    $script:DisconnectChoice = $null

    # --- [2/6] Connection & Session Verification ---
    Write-Log "=== [2/6] Checking Exchange Online Connection ===" "INFO"
    $existingConnections = Get-ConnectionInformation -ErrorAction SilentlyContinue
    $useExistingSession = $false

    if ($existingConnections -and $existingConnections.Count -gt 0) {
        $activeOrg = Get-OrganizationConfig -ErrorAction SilentlyContinue
        if ($activeOrg) {
            $orgName = if ($activeOrg.Identity) { $activeOrg.Identity } else { "Unknown Organization" }
            $activeUser = ($existingConnections | Select-Object -First 1).UserPrincipalName

            Write-Log "Active Exchange Online session detected:" "INFO"
            Write-Host "  * Connected Tenant : $orgName" -ForegroundColor Yellow
            if ($activeUser) {
                Write-Host "  * Account / UPN    : $activeUser" -ForegroundColor Yellow
            }

            if ($PSBoundParameters.ContainsKey('ReuseExistingConnection') -and $ReuseExistingConnection.IsPresent) {
                $useExistingSession = $true
            } else {
                Write-Host ""
                $prompt = Read-Host "Do you want to use this existing active connection? ([Y]/N)"
                if ([string]::IsNullOrWhiteSpace($prompt) -or ($prompt.Trim() -match '^(y|yes)$')) {
                    $useExistingSession = $true
                }
            }
        }
    }

    if ($useExistingSession) {
        Write-Log "Reusing active Exchange Online connection." "SUCCESS"
    } else {
        if ($existingConnections -and $existingConnections.Count -gt 0) {
            Write-Log "Disconnecting previous session before establishing a new one..." "INFO"
            Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
        }

        Write-Log "Authenticating to Exchange Online via Device Code..." "INFO"
        Write-Log "Follow the instructions below to complete sign-in in your browser." "WARN"

        $connectParams = @{
            Device = $true
        }
        if (-not [string]::IsNullOrWhiteSpace($Organization)) {
            $connectParams["Organization"] = $Organization
        }

        Connect-ExchangeOnline @connectParams
    }

    # Identify Connected Tenant
    $orgConfig = Get-OrganizationConfig -ErrorAction SilentlyContinue
    $tenantName = if ($orgConfig -and $orgConfig.Identity) {
        ($orgConfig.Identity.ToString() -replace '[\\/:*?"<>|\s]', '_')
    } else {
        "M365Tenant"
    }

    # Rename log file to include resolved tenant name
    $namedLogPath = Join-Path $OutputDir "DL_Export_${tenantName}_${timestamp}.log"
    try {
        Rename-Item -Path $script:LogFilePath -NewName (Split-Path $namedLogPath -Leaf) -Force
        $script:LogFilePath = $namedLogPath
    } catch { }

    Write-Log "Connected to tenant: $tenantName" "SUCCESS"

    $excelReportPath    = Join-Path $OutputDir "M365_Group_and_DL_Report_${tenantName}_${timestamp}.xlsx"
    $standardReportCsv  = Join-Path $OutputDir "DL_Members_${tenantName}_${timestamp}.csv"
    $m365ReportCsv      = Join-Path $OutputDir "M365_Group_Members_${tenantName}_${timestamp}.csv"
    $dynamicReportCsv   = Join-Path $OutputDir "Dynamic_DL_Members_${tenantName}_${timestamp}.csv"

    # Master Group Directory for Executive Reporting
    $masterDirectory = [System.Collections.Generic.List[PSObject]]::new()

    # --- [3/6] Query Standard Distribution Lists & Mail-Enabled Groups ---
    Write-Log "=== [3/6] Querying Distribution Lists & Mail-Enabled Groups ===" "INFO"
    $distGroups = Get-DistributionGroup -ResultSize Unlimited
    $distGroupCount = if ($distGroups) { $distGroups.Count } else { 0 }
    Write-Log "Found $distGroupCount standard / mail-enabled distribution groups." "SUCCESS"

    $standardResults = [System.Collections.Generic.List[PSObject]]::new()

    if ($distGroupCount -gt 0) {
        $groupCounter = 0
        foreach ($group in $distGroups) {
            $groupCounter++
            $percent = [math]::Round(($groupCounter / $distGroupCount) * 100, 1)
            Write-Progress -Activity "Processing Distribution Groups" `
                           -Status "[$groupCounter/$distGroupCount] ($percent%) - $($group.DisplayName)" `
                           -PercentComplete $percent

            # Determine Direct Member Count
            $directCount = 0
            try {
                $dm = Get-DistributionGroupMember -Identity $group.Identity -ResultSize Unlimited -ErrorAction Stop
                $directCount = if ($dm) { $dm.Count } else { 0 }
            } catch {
                $directCount = 0
            }

            $externalAccess = if ($group.RequireSenderAuthenticationEnabled) { "No (Internal Only)" } else { "Yes (Public Ingress)" }
            $managers = Get-CleanManagers $group.ManagedBy
            $status = if ($directCount -eq 0) { "Empty (Review)" } else { "Active" }

            # Add to Master Directory
            $masterDirectory.Add([PSCustomObject]@{
                "Group Name"            = $group.DisplayName
                "Email Address"         = $group.PrimarySmtpAddress
                "Group Type"            = (Get-FriendlyGroupType $group.RecipientTypeDetails)
                "External Mail Allowed" = $externalAccess
                "Member Count"          = $directCount
                "Owners / Managers"     = $managers
                "Status"                = $status
            })

            # If not dry-run, resolve full members recursively
            if (-not $DryRun) {
                $visitedSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
                $members = Get-RecursiveGroupMembers -RootGroup $group -CurrentGroup $group -VisitedGroupIds $visitedSet
                if ($members) {
                    foreach ($m in $members) {
                        $standardResults.Add($m)
                    }
                }
            }
        }
        Write-Progress -Activity "Processing Distribution Groups" -Completed
    }

    # --- [4/6] Query Microsoft 365 Groups (Unified Groups) ---
    $m365Results = [System.Collections.Generic.List[PSObject]]::new()
    $m365GroupCount = 0

    if (-not $SkipM365Groups) {
        Write-Log "=== [4/6] Querying Microsoft 365 Groups ===" "INFO"
        $m365Groups = Get-UnifiedGroup -ResultSize Unlimited
        $m365GroupCount = if ($m365Groups) { $m365Groups.Count } else { 0 }
        Write-Log "Found $m365GroupCount Microsoft 365 groups." "SUCCESS"

        if ($m365GroupCount -gt 0) {
            $ugCounter = 0
            foreach ($ug in $m365Groups) {
                $ugCounter++
                $percent = [math]::Round(($ugCounter / $m365GroupCount) * 100, 1)
                Write-Progress -Activity "Processing Microsoft 365 Groups" `
                               -Status "[$ugCounter/$m365GroupCount] ($percent%) - $($ug.DisplayName)" `
                               -PercentComplete $percent

                $members = @()
                $owners = @()

                try {
                    $members = Get-UnifiedGroupLinks -Identity $ug.Identity -LinkType Members -ResultSize Unlimited -ErrorAction Stop
                } catch { }

                try {
                    $owners = Get-UnifiedGroupLinks -Identity $ug.Identity -LinkType Owners -ResultSize Unlimited -ErrorAction Stop
                } catch { }

                $ownerCount = if ($owners) { $owners.Count } else { 0 }
                $memberCount = if ($members) { $members.Count } else { 0 }
                $ownerEmails = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
                $ownerDisplayList = [System.Collections.Generic.List[string]]::new()

                if ($owners) {
                    foreach ($o in $owners) {
                        if ($o.PrimarySmtpAddress) {
                            $ownerEmails.Add($o.PrimarySmtpAddress.ToString()) | Out-Null
                        }
                        $ownerDisplayList.Add($o.DisplayName)
                    }
                }

                $ownersText = if ($ownerDisplayList.Count -gt 0) { ($ownerDisplayList -join ", ") } else { "None" }
                $externalAccess = if ($ug.RequireSenderAuthenticationEnabled) { "No (Internal Only)" } else { "Yes (Public Ingress)" }
                $status = if ($memberCount -eq 0 -and $ownerCount -eq 0) { "Empty (Review)" } else { "Active" }

                # Add to Master Directory
                $masterDirectory.Add([PSCustomObject]@{
                    "Group Name"            = $ug.DisplayName
                    "Email Address"         = $ug.PrimarySmtpAddress
                    "Group Type"            = "Microsoft 365 Group"
                    "External Mail Allowed" = $externalAccess
                    "Member Count"          = $memberCount
                    "Owners / Managers"     = $ownersText
                    "Status"                = $status
                })

                if (-not $DryRun) {
                    if ((-not $members -or $members.Count -eq 0) -and (-not $owners -or $owners.Count -eq 0)) {
                        $m365Results.Add([PSCustomObject]@{
                            "Group Name"   = $ug.DisplayName
                            "Group Email"  = $ug.PrimarySmtpAddress
                            "Privacy"      = $ug.AccessType
                            "Member Name"  = "<No Members or Owners>"
                            "Member Email" = ""
                            "Role"         = ""
                            "Member Type"  = ""
                        })
                    } else {
                        $processedEmails = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

                        if ($members) {
                            foreach ($m in $members) {
                                $email = if ($m.PrimarySmtpAddress) { $m.PrimarySmtpAddress.ToString() } else { "" }
                                if ($email) { $processedEmails.Add($email) | Out-Null }
                                $role = if ($email -and $ownerEmails.Contains($email)) { "Owner & Member" } else { "Member" }

                                $m365Results.Add([PSCustomObject]@{
                                    "Group Name"   = $ug.DisplayName
                                    "Group Email"  = $ug.PrimarySmtpAddress
                                    "Privacy"      = $ug.AccessType
                                    "Member Name"  = $m.DisplayName
                                    "Member Email" = $email
                                    "Role"         = $role
                                    "Member Type"  = (Get-FriendlyRecipientType -RawType $m.RecipientType -Details $m.RecipientTypeDetails)
                                })
                            }
                        }

                        if ($owners) {
                            foreach ($o in $owners) {
                                $email = if ($o.PrimarySmtpAddress) { $o.PrimarySmtpAddress.ToString() } else { "" }
                                if ($email -and (-not $processedEmails.Contains($email))) {
                                    $m365Results.Add([PSCustomObject]@{
                                        "Group Name"   = $ug.DisplayName
                                        "Group Email"  = $ug.PrimarySmtpAddress
                                        "Privacy"      = $ug.AccessType
                                        "Member Name"  = $o.DisplayName
                                        "Member Email" = $email
                                        "Role"         = "Owner"
                                        "Member Type"  = (Get-FriendlyRecipientType -RawType $o.RecipientType -Details $o.RecipientTypeDetails)
                                    })
                                }
                            }
                        }
                    }
                }
            }
            Write-Progress -Activity "Processing Microsoft 365 Groups" -Completed
        }
    }

    # --- [5/6] Query Dynamic Distribution Lists ---
    $dynamicResults = [System.Collections.Generic.List[PSObject]]::new()
    $dynamicGroupCount = 0

    if (-not $SkipDynamicGroups) {
        Write-Log "=== [5/6] Querying Dynamic Distribution Groups ===" "INFO"
        $dynamicGroups = Get-DynamicDistributionGroup -ResultSize Unlimited
        $dynamicGroupCount = if ($dynamicGroups) { $dynamicGroups.Count } else { 0 }
        Write-Log "Found $dynamicGroupCount dynamic distribution groups." "SUCCESS"

        if ($dynamicGroupCount -gt 0) {
            $ddlCounter = 0
            foreach ($ddl in $dynamicGroups) {
                $ddlCounter++
                $percent = [math]::Round(($ddlCounter / $dynamicGroupCount) * 100, 1)
                Write-Progress -Activity "Processing Dynamic Groups" `
                               -Status "[$ddlCounter/$dynamicGroupCount] ($percent%) - $($ddl.DisplayName)" `
                               -PercentComplete $percent

                $recipients = @()
                $filter = $ddl.RecipientFilter

                if (-not [string]::IsNullOrWhiteSpace($filter)) {
                    try {
                        if ($ddl.RecipientContainer) {
                            $recipients = Get-Recipient -RecipientPreviewFilter $filter `
                                                       -OrganizationalUnit $ddl.RecipientContainer `
                                                       -ResultSize Unlimited `
                                                       -ErrorAction Stop
                        } else {
                            $recipients = Get-Recipient -RecipientPreviewFilter $filter `
                                                       -ResultSize Unlimited `
                                                       -ErrorAction Stop
                        }
                    } catch { }
                }

                $recipCount = if ($recipients) { $recipients.Count } else { 0 }
                $externalAccess = if ($ddl.RequireSenderAuthenticationEnabled) { "No (Internal Only)" } else { "Yes (Public Ingress)" }

                # Add to Master Directory
                $masterDirectory.Add([PSCustomObject]@{
                    "Group Name"            = $ddl.DisplayName
                    "Email Address"         = $ddl.PrimarySmtpAddress
                    "Group Type"            = "Dynamic Distribution Group"
                    "External Mail Allowed" = $externalAccess
                    "Member Count"          = $recipCount
                    "Owners / Managers"     = (Get-CleanManagers $ddl.ManagedBy)
                    "Status"                = if ($recipCount -eq 0) { "Empty / Filter Inactive" } else { "Active" }
                })

                if (-not $DryRun) {
                    if ($recipCount -eq 0) {
                        $dynamicResults.Add([PSCustomObject]@{
                            "Group Name"   = $ddl.DisplayName
                            "Group Email"  = $ddl.PrimarySmtpAddress
                            "Member Name"  = "<No evaluated members>"
                            "Member Email" = ""
                            "Member Type"  = ""
                        })
                    } else {
                        foreach ($r in $recipients) {
                            $dynamicResults.Add([PSCustomObject]@{
                                "Group Name"   = $ddl.DisplayName
                                "Group Email"  = $ddl.PrimarySmtpAddress
                                "Member Name"  = $r.DisplayName
                                "Member Email" = $r.PrimarySmtpAddress
                                "Member Type"  = (Get-FriendlyRecipientType -RawType $r.RecipientType -Details $r.RecipientTypeDetails)
                            })
                        }
                    }
                }
            }
            Write-Progress -Activity "Processing Dynamic Groups" -Completed
        }
    }

    # --- [6/6] Generating Decluttered Executive Excel Workbook ---
    if ($DryRun) {
        Write-Log "=== [6/6] Dry-Run Discovery Complete! ===" "DRYRUN"
        Write-Host "`n======================= MASTER GROUP CATALOG (PREVIEW) =======================" -ForegroundColor Yellow
        $masterDirectory | Format-Table -AutoSize | Out-String | ForEach-Object { Write-Host $_ -ForegroundColor White }

        Write-Log "Dry-run complete. In production mode, an executive workbook will be generated at: $excelReportPath" "INFO"
    } else {
        Write-Log "=== [6/6] Generating Executive-Ready Multi-Tab Excel Workbook ===" "INFO"

        if (Test-Path $excelReportPath) {
            Remove-Item $excelReportPath -Force
        }

        # Sheet 1: Executive KPI Metrics
        $totalGroups = $distGroupCount + $m365GroupCount + $dynamicGroupCount
        $kpiData = @(
            [PSCustomObject]@{ "Executive KPI Metric" = "Tenant Organization";        "Value" = $tenantName },
            [PSCustomObject]@{ "Executive KPI Metric" = "Audit Date & Time";           "Value" = (Get-Date -Format "yyyy-MM-dd HH:mm:ss") },
            [PSCustomObject]@{ "Executive KPI Metric" = "Total Email Groups Managed";  "Value" = $totalGroups },
            [PSCustomObject]@{ "Executive KPI Metric" = "Standard Distribution Lists"; "Value" = $distGroupCount },
            [PSCustomObject]@{ "Executive KPI Metric" = "Microsoft 365 Groups";       "Value" = if ($SkipM365Groups) { "Skipped" } else { $m365GroupCount } },
            [PSCustomObject]@{ "Executive KPI Metric" = "Dynamic Distribution Groups"; "Value" = if ($SkipDynamicGroups) { "Skipped" } else { $dynamicGroupCount } }
        )
        $kpiData | Export-Excel -Path $excelReportPath -WorksheetName "Executive_Summary" -AutoSize -BoldTopRow -TableStyle Medium2

        # Sheet 2: Master Group Directory
        if ($masterDirectory.Count -gt 0) {
            $masterDirectory | Export-Excel -Path $excelReportPath -WorksheetName "Group_Directory" -AutoSize -AutoFilter -BoldTopRow -TableStyle Medium9
        }

        # Sheet 3: Distribution Lists
        if ($standardResults.Count -gt 0) {
            $standardResults | Export-Excel -Path $excelReportPath -WorksheetName "Distribution_Lists" -AutoSize -AutoFilter -BoldTopRow -TableStyle Medium15
        } else {
            @([PSCustomObject]@{ "Status" = "No standard distribution list members found." }) | `
                Export-Excel -Path $excelReportPath -WorksheetName "Distribution_Lists" -AutoSize -BoldTopRow
        }

        # Sheet 4: Microsoft 365 Groups
        if (-not $SkipM365Groups) {
            if ($m365Results.Count -gt 0) {
                $m365Results | Export-Excel -Path $excelReportPath -WorksheetName "Microsoft_365_Groups" -AutoSize -AutoFilter -BoldTopRow -TableStyle Medium6
            } else {
                @([PSCustomObject]@{ "Status" = "No Microsoft 365 groups found." }) | `
                    Export-Excel -Path $excelReportPath -WorksheetName "Microsoft_365_Groups" -AutoSize -BoldTopRow
            }
        }

        # Sheet 5: Dynamic Distribution Lists
        if (-not $SkipDynamicGroups) {
            if ($dynamicResults.Count -gt 0) {
                $dynamicResults | Export-Excel -Path $excelReportPath -WorksheetName "Dynamic_Groups" -AutoSize -AutoFilter -BoldTopRow -TableStyle Medium4
            } else {
                @([PSCustomObject]@{ "Status" = "No dynamic distribution groups found in this tenant." }) | `
                    Export-Excel -Path $excelReportPath -WorksheetName "Dynamic_Groups" -AutoSize -BoldTopRow
            }
        }

        Write-Log "Executive Excel Workbook successfully generated: $excelReportPath" "SUCCESS"

        if ($ExportCsv) {
            Write-Log "Generating supplementary CSV files..." "INFO"
            if ($standardResults.Count -gt 0) { $standardResults | Export-Csv -Path $standardReportCsv -NoTypeInformation -Encoding utf8 }
            if (-not $SkipM365Groups -and $m365Results.Count -gt 0) { $m365Results | Export-Csv -Path $m365ReportCsv -NoTypeInformation -Encoding utf8 }
            if (-not $SkipDynamicGroups -and $dynamicResults.Count -gt 0) { $dynamicResults | Export-Csv -Path $dynamicReportCsv -NoTypeInformation -Encoding utf8 }
            Write-Log "Supplementary CSV files written." "SUCCESS"
        }

        Write-Log "=== Export Process Complete! ===" "SUCCESS"
        Write-Log "Output folder: $OutputDir" "INFO"
    }

    # Prompt user regarding session disconnect
    if ($PSBoundParameters.ContainsKey('KeepSessionConnected') -and $KeepSessionConnected.IsPresent) {
        $script:DisconnectChoice = $false
    } elseif ($PSBoundParameters.ContainsKey('DisconnectSession') -and $DisconnectSession.IsPresent) {
        $script:DisconnectChoice = $true
    } else {
        Write-Host ""
        $userChoice = Read-Host "Do you want to disconnect the Exchange Online session? (Y/[N], default is N to keep active)"
        if ($userChoice.Trim() -match '^(y|yes)$') {
            $script:DisconnectChoice = $true
        } else {
            $script:DisconnectChoice = $false
        }
    }
}
catch {
    Write-Log "An unexpected error occurred: $_" "ERROR"
    if ($_.ScriptStackTrace) {
        Write-Log "Stack Trace: $($_.ScriptStackTrace)" "ERROR"
    }
}
finally {
    if ($null -eq $script:DisconnectChoice) {
        if ($PSBoundParameters.ContainsKey('KeepSessionConnected') -and $KeepSessionConnected.IsPresent) {
            $script:DisconnectChoice = $false
        } elseif ($PSBoundParameters.ContainsKey('DisconnectSession') -and $DisconnectSession.IsPresent) {
            $script:DisconnectChoice = $true
        } else {
            Write-Host ""
            $errChoice = Read-Host "Disconnect Exchange Online session? (Y/[N], default is N to keep active)"
            if ($errChoice.Trim() -match '^(y|yes)$') {
                $script:DisconnectChoice = $true
            } else {
                $script:DisconnectChoice = $false
            }
        }
    }

    if ($script:DisconnectChoice -eq $true) {
        Write-Log "Disconnecting Exchange Online session..." "INFO"
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
        Write-Log "Session disconnected." "SUCCESS"
    } else {
        Write-Log "Exchange Online session remains ACTIVE." "SUCCESS"
        Write-Log "You can run subsequent commands or re-run this script without authenticating again." "INFO"
    }
}
