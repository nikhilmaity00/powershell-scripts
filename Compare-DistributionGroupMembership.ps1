<#
.SYNOPSIS
    Compares the direct Exchange Online distribution-group memberships of two users.

.DESCRIPTION
    The script checks for ExchangeOnlineManagement, signs in through the browser,
    then retrieves only the distribution lists (including mail-enabled security groups)
    that each supplied recipient is directly listed in. Dynamic distribution groups
    are not included because their membership is calculated from a query at delivery.
#>

[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Write-Title {
    param([string]$Text)
    $line = '═' * 72
    Write-Host "`n╔$line╗" -ForegroundColor Cyan
    Write-Host ("║" + $Text.PadLeft([Math]::Floor((72 + $Text.Length) / 2)).PadRight(72) + "║") -ForegroundColor Cyan
    Write-Host "╚$line╝" -ForegroundColor Cyan
}

function Write-Section {
    param([string]$Text)
    Write-Host "`n── $Text " -NoNewline -ForegroundColor Cyan
    Write-Host ('─' * [Math]::Max(1, 68 - $Text.Length)) -ForegroundColor DarkCyan
}

function Get-EmailAddress {
    param([string]$Prompt)

    do {
        $email = (Read-Host $Prompt).Trim()
        if ($email -notmatch '^[^\s@]+@[^\s@]+\.[^\s@]+$') {
            Write-Host 'Please enter a valid email address.' -ForegroundColor Yellow
        }
    } until ($email -match '^[^\s@]+@[^\s@]+\.[^\s@]+$')

    return $email
}

function Get-DirectDistributionGroupMembership {
    param(
        [Parameter(Mandatory)] [string]$EmailAddress
    )

    Write-Host "  Finding direct memberships for $EmailAddress..." -ForegroundColor DarkGray

    # Exchange evaluates this filter server-side: it returns only groups whose
    # multi-valued Members property contains this recipient's distinguished name.
    # This avoids enumerating every DL and every DL member in the tenant.
    $recipient = Get-Recipient -Identity $EmailAddress
    $distinguishedName = $recipient.DistinguishedName
    $filter = "Members -eq '$distinguishedName'"
    $groups = Get-DistributionGroup -Filter $filter -ResultSize Unlimited -WarningAction SilentlyContinue

    return @($groups | ForEach-Object {
        [PSCustomObject]@{
            DisplayName        = $_.DisplayName
            PrimarySmtpAddress = $_.PrimarySmtpAddress.ToString()
            GroupType          = $_.RecipientTypeDetails
        }
    } | Sort-Object DisplayName)
}

function Show-GroupTable {
    param(
        [string]$Heading,
        [array]$Groups,
        [ConsoleColor]$Color = 'White'
    )

    Write-Section $Heading
    if ($Groups.Count -eq 0) {
        Write-Host '  None' -ForegroundColor DarkGray
        return
    }

    Write-Host ($Groups | Select-Object DisplayName, PrimarySmtpAddress, GroupType | Format-Table -AutoSize | Out-String) -ForegroundColor $Color
}

function Export-Comparison {
    param(
        [string]$FirstEmail,
        [string]$SecondEmail,
        [array]$Common,
        [array]$OnlyFirst,
        [array]$OnlySecond
    )

    $defaultPath = Join-Path -Path ([Environment]::GetFolderPath('Desktop')) -ChildPath ("DL-Membership-Comparison_{0}.csv" -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
    $path = (Read-Host "Export path (press Enter for $defaultPath)").Trim()
    if ([string]::IsNullOrWhiteSpace($path)) { $path = $defaultPath }
    if ([IO.Path]::GetExtension($path) -ne '.csv') { $path += '.csv' }

    $rows = [System.Collections.Generic.List[object]]::new()
    foreach ($item in $Common) {
        $rows.Add([PSCustomObject]@{
            Category = 'Common'; DistributionList = $item.DisplayName; EmailAddress = $item.PrimarySmtpAddress; GroupType = $item.GroupType
            FirstUser = $FirstEmail; FirstUserMember = 'Yes'; SecondUser = $SecondEmail; SecondUserMember = 'Yes'
        })
    }
    foreach ($item in $OnlyFirst) {
        $rows.Add([PSCustomObject]@{
            Category = "Only $FirstEmail"; DistributionList = $item.DisplayName; EmailAddress = $item.PrimarySmtpAddress; GroupType = $item.GroupType
            FirstUser = $FirstEmail; FirstUserMember = 'Yes'; SecondUser = $SecondEmail; SecondUserMember = 'No'
        })
    }
    foreach ($item in $OnlySecond) {
        $rows.Add([PSCustomObject]@{
            Category = "Only $SecondEmail"; DistributionList = $item.DisplayName; EmailAddress = $item.PrimarySmtpAddress; GroupType = $item.GroupType
            FirstUser = $FirstEmail; FirstUserMember = 'No'; SecondUser = $SecondEmail; SecondUserMember = 'Yes'
        })
    }

    try {
        if ($rows.Count -eq 0) {
            # Keep the requested CSV useful even if neither user has any direct memberships.
            $rows = @([PSCustomObject]@{
                Category = 'No direct memberships found'; DistributionList = ''; EmailAddress = ''; GroupType = ''
                FirstUser = $FirstEmail; FirstUserMember = 'No'; SecondUser = $SecondEmail; SecondUserMember = 'No'
            })
        }
        $rows | Export-Csv -Path $path -NoTypeInformation -Encoding utf8
        Write-Host "  Exported $($rows.Count) rows to: $path" -ForegroundColor Green
    }
    catch {
        Write-Host "  Export failed: $($_.Exception.Message)" -ForegroundColor Red
    }
}

Write-Title 'DISTRIBUTION LIST MEMBERSHIP COMPARER'
Write-Host 'Compares direct memberships without scanning every DL in the tenant.' -ForegroundColor DarkGray
Write-Host 'Dynamic distribution groups are excluded (their memberships are query-based).' -ForegroundColor DarkGray

try {
    Write-Section 'Prerequisite check'
    $module = Get-Module -ListAvailable -Name ExchangeOnlineManagement |
        Sort-Object Version -Descending |
        Select-Object -First 1

    if (-not $module) {
        Write-Host 'ExchangeOnlineManagement is not installed. Installing for the current user...' -ForegroundColor Yellow
        Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Repository PSGallery -Force -AllowClobber
    }
    else {
        Write-Host "ExchangeOnlineManagement $($module.Version) is available." -ForegroundColor Green
    }
    Import-Module ExchangeOnlineManagement -ErrorAction Stop

    Write-Section 'Connect to Exchange Online'
    $adminId = Get-EmailAddress 'Exchange administrator ID'
    Write-Host 'Your browser will open for authentication. Complete sign-in, then return here.' -ForegroundColor Yellow
    Connect-ExchangeOnline -UserPrincipalName $adminId -ShowBanner:$false
    Write-Host 'Connected successfully.' -ForegroundColor Green

    do {
        Write-Section 'User comparison'
        $firstEmail = Get-EmailAddress 'First user email address'
        $secondEmail = Get-EmailAddress 'Second user email address'

        if ($firstEmail.Equals($secondEmail, [System.StringComparison]::OrdinalIgnoreCase)) {
            Write-Host 'Please provide two different email addresses.' -ForegroundColor Yellow
            continue
        }

        Write-Host 'Querying Exchange for each user memberships...' -ForegroundColor DarkGray
        $firstGroups = Get-DirectDistributionGroupMembership -EmailAddress $firstEmail
        $secondGroups = Get-DirectDistributionGroupMembership -EmailAddress $secondEmail

        $firstMap = @{}; $firstGroups | ForEach-Object { $firstMap[$_.PrimarySmtpAddress.ToLowerInvariant()] = $_ }
        $secondMap = @{}; $secondGroups | ForEach-Object { $secondMap[$_.PrimarySmtpAddress.ToLowerInvariant()] = $_ }

        $common = @($firstMap.Keys | Where-Object { $secondMap.ContainsKey($_) } | ForEach-Object { $firstMap[$_] } | Sort-Object DisplayName)
        $onlyFirst = @($firstMap.Keys | Where-Object { -not $secondMap.ContainsKey($_) } | ForEach-Object { $firstMap[$_] } | Sort-Object DisplayName)
        $onlySecond = @($secondMap.Keys | Where-Object { -not $firstMap.ContainsKey($_) } | ForEach-Object { $secondMap[$_] } | Sort-Object DisplayName)

        Write-Title 'COMPARISON RESULTS'
        Write-Host "  $firstEmail : $($firstGroups.Count) direct DL membership(s)" -ForegroundColor White
        Write-Host "  $secondEmail : $($secondGroups.Count) direct DL membership(s)" -ForegroundColor White
        Show-GroupTable -Heading 'Common distribution lists' -Groups $common -Color Green
        Show-GroupTable -Heading "Only $firstEmail is a member (missing for $secondEmail)" -Groups $onlyFirst -Color Yellow
        Show-GroupTable -Heading "Only $secondEmail is a member (missing for $firstEmail)" -Groups $onlySecond -Color Yellow

        if ((Read-Host "`nExport this comparison to CSV? (Y/N)") -match '^(y|yes)$') {
            Export-Comparison -FirstEmail $firstEmail -SecondEmail $secondEmail -Common $common -OnlyFirst $onlyFirst -OnlySecond $onlySecond
        }

        $again = Read-Host "`nCompare another pair of users? (Y/N)"
    } while ($again -match '^(y|yes)$')
}
catch {
    Write-Host "`nError: $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    if (Get-ConnectionInformation -ErrorAction SilentlyContinue) {
        $disconnect = Read-Host "`nDisconnect the Exchange Online session? (Y/N)"
        if ($disconnect -match '^(y|yes)$') {
            Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
            Write-Host 'Session disconnected.' -ForegroundColor Green
        }
        else {
            Write-Host 'Session remains connected in this PowerShell window.' -ForegroundColor Yellow
        }
    }
}

Write-Host "`nDone." -ForegroundColor Cyan
