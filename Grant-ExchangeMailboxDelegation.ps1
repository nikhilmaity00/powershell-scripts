<#
.SYNOPSIS
    Interactive Exchange Online mailbox delegation manager.

.DESCRIPTION
    Add, remove, and review Send As and Read & manage (Full Access) permissions.
    The delegate CSV must contain one of: EmailAddress, UserPrincipalName,
    PrimarySmtpAddress, or Email.
#>

$ErrorActionPreference = 'Stop'
$connectedByScript = $false
$recipientCache = @{}
$lastActionReport = $null

function Write-Title {
    Clear-Host
    Write-Host '╔══════════════════════════════════════════════════════════════╗' -ForegroundColor Cyan
    Write-Host '║        Exchange Online Mailbox Delegation Manager            ║' -ForegroundColor Cyan
    Write-Host '╚══════════════════════════════════════════════════════════════╝' -ForegroundColor Cyan
}

function Write-Section([string]$Text) {
    Write-Host "`n━━ $Text " -ForegroundColor Cyan -NoNewline
    Write-Host ('━' * [Math]::Max(1, 62 - $Text.Length)) -ForegroundColor DarkCyan
}

function Get-ExchangeConnection {
    if (Get-Command Get-ConnectionInformation -ErrorAction SilentlyContinue) {
        return @(Get-ConnectionInformation | Where-Object { $_.State -eq 'Connected' })
    }
    return @()
}

function Read-YesNo([string]$Prompt, [bool]$DefaultYes = $false) {
    $suffix = if ($DefaultYes) { '[Y/n]' } else { '[y/N]' }
    $answer = (Read-Host "$Prompt $suffix").Trim()
    if ([string]::IsNullOrWhiteSpace($answer)) { return $DefaultYes }
    return $answer -match '^(y|yes)$'
}

function Get-DelegateAddress {
    param([Parameter(Mandatory)][psobject]$Row)
    foreach ($column in 'EmailAddress', 'UserPrincipalName', 'PrimarySmtpAddress', 'Email') {
        $value = $Row.$column
        if (-not [string]::IsNullOrWhiteSpace([string]$value)) { return ([string]$value).Trim() }
    }
    return $null
}

function Get-DelegateInfo {
    param([Parameter(Mandatory)][string]$Identity)
    if ($recipientCache.ContainsKey($Identity)) { return $recipientCache[$Identity] }

    try {
        $recipient = Get-EXORecipient -Identity $Identity -Properties DisplayName,PrimarySmtpAddress,RecipientTypeDetails -ErrorAction Stop
        $info = [pscustomobject]@{
            DisplayName       = $recipient.DisplayName
            EmailAddress      = [string]$recipient.PrimarySmtpAddress
            RecipientType     = [string]$recipient.RecipientTypeDetails
            ResolvedIdentity  = $recipient.Identity
        }
    }
    catch {
        $info = [pscustomobject]@{
            DisplayName       = $Identity
            EmailAddress      = $Identity
            RecipientType     = 'Unresolved'
            ResolvedIdentity  = $Identity
        }
    }
    $recipientCache[$Identity] = $info
    return $info
}

function Get-RequestedAccess([string]$Verb) {
    while ($true) {
        Write-Host "`nWhich access should be $Verb?" -ForegroundColor Cyan
        Write-Host '  1. Send As'
        Write-Host '  2. Read & manage (Full Access)'
        Write-Host '  3. Both'
        switch ((Read-Host 'Enter 1, 2, or 3').Trim()) {
            '1' { return @('SendAs') }
            '2' { return @('FullAccess') }
            '3' { return @('SendAs', 'FullAccess') }
            default { Write-Host 'Please enter 1, 2, or 3.' -ForegroundColor Yellow }
        }
    }
}

function Get-Delegates {
    Write-Host "`nProvide a CSV path or one or more comma-separated email addresses." -ForegroundColor DarkCyan
    Write-Host 'CSV header supported: EmailAddress, UserPrincipalName, PrimarySmtpAddress, or Email' -ForegroundColor DarkCyan
    $inputValue = (Read-Host 'CSV path or email address(es)').Trim().Trim('"')
    if ([string]::IsNullOrWhiteSpace($inputValue)) { throw 'A CSV path or at least one email address is required.' }

    if (Test-Path -LiteralPath $inputValue -PathType Leaf) {
        $rows = @(Import-Csv -LiteralPath $inputValue -ErrorAction Stop)
        if ($rows.Count -eq 0) { throw 'The CSV contains no data rows.' }
        # -Path is intentionally used here for compatibility with Windows PowerShell and PowerShell 7.
        return [pscustomobject]@{ Rows = $rows; Directory = (Split-Path -Path $inputValue -Parent); Source = 'CSV' }
    }

    $addresses = @($inputValue -split ',' | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($addresses.Count -eq 0) { throw 'No usable email addresses were supplied.' }
    $rows = @($addresses | ForEach-Object { [pscustomobject]@{ EmailAddress = $_ } })
    $reportDirectory = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
    return [pscustomobject]@{ Rows = $rows; Directory = $reportDirectory; Source = 'Email addresses' }
}

function Test-DelegatePermission {
    param([string]$Mailbox, [string]$Delegate, [ValidateSet('SendAs', 'FullAccess')][string]$Permission)
    try {
        if ($Permission -eq 'FullAccess') {
            return @(
                Get-EXOMailboxPermission -Identity $Mailbox -User $Delegate -ErrorAction Stop |
                    Where-Object { -not $_.Deny -and $_.AccessRights -contains 'FullAccess' }
            ).Count -gt 0
        }
        return @(
            Get-RecipientPermission -Identity $Mailbox -Trustee $Delegate -ErrorAction Stop |
                Where-Object { -not $_.Deny -and $_.AccessRights -contains 'SendAs' }
        ).Count -gt 0
    }
    catch {
        # A failed lookup must not be treated as existing access; the change command will surface a real error.
        return $false
    }
}

function Export-Results($Results, [string]$Directory, [string]$Action) {
    $reportPath = Join-Path $Directory ("MailboxDelegation_{0}_{1}.csv" -f $Action, (Get-Date -Format 'yyyyMMdd-HHmmss'))
    $Results | Export-Csv -Path $reportPath -NoTypeInformation
    Write-Host "Report saved: $reportPath" -ForegroundColor Green
}

function Request-LastActionReport {
    if ($null -eq $script:lastActionReport) {
        Write-Host 'There is no grant or removal action to report for this session.' -ForegroundColor Yellow
        return
    }

    $actionName = $script:lastActionReport.Action.ToLowerInvariant()
    if (-not (Read-YesNo "Export a CSV report for the last $actionName action?")) { return }

    try {
        Export-Results -Results $script:lastActionReport.Results -Directory $script:lastActionReport.Directory -Action $script:lastActionReport.Action
    }
    catch {
        Write-Host "Could not export the report: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Set-Delegation {
    param([string]$Mailbox, [ValidateSet('Grant', 'Remove')][string]$Action)
    $input = Get-Delegates
    $pastTense = if ($Action -eq 'Grant') { 'granted' } else { 'removed' }
    $access = Get-RequestedAccess -Verb $pastTense
    if (-not (Read-YesNo "$Action the selected permissions for $($input.Rows.Count) delegate(s) from $($input.Source)?")) {
        Write-Host 'No changes were made.' -ForegroundColor Yellow
        return
    }

    $results = [System.Collections.Generic.List[object]]::new()
    $index = 0
    foreach ($row in $input.Rows) {
        $index++
        Write-Progress -Activity "$Action mailbox delegation" -Status "Processing $index of $($input.Rows.Count)" -PercentComplete (($index / $input.Rows.Count) * 100)
        $address = Get-DelegateAddress $row
        if ([string]::IsNullOrWhiteSpace($address)) {
            $results.Add([pscustomobject]@{ DisplayName = ''; EmailAddress = ''; SendAs = 'Skipped'; FullAccess = 'Skipped'; Details = 'No supported email column/value found.' })
            continue
        }
        $delegate = Get-DelegateInfo $address
        $entry = [ordered]@{ DisplayName = $delegate.DisplayName; EmailAddress = $delegate.EmailAddress; SendAs = 'Not requested'; FullAccess = 'Not requested'; Details = '' }
        foreach ($permission in $access) {
            try {
                $hasPermission = Test-DelegatePermission -Mailbox $Mailbox -Delegate $address -Permission $permission
                if ($Action -eq 'Grant' -and $hasPermission) {
                    if ($permission -eq 'SendAs') { $entry.SendAs = 'Already granted' } else { $entry.FullAccess = 'Already granted' }
                    continue
                }
                if ($Action -eq 'Remove' -and -not $hasPermission) {
                    if ($permission -eq 'SendAs') { $entry.SendAs = 'Not assigned' } else { $entry.FullAccess = 'Not assigned' }
                    continue
                }
                if ($permission -eq 'SendAs') {
                    if ($Action -eq 'Grant') { $null = Add-RecipientPermission -Identity $Mailbox -Trustee $address -AccessRights SendAs -Confirm:$false -ErrorAction Stop }
                    else { $null = Remove-RecipientPermission -Identity $Mailbox -Trustee $address -AccessRights SendAs -Confirm:$false -ErrorAction Stop }
                    $entry.SendAs = if ($Action -eq 'Grant') { 'Granted' } else { 'Removed' }
                }
                else {
                    if ($Action -eq 'Grant') { $null = Add-MailboxPermission -Identity $Mailbox -User $address -AccessRights FullAccess -InheritanceType All -AutoMapping:$true -Confirm:$false -ErrorAction Stop }
                    else { $null = Remove-MailboxPermission -Identity $Mailbox -User $address -AccessRights FullAccess -InheritanceType All -Confirm:$false -ErrorAction Stop }
                    $entry.FullAccess = if ($Action -eq 'Grant') { 'Granted' } else { 'Removed' }
                }
            }
            catch {
                $message = ($_.Exception.Message -replace '[\r\n]+', ' ').Trim()
                if ($permission -eq 'SendAs') { $entry.SendAs = "Failed: $message" } else { $entry.FullAccess = "Failed: $message" }
                $entry.Details = $message
            }
        }
        $results.Add([pscustomobject]$entry)
    }
    Write-Progress -Activity "$Action mailbox delegation" -Completed
    Write-Section "$Action results"
    $results | Format-Table DisplayName, EmailAddress, SendAs, FullAccess -AutoSize
    $script:lastActionReport = [pscustomobject]@{
        Results   = $results.ToArray()
        Directory = $input.Directory
        Action    = $Action
    }
    Write-Host 'The results are available for an optional CSV export when you exit.' -ForegroundColor DarkGray
}

function Show-Delegation {
    param([string]$Mailbox)
    Write-Section "Current delegation for $Mailbox"
    # The EXO cmdlet supplies richer, more reliable permission data than the legacy cmdlet.
    $fullAccess = @(Get-EXOMailboxPermission -Identity $Mailbox -ErrorAction Stop | Where-Object { -not $_.Deny -and -not $_.IsInherited -and $_.User.ToString() -notmatch 'NT AUTHORITY\\SELF' })
    $sendAs = @(Get-RecipientPermission -Identity $Mailbox -ErrorAction Stop | Where-Object { -not $_.Deny -and $_.AccessRights -contains 'SendAs' })
    $records = @{}

    foreach ($permission in $fullAccess) {
        $key = $permission.User.ToString(); $info = Get-DelegateInfo $key
        $records[$key] = [pscustomobject]@{ DisplayName = $info.DisplayName; EmailAddress = $info.EmailAddress; RecipientType = $info.RecipientType; SendAs = 'No'; FullAccess = 'Yes' }
    }
    foreach ($permission in $sendAs) {
        $key = $permission.Trustee.ToString(); $info = Get-DelegateInfo $key
        if (-not $records.ContainsKey($key)) { $records[$key] = [pscustomobject]@{ DisplayName = $info.DisplayName; EmailAddress = $info.EmailAddress; RecipientType = $info.RecipientType; SendAs = 'No'; FullAccess = 'No' } }
        $records[$key].SendAs = 'Yes'
    }

    if ($records.Count -eq 0) { Write-Host 'No explicit Send As or Full Access delegates were found.' -ForegroundColor Yellow; return }
    $records.Values | Sort-Object DisplayName | Format-Table DisplayName, EmailAddress, RecipientType, SendAs, FullAccess -AutoSize
}

try {
    Write-Title
    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        Write-Host 'Installing ExchangeOnlineManagement for the current user...' -ForegroundColor Yellow
        Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
    }
    Import-Module ExchangeOnlineManagement -ErrorAction Stop
    $connections = Get-ExchangeConnection
    if ($connections.Count -eq 0) {
        $adminUpn = (Read-Host 'Enter your Exchange Online admin email address').Trim()
        if ([string]::IsNullOrWhiteSpace($adminUpn)) { throw 'An admin email address is required.' }
        Write-Host 'Opening the browser for Exchange Online sign-in...' -ForegroundColor Cyan
        Connect-ExchangeOnline -UserPrincipalName $adminUpn -ShowBanner:$false -ErrorAction Stop
        $connectedByScript = $true
    }
    else { Write-Host "Connected as: $($connections[0].UserPrincipalName)" -ForegroundColor Green }

    $mailbox = (Read-Host "`nEnter the email address of the mailbox to manage").Trim()
    $mailboxInfo = Get-EXOMailbox -Identity $mailbox -Properties DisplayName,PrimarySmtpAddress -ErrorAction Stop
    $mailbox = [string]$mailboxInfo.PrimarySmtpAddress
    Write-Host "Managing: $($mailboxInfo.DisplayName) <$mailbox>" -ForegroundColor Green

    $keepRunning = $true
    while ($keepRunning) {
        Write-Section 'Choose an action'
        Write-Host '  1. Show current delegated access'
        Write-Host '  2. Grant access from CSV'
        Write-Host '  3. Remove access from CSV'
        Write-Host '  4. Change mailbox'
        Write-Host '  5. Exit'
        switch ((Read-Host 'Enter 1-5').Trim()) {
            '1' { Show-Delegation $mailbox }
            '2' { Set-Delegation $mailbox 'Grant' }
            '3' { Set-Delegation $mailbox 'Remove' }
            '4' {
                $newMailbox = (Read-Host 'Enter the email address of the mailbox to manage').Trim()
                $mailboxInfo = Get-EXOMailbox -Identity $newMailbox -Properties DisplayName,PrimarySmtpAddress -ErrorAction Stop
                $mailbox = [string]$mailboxInfo.PrimarySmtpAddress
                Write-Host "Now managing: $($mailboxInfo.DisplayName) <$mailbox>" -ForegroundColor Green
            }
            '5' {
                Request-LastActionReport
                $keepRunning = $false
            }
            default { Write-Host 'Please enter a number from 1 to 5.' -ForegroundColor Yellow }
        }
    }
}
catch { Write-Host "`nScript failed: $($_.Exception.Message)" -ForegroundColor Red }
finally {
    if ($connectedByScript -and (Read-YesNo 'Disconnect the Exchange Online session opened by this script?')) {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
        Write-Host 'Disconnected from Exchange Online.' -ForegroundColor DarkGray
    }
    elseif ($connectedByScript) { Write-Host 'Exchange Online session remains connected.' -ForegroundColor Yellow }
}
