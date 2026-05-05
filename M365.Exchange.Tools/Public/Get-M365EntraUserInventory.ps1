function Get-M365EntraUserInventory {
    [CmdletBinding()]
    param(
        [Parameter()]
        [ValidateSet('All', 'Member', 'Guest')]
        [string]$UserScope = 'All',

        [Parameter()]
        [string]$ExportPath
    )

    Assert-ExchangeOnlineConnected

    $users = @()
    $signInActivityStatus = 'Available'
    try {
        $users = Get-M365GraphCollection -Uri 'https://graph.microsoft.com/beta/users?$top=999&$select=id,displayName,userPrincipalName,mail,accountEnabled,userType,createdDateTime,department,jobTitle,companyName,officeLocation,city,state,country,onPremisesSyncEnabled,signInActivity,assignedLicenses,lastPasswordChangeDateTime,passwordPolicies'
    }
    catch {
        if ($_.Exception.Message -match 'Forbidden|insufficient privileges') {
            throw 'Missing Microsoft Graph permission AuditLog.Read.All for last sign-in activity. Reconnect using option 1 and consent the requested scopes.'
        }

        throw
    }

    $filteredUsers = switch ($UserScope) {
        'Member' { @($users | Where-Object { $_.userType -eq 'Member' }) }
        'Guest' { @($users | Where-Object { $_.userType -eq 'Guest' }) }
        Default { @($users) }
    }

    $domainPolicyBySuffix = @{}
    $defaultDomainPolicy = $null
    $domainPolicyStatus = 'Available'

    try {
        $domains = Get-M365GraphCollection -Uri 'https://graph.microsoft.com/v1.0/domains?$top=999&$select=id,isDefault,authenticationType,passwordValidityPeriodInDays'
        foreach ($domain in $domains) {
            $domainPolicyBySuffix[[string]$domain.id.ToLowerInvariant()] = $domain
            if ($domain.isDefault) {
                $defaultDomainPolicy = $domain
            }
        }
    }
    catch {
        $domainPolicyStatus = if ($_.Exception.Message -match 'Forbidden|insufficient privileges') { 'Domain policy unavailable (Domain.Read.All may be required)' } else { 'Domain policy lookup failed' }
    }

    $results = foreach ($user in ($filteredUsers | Sort-Object displayName, userPrincipalName)) {
        $mailboxKind = 'User'
        if ($user.userType -eq 'Guest') {
            $mailboxKind = 'Guest'
        }
        else {
            try {
                $mailboxSettings = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$($user.id)/mailboxSettings?`$select=userPurpose" -OutputType PSObject -ErrorAction Stop
                $userPurpose = [string]$mailboxSettings.userPurpose
                switch ($userPurpose.ToLowerInvariant()) {
                    'shared' {
                        $mailboxKind = 'Shared'
                    }
                    'room' {
                        $mailboxKind = 'Resource'
                    }
                    'equipment' {
                        $mailboxKind = 'Resource'
                    }
                    default {
                        $mailboxKind = 'User'
                    }
                }
            }
            catch {
                $mailboxKind = 'User'
            }
        }

        $domainSuffix = $null
        if ([string]$user.userPrincipalName -match '@') {
            $domainSuffix = ([string]$user.userPrincipalName).Split('@')[-1].ToLowerInvariant()
        }

        $domainPolicy = if ($domainSuffix -and $domainPolicyBySuffix.ContainsKey($domainSuffix)) {
            $domainPolicyBySuffix[$domainSuffix]
        }
        else {
            $defaultDomainPolicy
        }

        $passwordNeverExpires = ([string]$user.passwordPolicies -match 'DisablePasswordExpiration')
        $lastPasswordChangeDateTime = $user.lastPasswordChangeDateTime
        $passwordExpiryDateTime = $null
        $passwordDaysUntilExpiry = $null
        $passwordExpiryStatus = 'Unknown'

        if ($user.userType -eq 'Guest') {
            $passwordExpiryStatus = 'Not applicable (Guest account)'
        }
        elseif ($user.onPremisesSyncEnabled -eq $true) {
            $passwordExpiryStatus = 'Managed on-premises (not calculated)'
        }
        elseif ($passwordNeverExpires) {
            $passwordExpiryStatus = 'Does not expire'
        }
        elseif (-not $lastPasswordChangeDateTime) {
            $passwordExpiryStatus = 'Unknown (no last password change time)'
        }
        elseif (-not $domainPolicy) {
            $passwordExpiryStatus = 'Unknown (domain policy unavailable)'
        }
        elseif (($domainPolicy.authenticationType -ne 'Managed') -or (-not $domainPolicy.passwordValidityPeriodInDays) -or ([int]$domainPolicy.passwordValidityPeriodInDays -le 0)) {
            $passwordExpiryStatus = 'Unknown (no managed password validity policy)'
        }
        else {
            try {
                $passwordExpiryDateTime = ([datetime]$lastPasswordChangeDateTime).AddDays([int]$domainPolicy.passwordValidityPeriodInDays)
                $passwordDaysUntilExpiry = [int][Math]::Floor(($passwordExpiryDateTime - (Get-Date)).TotalDays)

                if ($passwordDaysUntilExpiry -lt 0) {
                    $passwordExpiryStatus = 'Expired'
                }
                elseif ($passwordDaysUntilExpiry -le 14) {
                    $passwordExpiryStatus = 'Expiring soon'
                }
                else {
                    $passwordExpiryStatus = 'Active'
                }
            }
            catch {
                $passwordExpiryStatus = 'Unknown (calculation failed)'
            }
        }

        [pscustomobject]@{
            DisplayName                       = $user.displayName
            UserPrincipalName                 = $user.userPrincipalName
            Mail                              = $user.mail
            UserId                            = $user.id
            AccountEnabled                    = $user.accountEnabled
            UserType                          = $user.userType
            MailboxKind                       = $mailboxKind
            WhenCreated                       = $user.createdDateTime
            LastSignInDateTime                = $user.signInActivity.lastSignInDateTime
            LastNonInteractiveSignInDateTime  = $user.signInActivity.lastNonInteractiveSignInDateTime
            LastSuccessfulSignInDateTime      = $user.signInActivity.lastSuccessfulSignInDateTime
            LastPasswordChangeDateTime        = $lastPasswordChangeDateTime
            PasswordPolicies                  = $user.passwordPolicies
            PasswordNeverExpires              = $passwordNeverExpires
            PasswordExpiryDateTime            = $passwordExpiryDateTime
            PasswordDaysUntilExpiry           = $passwordDaysUntilExpiry
            PasswordExpiryStatus              = $passwordExpiryStatus
            PasswordPolicyDomain              = $domainSuffix
            PasswordPolicyStatus              = $domainPolicyStatus
            Department                        = $user.department
            JobTitle                          = $user.jobTitle
            CompanyName                       = $user.companyName
            OfficeLocation                    = $user.officeLocation
            City                              = $user.city
            State                             = $user.state
            Country                           = $user.country
            OnPremisesSyncEnabled             = $user.onPremisesSyncEnabled
            AssignedLicenseCount              = @($user.assignedLicenses).Count
            SignInActivityStatus              = $signInActivityStatus
        }
    }

    Export-M365ReportData -InputObject $results -ExportPath $ExportPath
}
