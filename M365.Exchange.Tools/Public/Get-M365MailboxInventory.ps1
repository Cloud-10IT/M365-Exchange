function Get-M365MailboxInventory {
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$ExportPath
    )

    # Exchange must be connected for schema truth (Get-EXOMailbox).
    # Graph (Entra ID) is used for identity/state fields: AccountEnabled, Licenses,
    # PasswordLastSet, UsageLocation.  Sign-in activity comes from Graph beta
    # (signInActivity) which replaces the deprecated Exchange LastLogonTime field.
    Assert-M365ExchangePowerShellConnected

    # ── 1. Exchange Online: authoritative mailbox schema ────────────────────────
    Write-Host 'Retrieving mailbox list from Exchange Online (Get-EXOMailbox)...' -ForegroundColor Cyan
    $exoMailboxes = @(
        Get-EXOMailbox -ResultSize Unlimited `
            -Filter "RecipientTypeDetails -eq 'UserMailbox'" `
            -Properties DisplayName, PrimarySmtpAddress, UserPrincipalName,
                        ExternalDirectoryObjectId, RecipientTypeDetails,
                        ArchiveStatus, ArchiveName,
                        LitigationHoldEnabled, LitigationHoldDate,
                        RetentionHoldEnabled, RetentionPolicy,
                        ForwardingSmtpAddress, DeliverToMailboxAndForward,
                        HiddenFromAddressListsEnabled, ProhibitSendQuota,
                        ProhibitSendReceiveQuota, IssueWarningQuota `
            -ErrorAction Stop |
            Sort-Object DisplayName
    )

    $total = $exoMailboxes.Count
    Write-Host "Retrieved $total user mailboxes from Exchange. Enriching with Entra ID data..." -ForegroundColor Cyan

    # Build a lookup by ExternalDirectoryObjectId (= Entra Object ID / Graph user id)
    $exoById = @{}
    foreach ($mbx in $exoMailboxes) {
        $oid = [string]$mbx.ExternalDirectoryObjectId
        if (-not [string]::IsNullOrWhiteSpace($oid)) {
            $exoById[$oid] = $mbx
        }
    }

    # ── 2. Entra ID (Graph): identity, account state, licenses, password ────────
    # beta endpoint required for signInActivity (replaces Exchange LastLogonTime).
    # AuditLog.Read.All scope needed; fall back gracefully if missing.
    $signInActivityStatus = 'Available'
    $graphUsers = @()
    try {
        $graphUsers = Get-M365GraphCollection -Uri (
            'https://graph.microsoft.com/beta/users?$top=999&$select=' +
            'id,displayName,userPrincipalName,mail,accountEnabled,usageLocation,' +
            'assignedLicenses,lastPasswordChangeDateTime,passwordPolicies,' +
            'onPremisesSyncEnabled,signInActivity,userType'
        )
    }
    catch {
        if ($_.Exception.Message -match 'Forbidden|insufficient privileges|AuditLog') {
            $signInActivityStatus = 'Unavailable (AuditLog.Read.All scope required)'
            $graphUsers = Get-M365GraphCollection -Uri (
                'https://graph.microsoft.com/v1.0/users?$top=999&$select=' +
                'id,displayName,userPrincipalName,mail,accountEnabled,usageLocation,' +
                'assignedLicenses,lastPasswordChangeDateTime,passwordPolicies,' +
                'onPremisesSyncEnabled,userType'
            )
        }
        else {
            throw
        }
    }

    # Build lookup by Graph object id
    $graphById = @{}
    foreach ($u in $graphUsers) {
        $graphById[[string]$u.id] = $u
    }

    # ── 3. Merge: EXO (schema) + Graph (identity + activity) ───────────────────
    $counter = 0
    $results = foreach ($mbx in $exoMailboxes) {
        $counter++
        Write-Progress -Activity 'Building mailbox inventory' `
                       -Status "$counter / $total — $($mbx.DisplayName)" `
                       -PercentComplete ([int]($counter / $total * 100))

        $oid = [string]$mbx.ExternalDirectoryObjectId
        $g   = if ($graphById.ContainsKey($oid)) { $graphById[$oid] } else { $null }

        $passwordNeverExpires = ([string]$g.passwordPolicies -match 'DisablePasswordExpiration')
        $licenseCount         = if ($g) { @($g.assignedLicenses).Count } else { $null }

        # Sign-in activity — Graph beta signInActivity replaces Exchange LastLogonTime
        $lastSignIn              = if ($g) { $g.signInActivity.lastSignInDateTime }              else { $null }
        $lastNonInteractiveSignIn = if ($g) { $g.signInActivity.lastNonInteractiveSignInDateTime } else { $null }
        $lastSuccessfulSignIn    = if ($g) { $g.signInActivity.lastSuccessfulSignInDateTime }    else { $null }

        [pscustomobject]@{
            # ── Identity (Entra / Graph) ──────────────────────────────────────
            DisplayName                      = $mbx.DisplayName
            UserPrincipalName                = $mbx.UserPrincipalName
            PrimarySmtpAddress               = $mbx.PrimarySmtpAddress
            EntraObjectId                    = $oid
            AccountEnabled                   = if ($g) { $g.accountEnabled } else { $null }
            OnPremisesSyncEnabled            = if ($g) { $g.onPremisesSyncEnabled } else { $null }
            UsageLocation                    = if ($g) { $g.usageLocation } else { $null }
            AssignedLicenseCount             = $licenseCount
            PasswordLastSet                  = if ($g) { $g.lastPasswordChangeDateTime } else { $null }
            PasswordNeverExpires             = $passwordNeverExpires

            # ── Activity (Graph beta signInActivity — replaces Exchange LastLogonTime) ──
            LastSignInDateTime               = $lastSignIn
            LastNonInteractiveSignInDateTime = $lastNonInteractiveSignIn
            LastSuccessfulSignInDateTime     = $lastSuccessfulSignIn
            SignInActivityStatus             = $signInActivityStatus

            # ── Mailbox schema (Exchange — Get-EXOMailbox) ────────────────────
            ArchiveStatus                    = $mbx.ArchiveStatus
            LitigationHoldEnabled            = $mbx.LitigationHoldEnabled
            LitigationHoldDate               = $mbx.LitigationHoldDate
            RetentionHoldEnabled             = $mbx.RetentionHoldEnabled
            RetentionPolicy                  = $mbx.RetentionPolicy
            ForwardingSmtpAddress            = $mbx.ForwardingSmtpAddress
            DeliverToMailboxAndForward       = $mbx.DeliverToMailboxAndForward
            HiddenFromAddressLists           = $mbx.HiddenFromAddressListsEnabled
            ProhibitSendQuota                = $mbx.ProhibitSendQuota
            ProhibitSendReceiveQuota         = $mbx.ProhibitSendReceiveQuota
            MailboxKind                      = 'User'
        }
    }

    Write-Progress -Activity 'Building mailbox inventory' -Completed

    Export-M365ReportData -InputObject $results -ExportPath $ExportPath
}