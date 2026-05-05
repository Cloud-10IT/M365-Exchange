function Get-M365ResourceMailboxInventory {
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$ExportPath
    )

    Assert-ExchangeOnlineConnected

    $users = Get-M365GraphCollection -Uri 'https://graph.microsoft.com/v1.0/users?$top=999&$select=id,displayName,mail,officeLocation,createdDateTime'
    $results = foreach ($user in $users | Sort-Object displayName) {
        if ([string]::IsNullOrWhiteSpace($user.mail)) {
            continue
        }

        try {
            $mailboxSettings = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$($user.id)/mailboxSettings?`$select=userPurpose" -OutputType PSObject -ErrorAction Stop
        }
        catch {
            continue
        }

        if (($mailboxSettings.userPurpose -ne 'room') -and ($mailboxSettings.userPurpose -ne 'equipment')) {
            continue
        }

        [pscustomobject]@{
            DisplayName          = $user.displayName
            PrimarySmtpAddress   = $user.mail
            RecipientTypeDetails = $mailboxSettings.userPurpose
            Office               = $user.officeLocation
            WhenCreated          = $user.createdDateTime
            MailboxKind          = 'Resource'
        }
    }

    Export-M365ReportData -InputObject $results -ExportPath $ExportPath
}
