function Get-M365MailboxDelegationReport {
    [CmdletBinding()]
    param(
        [Parameter()]
        [string[]]$Identity,

        [Parameter()]
        [switch]$IncludeFolderPermissions,

        [Parameter()]
        [switch]$IncludeSelf,

        [Parameter()]
        [string[]]$FolderName = @('Calendar', 'Inbox'),

        [Parameter()]
        [string]$ExportPath
    )

    Assert-M365ExchangePowerShellConnected

    $mailboxes = if ($Identity) {
        foreach ($mailboxIdentity in $Identity) {
            Get-EXOMailbox -Identity $mailboxIdentity -Properties GrantSendOnBehalfTo -ErrorAction Stop
        }
    }
    else {
        Get-EXOMailbox -RecipientTypeDetails UserMailbox, SharedMailbox -ResultSize Unlimited -Properties GrantSendOnBehalfTo
    }

    $results = foreach ($mailbox in $mailboxes | Sort-Object DisplayName) {
        $excludeSelf = -not $IncludeSelf

        $fullAccessEntries = Get-MailboxPermission -Identity $mailbox.Identity -ErrorAction Stop |
            Where-Object {
                -not $_.IsInherited -and
                $_.AccessRights -contains 'FullAccess' -and
                (($excludeSelf -and $_.User -ne 'NT AUTHORITY\\SELF') -or (-not $excludeSelf)) -and
                $_.User -notmatch '^S-1-5-'
            }

        $sendAsEntries = Get-RecipientPermission -Identity $mailbox.Identity -ErrorAction SilentlyContinue |
            Where-Object {
                $_.AccessRights -contains 'SendAs' -and
                (($excludeSelf -and $_.Trustee -ne 'NT AUTHORITY\\SELF') -or (-not $excludeSelf)) -and
                $_.Trustee -notmatch '^S-1-5-'
            }

        $sendOnBehalfEntries = @($mailbox.GrantSendOnBehalfTo)
        $folderPermissionEntries = if ($IncludeFolderPermissions) {
            @(Get-M365MailboxFolderDelegationEntries -Mailbox $mailbox -FolderName $FolderName -IncludeSelf:$IncludeSelf)
        }
        else {
            @()
        }

        if (-not $fullAccessEntries -and -not $sendAsEntries -and -not $sendOnBehalfEntries -and -not $folderPermissionEntries) {
            [pscustomobject]@{
                MailboxDisplayName = $mailbox.DisplayName
                MailboxAddress     = $mailbox.PrimarySmtpAddress
                Scope              = 'Mailbox'
                FolderName         = ''
                PermissionType     = 'None'
                PermissionDetails  = ''
                GrantedTo          = ''
            }
            continue
        }

        foreach ($entry in $fullAccessEntries) {
            [pscustomobject]@{
                MailboxDisplayName = $mailbox.DisplayName
                MailboxAddress     = $mailbox.PrimarySmtpAddress
                Scope              = 'Mailbox'
                FolderName         = ''
                PermissionType     = 'FullAccess'
                PermissionDetails  = 'FullAccess'
                GrantedTo          = $entry.User.ToString()
            }
        }

        foreach ($entry in $sendAsEntries) {
            [pscustomobject]@{
                MailboxDisplayName = $mailbox.DisplayName
                MailboxAddress     = $mailbox.PrimarySmtpAddress
                Scope              = 'Mailbox'
                FolderName         = ''
                PermissionType     = 'SendAs'
                PermissionDetails  = 'SendAs'
                GrantedTo          = $entry.Trustee.ToString()
            }
        }

        foreach ($entry in $sendOnBehalfEntries) {
            [pscustomobject]@{
                MailboxDisplayName = $mailbox.DisplayName
                MailboxAddress     = $mailbox.PrimarySmtpAddress
                Scope              = 'Mailbox'
                FolderName         = ''
                PermissionType     = 'SendOnBehalf'
                PermissionDetails  = 'SendOnBehalf'
                GrantedTo          = $entry.Name
            }
        }

        foreach ($entry in $folderPermissionEntries) {
            $entry
        }
    }

    Export-M365ReportData -InputObject $results -ExportPath $ExportPath
}