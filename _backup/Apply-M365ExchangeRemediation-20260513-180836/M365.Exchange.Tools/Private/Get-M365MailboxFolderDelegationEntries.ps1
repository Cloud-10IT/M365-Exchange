function Get-M365MailboxFolderDelegationEntries {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$Mailbox,

        [Parameter()]
        [switch]$IncludeSelf,

        [Parameter()]
        [string[]]$FolderName = @('Calendar', 'Inbox')
    )

    foreach ($name in $FolderName) {
        $folderIdentity = '{0}:\{1}' -f $Mailbox.PrimarySmtpAddress, $name

        try {
            $permissions = Get-MailboxFolderPermission -Identity $folderIdentity -ErrorAction Stop |
                Where-Object {
                    $_.User -ne 'Default' -and
                    $_.User -ne 'Anonymous' -and
                    ($IncludeSelf -or $_.User -ne 'NT AUTHORITY\\SELF') -and
                    $_.User -notmatch '^S-1-5-'
                }
        }
        catch {
            if ($_.Exception.Message -match 'cannot be found|doesn''t exist|Cannot process argument') {
                continue
            }

            throw
        }

        foreach ($permission in $permissions) {
            [pscustomobject]@{
                MailboxDisplayName = $Mailbox.DisplayName
                MailboxAddress     = $Mailbox.PrimarySmtpAddress
                Scope              = 'Folder'
                FolderName         = $name
                PermissionType     = 'MailboxFolderPermission'
                PermissionDetails  = ($permission.AccessRights -join ', ')
                GrantedTo          = $permission.User.ToString()
            }
        }
    }
}