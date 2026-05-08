function Get-M365ResourceMailboxInventory {
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$ExportPath
    )

    Assert-M365ExchangePowerShellConnected

    Write-Host 'Retrieving resource mailboxes from Exchange Online...' -ForegroundColor Cyan

    # RoomMailbox and EquipmentMailbox are the two RecipientTypeDetails values for resources.
    $mailboxes = @(
        Get-EXOMailbox -ResultSize Unlimited `
            -Filter "RecipientTypeDetails -eq 'RoomMailbox' -or RecipientTypeDetails -eq 'EquipmentMailbox'" `
            -Properties DisplayName, PrimarySmtpAddress, Alias,
                        RecipientTypeDetails, Office,
                        HiddenFromAddressListsEnabled, WhenCreated,
                        ResourceCapacity, LitigationHoldEnabled, RetentionPolicy `
            -ErrorAction Stop |
            Sort-Object DisplayName
    )

    Write-Host "Found $($mailboxes.Count) resource mailbox(es)." -ForegroundColor DarkCyan

    $results = foreach ($mbx in $mailboxes) {
        [pscustomobject]@{
            DisplayName                   = $mbx.DisplayName
            PrimarySmtpAddress            = $mbx.PrimarySmtpAddress
            Alias                         = $mbx.Alias
            ResourceType                  = $mbx.RecipientTypeDetails  # RoomMailbox / EquipmentMailbox
            Office                        = $mbx.Office
            ResourceCapacity              = $mbx.ResourceCapacity
            HiddenFromAddressListsEnabled = $mbx.HiddenFromAddressListsEnabled
            LitigationHoldEnabled         = $mbx.LitigationHoldEnabled
            RetentionPolicy               = $mbx.RetentionPolicy
            WhenCreated                   = $mbx.WhenCreated
            MailboxKind                   = 'Resource'
        }
    }

    Export-M365ReportData -InputObject $results -ExportPath $ExportPath
}
