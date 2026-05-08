function Get-M365SharedMailboxInventory {
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$ExportPath
    )

    Assert-M365ExchangePowerShellConnected

    Write-Host 'Retrieving shared mailboxes from Exchange Online...' -ForegroundColor Cyan

    $mailboxes = @(
        Get-EXOMailbox -ResultSize Unlimited `
            -Filter "RecipientTypeDetails -eq 'SharedMailbox'" `
            -Properties DisplayName, PrimarySmtpAddress, Alias,
                        HiddenFromAddressListsEnabled, WhenCreated,
                        ForwardingSmtpAddress, DeliverToMailboxAndForward,
                        LitigationHoldEnabled, RetentionPolicy `
            -ErrorAction Stop |
            Sort-Object DisplayName
    )

    Write-Host "Found $($mailboxes.Count) shared mailbox(es)." -ForegroundColor DarkCyan

    $results = foreach ($mbx in $mailboxes) {
        [pscustomobject]@{
            DisplayName                   = $mbx.DisplayName
            PrimarySmtpAddress            = $mbx.PrimarySmtpAddress
            Alias                         = $mbx.Alias
            HiddenFromAddressListsEnabled = $mbx.HiddenFromAddressListsEnabled
            ForwardingSmtpAddress         = $mbx.ForwardingSmtpAddress
            DeliverToMailboxAndForward    = $mbx.DeliverToMailboxAndForward
            LitigationHoldEnabled         = $mbx.LitigationHoldEnabled
            RetentionPolicy               = $mbx.RetentionPolicy
            WhenCreated                   = $mbx.WhenCreated
            MailboxKind                   = 'Shared'
        }
    }

    Export-M365ReportData -InputObject $results -ExportPath $ExportPath
}