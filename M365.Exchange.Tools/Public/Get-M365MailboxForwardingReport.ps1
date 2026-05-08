function Get-M365MailboxForwardingReport {
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$ExportPath
    )

    Assert-M365ExchangePowerShellConnected

    Write-Host 'Fetching mailboxes with forwarding rules...' -ForegroundColor Cyan

    # PropertySets Delivery is required for ForwardingSmtpAddress / ForwardingAddress to be populated.
    # Use a server-side filter so only mailboxes with external SMTP forwarding are returned in the first pass;
    # then fetch internal ForwardingAddress separately to avoid loading all mailboxes.
    $externalFwd = @(Get-EXOMailbox -Filter 'ForwardingSmtpAddress -ne $null' -ResultSize Unlimited -PropertySets Minimum,Delivery)

    # Internal AD-object forwarding (ForwardingAddress) cannot always be filtered server-side reliably,
    # so fetch all and filter client-side — only load Minimum+Delivery to keep it fast.
    $internalFwd = @(
        Get-EXOMailbox -ResultSize Unlimited -PropertySets Minimum,Delivery |
            Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.ForwardingAddress) -and
                           [string]::IsNullOrWhiteSpace([string]$_.ForwardingSmtpAddress) }
    )

    $allMailboxes = @($externalFwd) + @($internalFwd) | Sort-Object -Property PrimarySmtpAddress -Unique

    $rows = @(
        $allMailboxes |
            Where-Object {
                -not [string]::IsNullOrWhiteSpace([string]$_.ForwardingSmtpAddress) -or
                -not [string]::IsNullOrWhiteSpace([string]$_.ForwardingAddress)
            } |
            ForEach-Object {
                [pscustomobject]@{
                    DisplayName                = [string]$_.DisplayName
                    UserPrincipalName          = [string]$_.UserPrincipalName
                    PrimarySmtpAddress         = [string]$_.PrimarySmtpAddress
                    RecipientType              = [string]$_.RecipientTypeDetails
                    ForwardingSmtpAddress      = [string]$_.ForwardingSmtpAddress
                    ForwardingAddress          = [string]$_.ForwardingAddress
                    DeliverToMailboxAndForward = $_.DeliverToMailboxAndForward
                    HiddenFromAddressLists     = $_.HiddenFromAddressListsEnabled
                }
            }
    )

    if (-not [string]::IsNullOrWhiteSpace($ExportPath)) {
        Export-M365ReportData -InputObject $rows -ExportPath $ExportPath | Out-Null
    }

    return $rows
}
