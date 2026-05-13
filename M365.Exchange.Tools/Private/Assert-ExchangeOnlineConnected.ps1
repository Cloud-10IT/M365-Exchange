function Assert-ExchangeOnlineConnected {
    [CmdletBinding()]
    param()

    # Legacy compatibility wrapper.
    Assert-M365GraphConnected
}