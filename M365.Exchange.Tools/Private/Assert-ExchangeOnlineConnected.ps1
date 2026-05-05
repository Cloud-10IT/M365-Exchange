function Assert-ExchangeOnlineConnected {
    if (-not (Test-ExchangeOnlineConnection)) {
        throw 'Not connected to Microsoft Graph. Run Connect-M365ExchangeTools first.'
    }
}