function Assert-M365ExchangePowerShellConnected {
    if (-not (Test-M365ExchangePowerShellConnection)) {
        throw 'Not connected to Exchange Online PowerShell. Run Connect-M365ExchangePowerShell first.'
    }
}
