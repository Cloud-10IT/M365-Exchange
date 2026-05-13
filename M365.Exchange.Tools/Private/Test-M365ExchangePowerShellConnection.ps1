function Test-M365ExchangePowerShellConnection {
    [CmdletBinding()]
    param()

    $connectionCommand = Get-Command -Name Get-ConnectionInformation -ErrorAction SilentlyContinue
    if ($connectionCommand) {
        try {
            $connections = @(Get-ConnectionInformation -ErrorAction Stop)
            foreach ($connection in $connections) {
                if (
                    $connection.State -eq 'Connected' -and
                    (
                        $connection.Name -match 'ExchangeOnline' -or
                        $connection.ConnectionUri -match 'outlook\.office365\.com'
                    )
                ) {
                    return $true
                }
            }
        }
        catch {
        }
    }

    return $false
}