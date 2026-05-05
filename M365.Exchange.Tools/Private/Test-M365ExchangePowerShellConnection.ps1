function Test-M365ExchangePowerShellConnection {
    $connectionCommand = Get-Command -Name Get-ConnectionInformation -ErrorAction SilentlyContinue
    if ($connectionCommand) {
        try {
            $connections = @(Get-ConnectionInformation -ErrorAction Stop)
        }
        catch {
            $connections = @()
        }

        $hasConnectedExoSession = $connections | Where-Object {
            (($_.State -eq 'Connected') -or ($_.State -eq 'Open')) -and (
                ($_.Name -match 'Exchange') -or
                ($_.ModuleName -match 'ExchangeOnline') -or
                ([string]$_.ConnectionUri -match 'outlook\.office365\.com|ps\.outlook\.com')
            )
        }

        if ($hasConnectedExoSession) {
            return $true
        }
    }

    try {
        $psSessions = @(Get-PSSession -ErrorAction Stop)
    }
    catch {
        $psSessions = @()
    }

    return [bool]($psSessions | Where-Object {
        $_.State -eq 'Opened' -and [string]$_.ComputerName -match 'outlook\.office365\.com|ps\.outlook\.com'
    })
}
