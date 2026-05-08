function Connect-M365ExchangePowerShell {
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$UserPrincipalName
    )

    function Test-M365ExchangeWamBrokerError {
        param(
            [Parameter(Mandatory)]
            [System.Management.Automation.ErrorRecord]$ErrorRecord
        )

        $parts = [System.Collections.Generic.List[string]]::new()
        $parts.Add($ErrorRecord.ToString())

        $ex = $ErrorRecord.Exception
        while ($null -ne $ex) {
            $parts.Add($ex.Message)
            $parts.Add($ex.GetType().FullName)
            if ($ex.StackTrace) { $parts.Add($ex.StackTrace) }
            $ex = $ex.InnerException
        }

        $combined = $parts -join ' '
        return $combined -match 'RuntimeBroker|NullReferenceException.*Broker|Error Acquiring Token'
    }

    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        throw "ExchangeOnlineManagement is not installed. Run 'Install-Module ExchangeOnlineManagement -Scope CurrentUser'."
    }

    Import-Module ExchangeOnlineManagement -ErrorAction Stop | Out-Null
    $uiSettings = Get-M365UiSettings
    $preferredAuthMode = [string]$uiSettings.ExchangeAuthMode

    if (Test-M365ExchangePowerShellConnection) {
        Write-Host 'Exchange Online PowerShell connection already active.' -ForegroundColor Green
        return
    }

    $resolvedUserPrincipalName = $UserPrincipalName
    if ([string]::IsNullOrWhiteSpace($resolvedUserPrincipalName)) {
        $getMgContextCommand = Get-Command -Name Get-MgContext -ErrorAction SilentlyContinue
        if ($getMgContextCommand) {
            try {
                $graphContext = Get-MgContext
                if ($graphContext -and -not [string]::IsNullOrWhiteSpace($graphContext.Account)) {
                    $resolvedUserPrincipalName = $graphContext.Account
                    Write-Host "Using signed-in Microsoft Graph account: $resolvedUserPrincipalName" -ForegroundColor DarkCyan
                }
            }
            catch {
                # Continue with default interactive sign-in if Graph context cannot be read.
            }
        }
    }

    $connectParams = @{
        ShowBanner        = $false
        UseMultithreading = $true
    }

    if (-not [string]::IsNullOrWhiteSpace($resolvedUserPrincipalName)) {
        $connectParams['UserPrincipalName'] = $resolvedUserPrincipalName
    }

    $connectCommand = Get-Command -Name Connect-ExchangeOnline -ErrorAction Stop

    $attemptDefinitions = [System.Collections.Generic.List[object]]::new()
    $attemptDefinitions.Add(@{
        Name = 'UserPrincipalName'
        Label = if ($connectParams.ContainsKey('UserPrincipalName')) { 'signed-in account' } else { 'default interactive sign-in' }
        Enabled = $connectParams.ContainsKey('UserPrincipalName')
        Params = @{} + $connectParams
    })

    $interactiveParams = @{} + $connectParams
    $interactiveParams.Remove('UserPrincipalName') | Out-Null
    $attemptDefinitions.Add(@{
        Name = 'Interactive'
        Label = 'default interactive sign-in'
        Enabled = $true
        Params = $interactiveParams
    })

    $disableWamParams = @{} + $connectParams
    $disableWamParams.Remove('UserPrincipalName') | Out-Null
    $disableWamParams['DisableWAM'] = $true
    $attemptDefinitions.Add(@{
        Name = 'DisableWAM'
        Label = 'interactive sign-in with DisableWAM'
        Enabled = $connectCommand.Parameters.ContainsKey('DisableWAM')
        Params = $disableWamParams
    })

    $deviceParams = @{} + $connectParams
    $deviceParams.Remove('UserPrincipalName') | Out-Null
    $deviceParams['Device'] = $true
    $attemptDefinitions.Add(@{
        Name = 'Device'
        Label = 'device code sign-in'
        Enabled = $connectCommand.Parameters.ContainsKey('Device')
        Params = $deviceParams
    })

    $preferredOrder = switch ($preferredAuthMode) {
        'Interactive' { @('Interactive', 'DisableWAM', 'Device') }
        'DisableWAM' { @('DisableWAM', 'Device', 'Interactive') }
        'Device' { @('Device', 'DisableWAM', 'Interactive') }
        Default { @('UserPrincipalName', 'Interactive', 'DisableWAM', 'Device') }
    }

    $attempts = @()
    foreach ($attemptName in $preferredOrder) {
        $match = $attemptDefinitions | Where-Object { $_.Name -eq $attemptName -and $_.Enabled } | Select-Object -First 1
        if ($match) {
            $attempts += $match
        }
    }

    $attemptIndex = 0
    $lastError = $null
    $skipInteractiveRetries = $false
    foreach ($attempt in $attempts) {
        if ($skipInteractiveRetries -and $attempt.Name -ne 'Device') {
            continue
        }

        try {
            if ($attemptIndex -gt 0) {
                Write-Host "Connect-ExchangeOnline failed previously. Retrying with $($attempt.Label)." -ForegroundColor Yellow
            }

            $attemptParams = $attempt.Params
            Connect-ExchangeOnline @attemptParams -ErrorAction Stop
            $lastError = $null
            break
        }
        catch {
            $lastError = $_
            if ($attempt.Name -ne 'Device' -and (Test-M365ExchangeWamBrokerError -ErrorRecord $_)) {
                $skipInteractiveRetries = $true
                Write-Host 'Detected Windows broker token acquisition failure. Switching directly to device code sign-in.' -ForegroundColor Yellow
            }

            $attemptIndex++
        }
    }

    if ($lastError) {
        throw $lastError
    }

    if (-not (Test-M365ExchangePowerShellConnection)) {
        throw 'Sign-in completed but an Exchange Online PowerShell session was not detected. Try reconnecting.'
    }

    Write-Host 'Connected to Exchange Online PowerShell.' -ForegroundColor Green
}
