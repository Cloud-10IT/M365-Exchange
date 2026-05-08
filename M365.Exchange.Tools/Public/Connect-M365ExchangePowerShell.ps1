function Connect-M365ExchangePowerShell {
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$UserPrincipalName
    )

    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        throw "ExchangeOnlineManagement is not installed. Run 'Install-Module ExchangeOnlineManagement -Scope CurrentUser'."
    }

    Import-Module ExchangeOnlineManagement -ErrorAction Stop | Out-Null

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
    $attemptedWithUserPrincipalName = $connectParams.ContainsKey('UserPrincipalName')

    try {
        Connect-ExchangeOnline @connectParams -ErrorAction Stop
    }
    catch {
        if ($attemptedWithUserPrincipalName) {
            Write-Host 'Connect-ExchangeOnline failed with UserPrincipalName. Retrying with default interactive sign-in.' -ForegroundColor Yellow
            $connectParams.Remove('UserPrincipalName') | Out-Null
            Connect-ExchangeOnline @connectParams -ErrorAction Stop
        }
        elseif ($connectCommand.Parameters.ContainsKey('UseDeviceAuthentication')) {
            Write-Host 'Interactive sign-in failed. Retrying with device authentication.' -ForegroundColor Yellow
            $connectParams['UseDeviceAuthentication'] = $true
            Connect-ExchangeOnline @connectParams -ErrorAction Stop
        }
        else {
            throw
        }
    }

    if (-not (Test-M365ExchangePowerShellConnection)) {
        throw 'Sign-in completed but an Exchange Online PowerShell session was not detected. Try reconnecting.'
    }

    Write-Host 'Connected to Exchange Online PowerShell.' -ForegroundColor Green
}
