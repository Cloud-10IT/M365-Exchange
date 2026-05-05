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

    $connectParams = @{
        ShowBanner        = $false
        UseMultithreading = $true
    }

    if ($UserPrincipalName) {
        $connectParams['UserPrincipalName'] = $UserPrincipalName
    }

    Connect-ExchangeOnline @connectParams

    if (-not (Test-M365ExchangePowerShellConnection)) {
        throw 'Sign-in completed but an Exchange Online PowerShell session was not detected. Try reconnecting.'
    }

    Write-Host 'Connected to Exchange Online PowerShell.' -ForegroundColor Green
}
