function Connect-M365ExchangePowerShell {
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$UserPrincipalName
    )

    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        throw "ExchangeOnlineManagement is not installed. Run 'Install-Module ExchangeOnlineManagement -Scope CurrentUser'."
    }

    Import-Module ExchangeOnlineManagement -Force -ErrorAction Stop | Out-Null

    if (Test-M365ExchangePowerShellConnection) {
        Write-Host 'Exchange Online PowerShell connection already active.' -ForegroundColor Green
        return
    }

    $settingsPath = Join-Path $PSScriptRoot '..\Config\M365.Exchange.Tools.Settings.json'
    $settingsPath = [System.IO.Path]::GetFullPath($settingsPath)

    $authMode = 'Device'
    if (Test-Path $settingsPath) {
        try {
            $settings = Get-Content -LiteralPath $settingsPath -Raw | ConvertFrom-Json
            if ($settings.ExchangeAuthMode) {
                $authMode = [string]$settings.ExchangeAuthMode
            }
        }
        catch {
        }
    }

    try {
        switch ($authMode) {
            'Device' {
                Write-Host 'Connecting to Exchange Online PowerShell using device code sign-in...' -ForegroundColor Yellow
                Connect-ExchangeOnline -Device -ShowBanner:$false -ErrorAction Stop
            }

            'DisableWAM' {
                Write-Host 'Connecting to Exchange Online PowerShell using DisableWAM...' -ForegroundColor Yellow
                Connect-ExchangeOnline -DisableWAM -ShowBanner:$false -ErrorAction Stop
            }

            default {
                Write-Host 'Connecting to Exchange Online PowerShell using interactive sign-in...' -ForegroundColor Yellow
                Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
            }
        }
    }
    catch {
        throw "Exchange Online PowerShell connection failed. $($_.Exception.Message)"
    }

    if (-not (Test-M365ExchangePowerShellConnection)) {
        throw 'Sign-in completed but an Exchange Online PowerShell session was not detected.'
    }

    Write-Host 'Connected to Exchange Online PowerShell.' -ForegroundColor Green
}