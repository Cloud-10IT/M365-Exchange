function Set-M365UiSettings {
    [CmdletBinding()]
    param(
        [Parameter()]
        [ValidateSet('None', 'Edge', 'Firefox', 'Chrome', 'Brave', 'Default')]
        [string]$BrowserPopout,

        [Parameter()]
        [string]$CompanyName,

        [Parameter()]
        [string]$LogoPath
    )

    $moduleRoot = Split-Path -Path $PSScriptRoot -Parent
    $settingsPath = Join-Path -Path $moduleRoot -ChildPath 'Config\M365.Exchange.Tools.Settings.json'

    $settingsDirectory = Split-Path -Path $settingsPath -Parent
    if (-not (Test-Path -Path $settingsDirectory)) {
        New-Item -Path $settingsDirectory -ItemType Directory -Force | Out-Null
    }

    $currentSettings = Get-M365UiSettings

    $settings = @{
        BrowserPopout = if ($PSBoundParameters.ContainsKey('BrowserPopout')) { [string]$BrowserPopout } else { [string]$currentSettings.BrowserPopout }
        CompanyName   = if ($PSBoundParameters.ContainsKey('CompanyName')) { [string]$CompanyName } else { [string]$currentSettings.CompanyName }
        LogoPath      = if ($PSBoundParameters.ContainsKey('LogoPath')) { [string]$LogoPath } else { [string]$currentSettings.LogoPath }
    }

    $settings | ConvertTo-Json -Depth 5 | Set-Content -Path $settingsPath -Encoding UTF8

    return $settings
}
