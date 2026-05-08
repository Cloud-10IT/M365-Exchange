function Set-M365UiSettings {
    [CmdletBinding()]
    param(
        [Parameter()]
        [ValidateSet('None', 'Edge', 'Firefox', 'Chrome', 'Brave', 'Default')]
        [string]$BrowserPopout,

        [Parameter()]
        [string]$CompanyName,

        [Parameter()]
        [string]$LogoPath,

        [Parameter()]
        [string]$ReportSavePath,

        [Parameter()]
        [string]$FileNameTemplate,

        [Parameter()]
        [bool]$HtmlBrandingEnabled,

        [Parameter()]
        [bool]$HtmlShowCompanyName,

        [Parameter()]
        [bool]$HtmlShowCompanyLogo,

        [Parameter()]
        [string]$ThemePrimaryColor,

        [Parameter()]
        [string]$ThemeSecondaryColor,

        [Parameter()]
        [string]$ReportFontFamily
    )

    $moduleRoot = Split-Path -Path $PSScriptRoot -Parent
    $settingsPath = Join-Path -Path $moduleRoot -ChildPath 'Config\M365.Exchange.Tools.Settings.json'

    $settingsDirectory = Split-Path -Path $settingsPath -Parent
    if (-not (Test-Path -Path $settingsDirectory)) {
        New-Item -Path $settingsDirectory -ItemType Directory -Force | Out-Null
    }

    $currentSettings = Get-M365UiSettings

    $settings = @{
        BrowserPopout       = if ($PSBoundParameters.ContainsKey('BrowserPopout')) { [string]$BrowserPopout } else { [string]$currentSettings.BrowserPopout }
        CompanyName         = if ($PSBoundParameters.ContainsKey('CompanyName')) { [string]$CompanyName } else { [string]$currentSettings.CompanyName }
        LogoPath            = if ($PSBoundParameters.ContainsKey('LogoPath')) { [string]$LogoPath } else { [string]$currentSettings.LogoPath }
        ReportSavePath      = if ($PSBoundParameters.ContainsKey('ReportSavePath')) { [string]$ReportSavePath } else { [string]$currentSettings.ReportSavePath }
        FileNameTemplate    = if ($PSBoundParameters.ContainsKey('FileNameTemplate')) { [string]$FileNameTemplate } else { [string]$currentSettings.FileNameTemplate }
        HtmlBrandingEnabled = if ($PSBoundParameters.ContainsKey('HtmlBrandingEnabled')) { [bool]$HtmlBrandingEnabled } else { [bool]$currentSettings.HtmlBrandingEnabled }
        HtmlShowCompanyName = if ($PSBoundParameters.ContainsKey('HtmlShowCompanyName')) { [bool]$HtmlShowCompanyName } else { [bool]$currentSettings.HtmlShowCompanyName }
        HtmlShowCompanyLogo = if ($PSBoundParameters.ContainsKey('HtmlShowCompanyLogo')) { [bool]$HtmlShowCompanyLogo } else { [bool]$currentSettings.HtmlShowCompanyLogo }
        ThemePrimaryColor   = if ($PSBoundParameters.ContainsKey('ThemePrimaryColor')) { [string]$ThemePrimaryColor } else { [string]$currentSettings.ThemePrimaryColor }
        ThemeSecondaryColor = if ($PSBoundParameters.ContainsKey('ThemeSecondaryColor')) { [string]$ThemeSecondaryColor } else { [string]$currentSettings.ThemeSecondaryColor }
        ReportFontFamily    = if ($PSBoundParameters.ContainsKey('ReportFontFamily')) { [string]$ReportFontFamily } else { [string]$currentSettings.ReportFontFamily }
    }

    $settings | ConvertTo-Json -Depth 5 | Set-Content -Path $settingsPath -Encoding UTF8

    return $settings
}
