function Get-M365UiSettings {
    [CmdletBinding()]
    param(
        [Parameter()]
        [switch]$ForceRefresh
    )

    $moduleRoot = Split-Path -Path $PSScriptRoot -Parent
    $settingsPath = Join-Path -Path $moduleRoot -ChildPath 'Config\M365.Exchange.Tools.Settings.json'

    $defaultSettings = @{
        BrowserPopout = 'Edge'
        CompanyName = ''
        LogoPath = ''
        LogoDataUri = ''
        ReportSavePath = ''
        FileNameTemplate = '{Title}-{Timestamp}'
        HtmlBrandingEnabled = $true
        HtmlShowCompanyName = $true
        HtmlShowCompanyLogo = $true
        ThemePrimaryColor = '#0f766e'
        ThemeSecondaryColor = '#1e293b'
        ReportFontFamily = 'Segoe UI'
        ExchangeAuthMode = 'Auto'
    }

    if (-not $ForceRefresh -and $script:M365UiSettingsCache -and $script:M365UiSettingsCachePath -eq $settingsPath) {
        return @{} + $script:M365UiSettingsCache
    }

    if (-not (Test-Path -Path $settingsPath)) {
        $settingsDirectory = Split-Path -Path $settingsPath -Parent
        if (-not (Test-Path -Path $settingsDirectory)) {
            New-Item -Path $settingsDirectory -ItemType Directory -Force | Out-Null
        }

        $defaultSettings |
            ConvertTo-Json -Depth 5 |
            Set-Content -Path $settingsPath -Encoding UTF8

        $script:M365UiSettingsCache = @{} + $defaultSettings
        $script:M365UiSettingsCachePath = $settingsPath
        return @{} + $script:M365UiSettingsCache
    }

    try {
        $settings = Get-Content -Path $settingsPath -Raw | ConvertFrom-Json -ErrorAction Stop
    }
    catch {
        $settings = [pscustomobject]$defaultSettings
    }

    if ($settings.PSObject.Properties.Name -contains 'UseEdgePopout' -and -not ($settings.PSObject.Properties.Name -contains 'BrowserPopout')) {
        $settings | Add-Member -NotePropertyName 'BrowserPopout' -NotePropertyValue (if ($settings.UseEdgePopout) { 'Edge' } else { 'None' }) -Force
    }

    $browserPopout = [string]$settings.BrowserPopout
    if ([string]::IsNullOrWhiteSpace($browserPopout) -or $browserPopout -in @('True', 'False')) {
        $browserPopout = 'Edge'
    }

    $resolvedSettings = @{
        BrowserPopout = $browserPopout
        CompanyName = [string]$settings.CompanyName
        LogoPath = [string]$settings.LogoPath
        LogoDataUri = if ($settings.PSObject.Properties.Name -contains 'LogoDataUri') { [string]$settings.LogoDataUri } else { '' }
        ReportSavePath = [string]$settings.ReportSavePath
        FileNameTemplate = if ([string]::IsNullOrWhiteSpace([string]$settings.FileNameTemplate)) { '{Title}-{Timestamp}' } else { [string]$settings.FileNameTemplate }
        HtmlBrandingEnabled = if ($null -eq $settings.HtmlBrandingEnabled) { $true } else { [bool]$settings.HtmlBrandingEnabled }
        HtmlShowCompanyName = if ($null -eq $settings.HtmlShowCompanyName) { $true } else { [bool]$settings.HtmlShowCompanyName }
        HtmlShowCompanyLogo = if ($null -eq $settings.HtmlShowCompanyLogo) { $true } else { [bool]$settings.HtmlShowCompanyLogo }
        ThemePrimaryColor = if ([string]::IsNullOrWhiteSpace([string]$settings.ThemePrimaryColor)) { '#0f766e' } else { [string]$settings.ThemePrimaryColor }
        ThemeSecondaryColor = if ([string]::IsNullOrWhiteSpace([string]$settings.ThemeSecondaryColor)) { '#1e293b' } else { [string]$settings.ThemeSecondaryColor }
        ReportFontFamily = if ([string]::IsNullOrWhiteSpace([string]$settings.ReportFontFamily)) { 'Segoe UI' } else { [string]$settings.ReportFontFamily }
        ExchangeAuthMode = if ([string]::IsNullOrWhiteSpace([string]$settings.ExchangeAuthMode)) { 'Auto' } else { [string]$settings.ExchangeAuthMode }
    }

    $script:M365UiSettingsCache = @{} + $resolvedSettings
    $script:M365UiSettingsCachePath = $settingsPath

    return @{} + $script:M365UiSettingsCache
}