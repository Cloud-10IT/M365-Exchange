function Get-M365UiSettings {
    $moduleRoot = Split-Path -Path $PSScriptRoot -Parent
    $settingsPath = Join-Path -Path $moduleRoot -ChildPath 'Config\M365.Exchange.Tools.Settings.json'

    $defaultSettings = @{
        BrowserPopout        = 'Edge'
        CompanyName          = ''
        LogoPath             = ''
        ReportSavePath       = ''
        FileNameTemplate     = '{Title}-{Timestamp}'
        HtmlBrandingEnabled  = $true
        HtmlShowCompanyName  = $true
        HtmlShowCompanyLogo  = $true
        ThemePrimaryColor    = '#0f766e'
        ThemeSecondaryColor  = '#1e293b'
        ReportFontFamily     = 'Segoe UI'
        ExchangeAuthMode     = 'Auto'
    }

    if (-not (Test-Path -Path $settingsPath)) {
        $settingsDirectory = Split-Path -Path $settingsPath -Parent
        if (-not (Test-Path -Path $settingsDirectory)) {
            New-Item -Path $settingsDirectory -ItemType Directory -Force | Out-Null
        }

        $defaultSettings | ConvertTo-Json -Depth 5 | Set-Content -Path $settingsPath -Encoding UTF8
        return $defaultSettings
    }

    try {
        $settings = Get-Content -Path $settingsPath -Raw | ConvertFrom-Json -ErrorAction Stop
    }
    catch {
        $settings = [pscustomobject]$defaultSettings
    }

    # Handle backward compatibility: if old UseEdgePopout exists, convert to BrowserPopout
    if ($settings.PSObject.Properties.Name -contains 'UseEdgePopout' -and -not ($settings.PSObject.Properties.Name -contains 'BrowserPopout')) {
        $settings | Add-Member -NotePropertyName 'BrowserPopout' -NotePropertyValue (if ($settings.UseEdgePopout) { 'Edge' } else { 'None' }) -Force
    }

    $browserPopout = [string]$settings.BrowserPopout
    if (-not $browserPopout -or $browserPopout -eq '' -or $browserPopout -eq 'True' -or $browserPopout -eq 'False') {
        $browserPopout = 'Edge'
    }

    return @{
        BrowserPopout       = $browserPopout
        CompanyName         = [string]$settings.CompanyName
        LogoPath            = [string]$settings.LogoPath
        ReportSavePath      = [string]$settings.ReportSavePath
        FileNameTemplate    = if ([string]::IsNullOrWhiteSpace([string]$settings.FileNameTemplate)) { '{Title}-{Timestamp}' } else { [string]$settings.FileNameTemplate }
        HtmlBrandingEnabled = if ($null -eq $settings.HtmlBrandingEnabled) { $true } else { [bool]$settings.HtmlBrandingEnabled }
        HtmlShowCompanyName = if ($null -eq $settings.HtmlShowCompanyName) { $true } else { [bool]$settings.HtmlShowCompanyName }
        HtmlShowCompanyLogo = if ($null -eq $settings.HtmlShowCompanyLogo) { $true } else { [bool]$settings.HtmlShowCompanyLogo }
        ThemePrimaryColor   = if ([string]::IsNullOrWhiteSpace([string]$settings.ThemePrimaryColor)) { '#0f766e' } else { [string]$settings.ThemePrimaryColor }
        ThemeSecondaryColor = if ([string]::IsNullOrWhiteSpace([string]$settings.ThemeSecondaryColor)) { '#1e293b' } else { [string]$settings.ThemeSecondaryColor }
        ReportFontFamily    = if ([string]::IsNullOrWhiteSpace([string]$settings.ReportFontFamily)) { 'Segoe UI' } else { [string]$settings.ReportFontFamily }
        ExchangeAuthMode    = if ([string]::IsNullOrWhiteSpace([string]$settings.ExchangeAuthMode)) { 'Auto' } else { [string]$settings.ExchangeAuthMode }
    }
}
