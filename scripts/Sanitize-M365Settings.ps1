param(
    [Parameter()]
    [string]$SettingsPath = (Join-Path -Path $PSScriptRoot -ChildPath '..\M365.Exchange.Tools\Config\M365.Exchange.Tools.Settings.json')
)

$resolvedPath = [System.IO.Path]::GetFullPath($SettingsPath)

if (-not (Test-Path -Path $resolvedPath)) {
    Write-Host "Settings file not found, skipping sanitization: $resolvedPath" -ForegroundColor Yellow
    exit 0
}

$sanitized = [ordered]@{
    FileNameTemplate    = '{Title}-{Timestamp}'
    HtmlShowCompanyLogo = $true
    HtmlShowCompanyName = $true
    LogoPath            = ''
    BrowserPopout       = 'Default'
    HtmlBrandingEnabled = $true
    ReportSavePath      = ''
    CompanyName         = 'Contoso'
    ThemePrimaryColor   = '#0f766e'
    ThemeSecondaryColor = '#1e293b'
    ReportFontFamily    = 'Segoe UI'
}

$json = $sanitized | ConvertTo-Json -Depth 5
Set-Content -Path $resolvedPath -Value $json -Encoding UTF8

Write-Host "Sanitized settings: $resolvedPath" -ForegroundColor Green