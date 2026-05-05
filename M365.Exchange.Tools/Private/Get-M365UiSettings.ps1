function Get-M365UiSettings {
    $moduleRoot = Split-Path -Path $PSScriptRoot -Parent
    $settingsPath = Join-Path -Path $moduleRoot -ChildPath 'Config\M365.Exchange.Tools.Settings.json'

    $defaultSettings = @{
        BrowserPopout = 'Edge'
        CompanyName   = ''
        LogoPath      = ''
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
        BrowserPopout = $browserPopout
        CompanyName   = [string]$settings.CompanyName
        LogoPath      = [string]$settings.LogoPath
    }
}
