# GPT-4.1 APPLY FILE — EXECUTE, DO NOT SUMMARIZE

## Required behavior
You must **apply** the file operations in this document.

- Read this file completely.
- Create/update/delete the files exactly as specified.
- Do not stop after reading.
- Do not summarize the changes instead of writing files.
- Do not output diffs.
- Do not ask for confirmation.
- Do not change files not listed here.
- Remove the checkpoint API files listed here.
- Complete the whole operation in one pass.

## Final response requirement
After the workspace changes are actually applied, reply with exactly:

`Applied remediation pack successfully.`

If any file operation fails, reply with exactly:

`Remediation pack failed.`

No other text.

---

# DELETE THESE FILES

Delete these files completely:

- `reference/checkpoint_api_credentials.json`
- `reference/checkpoint_api_credentials_endpoint.json`
- `reference/checkpoint_api_credentials_monitoring.json`
- `reference/checkpoint_api_reference.json`
- `scripts/Test-CheckpointAIBrowserEvents.ps1`

If the `reference` folder becomes empty, delete the empty folder.

---

# CREATE THIS FILE

## `M365.Exchange.Tools\Private\Test-M365GraphConnection.ps1`

```powershell
function Test-M365GraphConnection {
    [CmdletBinding()]
    param(
        [Parameter()]
        [string[]]$RequiredScopes = @(),

        [Parameter()]
        [switch]$RequireAllScopes
    )

    $getContextCommand = Get-Command -Name Get-MgContext -ErrorAction SilentlyContinue
    if (-not $getContextCommand) {
        return $false
    }

    try {
        $context = Get-MgContext -ErrorAction Stop
    }
    catch {
        return $false
    }

    if (-not $context) {
        return $false
    }

    if ([string]::IsNullOrWhiteSpace([string]$context.Account) -or [string]::IsNullOrWhiteSpace([string]$context.TenantId)) {
        return $false
    }

    if (-not $RequiredScopes -or $RequiredScopes.Count -eq 0) {
        return $true
    }

    $grantedScopes = @($context.Scopes)
    if (-not $grantedScopes -or $grantedScopes.Count -eq 0) {
        return $false
    }

    $scopeLookup = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($scope in $grantedScopes) {
        if (-not [string]::IsNullOrWhiteSpace([string]$scope)) {
            [void]$scopeLookup.Add([string]$scope)
        }
    }

    if ($RequireAllScopes) {
        foreach ($requiredScope in $RequiredScopes) {
            if ([string]::IsNullOrWhiteSpace([string]$requiredScope)) {
                continue
            }

            if (-not $scopeLookup.Contains([string]$requiredScope)) {
                return $false
            }
        }

        return $true
    }

    foreach ($requiredScope in $RequiredScopes) {
        if ([string]::IsNullOrWhiteSpace([string]$requiredScope)) {
            continue
        }

        if ($scopeLookup.Contains([string]$requiredScope)) {
            return $true
        }
    }

    return $false
}
```

---

# CREATE THIS FILE

## `M365.Exchange.Tools\Private\Assert-M365GraphConnected.ps1`

```powershell
function Assert-M365GraphConnected {
    [CmdletBinding()]
    param(
        [Parameter()]
        [string[]]$RequiredScopes = @(
            'User.Read.All',
            'Group.Read.All',
            'Directory.Read.All',
            'Organization.Read.All',
            'Policy.Read.All',
            'OrgContact.Read.All',
            'MailboxSettings.Read',
            'AuditLog.Read.All'
        )
    )

    if (-not (Test-M365GraphConnection -RequiredScopes $RequiredScopes -RequireAllScopes)) {
        throw 'Not connected to Microsoft Graph with the required scopes. Run Connect-M365ExchangeTools first.'
    }
}
```

---

# OVERWRITE THIS FILE

## `M365.Exchange.Tools\Private\Test-ExchangeOnlineConnection.ps1`

```powershell
function Test-ExchangeOnlineConnection {
    [CmdletBinding()]
    param(
        [Parameter()]
        [string[]]$RequiredScopes = @(
            'User.Read.All',
            'Group.Read.All',
            'Directory.Read.All',
            'Organization.Read.All',
            'Policy.Read.All',
            'OrgContact.Read.All',
            'MailboxSettings.Read',
            'AuditLog.Read.All'
        )
    )

    # Legacy compatibility wrapper.
    return (Test-M365GraphConnection -RequiredScopes $RequiredScopes -RequireAllScopes)
}
```

---

# OVERWRITE THIS FILE

## `M365.Exchange.Tools\Private\Assert-ExchangeOnlineConnected.ps1`

```powershell
function Assert-ExchangeOnlineConnected {
    [CmdletBinding()]
    param()

    # Legacy compatibility wrapper.
    Assert-M365GraphConnected
}
```

---

# OVERWRITE THIS FILE

## `M365.Exchange.Tools\Private\Get-M365GraphCollection.ps1`

```powershell
function Get-M365GraphCollection {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Uri,

        [Parameter()]
        [ValidateRange(1, 10)]
        [int]$MaxRetryCount = 5,

        [Parameter()]
        [ValidateRange(1, 300)]
        [int]$DefaultRetryDelaySeconds = 3
    )

    function Get-M365GraphStatusCode {
        param(
            [Parameter(Mandatory)]
            [System.Management.Automation.ErrorRecord]$ErrorRecord
        )

        try {
            if ($ErrorRecord.Exception -and $ErrorRecord.Exception.ResponseStatusCode) {
                return [int]$ErrorRecord.Exception.ResponseStatusCode
            }
        }
        catch {
        }

        try {
            if ($ErrorRecord.Exception -and $ErrorRecord.Exception.Response -and $ErrorRecord.Exception.Response.StatusCode) {
                return [int]$ErrorRecord.Exception.Response.StatusCode
            }
        }
        catch {
        }

        $message = [string]$ErrorRecord.Exception.Message
        if ($message -match '\b429\b') {
            return 429
        }
        if ($message -match '\b503\b') {
            return 503
        }

        return $null
    }

    function Get-M365GraphRetryDelaySeconds {
        param(
            [Parameter(Mandatory)]
            [System.Management.Automation.ErrorRecord]$ErrorRecord,

            [Parameter(Mandatory)]
            [int]$FallbackDelaySeconds
        )

        try {
            $exception = $ErrorRecord.Exception
            if ($exception -and $exception.Response -and $exception.Response.Headers) {
                $retryAfter = $exception.Response.Headers['Retry-After']
                if ($retryAfter) {
                    $retryAfterValue = [string]($retryAfter | Select-Object -First 1)
                    $parsedSeconds = 0
                    if ([int]::TryParse($retryAfterValue, [ref]$parsedSeconds)) {
                        return [Math]::Max($parsedSeconds, 1)
                    }
                }
            }
        }
        catch {
        }

        return $FallbackDelaySeconds
    }

    $results = [System.Collections.Generic.List[object]]::new()
    $nextUri = $Uri

    while (-not [string]::IsNullOrWhiteSpace($nextUri)) {
        $response = $null
        $attempt = 0

        while ($true) {
            try {
                $response = Invoke-MgGraphRequest -Method GET -Uri $nextUri -OutputType PSObject -ErrorAction Stop
                break
            }
            catch {
                $attempt++
                $statusCode = Get-M365GraphStatusCode -ErrorRecord $_
                $isRetryable = ($statusCode -in @(429, 503))

                if (-not $isRetryable -or $attempt -ge $MaxRetryCount) {
                    throw
                }

                $delaySeconds = Get-M365GraphRetryDelaySeconds -ErrorRecord $_ -FallbackDelaySeconds ($DefaultRetryDelaySeconds * $attempt)
                Start-Sleep -Seconds $delaySeconds
            }
        }

        if ($null -eq $response) {
            break
        }

        if ($response.PSObject.Properties.Name -contains 'value' -and $null -ne $response.value) {
            foreach ($item in @($response.value)) {
                [void]$results.Add($item)
            }
        }
        elseif ($response -is [System.Collections.IEnumerable] -and -not ($response -is [string])) {
            foreach ($item in $response) {
                [void]$results.Add($item)
            }
        }
        else {
            [void]$results.Add($response)
        }

        $nextUri = $null
        if ($response.PSObject.Properties.Name -contains '@odata.nextLink') {
            $nextUri = [string]$response.'@odata.nextLink'
        }
    }

    return @($results)
}
```

---

# OVERWRITE THIS FILE

## `M365.Exchange.Tools\Private\Get-M365MailboxFolderDelegationEntries.ps1`

```powershell
function Get-M365MailboxFolderDelegationEntries {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$Mailbox,

        [Parameter()]
        [switch]$IncludeSelf,

        [Parameter()]
        [string[]]$FolderName = @('Calendar', 'Inbox')
    )

    foreach ($name in $FolderName) {
        $folderIdentity = '{0}:\{1}' -f $Mailbox.PrimarySmtpAddress, $name

        try {
            $permissions = Get-MailboxFolderPermission -Identity $folderIdentity -ErrorAction Stop |
                Where-Object {
                    $_.User -ne 'Default' -and
                    $_.User -ne 'Anonymous' -and
                    ($IncludeSelf -or $_.User -ne 'NT AUTHORITY\SELF') -and
                    $_.User -notmatch '^S-1-5-'
                }
        }
        catch {
            $fullyQualifiedErrorId = [string]$_.FullyQualifiedErrorId
            $category = [string]$_.CategoryInfo.Category
            $message = [string]$_.Exception.Message

            $isExpectedMissingFolderError = $false

            if ($category -in @('ObjectNotFound', 'InvalidArgument')) {
                $isExpectedMissingFolderError = $true
            }
            elseif ($fullyQualifiedErrorId -match 'ObjectNotFound|ManagementObjectNotFoundException|FolderNotFound') {
                $isExpectedMissingFolderError = $true
            }
            elseif ($message -match 'cannot be found|doesn''t exist|Cannot process argument') {
                $isExpectedMissingFolderError = $true
            }

            if ($isExpectedMissingFolderError) {
                continue
            }

            throw
        }

        foreach ($permission in $permissions) {
            [pscustomobject]@{
                MailboxDisplayName = $Mailbox.DisplayName
                MailboxAddress = $Mailbox.PrimarySmtpAddress
                Scope = 'Folder'
                FolderName = $name
                PermissionType = 'MailboxFolderPermission'
                PermissionDetails = ($permission.AccessRights -join ', ')
                GrantedTo = $permission.User.ToString()
            }
        }
    }
}
```

---

# OVERWRITE THIS FILE

## `M365.Exchange.Tools\Private\Get-M365UiSettings.ps1`

```powershell
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
```

---

# OVERWRITE THIS FILE

## `M365.Exchange.Tools\Private\Set-M365UiSettings.ps1`

```powershell
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
        [string]$ReportFontFamily,

        [Parameter()]
        [ValidateSet('Auto', 'Interactive', 'DisableWAM', 'Device')]
        [string]$ExchangeAuthMode
    )

    $moduleRoot = Split-Path -Path $PSScriptRoot -Parent
    $settingsPath = Join-Path -Path $moduleRoot -ChildPath 'Config\M365.Exchange.Tools.Settings.json'
    $settingsDirectory = Split-Path -Path $settingsPath -Parent

    if (-not (Test-Path -Path $settingsDirectory)) {
        New-Item -Path $settingsDirectory -ItemType Directory -Force | Out-Null
    }

    $currentSettings = Get-M365UiSettings -ForceRefresh

    $settings = @{
        BrowserPopout = if ($PSBoundParameters.ContainsKey('BrowserPopout')) { [string]$BrowserPopout } else { [string]$currentSettings.BrowserPopout }
        CompanyName = if ($PSBoundParameters.ContainsKey('CompanyName')) { [string]$CompanyName } else { [string]$currentSettings.CompanyName }
        LogoPath = if ($PSBoundParameters.ContainsKey('LogoPath')) { [string]$LogoPath } else { [string]$currentSettings.LogoPath }
        ReportSavePath = if ($PSBoundParameters.ContainsKey('ReportSavePath')) { [string]$ReportSavePath } else { [string]$currentSettings.ReportSavePath }
        FileNameTemplate = if ($PSBoundParameters.ContainsKey('FileNameTemplate')) { [string]$FileNameTemplate } else { [string]$currentSettings.FileNameTemplate }
        HtmlBrandingEnabled = if ($PSBoundParameters.ContainsKey('HtmlBrandingEnabled')) { [bool]$HtmlBrandingEnabled } else { [bool]$currentSettings.HtmlBrandingEnabled }
        HtmlShowCompanyName = if ($PSBoundParameters.ContainsKey('HtmlShowCompanyName')) { [bool]$HtmlShowCompanyName } else { [bool]$currentSettings.HtmlShowCompanyName }
        HtmlShowCompanyLogo = if ($PSBoundParameters.ContainsKey('HtmlShowCompanyLogo')) { [bool]$HtmlShowCompanyLogo } else { [bool]$currentSettings.HtmlShowCompanyLogo }
        ThemePrimaryColor = if ($PSBoundParameters.ContainsKey('ThemePrimaryColor')) { [string]$ThemePrimaryColor } else { [string]$currentSettings.ThemePrimaryColor }
        ThemeSecondaryColor = if ($PSBoundParameters.ContainsKey('ThemeSecondaryColor')) { [string]$ThemeSecondaryColor } else { [string]$currentSettings.ThemeSecondaryColor }
        ReportFontFamily = if ($PSBoundParameters.ContainsKey('ReportFontFamily')) { [string]$ReportFontFamily } else { [string]$currentSettings.ReportFontFamily }
        ExchangeAuthMode = if ($PSBoundParameters.ContainsKey('ExchangeAuthMode')) { [string]$ExchangeAuthMode } else { [string]$currentSettings.ExchangeAuthMode }
    }

    $settings |
        ConvertTo-Json -Depth 5 |
        Set-Content -Path $settingsPath -Encoding UTF8

    $script:M365UiSettingsCache = @{} + $settings
    $script:M365UiSettingsCachePath = $settingsPath

    return @{} + $settings
}
```

---

# OVERWRITE THIS FILE

## `M365.Exchange.Tools\Public\Connect-M365ExchangeTools.ps1`

```powershell
function Connect-M365ExchangeTools {
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$UserPrincipalName
    )

    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
        throw "Microsoft.Graph.Authentication is not installed. Run 'Install-Module Microsoft.Graph -Scope CurrentUser'."
    }

    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop | Out-Null

    $requiredScopes = @(
        'User.Read.All',
        'Group.Read.All',
        'Directory.Read.All',
        'Organization.Read.All',
        'Policy.Read.All',
        'OrgContact.Read.All',
        'MailboxSettings.Read',
        'AuditLog.Read.All'
    )

    if (Test-M365GraphConnection -RequiredScopes $requiredScopes -RequireAllScopes) {
        Write-Host 'Microsoft Graph connection already active.' -ForegroundColor Green
        return
    }

    $connectParams = @{
        Scopes = $requiredScopes
        NoWelcome = $true
    }

    if ($UserPrincipalName) {
        Write-Host 'UserPrincipalName is not used by Microsoft Graph interactive sign-in and will be ignored.' -ForegroundColor DarkYellow
    }

    Connect-MgGraph @connectParams -ErrorAction Stop

    $selectProfileCommand = Get-Command -Name Select-MgProfile -ErrorAction SilentlyContinue
    if ($selectProfileCommand) {
        Select-MgProfile -Name 'v1.0' | Out-Null
    }

    if (-not (Test-M365GraphConnection -RequiredScopes $requiredScopes -RequireAllScopes)) {
        throw 'Sign-in completed but a Microsoft Graph context with the required scopes was not detected. Try reconnecting, then run Check prerequisites to verify module and session state.'
    }

    Write-Host 'Connected to Microsoft Graph.' -ForegroundColor Green
}
```

---

# OVERWRITE THIS FILE

## `M365.Exchange.Tools\Public\Connect-M365ExchangePowerShell.ps1`

```powershell
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

        $exception = $ErrorRecord.Exception
        while ($null -ne $exception) {
            $parts.Add([string]$exception.Message)
            $parts.Add([string]$exception.GetType().FullName)
            if ($exception.StackTrace) {
                $parts.Add([string]$exception.StackTrace)
            }
            $exception = $exception.InnerException
        }

        $combined = $parts -join ' '
        return $combined -match 'RuntimeBroker|NullReferenceException.*Broker|Error Acquiring Token'
    }

    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        throw "ExchangeOnlineManagement is not installed. Run 'Install-Module ExchangeOnlineManagement -Scope CurrentUser'."
    }

    Import-Module ExchangeOnlineManagement -ErrorAction Stop | Out-Null

    if (Test-M365ExchangePowerShellConnection) {
        Write-Host 'Exchange Online PowerShell connection already active.' -ForegroundColor Green
        return
    }

    $uiSettings = Get-M365UiSettings
    $preferredAuthMode = [string]$uiSettings.ExchangeAuthMode

    $resolvedUserPrincipalName = $UserPrincipalName
    if ([string]::IsNullOrWhiteSpace($resolvedUserPrincipalName)) {
        $getMgContextCommand = Get-Command -Name Get-MgContext -ErrorAction SilentlyContinue
        if ($getMgContextCommand) {
            try {
                $graphContext = Get-MgContext -ErrorAction Stop
                if ($graphContext -and -not [string]::IsNullOrWhiteSpace([string]$graphContext.Account)) {
                    $resolvedUserPrincipalName = [string]$graphContext.Account
                    Write-Host "Using signed-in Microsoft Graph account: $resolvedUserPrincipalName" -ForegroundColor DarkCyan
                }
            }
            catch {
            }
        }
    }

    $baseConnectParams = @{
        ShowBanner = $false
        UseMultithreading = $true
    }

    if (-not [string]::IsNullOrWhiteSpace($resolvedUserPrincipalName)) {
        $baseConnectParams['UserPrincipalName'] = $resolvedUserPrincipalName
    }

    $connectCommand = Get-Command -Name Connect-ExchangeOnline -ErrorAction Stop
    $supportsDisableWAM = $connectCommand.Parameters.ContainsKey('DisableWAM')
    $supportsDevice = $connectCommand.Parameters.ContainsKey('Device')

    $attemptDefinitions = @{}

    $attemptDefinitions['UserPrincipalName'] = [pscustomobject]@{
        Name = 'UserPrincipalName'
        Label = 'signed-in account'
        Enabled = $baseConnectParams.ContainsKey('UserPrincipalName')
        Params = @{} + $baseConnectParams
    }

    $interactiveParams = @{} + $baseConnectParams
    [void]$interactiveParams.Remove('UserPrincipalName')
    $attemptDefinitions['Interactive'] = [pscustomobject]@{
        Name = 'Interactive'
        Label = 'default interactive sign-in'
        Enabled = $true
        Params = $interactiveParams
    }

    $disableWamParams = @{} + $interactiveParams
    $disableWamParams['DisableWAM'] = $true
    $attemptDefinitions['DisableWAM'] = [pscustomobject]@{
        Name = 'DisableWAM'
        Label = 'interactive sign-in with DisableWAM'
        Enabled = $supportsDisableWAM
        Params = $disableWamParams
    }

    $deviceParams = @{} + $interactiveParams
    $deviceParams['Device'] = $true
    $attemptDefinitions['Device'] = [pscustomobject]@{
        Name = 'Device'
        Label = 'device code sign-in'
        Enabled = $supportsDevice
        Params = $deviceParams
    }

    $attemptOrder = switch ($preferredAuthMode) {
        'Interactive' { @('Interactive', 'DisableWAM', 'Device') }
        'DisableWAM' { @('DisableWAM', 'Device', 'Interactive') }
        'Device' { @('Device', 'DisableWAM', 'Interactive') }
        default { @('UserPrincipalName', 'Interactive', 'DisableWAM', 'Device') }
    }

    $attempts = New-Object System.Collections.Generic.List[object]
    foreach ($attemptName in $attemptOrder) {
        if ($attemptDefinitions.ContainsKey($attemptName)) {
            $attempt = $attemptDefinitions[$attemptName]
            if ($attempt.Enabled) {
                [void]$attempts.Add($attempt)
            }
        }
    }

    $lastError = $null
    $brokerFailureSeen = $false
    $attemptNumber = 0

    foreach ($attempt in $attempts) {
        if ($brokerFailureSeen -and $attempt.Name -in @('UserPrincipalName', 'Interactive')) {
            continue
        }

        try {
            if ($attemptNumber -gt 0) {
                Write-Host "Connect-ExchangeOnline failed previously. Retrying with $($attempt.Label)." -ForegroundColor Yellow
            }

            Connect-ExchangeOnline @($attempt.Params) -ErrorAction Stop
            $lastError = $null
            break
        }
        catch {
            $lastError = $_
            $attemptNumber++

            if ($attempt.Name -in @('UserPrincipalName', 'Interactive') -and (Test-M365ExchangeWamBrokerError -ErrorRecord $_)) {
                $brokerFailureSeen = $true
                if ($supportsDisableWAM) {
                    Write-Host 'Detected Windows broker token acquisition failure. Retrying with DisableWAM before falling back to device code sign-in.' -ForegroundColor Yellow
                }
                else {
                    Write-Host 'Detected Windows broker token acquisition failure. Retrying with device code sign-in.' -ForegroundColor Yellow
                }
            }
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
```

---

# OVERWRITE THIS FILE

## `M365.Exchange.Tools\Public\Get-ADDNSHealth.ps1`

```powershell
function Get-ADDNSHealth {
    <#
    .SYNOPSIS
    Returns DNS zone health for AD-integrated DNS servers.
    Uses the DnsServer module when available; falls back to AD zone enumeration.

    .PARAMETER Server
    Optional. DNS/DC server to query. Defaults to current DC.

    .PARAMETER ExportPath
    Optional. Path to export results.

    .PARAMETER IncludeRecordCount
    Optional. When set, counts DNS records per zone. This can be slow in large environments.
    #>
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$Server,

        [Parameter()]
        [string]$ExportPath,

        [Parameter()]
        [switch]$IncludeRecordCount
    )

    if (-not (Get-Module ActiveDirectory -ErrorAction SilentlyContinue)) {
        if (Get-Module ActiveDirectory -ListAvailable -ErrorAction SilentlyContinue) {
            Import-Module ActiveDirectory -ErrorAction Stop
        }
        else {
            Write-Error 'ActiveDirectory module not found.'
            return @()
        }
    }

    $sp = if ($Server) { @{ Server = $Server } } else { @{} }
    $dnsAvail = [bool](Get-Module DnsServer -ListAvailable -ErrorAction SilentlyContinue)
    $dnsServer = if ($Server) { $Server } else { (Get-ADDomainController -Discover @sp -ErrorAction SilentlyContinue).HostName }

    $rows = [System.Collections.Generic.List[object]]::new()

    if ($dnsAvail) {
        Import-Module DnsServer -ErrorAction SilentlyContinue

        try {
            $dnsSettings = Get-DnsServer -ComputerName $dnsServer -ErrorAction Stop

            $fwdList = @($dnsSettings.ServerForwarder.IPAddress) | Where-Object { $_ }
            $fwdDisplay = if ($fwdList.Count -gt 0) { $fwdList -join ', ' } else { 'None configured' }
            $fwdStatus = if ($fwdList.Count -gt 0) { 'Info' } else { 'Warning' }

            $rows.Add([pscustomobject]@{
                ZoneName = '(Server)'
                ZoneType = 'Global Setting'
                DynamicUpdate = ''
                IsADIntegrated = ''
                ScavengingEnabled = ''
                RecordCount = ''
                Status = $fwdStatus
                Notes = "Forwarders: $fwdDisplay"
            })

            $rootHints = @($dnsSettings.ServerRootHint)
            $rows.Add([pscustomobject]@{
                ZoneName = '(Server)'
                ZoneType = 'Global Setting'
                DynamicUpdate = ''
                IsADIntegrated = ''
                ScavengingEnabled = ''
                RecordCount = ''
                Status = 'Info'
                Notes = "Root hints: $($rootHints.Count) configured"
            })
        }
        catch {
        }

        try {
            $zones = @(Get-DnsServerZone -ComputerName $dnsServer -ErrorAction Stop)
            $i = 0

            foreach ($z in $zones) {
                $i++
                Write-Progress -Activity 'Collecting DNS zone data' -Status $z.ZoneName -PercentComplete (($i / [math]::Max($zones.Count, 1)) * 100)

                if ($z.ZoneType -eq 'Cache') {
                    continue
                }

                $scavEnabled = $false
                try {
                    $aging = Get-DnsServerZoneAging -ZoneName $z.ZoneName -ComputerName $dnsServer -ErrorAction SilentlyContinue
                    $scavEnabled = [bool]$aging.AgingEnabled
                }
                catch {
                }

                $recordCount = ''
                if ($IncludeRecordCount) {
                    $recordCount = try {
                        @(Get-DnsServerResourceRecord -ZoneName $z.ZoneName -ComputerName $dnsServer -ErrorAction SilentlyContinue).Count
                    }
                    catch {
                        '?'
                    }
                }

                $isADIntegrated = $null
                if ($z.PSObject.Properties.Name -contains 'IsDsIntegrated') {
                    $isADIntegrated = [bool]$z.IsDsIntegrated
                }
                elseif ($z.PSObject.Properties.Name -contains 'IsDirectoryIntegrated') {
                    $isADIntegrated = [bool]$z.IsDirectoryIntegrated
                }

                $issues = [System.Collections.Generic.List[string]]::new()

                if (-not $scavEnabled -and $z.ZoneType -eq 'Primary') {
                    $issues.Add('Scavenging disabled — stale DNS records accumulate over time')
                }

                if ($z.DynamicUpdate -eq 'None' -and $z.ZoneType -eq 'Primary' -and $z.ZoneName -notlike '*.arpa') {
                    $issues.Add('Dynamic update disabled — clients cannot auto-register')
                }

                $status = if ($issues.Count -ge 1) { 'Warning' } else { 'OK' }

                $rows.Add([pscustomobject]@{
                    ZoneName = $z.ZoneName
                    ZoneType = [string]$z.ZoneType
                    DynamicUpdate = [string]$z.DynamicUpdate
                    IsADIntegrated = $isADIntegrated
                    ScavengingEnabled = $scavEnabled
                    RecordCount = $recordCount
                    Status = $status
                    Notes = ($issues -join '; ')
                })
            }

            Write-Progress -Activity 'Collecting DNS zone data' -Completed
        }
        catch {
            $rows.Add([pscustomobject]@{
                ZoneName = 'Error'
                ZoneType = ''
                DynamicUpdate = ''
                IsADIntegrated = ''
                ScavengingEnabled = ''
                RecordCount = ''
                Status = 'Warning'
                Notes = "DnsServer zone query failed: $($_.Exception.Message)"
            })
        }
    }
    else {
        Write-Host 'DnsServer module not available. Falling back to AD zone enumeration.' -ForegroundColor Yellow

        try {
            $dom = Get-ADDomain @sp
            $domainDN = $dom.DistinguishedName
            $forestRootDN = ($dom.Forest -split '\.' | ForEach-Object { "DC=$_" }) -join ','
            $searchBases = @(
                "CN=MicrosoftDNS,DC=DomainDnsZones,$domainDN",
                "CN=MicrosoftDNS,DC=ForestDnsZones,$forestRootDN",
                "CN=MicrosoftDNS,CN=System,$domainDN"
            )

            foreach ($base in $searchBases) {
                try {
                    $zoneObjects = @(Get-ADObject -SearchBase $base -Filter { objectClass -eq 'dnsZone' } @sp -ErrorAction SilentlyContinue)
                    foreach ($z in $zoneObjects) {
                        $zName = $z.Name
                        if ($zName -in @('RootDNSServers', '..TrustAnchors')) {
                            continue
                        }

                        $partition = if ($base -like '*DomainDnsZones*') {
                            'DomainDnsZones'
                        }
                        elseif ($base -like '*ForestDnsZones*') {
                            'ForestDnsZones'
                        }
                        else {
                            'Legacy/System'
                        }

                        $rows.Add([pscustomobject]@{
                            ZoneName = $zName
                            ZoneType = "AD-Integrated ($partition)"
                            DynamicUpdate = 'Unknown (DnsServer module required)'
                            IsADIntegrated = $true
                            ScavengingEnabled = 'Unknown'
                            RecordCount = if ($IncludeRecordCount) { '?' } else { '' }
                            Status = 'Info'
                            Notes = 'Install DnsServer RSAT for full DNS health data.'
                        })
                    }
                }
                catch {
                }
            }
        }
        catch {
            $rows.Add([pscustomobject]@{
                ZoneName = 'Error'
                ZoneType = ''
                DynamicUpdate = ''
                IsADIntegrated = ''
                ScavengingEnabled = ''
                RecordCount = ''
                Status = 'Warning'
                Notes = "AD zone enumeration failed: $($_.Exception.Message)"
            })
        }
    }

    $result = @($rows | Sort-Object { switch ($_.Status) { 'Critical' { 0 } 'Warning' { 1 } default { 2 } } }, ZoneName)

    if (-not [string]::IsNullOrWhiteSpace($ExportPath)) {
        Export-M365ReportData -InputObject $result -ExportPath $ExportPath | Out-Null
    }

    return $result
}
```

---

# OPTIONAL SINGLE-LINE CLARITY UPDATE

If you choose to update `M365.Exchange.Tools\Private\Get-M365TenantCapabilities.ps1`, only change this line:

```powershell
$isGraphConnected = Test-M365GraphConnection
```

Do not change anything else in that file.

---

# VALIDATION CHECKLIST

After applying the changes, verify:

- PowerShell syntax loads cleanly.
- `Connect-M365ExchangeTools` still validates Graph scopes.
- `Connect-M365ExchangePowerShell` now tries `DisableWAM` before `Device` after broker/WAM errors.
- `Get-M365GraphCollection` no longer uses array `+=` in the pagination loop.
- `Get-M365MailboxFolderDelegationEntries` still skips expected missing-folder cases.
- `Get-M365UiSettings` caches settings.
- `Set-M365UiSettings` refreshes the cache after save.
- `Get-ADDNSHealth` no longer infers AD integration from `IsAutoCreated`.
- All checkpoint API artifacts are deleted.
