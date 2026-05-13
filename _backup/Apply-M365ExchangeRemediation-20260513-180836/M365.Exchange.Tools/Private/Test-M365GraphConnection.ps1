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
    } else {
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
}
