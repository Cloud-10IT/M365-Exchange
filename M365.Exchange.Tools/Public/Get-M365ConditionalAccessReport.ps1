function Get-M365ConditionalAccessReport {
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$ExportPath
    )

    Assert-ExchangeOnlineConnected

    Write-Host 'Fetching Conditional Access policies...' -ForegroundColor Cyan
    $policies = @(Get-M365GraphCollection -Uri '/v1.0/policies/conditionalAccessPolicies')

    if ($policies.Count -eq 0) {
        Write-Host 'No Conditional Access policies found. Ensure the tenant has an Entra ID P1 or P2 license and the account has Policy.Read.All scope.' -ForegroundColor Yellow
        return @()
    }

    $rows = @(
        $policies | ForEach-Object {
            $p = $_

            # Users / groups / roles
            $includeUsers  = if ($p.conditions.users.includeUsers)  { $p.conditions.users.includeUsers  -join '; ' } else { '' }
            $excludeUsers  = if ($p.conditions.users.excludeUsers)  { $p.conditions.users.excludeUsers  -join '; ' } else { '' }
            $includeGroups = if ($p.conditions.users.includeGroups) { $p.conditions.users.includeGroups -join '; ' } else { '' }
            $excludeGroups = if ($p.conditions.users.excludeGroups) { $p.conditions.users.excludeGroups -join '; ' } else { '' }
            $includeRoles  = if ($p.conditions.users.includeRoles)  { $p.conditions.users.includeRoles  -join '; ' } else { '' }
            $excludeRoles  = if ($p.conditions.users.excludeRoles)  { $p.conditions.users.excludeRoles  -join '; ' } else { '' }

            # Applications
            $includeApps = if ($p.conditions.applications.includeApplications) { $p.conditions.applications.includeApplications -join '; ' } else { '' }
            $excludeApps = if ($p.conditions.applications.excludeApplications) { $p.conditions.applications.excludeApplications -join '; ' } else { '' }

            # Platforms
            $includePlatforms = if ($p.conditions.platforms) { $p.conditions.platforms.includePlatforms -join '; ' } else { '' }
            $excludePlatforms = if ($p.conditions.platforms) { $p.conditions.platforms.excludePlatforms -join '; ' } else { '' }

            # Risk + client types
            $signInRisk     = if ($p.conditions.signInRiskLevels) { $p.conditions.signInRiskLevels -join '; ' } else { '' }
            $userRisk       = if ($p.conditions.userRiskLevels)   { $p.conditions.userRiskLevels   -join '; ' } else { '' }
            $clientAppTypes = if ($p.conditions.clientAppTypes)   { $p.conditions.clientAppTypes   -join '; ' } else { '' }

            # Grant controls
            $grantSummary = ''
            if ($p.grantControls) {
                $controls = if ($p.grantControls.builtInControls) { $p.grantControls.builtInControls -join '; ' } else { '' }
                $grantSummary = if ($controls) { "$($p.grantControls.operator): $controls" } else { [string]$p.grantControls.operator }
            }

            # Session controls
            $sessionParts = [System.Collections.Generic.List[string]]::new()
            if ($p.sessionControls) {
                if ($p.sessionControls.signInFrequency -and $p.sessionControls.signInFrequency.isEnabled) {
                    $sessionParts.Add("SignInFreq: $($p.sessionControls.signInFrequency.value) $($p.sessionControls.signInFrequency.type)")
                }
                if ($p.sessionControls.persistentBrowser -and $p.sessionControls.persistentBrowser.isEnabled) {
                    $sessionParts.Add("PersistentBrowser: $($p.sessionControls.persistentBrowser.mode)")
                }
                if ($p.sessionControls.cloudAppSecurity -and $p.sessionControls.cloudAppSecurity.isEnabled) {
                    $sessionParts.Add("MCAS: $($p.sessionControls.cloudAppSecurity.cloudAppSecurityType)")
                }
                if ($p.sessionControls.applicationEnforcedRestrictions -and $p.sessionControls.applicationEnforcedRestrictions.isEnabled) {
                    $sessionParts.Add('AppRestrictions')
                }
                if ($p.sessionControls.continuousAccessEvaluation -and $p.sessionControls.continuousAccessEvaluation.mode -ne 'disabled') {
                    $sessionParts.Add("CAE: $($p.sessionControls.continuousAccessEvaluation.mode)")
                }
            }
            $sessionSummary = $sessionParts -join '; '

            [pscustomobject]@{
                PolicyName       = [string]$p.displayName
                State            = [string]$p.state
                IncludeUsers     = $includeUsers
                ExcludeUsers     = $excludeUsers
                IncludeGroups    = $includeGroups
                ExcludeGroups    = $excludeGroups
                IncludeRoles     = $includeRoles
                ExcludeRoles     = $excludeRoles
                IncludeApps      = $includeApps
                ExcludeApps      = $excludeApps
                IncludePlatforms = $includePlatforms
                ExcludePlatforms = $excludePlatforms
                ClientAppTypes   = $clientAppTypes
                SignInRiskLevels = $signInRisk
                UserRiskLevels   = $userRisk
                GrantControls    = $grantSummary
                SessionControls  = $sessionSummary
                PolicyId         = [string]$p.id
            }
        }
    )

    if (-not [string]::IsNullOrWhiteSpace($ExportPath)) {
        Export-M365ReportData -InputObject $rows -ExportPath $ExportPath | Out-Null
    }

    return $rows
}
