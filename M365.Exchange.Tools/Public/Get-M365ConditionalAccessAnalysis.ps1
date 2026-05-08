function Get-M365ConditionalAccessAnalysis {
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$ExportPath
    )

    Assert-ExchangeOnlineConnected

    Write-Host 'Fetching Conditional Access policies for analysis...' -ForegroundColor Cyan
    $policies = @(Get-M365GraphCollection -Uri '/v1.0/policies/conditionalAccessPolicies')

    if ($policies.Count -eq 0) {
        Write-Host 'No Conditional Access policies found. Ensure Policy.Read.All scope is granted.' -ForegroundColor Yellow
    }

    # ── Helper functions ─────────────────────────────────────────────

    function Get-GrantControls ($p) {
        if ($p.grantControls -and $p.grantControls.builtInControls) { return @($p.grantControls.builtInControls) }
        return @()
    }

    function Test-MfaGrant          ($p) { (Get-GrantControls $p) -contains 'mfa' }
    function Test-BlockGrant         ($p) { (Get-GrantControls $p) -contains 'block' }
    function Test-CompliantDevice    ($p) { $c = Get-GrantControls $p; $c -contains 'compliantDevice' -or $c -contains 'domainJoinedDevice' }
    function Test-AllUsers           ($p) { @($p.conditions.users.includeUsers) -contains 'All' }
    function Test-AllApps            ($p) { @($p.conditions.applications.includeApplications) -contains 'All' }

    function Test-Office365 ($p) {
        $apps = @($p.conditions.applications.includeApplications)
        $apps -contains 'All' -or $apps -contains 'Office365'
    }

    # Well-known admin role GUIDs
    $adminRoleIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    @(
        '62e90394-69f5-4237-9190-012177145e10', # Global Administrator
        '29232cdf-9323-42fd-ade2-1d097af3e4de', # Exchange Administrator
        'f28a1f50-f6e7-4571-818b-6a12f2af6b6c', # SharePoint Administrator
        '69091246-20e8-4a56-aa4d-066075b2a7a8', # Teams Administrator
        '194ae4cb-b126-40b2-bd5b-6091b380977d', # Security Administrator
        'b1be1c3e-b65d-4f19-8427-f6fa0d97feb9', # Conditional Access Administrator
        'c4e39bd9-1100-46d3-8c65-fb160da0071f', # Authentication Administrator
        '9b895d92-2cd3-44c7-9d02-a6ac2d5ea5c3', # Application Administrator
        '158c047a-c907-4556-b7ef-446551a6b5f7', # Cloud Application Administrator
        '7be44c8a-adaf-4e2a-84d6-ab2649e08a13', # Privileged Authentication Administrator
        'e8611ab8-c189-46e8-94e1-60213ab1f814', # Privileged Role Administrator
        'aaf43236-0c0d-4d5f-883a-6955382ac081'  # Security Operator
    ) | ForEach-Object { [void]$adminRoleIds.Add($_) }

    function Test-HasAdminRoles ($p) {
        foreach ($r in @($p.conditions.users.includeRoles)) {
            if ($adminRoleIds.Contains([string]$r)) { return $true }
        }
        return $false
    }

    function Test-GuestsOrExternal ($p) {
        if (@($p.conditions.users.includeUsers) -contains 'GuestsOrExternalUsers') { return $true }
        # Newer Graph property (Entra External Identities)
        if ($p.conditions.users.includeGuestsOrExternalUsers) { return $true }
        return $false
    }

    function Test-LegacyAuth ($p) {
        $types = @($p.conditions.clientAppTypes)
        ($types -contains 'exchangeActiveSync') -or ($types -contains 'other')
    }

    $azureMgmtAppId = '797f4846-ba00-4fd7-ba43-dac1f8f63013'
    function Test-AzureMgmt ($p) {
        $apps = @($p.conditions.applications.includeApplications)
        $apps -contains $azureMgmtAppId
    }

    $m365WorkloadIds = @(
        '00000002-0000-0ff1-ce00-000000000000', # Exchange Online
        '00000003-0000-0ff1-ce00-000000000000', # SharePoint Online
        'cc15fd57-2c6c-4117-a88c-83b1d56b4bbe'  # Microsoft Teams
    )
    function Test-M365Workloads ($p) {
        $apps = @($p.conditions.applications.includeApplications)
        if ($apps -contains 'All' -or $apps -contains 'Office365') { return $true }
        $hits = ($apps | Where-Object { $m365WorkloadIds -contains $_ }).Count
        return $hits -ge 2
    }

    function Test-SignInRisk ($p, $levels) {
        $risk = @($p.conditions.signInRiskLevels)
        foreach ($l in $levels) { if ($risk -contains $l) { return $true } }
        return $false
    }

    function Test-SessionControls ($p) {
        if (-not $p.sessionControls) { return $false }
        $sc = $p.sessionControls
        ($sc.signInFrequency -and $sc.signInFrequency.isEnabled) -or
        ($sc.persistentBrowser -and $sc.persistentBrowser.isEnabled)
    }

    function Get-PolicyStateSummary ($matched) {
        if (-not $matched -or $matched.Count -eq 0) { return '—' }
        $states = @($matched | ForEach-Object { [string]$_.state }) | Sort-Object -Unique
        if ($states -contains 'enabled')                    { return 'Enabled' }
        if ($states -contains 'enabledForReportingOnly')    { return 'Report Only' }
        return 'Disabled'
    }

    function New-AnalysisRow ($num, $name, $matched, $risk, $guidance) {
        $present = $matched -and $matched.Count -gt 0
        $state   = Get-PolicyStateSummary $matched
        $names   = if ($present) { ($matched | ForEach-Object { [string]$_.displayName }) -join '; ' } else { '' }
        [pscustomobject]@{
            '#'                = $num
            PolicyCheck        = $name
            Status             = if ($present) { 'Present' } else { 'Missing' }
            PolicyState        = $state
            MatchedPolicies    = $names
            Risk               = $risk
            Guidance           = $guidance
        }
    }

    # ── 10 checks ────────────────────────────────────────────────────

    $rows = [System.Collections.Generic.List[object]]::new()

    # 1. MFA for all users
    $m = @($policies | Where-Object { (Test-AllUsers $_) -and (Test-AllApps $_) -and (Test-MfaGrant $_) })
    $rows.Add((New-AnalysisRow 1 'Require MFA for all users' $m `
        'Any stolen credential grants full account access. This is the single highest-impact control.' `
        'Applies to: All users / All cloud apps / Grant: Require MFA. Exclude only break-glass accounts.'))

    # 2. MFA for administrative roles
    $m = @($policies | Where-Object { (Test-HasAdminRoles $_) -and (Test-MfaGrant $_) })
    $rows.Add((New-AnalysisRow 2 'Require MFA for all administrative roles' $m `
        'Privileged accounts without enforced MFA are the primary target for account takeover and tenant compromise.' `
        'Applies to: Directory roles (Global Admin, Exchange Admin, etc.) / Grant: Require MFA. Prefer phishing-resistant MFA.'))

    # 3. Block legacy authentication
    $m = @($policies | Where-Object { (Test-LegacyAuth $_) -and (Test-BlockGrant $_) })
    $rows.Add((New-AnalysisRow 3 'Block legacy authentication protocols' $m `
        'Legacy auth bypasses MFA entirely. Password spray attacks almost exclusively target legacy auth endpoints.' `
        'Applies to: All users / Condition: Client apps = Exchange ActiveSync + Other (legacy) / Grant: Block access.'))

    # 4. MFA for guest and external users
    $m = @($policies | Where-Object { (Test-GuestsOrExternal $_) -and (Test-MfaGrant $_) })
    $rows.Add((New-AnalysisRow 4 'Require MFA for guest and external users' $m `
        'Guest accounts are outside the organization identity boundary and cannot be assumed safe.' `
        'Applies to: Guest or external users / All cloud apps / Grant: Require MFA.'))

    # 5. Require compliant or hybrid-joined device
    $m = @($policies | Where-Object { Test-CompliantDevice $_ })
    $rows.Add((New-AnalysisRow 5 'Require compliant or hybrid Azure AD joined device' $m `
        'Unmanaged personal devices accessing company data are invisible to IT — malware, key-loggers, unpatched OS.' `
        'Applies to: Member users / Office 365 apps / Grant: Compliant device OR Hybrid Azure AD joined.'))

    # 6. Block high-risk sign-ins
    $m = @($policies | Where-Object { (Test-SignInRisk $_ @('high')) -and (Test-BlockGrant $_) })
    $rows.Add((New-AnalysisRow 6 'Block high-risk sign-ins (Identity Protection)' $m `
        'High-risk sign-ins flagged by Entra Identity Protection proceed without intervention if no policy blocks them.' `
        'Requires Entra ID P2. Applies to: All users / Sign-in risk = High / Grant: Block access.'))

    # 7. MFA for medium-risk sign-ins
    $m = @($policies | Where-Object { (Test-SignInRisk $_ @('medium')) -and (Test-MfaGrant $_) })
    $rows.Add((New-AnalysisRow 7 'Require MFA for medium-risk sign-ins (Identity Protection)' $m `
        'Medium-risk sign-ins are suspicious but not definitively malicious — MFA step-up closes the gap.' `
        'Requires Entra ID P2. Applies to: All users / Sign-in risk = Medium / Grant: Require MFA.'))

    # 8. MFA for Azure management
    $m = @($policies | Where-Object { (Test-AzureMgmt $_) -and (Test-MfaGrant $_) })
    if ($m.Count -eq 0) {
        # Broad admin + all-apps MFA also satisfies this
        $m = @($policies | Where-Object { (Test-AllApps $_) -and (Test-MfaGrant $_) -and (Test-HasAdminRoles $_) })
    }
    $rows.Add((New-AnalysisRow 8 'Require MFA for Azure management access' $m `
        'Unrestricted Azure portal access without MFA for admins exposes subscriptions, VMs, and storage.' `
        'Applies to: Admin roles / App: Microsoft Azure Management (797f4846...) or All / Grant: Require MFA.'))

    # 9. MFA for core M365 workloads
    $m = @($policies | Where-Object { (Test-M365Workloads $_) -and (Test-MfaGrant $_) -and (Test-AllUsers $_) })
    if ($m.Count -eq 0) {
        # All-users + All-apps + MFA satisfies this
        $m = @($policies | Where-Object { (Test-AllApps $_) -and (Test-MfaGrant $_) -and (Test-AllUsers $_) })
    }
    $rows.Add((New-AnalysisRow 9 'Require MFA for core M365 workloads (Exchange / SharePoint / Teams)' $m `
        'Core collaboration workloads without MFA protection expose mail, files, and chat to credential attacks.' `
        'Applies to: All users / Apps: Exchange Online, SharePoint Online, Microsoft Teams (or Office365 / All) / Grant: Require MFA.'))

    # 10. Session controls for browser access
    $m = @($policies | Where-Object { (Test-SessionControls $_) -and (Test-Office365 $_) -and (Test-AllUsers $_) })
    if ($m.Count -eq 0) {
        # Any policy with session controls at all counts
        $m = @($policies | Where-Object { Test-SessionControls $_ })
    }
    $rows.Add((New-AnalysisRow 10 'Enforce session controls for browser access' $m `
        'Without sign-in frequency limits, a stolen browser session token grants persistent access indefinitely.' `
        'Applies to: All users / Office 365 / Session: Sign-in frequency enforced + Persistent browser = Never.'))

    # ── Summary ───────────────────────────────────────────────────────
    $missingCount  = @($rows | Where-Object { $_.Status -eq 'Missing' }).Count
    $disabledCount = @($rows | Where-Object { $_.Status -eq 'Present' -and $_.PolicyState -ne 'Enabled' }).Count
    $summaryColor = if ($missingCount -gt 0) { 'Yellow' } else { 'Green' }
    Write-Host "Analysis complete: $($rows.Count) checks | Missing: $missingCount | Present but not enabled: $disabledCount" -ForegroundColor $summaryColor

    $result = @($rows)

    if (-not [string]::IsNullOrWhiteSpace($ExportPath)) {
        Export-M365ReportData -InputObject $result -ExportPath $ExportPath | Out-Null
    }

    return $result
}
