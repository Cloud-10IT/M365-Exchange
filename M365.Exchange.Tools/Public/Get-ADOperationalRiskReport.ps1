function Get-ADOperationalRiskReport {
    <#
    .SYNOPSIS
        Identifies operational risk indicators across AD DS: SYSVOL replication, GPO health,
        schema/functional level, stale infrastructure, and site topology gaps.
        Read-only — no environment changes are made.
    .PARAMETER Server
        Optional. Target domain controller FQDN or domain name.
    .PARAMETER ExportPath
        Optional. Path to export results.
    #>
    [CmdletBinding()]
    param(
        [Parameter()][string]$Server,
        [Parameter()][string]$ExportPath
    )

    if (-not (Get-Module ActiveDirectory -ErrorAction SilentlyContinue)) {
        if (Get-Module ActiveDirectory -ListAvailable -ErrorAction SilentlyContinue) {
            Import-Module ActiveDirectory -ErrorAction Stop
        } else {
            Write-Error 'ActiveDirectory module not found.'
            return @()
        }
    }

    $sp       = if ($Server) { @{ Server = $Server } } else { @{} }
    $gpAvail  = [bool](Get-Module GroupPolicy -ListAvailable -ErrorAction SilentlyContinue)
    $rows     = [System.Collections.Generic.List[object]]::new()
    $num      = 0

    function Add-Risk ($area, $finding, $severity, $details, $recommendation) {
        $script:num++
        $rows.Add([pscustomobject]@{
            '#'            = $script:num
            RiskArea       = $area
            Finding        = $finding
            Severity       = $severity
            Details        = $details
            Recommendation = $recommendation
        })
    }

    # ── SYSVOL Replication mode (DFSR vs FRS) ─────────────────────────
    try {
        $domDN   = (Get-ADDomain @sp).DistinguishedName
        $dcsOU   = "OU=Domain Controllers,$domDN"
        $dfsrSubs = @(Get-ADObject -SearchBase $dcsOU -Filter { objectClass -eq 'msDFSR-Subscription' } @sp -ErrorAction SilentlyContinue)
        $usesDFSR = $dfsrSubs.Count -gt 0

        if ($usesDFSR) {
            Add-Risk 'SYSVOL Replication' 'SYSVOL uses DFS-R (modern)' 'Info' `
                "DFS-R is the current supported SYSVOL replication engine. $($dfsrSubs.Count) subscription object(s) found." `
                'No action required. Ensure DFS-R service (DFSR) is running on all DCs.'
        } else {
            Add-Risk 'SYSVOL Replication' 'SYSVOL may be using legacy FRS replication' 'High' `
                'No msDFSR-Subscription objects found under Domain Controllers OU. SYSVOL may still use File Replication Service (FRS), which was deprecated in Windows Server 2008 R2 and is unsupported on Server 2016+.' `
                'Migrate SYSVOL from FRS to DFSR using the dfsrmig.exe tool. This is a prerequisite for raising domain functional level to 2008 R2 or higher.'
        }
    }
    catch { Add-Risk 'SYSVOL Replication' 'SYSVOL replication mode check failed' 'Medium' $_.Exception.Message 'Manually verify SYSVOL replication mode using: dfsrmig /getglobalstate' }

    # ── Domain/Forest Functional Level ────────────────────────────────
    try {
        $dom    = Get-ADDomain @sp
        $forest = Get-ADForest @sp
        $dfl    = [string]$dom.DomainMode
        $ffl    = [string]$forest.ForestMode

        $dflScore = switch -Wildcard ($dfl) {
            '*2016*' { 4 }; '*2012R2*' { 3 }; '*2012*' { 2 }; '*2008R2*' { 1 }; '*2008*' { 0 }; default { 0 }
        }
        if ($dflScore -lt 3) {
            Add-Risk 'Functional Level' 'Domain Functional Level below Windows Server 2016' 'Medium' `
                "Current DFL: $dfl. Features requiring newer DFL include Privileged Access Management (PAM), Kerberos armoring improvements, and full Protected Users support." `
                'Plan a DFL/FFL raise after ensuring all DCs run Windows Server 2016 or later. Back up all DCs and the system state before raising the level.'
        } else {
            Add-Risk 'Functional Level' 'Domain Functional Level is current' 'Info' `
                "DFL: $dfl | FFL: $ffl. All modern AD DS features are available." ''
        }
    }
    catch { Add-Risk 'Functional Level' 'Functional level check failed' 'Medium' $_.Exception.Message '' }

    # ── DCs on end-of-support OS ───────────────────────────────────────
    try {
        $eosMap = @{ '2003' = [datetime]'2015-07-14'; '2008' = [datetime]'2020-01-14'; '2012' = [datetime]'2023-10-10' }
        $dcs    = @(Get-ADDomainController -Filter * @sp)
        $eolDCs = $dcs | Where-Object {
            $os = [string]$_.OperatingSystem
            $eosMap.Keys | Where-Object { $os -like "*$_*" -and $eosMap[$_] -lt (Get-Date) }
        }
        if (@($eolDCs).Count -gt 0) {
            $names = (@($eolDCs) | ForEach-Object { "$($_.HostName) ($($_.OperatingSystem))" }) -join '; '
            Add-Risk 'Domain Controllers' 'Domain controllers running end-of-life OS' 'Critical' `
                "$(@($eolDCs).Count) DC(s) on unsupported OS: $names" `
                'Upgrade or decommission EOL DCs immediately. EOL OS receives no security patches — DCs are the highest-value target in any Windows environment and must run supported OS versions.'
        } else {
            Add-Risk 'Domain Controllers' 'All DCs running supported OS' 'Info' `
                "$($dcs.Count) DC(s) — none on end-of-life OS." ''
        }
    }
    catch { Add-Risk 'Domain Controllers' 'DC OS check failed' 'Medium' $_.Exception.Message '' }

    # ── Single point of failure (only one DC) ─────────────────────────
    try {
        $dcs = @(Get-ADDomainController -Filter * @sp)
        if ($dcs.Count -eq 1) {
            Add-Risk 'Domain Controllers' 'Single domain controller in domain' 'High' `
                "Only 1 DC found ($($dcs[0].HostName)). If this DC is unavailable, no authentication, GPO processing, or AD queries will function." `
                'Deploy a second DC (physical or VM) in a separate failure domain. For small environments, a Server Core DC on an independent hypervisor host is acceptable.'
        } elseif ($dcs.Count -eq 2) {
            Add-Risk 'Domain Controllers' 'Only two domain controllers' 'Medium' `
                "$($dcs.Count) DCs detected. Two DCs provide basic redundancy but patching or outage of one DC temporarily reduces fault tolerance to zero." `
                'Consider a third DC for environments requiring continuous authentication availability during maintenance windows.'
        } else {
            Add-Risk 'Domain Controllers' 'Domain controller count' 'Info' `
                "$($dcs.Count) DCs provide fault tolerance. Verify at least one GC and PDC Emulator are in primary sites." ''
        }
    }
    catch { Add-Risk 'Domain Controllers' 'DC count check failed' 'Medium' $_.Exception.Message '' }

    # ── AD Recycle Bin ─────────────────────────────────────────────────
    try {
        $rb = Get-ADOptionalFeature -Filter { Name -eq 'Recycle Bin Feature' } @sp -ErrorAction SilentlyContinue
        $en = $rb -and @($rb.EnabledScopes).Count -gt 0
        if (-not $en) {
            Add-Risk 'Recovery' 'AD Recycle Bin is disabled' 'High' `
                'Accidentally deleted AD objects (users, computers, groups) cannot be recovered without an authoritative restore, requiring DC downtime and replication reconfiguration.' `
                'Enable the AD Recycle Bin: Enable-ADOptionalFeature -Identity "Recycle Bin Feature" -Scope ForestOrConfigurationSet -Target <forest-FQDN>. Requires DFL >= Windows Server 2008 R2. This is a one-way change.'
        } else {
            Add-Risk 'Recovery' 'AD Recycle Bin is enabled' 'Info' 'Deleted objects are recoverable for tombstone lifetime period.' ''
        }
    }
    catch { Add-Risk 'Recovery' 'Recycle Bin check failed' 'Medium' $_.Exception.Message '' }

    # ── Tombstone lifetime ─────────────────────────────────────────────
    try {
        $cn = (Get-ADRootDSE @sp).configurationNamingContext
        $ts = (Get-ADObject "CN=Directory Service,CN=Windows NT,CN=Services,$cn" -Properties tombstoneLifetime @sp).tombstoneLifetime
        if (-not $ts -or $ts -eq 0) { $ts = 60 }
        $ts = [int]$ts
        if ($ts -lt 90) {
            Add-Risk 'Recovery' "Tombstone lifetime too short ($ts days)" 'High' `
                "Tombstone lifetime of $ts days. Any DC offline longer than $ts days cannot rejoin replication — it must be rebuilt from scratch or restored. Backup windows, disaster recovery scenarios, and long-term outages are all affected." `
                'Increase tombstone lifetime to 180 days: Set-ADObject "CN=Directory Service,..." -Replace @{tombstoneLifetime=180}. This also allows the Recycle Bin to retain deleted objects for 180 days.'
        } elseif ($ts -lt 180) {
            Add-Risk 'Recovery' "Tombstone lifetime below recommended ($ts days)" 'Medium' `
                "Tombstone lifetime of $ts days is functional but below Microsoft's recommended 180 days." `
                'Consider increasing to 180 days for improved recovery flexibility. Verify backup schedules align with the current tombstone lifetime.'
        } else {
            Add-Risk 'Recovery' 'Tombstone lifetime is adequate' 'Info' "$ts days — within Microsoft recommendation." ''
        }
    }
    catch { Add-Risk 'Recovery' 'Tombstone lifetime check failed' 'Medium' $_.Exception.Message '' }

    # ── GPO Health (requires GroupPolicy module) ───────────────────────
    if ($gpAvail) {
        Import-Module GroupPolicy -ErrorAction SilentlyContinue
        try {
            $domain   = Get-ADDomain @sp
            $allGPOs  = @(Get-GPO -All -Domain $domain.DNSRoot -ErrorAction Stop)

            # Unlinked GPOs
            $unlinked = @($allGPOs | Where-Object {
                try {
                    $report = [xml](Get-GPOReport -Guid $_.Id -ReportType Xml -Domain $domain.DNSRoot -ErrorAction SilentlyContinue)
                    $links  = @($report.GPO.LinksTo)
                    $links.Count -eq 0
                } catch { $false }
            })
            if ($unlinked.Count -gt 0) {
                $names = ($unlinked | Select-Object -First 10 | ForEach-Object { $_.DisplayName }) -join ', '
                Add-Risk 'Group Policy' "Unlinked GPOs ($($unlinked.Count))" 'Low' `
                    "GPOs not linked to any OU, domain, or site: $names" `
                    'Review and remove unlinked GPOs. Unlinked GPOs still create AD objects and may confuse administrators reviewing policy scope.'
            } else {
                Add-Risk 'Group Policy' 'No unlinked GPOs found' 'Info' 'All GPOs are linked to at least one scope of management.' ''
            }

            # GPOs with version mismatch (AD version != SYSVOL version)
            $mismatch = @($allGPOs | Where-Object {
                $_.Computer.DSVersion    -ne $_.Computer.SysvolVersion -or
                $_.User.DSVersion        -ne $_.User.SysvolVersion
            })
            if ($mismatch.Count -gt 0) {
                $names = ($mismatch | Select-Object -First 10 | ForEach-Object { $_.DisplayName }) -join ', '
                Add-Risk 'Group Policy' "GPO version mismatch — AD vs SYSVOL ($($mismatch.Count))" 'High' `
                    "GPOs where AD version differs from SYSVOL version: $names" `
                    'Version mismatches indicate SYSVOL replication problems. DCs with stale SYSVOL copies will apply outdated policy. Run: repadmin /showrepl and check DFS-R health.'
            } else {
                Add-Risk 'Group Policy' 'No GPO version mismatches' 'Info' 'AD and SYSVOL GPO versions are consistent.' ''
            }
        }
        catch {
            Add-Risk 'Group Policy' 'GPO health check failed' 'Medium' $_.Exception.Message 'Ensure GroupPolicy module is available and the account has GPO read permissions.'
        }
    } else {
        Add-Risk 'Group Policy' 'GroupPolicy module not available — GPO checks skipped' 'Info' `
            'Install RSAT: Group Policy Management Tools to enable GPO health checks.' ''
    }

    # ── Schema version warning ─────────────────────────────────────────
    try {
        $sv = try { (Get-ADObject (Get-ADRootDSE @sp).schemaNamingContext -Properties objectVersion @sp).objectVersion } catch { 0 }
        if ([int]$sv -lt 87) {
            Add-Risk 'Schema' "AD schema version $sv is below Windows Server 2016 level" 'Medium' `
                "Current schema version: $sv. Modern features including Windows LAPS (v89+), advanced Kerberos settings, and current GPO ADMX templates require schema v87+." `
                'Extend the schema by introducing a Windows Server 2016+ DC or running adprep /forestprep from Server 2016+ media.'
        }
    }
    catch { }

    # ── Sites with no DCs ──────────────────────────────────────────────
    try {
        $sites = @(Get-ADReplicationSite -Filter * @sp)
        foreach ($site in $sites) {
            $siteHasDC = try {
                @(Get-ADDomainController -Filter { Site -eq $site.Name } @sp -ErrorAction SilentlyContinue).Count -gt 0
            } catch { $true }  # default to true on error to avoid false positives
            if (-not $siteHasDC) {
                Add-Risk 'Sites & Services' "Site '$($site.Name)' has no domain controllers" 'Medium' `
                    "AD site '$($site.Name)' has subnets defined but no DCs. Clients in this site will authenticate cross-site, increasing latency and WAN dependency." `
                    'Deploy a DC in this site or verify whether the site is still needed. If the site is decommissioned, remove it and its subnets from AD Sites and Services.'
            }
        }
    }
    catch { }

    $critical = @($rows | Where-Object { $_.Severity -in 'Critical','High' }).Count
    $summaryColor = if ($critical -gt 0) { 'Yellow' } else { 'Green' }
    Write-Host "Operational risk analysis complete: $($rows.Count) findings | Critical/High: $critical" `
        -ForegroundColor $summaryColor

    $result = @($rows | Sort-Object {
        switch ($_.Severity) {
            'Critical' { 0 }; 'High' { 1 }; 'Medium' { 2 }; 'Low' { 3 }; default { 4 }
        }
    })
    if (-not [string]::IsNullOrWhiteSpace($ExportPath)) {
        Export-M365ReportData -InputObject $result -ExportPath $ExportPath | Out-Null
    }
    return $result
}
