function Get-ADSitesAndServicesReport {
    <#
    .SYNOPSIS
        Returns AD Sites and Services topology: sites, subnets, and site links with risk flags.
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

    $sp   = if ($Server) { @{ Server = $Server } } else { @{} }
    $rows = [System.Collections.Generic.List[object]]::new()

    # ── Sites ──────────────────────────────────────────────────────────
    try {
        $sites   = @(Get-ADReplicationSite -Filter * @sp -Properties Description, Location)
        $subnets = @(Get-ADReplicationSubnet -Filter * @sp -Properties Site, Location)

        # Map site name → subnet count
        $subnetsBySite = @{}
        foreach ($sn in $subnets) {
            $siteName = if ($sn.Site) { ($sn.Site -replace 'CN=([^,]+),.*','$1') } else { '(unassigned)' }
            if (-not $subnetsBySite.ContainsKey($siteName)) { $subnetsBySite[$siteName] = [System.Collections.Generic.List[string]]::new() }
            $subnetsBySite[$siteName].Add($sn.Name)
        }

        foreach ($site in $sites) {
            $siteName    = $site.Name
            $siteSubnets = if ($subnetsBySite.ContainsKey($siteName)) { $subnetsBySite[$siteName] } else { @() }

            $dcCount = try {
                @(Get-ADDomainController -Filter { Site -eq $siteName } @sp -ErrorAction SilentlyContinue).Count
            } catch { '?' }

            $issues = [System.Collections.Generic.List[string]]::new()
            if ($siteSubnets.Count -eq 0) { $issues.Add('No subnets assigned — clients may authenticate against wrong site DC') }
            if ($dcCount -eq 0)           { $issues.Add('No DCs in site — authentication falls back to other site (cross-site auth latency)') }
            if ($dcCount -eq 1)           { $issues.Add('Single DC in site — no fault tolerance; DC failure impacts all site clients') }

            $status = if ($issues.Count -gt 0) { 'Warning' } else { 'OK' }

            $rows.Add([pscustomobject]@{
                Type               = 'Site'
                Name               = $siteName
                Location           = [string]$site.Location
                Description        = [string]$site.Description
                SubnetCount        = $siteSubnets.Count
                Subnets            = $siteSubnets -join '; '
                DCsInSite          = $dcCount
                ReplicationInterval = ''
                Cost               = ''
                ChangeNotification = ''
                Status             = $status
                Notes              = $issues -join '; '
            })
        }

        # Subnets with no site assignment
        $orphanSubnets = @($subnets | Where-Object { -not $_.Site })
        foreach ($sn in $orphanSubnets) {
            $rows.Add([pscustomobject]@{
                Type               = 'Subnet'
                Name               = $sn.Name
                Location           = [string]$sn.Location
                Description        = ''
                SubnetCount        = ''
                Subnets            = ''
                DCsInSite          = ''
                ReplicationInterval = ''
                Cost               = ''
                ChangeNotification = ''
                Status             = 'Warning'
                Notes              = 'Subnet not assigned to any site — clients on this subnet will use site selection heuristics and may authenticate against a distant DC.'
            })
        }
    }
    catch {
        $rows.Add([pscustomobject]@{
            Type = 'Site'; Name = 'Error'; Location = ''; Description = ''; SubnetCount = ''
            Subnets = ''; DCsInSite = ''; ReplicationInterval = ''; Cost = ''
            ChangeNotification = ''; Status = 'Warning'
            Notes = "Site query failed: $($_.Exception.Message)"
        })
    }

    # ── Site Links ─────────────────────────────────────────────────────
    try {
        $links = @(Get-ADReplicationSiteLink -Filter * @sp -Properties Cost, ReplicationFrequencyInMinutes, Options, SitesIncluded, Description)

        foreach ($link in $links) {
            $interval = [int]$link.ReplicationFrequencyInMinutes
            $cost     = [int]$link.Cost
            # Options bit 1 = change notification (SMTP-based or IP with change notification enabled)
            $changeNotification = ($link.Options -band 1) -eq 1

            $issues = [System.Collections.Generic.List[string]]::new()
            if ($interval -gt 60)  { $issues.Add("Replication interval $interval min (>60) — changes may take hours to propagate across the WAN") }
            if ($interval -gt 180) { $issues.Add("Interval $interval min is very high — urgent password resets and account lockouts may lag significantly") }

            # Parse site names from SitesIncluded DNs
            $linkedSites = @($link.SitesIncluded | ForEach-Object { ($_ -replace 'CN=([^,]+),.*','$1') }) -join ', '

            $status = if ($issues.Count -gt 0) { 'Warning' } else { 'OK' }

            $rows.Add([pscustomobject]@{
                Type               = 'Site Link'
                Name               = $link.Name
                Location           = ''
                Description        = [string]$link.Description
                SubnetCount        = ''
                Subnets            = ''
                DCsInSite          = ''
                ReplicationInterval = "$interval minutes"
                Cost               = $cost
                ChangeNotification = $changeNotification
                Status             = $status
                Notes              = if ($issues.Count -gt 0) { $issues -join '; ' } else { "Links: $linkedSites" }
            })
        }
    }
    catch {
        $rows.Add([pscustomobject]@{
            Type = 'Site Link'; Name = 'Error'; Location = ''; Description = ''; SubnetCount = ''
            Subnets = ''; DCsInSite = ''; ReplicationInterval = ''; Cost = ''
            ChangeNotification = ''; Status = 'Warning'
            Notes = "Site link query failed: $($_.Exception.Message)"
        })
    }

    $result = @($rows | Sort-Object Type, { switch ($_.Status) { 'Critical' { 0 } 'Warning' { 1 } default { 2 } } }, Name)
    if (-not [string]::IsNullOrWhiteSpace($ExportPath)) {
        Export-M365ReportData -InputObject $result -ExportPath $ExportPath | Out-Null
    }
    return $result
}
