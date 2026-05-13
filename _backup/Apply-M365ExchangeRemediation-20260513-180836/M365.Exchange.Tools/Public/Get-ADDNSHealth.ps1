function Get-ADDNSHealth {
    <#
    .SYNOPSIS
        Returns DNS zone health for AD-integrated DNS servers.
        Uses the DnsServer module when available; falls back to AD zone enumeration.
    .PARAMETER Server
        Optional. DNS/DC server to query. Defaults to current DC.
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

    $sp          = if ($Server) { @{ Server = $Server } } else { @{} }
    $dnsAvail    = [bool](Get-Module DnsServer -ListAvailable -ErrorAction SilentlyContinue)
    $dnsServer   = if ($Server) { $Server } else { (Get-ADDomainController -Discover @sp -ErrorAction SilentlyContinue).HostName }
    $rows        = [System.Collections.Generic.List[object]]::new()

    if ($dnsAvail) {
        Import-Module DnsServer -ErrorAction SilentlyContinue

        # Global DNS server settings
        try {
            $dnsSettings = Get-DnsServer -ComputerName $dnsServer -ErrorAction Stop

            # Forwarders
            $fwdList = @($dnsSettings.ServerForwarder.IPAddress) | Where-Object { $_ }
            $fwdDisplay = if ($fwdList.Count -gt 0) { $fwdList -join ', ' } else { 'None configured' }
            $fwdStatus  = if ($fwdList.Count -gt 0) { 'Info' } else { 'Warning' }
            $rows.Add([pscustomobject]@{
                ZoneName         = '(Server)'
                ZoneType         = 'Global Setting'
                DynamicUpdate    = ''
                IsADIntegrated   = ''
                ScavengingEnabled = ''
                RecordCount      = ''
                Status           = $fwdStatus
                Notes            = "Forwarders: $fwdDisplay"
            })

            # Root hints
            $rootHints = @($dnsSettings.ServerRootHint)
            $rows.Add([pscustomobject]@{
                ZoneName         = '(Server)'
                ZoneType         = 'Global Setting'
                DynamicUpdate    = ''
                IsADIntegrated   = ''
                ScavengingEnabled = ''
                RecordCount      = ''
                Status           = 'Info'
                Notes            = "Root hints: $($rootHints.Count) configured"
            })
        }
        catch { }

        # Zones
        try {
            $zones = @(Get-DnsServerZone -ComputerName $dnsServer -ErrorAction Stop)
            $i = 0
            foreach ($z in $zones) {
                $i++
                Write-Progress -Activity 'Collecting DNS zone data' -Status $z.ZoneName -PercentComplete (($i / [math]::Max($zones.Count, 1)) * 100)

                # Skip cache zones
                if ($z.ZoneType -eq 'Cache') { continue }

                $scavEnabled = $false
                try {
                    $aging = Get-DnsServerZoneAging -ZoneName $z.ZoneName -ComputerName $dnsServer -ErrorAction SilentlyContinue
                    $scavEnabled = [bool]$aging.AgingEnabled
                } catch { }

                $recCount = try {
                    @(Get-DnsServerResourceRecord -ZoneName $z.ZoneName -ComputerName $dnsServer -ErrorAction SilentlyContinue).Count
                } catch { '?' }

                # Status logic
                $issues = [System.Collections.Generic.List[string]]::new()
                if (-not $scavEnabled -and $z.ZoneType -eq 'Primary') {
                    $issues.Add('Scavenging disabled — stale DNS records accumulate over time')
                }
                if ($z.DynamicUpdate -eq 'None' -and $z.ZoneType -eq 'Primary' -and $z.ZoneName -notlike '*.arpa') {
                    $issues.Add('Dynamic update disabled — clients cannot auto-register')
                }

                $status = if ($issues.Count -ge 2) { 'Warning' }
                          elseif ($issues.Count -eq 1) { 'Warning' }
                          else { 'OK' }

                $rows.Add([pscustomobject]@{
                    ZoneName         = $z.ZoneName
                    ZoneType         = [string]$z.ZoneType
                    DynamicUpdate    = [string]$z.DynamicUpdate
                    IsADIntegrated   = $z.IsAutoCreated -eq $false -and $z.ZoneType -ne 'Stub'
                    ScavengingEnabled = $scavEnabled
                    RecordCount      = $recCount
                    Status           = $status
                    Notes            = $issues -join '; '
                })
            }
            Write-Progress -Activity 'Collecting DNS zone data' -Completed
        }
        catch {
            $rows.Add([pscustomobject]@{
                ZoneName = 'Error'
                ZoneType = ''; DynamicUpdate = ''; IsADIntegrated = ''; ScavengingEnabled = ''
                RecordCount = ''; Status = 'Warning'
                Notes = "DnsServer zone query failed: $($_.Exception.Message)"
            })
        }
    }
    else {
        # Fallback: enumerate AD-integrated DNS zones from directory partitions
        Write-Host 'DnsServer module not available. Falling back to AD zone enumeration.' -ForegroundColor Yellow
        try {
            $dom        = Get-ADDomain @sp
            $domainDN   = $dom.DistinguishedName
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
                        if ($zName -eq 'RootDNSServers' -or $zName -eq '..TrustAnchors') { continue }

                        $partition = if ($base -like '*DomainDnsZones*') { 'DomainDnsZones' }
                                     elseif ($base -like '*ForestDnsZones*') { 'ForestDnsZones' }
                                     else { 'Legacy/System' }

                        $rows.Add([pscustomobject]@{
                            ZoneName          = $zName
                            ZoneType          = "AD-Integrated ($partition)"
                            DynamicUpdate     = 'Unknown (DnsServer module required)'
                            IsADIntegrated    = $true
                            ScavengingEnabled = 'Unknown'
                            RecordCount       = '?'
                            Status            = 'Info'
                            Notes             = 'Install DnsServer RSAT for full DNS health data.'
                        })
                    }
                } catch { }
            }
        }
        catch {
            $rows.Add([pscustomobject]@{
                ZoneName = 'Error'; ZoneType = ''; DynamicUpdate = ''; IsADIntegrated = ''
                ScavengingEnabled = ''; RecordCount = ''; Status = 'Warning'
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
