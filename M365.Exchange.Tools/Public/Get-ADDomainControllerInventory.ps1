function Get-ADDomainControllerInventory {
    <#
    .SYNOPSIS
        Returns a per-DC inventory with OS, roles, end-of-support status, and health indicators.
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

    $sp = if ($Server) { @{ Server = $Server } } else { @{} }

    # OS end-of-support dates (mainstream support end; extended where relevant)
    $eosMap = @{
        '2003'    = [datetime]'2015-07-14'
        '2008'    = [datetime]'2020-01-14'
        '2012'    = [datetime]'2023-10-10'
        '2016'    = [datetime]'2027-01-12'
        '2019'    = [datetime]'2029-01-09'
        '2022'    = [datetime]'2031-10-14'
        '2025'    = [datetime]'2034-10-10'
    }

    $dcs  = @(Get-ADDomainController -Filter * @sp)
    $rows = [System.Collections.Generic.List[object]]::new()
    $i    = 0

    foreach ($dc in $dcs) {
        $i++
        Write-Progress -Activity 'Inventorying Domain Controllers' -Status $dc.HostName -PercentComplete (($i / [math]::Max($dcs.Count, 1)) * 100)

        $os    = [string]$dc.OperatingSystem
        $osVer = [string]$dc.OperatingSystemServicePack

        # Match end-of-support year
        $eosYear = $null
        foreach ($yr in @('2025','2022','2019','2016','2012','2008','2003')) {
            if ($os -like "*$yr*") { $eosYear = $yr; break }
        }
        $eosDate    = if ($eosYear) { $eosMap[$eosYear] } else { $null }
        $eosDisplay = if ($eosDate) { $eosDate.ToString('yyyy-MM-dd') } else { 'Unknown' }
        $eosStatus  = if (-not $eosDate) { 'Info' }
                      elseif ($eosDate -lt (Get-Date))                    { 'Critical' }
                      elseif ($eosDate -lt (Get-Date).AddMonths(12))      { 'Warning'  }
                      else                                                 { 'OK'       }

        # FSMO roles held by this DC
        $roles = @($dc.OperationMasterRoles | ForEach-Object {
            switch ($_) {
                'PDCEmulator'          { 'PDC Emulator' }
                'RIDMaster'            { 'RID Master' }
                'InfrastructureMaster' { 'Infrastructure Master' }
                'SchemaMaster'         { 'Schema Master' }
                'DomainNamingMaster'   { 'Domain Naming Master' }
            }
        })

        # SYSVOL replication type (DFSR vs FRS) — check for msDFSR-Subscription under DC computer object
        $replicationMode = try {
            $dcCompDN = $dc.ComputerObjectDN
            $dfsrSub  = Get-ADObject -SearchBase $dcCompDN -Filter { objectClass -eq 'msDFSR-Subscription' } @sp -ErrorAction SilentlyContinue
            if ($dfsrSub) { 'DFSR' } else { 'FRS (legacy)' }
        } catch { 'Unknown' }

        # Try to ping
        $pingable = try { (Test-Connection -ComputerName $dc.HostName -Count 1 -Quiet -ErrorAction SilentlyContinue) } catch { $false }

        # Overall status
        $status = if ($eosStatus -eq 'Critical') { 'Critical' }
                  elseif (-not $pingable)         { 'Warning'  }
                  elseif ($eosStatus -eq 'Warning') { 'Warning' }
                  else                             { 'OK' }

        $rows.Add([pscustomobject]@{
            Name              = $dc.HostName
            Site              = $dc.Site
            IPAddress         = $dc.IPv4Address
            OperatingSystem   = $os
            OSVersion         = $dc.OperatingSystemVersion
            IsGlobalCatalog   = $dc.IsGlobalCatalog
            IsReadOnlyDC      = $dc.IsReadOnly
            FSMORoles         = if ($roles.Count -gt 0) { $roles -join '; ' } else { '' }
            SYSVOLReplication = $replicationMode
            Pingable          = $pingable
            EOSDate           = $eosDisplay
            OSStatus          = $eosStatus
            Status            = $status
        })
    }

    Write-Progress -Activity 'Inventorying Domain Controllers' -Completed

    $result = @($rows | Sort-Object Status, Name)
    if (-not [string]::IsNullOrWhiteSpace($ExportPath)) {
        Export-M365ReportData -InputObject $result -ExportPath $ExportPath | Out-Null
    }
    return $result
}
