function Get-ADReplicationHealth {
    <#
    .SYNOPSIS
        Returns per-DC replication partnership health including failure counts and lag.
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

    $dcs = @(Get-ADDomainController -Filter * @sp)
    $i   = 0

    foreach ($dc in $dcs) {
        $i++
        Write-Progress -Activity 'Collecting replication metadata' -Status $dc.HostName -PercentComplete (($i / [math]::Max($dcs.Count, 1)) * 100)

        try {
            $partners = @(Get-ADReplicationPartnerMetadata -Target $dc.HostName -ErrorAction Stop)

            foreach ($p in $partners) {
                # Parse partner DC name from the full DN
                # DN format: CN=NTDS Settings,CN=<DCName>,CN=<SiteName>,CN=Sites,...
                $partnerName = if ([string]$p.Partner -match 'CN=NTDS Settings,CN=([^,]+)') {
                    $Matches[1]
                } else {
                    [string]$p.Partner
                }

                $consecutive  = [int]$p.ConsecutiveReplicationFailures
                $lastSuccess  = $p.LastReplicationSuccess
                $lastAttempt  = $p.LastReplicationAttempt
                $sinceDays    = if ($lastSuccess) { [int]((Get-Date) - $lastSuccess).TotalDays } else { 9999 }
                $result       = [int]$p.LastReplicationResult

                # Friendly result string
                $resultDesc = switch ($result) {
                    0    { 'Success' }
                    8606 { 'Insufficient attributes' }
                    8453 { 'Access denied' }
                    8456 { 'Source DC unavailable' }
                    8457 { 'Destination DC unavailable' }
                    8461 { 'Replication preempted' }
                    8464 { 'Sync from source disabled' }
                    8418 { 'Schema mismatch' }
                    8545 { 'Replication object not found' }
                    default { if ($result -eq 0) { 'Success' } else { "Error 0x$($result.ToString('X'))($result)" } }
                }

                $status = if ($consecutive -eq 0 -and $result -eq 0) { 'OK' }
                          elseif ($consecutive -le 2)                 { 'Warning' }
                          else                                         { 'Critical' }

                $rows.Add([pscustomobject]@{
                    Server               = $dc.HostName
                    Partner              = $partnerName
                    Partition            = [string]$p.Partition
                    LastAttempt          = $lastAttempt
                    LastSuccess          = $lastSuccess
                    DaysSinceLastSuccess = if ($sinceDays -ge 9999) { 'Never' } else { [string]$sinceDays }
                    ConsecutiveFailures  = $consecutive
                    LastSyncResult       = $resultDesc
                    Status               = $status
                })
            }
        }
        catch {
            $rows.Add([pscustomobject]@{
                Server               = $dc.HostName
                Partner              = 'Query failed'
                Partition            = ''
                LastAttempt          = $null
                LastSuccess          = $null
                DaysSinceLastSuccess = ''
                ConsecutiveFailures  = -1
                LastSyncResult       = $_.Exception.Message
                Status               = 'Warning'
            })
        }
    }

    Write-Progress -Activity 'Collecting replication metadata' -Completed

    # Surface recent replication failures as additional rows
    try {
        $failures = @(Get-ADReplicationFailure -Target * @sp -ErrorAction SilentlyContinue)
        foreach ($f in $failures) {
            # Only add if not already represented above
            $alreadyCovered = $rows | Where-Object {
                $_.Server -eq $f.Server -and $_.Status -ne 'OK'
            }
            if (-not $alreadyCovered) {
                $rows.Add([pscustomobject]@{
                    Server               = $f.Server
                    Partner              = [string]$f.Partner
                    Partition            = ''
                    LastAttempt          = $f.FirstFailureTime
                    LastSuccess          = $null
                    DaysSinceLastSuccess = ''
                    ConsecutiveFailures  = $f.FailureCount
                    LastSyncResult       = [string]$f.FailureType
                    Status               = 'Critical'
                })
            }
        }
    }
    catch { }

    $result = @($rows | Sort-Object { switch ($_.Status) { 'Critical' { 0 } 'Warning' { 1 } default { 2 } } }, Server)
    if (-not [string]::IsNullOrWhiteSpace($ExportPath)) {
        Export-M365ReportData -InputObject $result -ExportPath $ExportPath | Out-Null
    }
    return $result
}
