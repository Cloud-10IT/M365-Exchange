function Get-ADDomainSummary {
    <#
    .SYNOPSIS
        Returns a structured domain/forest overview suitable for MSP assessment reports.
    .DESCRIPTION
        Evaluates forest configuration, domain settings, default domain password policy,
        and inventory counts. Does not modify any AD objects.
    .PARAMETER Server
        Optional. Target domain controller FQDN or domain name.
    .PARAMETER ExportPath
        Optional. Path to export CSV/JSON results.
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
            Write-Error 'ActiveDirectory module not found. Install RSAT: Active Directory Domain Services and Lightweight Directory Services Tools.'
            return @()
        }
    }

    $sp   = if ($Server) { @{ Server = $Server } } else { @{} }
    $rows = [System.Collections.Generic.List[object]]::new()

    function Add-Row ($cat, $item, $val, $status, $risk, $notes) {
        $rows.Add([pscustomobject]@{
            Category = $cat
            Item     = $item
            Value    = [string]$val
            Status   = $status
            Risk     = $risk
            Notes    = $notes
        })
    }

    # ── Forest ────────────────────────────────────────────────────────
    try {
        $forest = Get-ADForest @sp
        $rde    = Get-ADRootDSE @sp

        Add-Row 'Forest' 'Forest Name'            $forest.Name       'Info' '' "Root domain: $($forest.RootDomain)"
        Add-Row 'Forest' 'Forest Functional Level' $forest.ForestMode 'Info' '' ''

        $sv = try { (Get-ADObject $rde.schemaNamingContext -Properties objectVersion @sp).objectVersion } catch { 0 }
        $svDesc = switch ([int]$sv) {
            89 { "Windows Server 2022 (v$sv)" }
            88 { "Windows Server 2019 (v$sv)" }
            87 { "Windows Server 2016 (v$sv)" }
            72 { "Windows Server 2012 R2 (v$sv)" }
            69 { "Windows Server 2012 (v$sv)" }
            47 { "Windows Server 2008 R2 (v$sv)" }
            44 { "Windows Server 2008 (v$sv)" }
            default { "Version $sv" }
        }
        $svSt = if ([int]$sv -ge 87) { 'OK' } elseif ([int]$sv -ge 69) { 'Warning' } else { 'Critical' }
        $svRk = if ([int]$sv -lt 87) { "Schema version $sv (pre-2016). Windows LAPS, Kerberos Claims, and other modern controls require schema version >=87. Plan an in-place upgrade or forest raise." } else { '' }
        Add-Row 'Forest' 'Schema Version' $svDesc $svSt $svRk ''

        $rb = try {
            $f = Get-ADOptionalFeature -Filter { Name -eq 'Recycle Bin Feature' } @sp -ErrorAction SilentlyContinue
            if ($f -and @($f.EnabledScopes).Count -gt 0) { 'Enabled' } else { 'Disabled' }
        } catch { 'Unknown' }
        $rbSt = if ($rb -eq 'Enabled') { 'OK' } elseif ($rb -eq 'Disabled') { 'Warning' } else { 'Info' }
        $rbRk = if ($rb -ne 'Enabled') { 'AD Recycle Bin is disabled. Accidental deletions require authoritative restore from backup with domain controller downtime. Enable immediately (requires DFL >= 2008 R2).' } else { '' }
        Add-Row 'Forest' 'AD Recycle Bin' $rb $rbSt $rbRk ''
    }
    catch { Add-Row 'Forest' 'Forest Information' "Error: $($_.Exception.Message)" 'Warning' '' '' }

    # ── Domain ────────────────────────────────────────────────────────
    try {
        $dom = Get-ADDomain @sp
        Add-Row 'Domain' 'Domain Name (FQDN)'       $dom.DNSRoot    'Info' '' "NetBIOS: $($dom.NetBIOSName)"
        Add-Row 'Domain' 'Domain Functional Level'   $dom.DomainMode 'Info' '' ''

        $ts = try {
            $cn = (Get-ADRootDSE @sp).configurationNamingContext
            $v  = (Get-ADObject "CN=Directory Service,CN=Windows NT,CN=Services,$cn" -Properties tombstoneLifetime @sp).tombstoneLifetime
            if ($null -ne $v -and $v -ne 0) { [int]$v } else { 60 }
        } catch { 60 }
        $tsSt = if ($ts -ge 180) { 'OK' } elseif ($ts -ge 90) { 'Warning' } else { 'Critical' }
        $tsRk = if ($ts -lt 180) { "Tombstone lifetime of $ts days limits the disaster recovery window. A domain controller offline longer than $ts days cannot rejoin replication. Microsoft recommends 180 days minimum." } else { '' }
        Add-Row 'Domain' 'Tombstone Lifetime' "$ts days" $tsSt $tsRk ''

        $dcCt = try { @(Get-ADDomainController -Filter * @sp).Count } catch { '?' }
        Add-Row 'Domain' 'Domain Controller Count' $dcCt 'Info' '' ''

        try {
            $krb = Get-ADUser 'krbtgt' -Properties PasswordLastSet @sp
            $age = if ($krb.PasswordLastSet) { [int]((Get-Date) - $krb.PasswordLastSet).TotalDays } else { 9999 }
            $kSt = if ($age -le 180) { 'OK' } elseif ($age -le 365) { 'Warning' } else { 'Critical' }
            $kRk = if ($age -gt 180) { "KRBTGT password is $age days old. A stolen KRBTGT key enables Golden Ticket attacks — forged Kerberos tickets valid for up to 10 years that bypass all authentication controls. Rotate every 180 days." } else { '' }
            Add-Row 'Domain' 'KRBTGT Password Age' "$age days" $kSt $kRk "Last set: $($krb.PasswordLastSet)"
        }
        catch { Add-Row 'Domain' 'KRBTGT Password Age' 'Query failed' 'Warning' '' '' }
    }
    catch { Add-Row 'Domain' 'Domain Information' "Error: $($_.Exception.Message)" 'Warning' '' '' }

    # ── Default Domain Password Policy ────────────────────────────────
    try {
        $pp = Get-ADDefaultDomainPasswordPolicy @sp

        $ml = $pp.MinPasswordLength
        Add-Row 'Password Policy' 'Minimum Password Length' $ml `
            (if ($ml -ge 14) { 'OK' } elseif ($ml -ge 8) { 'Warning' } else { 'Critical' }) `
            (if ($ml -lt 14) { "Length $ml does not meet CIS Benchmark Level 1 (14+) or NIST SP 800-63B guidance. Short passwords are vulnerable to brute-force and credential stuffing attacks." } else { '' }) ''

        $lt = $pp.LockoutThreshold
        Add-Row 'Password Policy' 'Account Lockout Threshold' (if ($lt -eq 0) { 'Disabled' } else { [string]$lt }) `
            (if ($lt -gt 0 -and $lt -le 10) { 'OK' } elseif ($lt -eq 0) { 'Critical' } else { 'Warning' }) `
            (if ($lt -eq 0) { 'Account lockout is DISABLED. Online brute-force and password spray attacks will not be blocked by lockout.' } elseif ($lt -gt 10) { "Lockout threshold of $lt allows many guesses per account. An attacker can test $lt passwords before lockout triggers." } else { '' }) ''

        $ld = $pp.LockoutDuration
        Add-Row 'Password Policy' 'Lockout Duration' (if ($ld.TotalMinutes -eq 0) { 'Until admin unlock' } else { "$([int]$ld.TotalMinutes) minutes" }) 'Info' '' ''

        $lo = $pp.LockoutObservationWindow
        Add-Row 'Password Policy' 'Observation Window' "$([int]$lo.TotalMinutes) minutes" 'Info' '' ''

        $cx = $pp.ComplexityEnabled
        Add-Row 'Password Policy' 'Password Complexity' (if ($cx) { 'Enabled' } else { 'Disabled' }) `
            (if ($cx) { 'OK' } else { 'Warning' }) `
            (if (-not $cx) { 'Complexity disabled — users may set simple, easily guessable passwords.' } else { '' }) ''

        $rev = $pp.ReversibleEncryptionEnabled
        Add-Row 'Password Policy' 'Reversible Encryption' (if ($rev) { 'ENABLED' } else { 'Disabled' }) `
            (if (-not $rev) { 'OK' } else { 'Critical' }) `
            (if ($rev) { 'Reversible encryption stores passwords in a recoverable form — functionally equivalent to plaintext. Any DC compromise yields all user passwords. Disable immediately.' } else { '' }) ''

        $mx = [int]$pp.MaxPasswordAge.TotalDays
        Add-Row 'Password Policy' 'Maximum Password Age' (if ($mx -eq 0) { 'No expiry' } else { "$mx days" }) `
            (if ($mx -eq 0) { 'Warning' } elseif ($mx -le 365) { 'OK' } else { 'Warning' }) `
            (if ($mx -eq 0) { 'Passwords never expire at the domain policy level. Verify Fine-Grained Password Policies are compensating.' } else { '' }) ''

        $hist = $pp.PasswordHistoryCount
        Add-Row 'Password Policy' 'Password History Count' $hist `
            (if ($hist -ge 24) { 'OK' } elseif ($hist -ge 5) { 'Warning' } else { 'Critical' }) `
            (if ($hist -lt 24) { "History count $hist is below CIS recommendation of 24. Users can recycle previously used passwords." } else { '' }) ''
    }
    catch { Add-Row 'Password Policy' 'Default Domain Password Policy' "Error: $($_.Exception.Message)" 'Warning' '' '' }

    # ── Inventory snapshot ─────────────────────────────────────────────
    try {
        $eu  = (Get-ADUser     -Filter { Enabled -eq $true  } -ResultSetSize $null @sp | Measure-Object).Count
        $du  = (Get-ADUser     -Filter { Enabled -eq $false } -ResultSetSize $null @sp | Measure-Object).Count
        $ec  = (Get-ADComputer -Filter { Enabled -eq $true  } -ResultSetSize $null @sp | Measure-Object).Count
        $dc2 = (Get-ADComputer -Filter { Enabled -eq $false } -ResultSetSize $null @sp | Measure-Object).Count
        $gc  = try { (Get-ADGroup -Filter * -ResultSetSize $null @sp | Measure-Object).Count } catch { '?' }
        Add-Row 'Inventory' 'Enabled User Accounts'      $eu  'Info' '' ''
        Add-Row 'Inventory' 'Disabled User Accounts'     $du  'Info' '' ''
        Add-Row 'Inventory' 'Enabled Computer Accounts'  $ec  'Info' '' ''
        Add-Row 'Inventory' 'Disabled Computer Accounts' $dc2 'Info' '' ''
        Add-Row 'Inventory' 'Total Security Groups'      $gc  'Info' '' ''
    }
    catch { Add-Row 'Inventory' 'Account Counts' "Error: $($_.Exception.Message)" 'Warning' '' '' }

    $result = @($rows)
    if (-not [string]::IsNullOrWhiteSpace($ExportPath)) {
        Export-M365ReportData -InputObject $result -ExportPath $ExportPath | Out-Null
    }
    return $result
}
