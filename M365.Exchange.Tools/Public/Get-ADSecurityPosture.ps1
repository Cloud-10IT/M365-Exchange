function Get-ADSecurityPosture {
    <#
    .SYNOPSIS
        Performs a structured security posture assessment against the Active Directory environment.
        Returns one row per security check with status (Pass/Warning/Critical/Info), risk description,
        and remediation guidance. Read-only — no environment changes are made.
    .PARAMETER Server
        Optional. Target domain controller FQDN or domain name.
    .PARAMETER StaleAccountDays
        Number of days without logon before a user account is considered stale. Default: 90.
    .PARAMETER ExportPath
        Optional. Path to export results.
    #>
    [CmdletBinding()]
    param(
        [Parameter()][string]$Server,
        [Parameter()][int]$StaleAccountDays = 90,
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
    $num  = 0

    function Add-Check ($name, $status, $details, $risk, $recommendation) {
        $script:num++
        $rows.Add([pscustomobject]@{
            '#'              = $script:num
            SecurityCheck    = $name
            Status           = $status
            Details          = $details
            Risk             = $risk
            Recommendation   = $recommendation
        })
    }

    # ── 1. Domain Admins membership ────────────────────────────────────
    try {
        $da = @(Get-ADGroupMember 'Domain Admins' -Recursive @sp -ErrorAction Stop | Where-Object { $_.objectClass -eq 'user' })
        $ct = $da.Count
        $st = if ($ct -le 3) { 'Pass' } elseif ($ct -le 5) { 'Warning' } else { 'Critical' }
        $rk = if ($ct -gt 5) { 'Excessive Domain Admin membership dramatically increases blast radius. Each additional DA account is a potential pivot point to full domain compromise.' } else { '' }
        $names = ($da | Select-Object -First 10 | ForEach-Object { $_.SamAccountName }) -join ', '
        Add-Check 'Domain Admins membership count' $st "$ct member(s): $names" $rk 'Target <=3 members. Named DA accounts should be used only for specific AD tasks, not for daily workstation use. Use tiered admin model.'
    }
    catch { Add-Check 'Domain Admins membership count' 'Warning' "Query failed: $($_.Exception.Message)" '' '' }

    # ── 2. Enterprise Admins should be empty ──────────────────────────
    try {
        $ea  = @(Get-ADGroupMember 'Enterprise Admins' -Recursive @sp -ErrorAction Stop | Where-Object { $_.objectClass -eq 'user' })
        $ct  = $ea.Count
        $st  = if ($ct -eq 0) { 'Pass' } else { 'Warning' }
        $rk  = if ($ct -gt 0) { "Enterprise Admins grants forest-wide control. Permanent membership is unnecessary — this group should be empty except during forest-level operations." } else { '' }
        $names = ($ea | ForEach-Object { $_.SamAccountName }) -join ', '
        Add-Check 'Enterprise Admins should be empty' $st "$ct member(s)$(if ($names) { ': ' + $names })" $rk 'Remove all permanent members. Use time-limited membership via PAM or JIT access when forest operations are required.'
    }
    catch { Add-Check 'Enterprise Admins should be empty' 'Warning' "Query failed: $($_.Exception.Message)" '' '' }

    # ── 3. Schema Admins should be empty ──────────────────────────────
    try {
        $sa  = @(Get-ADGroupMember 'Schema Admins' -Recursive @sp -ErrorAction Stop | Where-Object { $_.objectClass -eq 'user' })
        $ct  = $sa.Count
        $st  = if ($ct -eq 0) { 'Pass' } else { 'Warning' }
        $rk  = if ($ct -gt 0) { "Schema Admins can extend and modify the AD schema — a destructive or misconfigured change is difficult or impossible to reverse." } else { '' }
        $names = ($sa | ForEach-Object { $_.SamAccountName }) -join ', '
        Add-Check 'Schema Admins should be empty' $st "$ct member(s)$(if ($names) { ': ' + $names })" $rk 'Remove all permanent members. Add only during schema extension operations, then remove immediately after.'
    }
    catch { Add-Check 'Schema Admins should be empty' 'Warning' "Query failed: $($_.Exception.Message)" '' '' }

    # ── 4. Built-in Administrator account (SID -500) ───────────────────
    try {
        $dom     = Get-ADDomain @sp
        $adminSid = "$($dom.DomainSID)-500"
        $adminAcct = Get-ADUser -Identity $adminSid -Properties Name, Enabled, PasswordNeverExpires, LastLogonDate @sp
        $issues = [System.Collections.Generic.List[string]]::new()
        if ($adminAcct.Name -eq 'Administrator') { $issues.Add('Account has default name "Administrator" — should be renamed') }
        if ($adminAcct.Enabled)                  { $issues.Add('Account is enabled — should be disabled when not in active use') }
        if ($adminAcct.PasswordNeverExpires)     { $issues.Add('Password never expires') }
        $st = if ($issues.Count -eq 0) { 'Pass' } elseif ($issues.Count -eq 1) { 'Warning' } else { 'Critical' }
        $rk = if ($issues.Count -gt 0) { 'The built-in Administrator is a high-value target — its SID is well-known and it cannot be locked out. Attackers actively target this account.' } else { '' }
        Add-Check 'Built-in Administrator account hardening' $st ($issues -join '; ') $rk 'Rename account, disable when not in use, set a long complex password, and enable auditing on all uses.'
    }
    catch { Add-Check 'Built-in Administrator account hardening' 'Warning' "Query failed: $($_.Exception.Message)" '' '' }

    # ── 5. Built-in Guest account (SID -501) ───────────────────────────
    try {
        $dom      = Get-ADDomain @sp
        $guestSid = "$($dom.DomainSID)-501"
        $guest    = Get-ADUser -Identity $guestSid -Properties Enabled @sp
        $st  = if (-not $guest.Enabled) { 'Pass' } else { 'Critical' }
        $rk  = if ($guest.Enabled) { 'The Guest account provides unauthenticated access to domain resources. It should always be disabled.' } else { '' }
        Add-Check 'Built-in Guest account is disabled' $st "Enabled: $($guest.Enabled)" $rk 'Disable the Guest account. Verify it has no group memberships.'
    }
    catch { Add-Check 'Built-in Guest account is disabled' 'Warning' "Query failed: $($_.Exception.Message)" '' '' }

    # ── 6. Accounts with DoNotRequirePreAuth (AS-REP Roasting) ────────
    try {
        $asrep = @(Get-ADUser -Filter { DoesNotRequirePreAuth -eq $true -and Enabled -eq $true } -Properties DoesNotRequirePreAuth @sp)
        $ct    = $asrep.Count
        $st    = if ($ct -eq 0) { 'Pass' } else { 'Critical' }
        $names = ($asrep | ForEach-Object { $_.SamAccountName }) -join ', '
        $rk    = if ($ct -gt 0) { "AS-REP roastable accounts allow offline Kerberos hash cracking without credentials. An attacker can request AS-REP for any of these accounts and crack the hash offline." } else { '' }
        Add-Check 'No accounts with Kerberos pre-auth disabled (AS-REP Roasting)' $st "$ct account(s)$(if ($names) { ': ' + $names })" $rk 'Enable pre-authentication for all user accounts unless specifically required for legacy applications. Rotate passwords for any affected accounts.'
    }
    catch { Add-Check 'No accounts with Kerberos pre-auth disabled (AS-REP Roasting)' 'Warning' "Query failed: $($_.Exception.Message)" '' '' }

    # ── 7. Kerberoastable accounts (user accounts with SPNs) ───────────
    try {
        $kerb = @(Get-ADUser -Filter { ServicePrincipalName -like '*' -and Enabled -eq $true } -Properties ServicePrincipalName @sp |
                  Where-Object { $_.SamAccountName -ne 'krbtgt' })
        $ct   = $kerb.Count
        $st   = if ($ct -eq 0) { 'Pass' } elseif ($ct -le 3) { 'Warning' } else { 'Warning' }
        $names = ($kerb | ForEach-Object { $_.SamAccountName }) -join ', '
        $rk   = if ($ct -gt 0) { "User accounts with SPNs are Kerberoastable — any domain user can request their Kerberos service tickets and crack them offline to obtain plaintext passwords." } else { '' }
        Add-Check 'User accounts with SPNs (Kerberoastable)' $st "$ct account(s)$(if ($names) { ': ' + $names })" $rk 'Migrate service accounts to Group Managed Service Accounts (gMSA) — gMSA passwords are 240-byte random strings automatically rotated. Ensure all service account passwords are >=25 characters.'
    }
    catch { Add-Check 'User accounts with SPNs (Kerberoastable)' 'Warning' "Query failed: $($_.Exception.Message)" '' '' }

    # ── 8. Unconstrained delegation on non-DC computers ───────────────
    try {
        $uncon = @(Get-ADComputer -Filter { TrustedForDelegation -eq $true } -Properties TrustedForDelegation, PrimaryGroupID @sp |
                   Where-Object { $_.PrimaryGroupID -ne 516 -and $_.PrimaryGroupID -ne 521 })
        $ct    = $uncon.Count
        $st    = if ($ct -eq 0) { 'Pass' } else { 'Critical' }
        $names = ($uncon | ForEach-Object { $_.Name }) -join ', '
        $rk    = if ($ct -gt 0) { "Unconstrained delegation allows these machines to impersonate any user to any service. Compromise of these hosts plus a privileged user connection = full domain takeover via Pass-the-Ticket." } else { '' }
        Add-Check 'No non-DC computers with unconstrained delegation' $st "$ct computer(s)$(if ($names) { ': ' + $names })" $rk 'Replace with constrained delegation (KCD) or resource-based constrained delegation (RBCD). Remove unconstrained delegation from all non-DC machines.'
    }
    catch { Add-Check 'No non-DC computers with unconstrained delegation' 'Warning' "Query failed: $($_.Exception.Message)" '' '' }

    # ── 9. Password Never Expires on enabled user accounts ────────────
    try {
        $allEnabled = (Get-ADUser -Filter { Enabled -eq $true } -ResultSetSize $null @sp | Measure-Object).Count
        $pne        = @(Get-ADUser -Filter { Enabled -eq $true -and PasswordNeverExpires -eq $true } -ResultSetSize $null @sp)
        $ct         = $pne.Count
        $pct        = if ($allEnabled -gt 0) { [int]($ct / $allEnabled * 100) } else { 0 }
        $st         = if ($pct -le 5) { 'Pass' } elseif ($pct -le 15) { 'Warning' } else { 'Critical' }
        $rk         = if ($pct -gt 5) { "Accounts that never expire accumulate over time — former employees, service stubs, and test accounts remain indefinitely with unchanged passwords vulnerable to credential stuffing." } else { '' }
        Add-Check 'Enabled accounts with password-never-expires' $st "$ct accounts ($pct% of enabled users)" $rk 'Review and clean up. Use Fine-Grained Password Policies for service accounts that legitimately need non-expiring passwords. Ensure all such accounts have long, complex, unique passwords.'
    }
    catch { Add-Check 'Enabled accounts with password-never-expires' 'Warning' "Query failed: $($_.Exception.Message)" '' '' }

    # ── 10. Stale enabled user accounts ───────────────────────────────
    try {
        $cutoff     = (Get-Date).AddDays(-$StaleAccountDays)
        $allEnabled = (Get-ADUser -Filter { Enabled -eq $true } -ResultSetSize $null @sp | Measure-Object).Count
        # LastLogonDate is the replicated timestamp (replicated every 14 days by default)
        $stale = @(Get-ADUser -Filter { Enabled -eq $true -and LastLogonDate -lt $cutoff } -Properties LastLogonDate -ResultSetSize $null @sp |
                   Where-Object { $_.LastLogonDate -ne $null })
        # Also include accounts that have never logged on (LastLogonDate is null/zero)
        $neverLogon = @(Get-ADUser -Filter { Enabled -eq $true -and LastLogonDate -notlike '*' } -ResultSetSize $null @sp)
        $totalStale = $stale.Count + $neverLogon.Count
        $pct        = if ($allEnabled -gt 0) { [int]($totalStale / $allEnabled * 100) } else { 0 }
        $st         = if ($pct -le 5) { 'Pass' } elseif ($pct -le 15) { 'Warning' } else { 'Critical' }
        $rk         = if ($totalStale -gt 0) { "Stale enabled accounts represent unmonitored credential exposure. Attackers can use dormant accounts for persistent access — anomalous logins may go unnoticed for months." } else { '' }
        Add-Check "Stale enabled user accounts (no login >$StaleAccountDays days)" $st "$totalStale accounts ($pct% of enabled users; includes $($neverLogon.Count) never logged on)" $rk 'Implement an automated stale account detection and disablement policy. Consider disabling accounts after 60 days of inactivity and deleting after 180 days.'
    }
    catch { Add-Check "Stale enabled user accounts (no login >$StaleAccountDays days)" 'Warning' "Query failed: $($_.Exception.Message)" '' '' }

    # ── 11. Protected Users group usage ────────────────────────────────
    try {
        $pu  = @(Get-ADGroupMember 'Protected Users' -Recursive @sp -ErrorAction Stop | Where-Object { $_.objectClass -eq 'user' })
        $ct  = $pu.Count
        $st  = if ($ct -gt 0) { 'Pass' } else { 'Warning' }
        $rk  = if ($ct -eq 0) { 'Protected Users group is empty. Members receive enhanced protections: no NTLM fallback, no DES/RC4 Kerberos, credential cache cleared on logoff, short ticket lifetimes. All privileged accounts should be members.' } else { '' }
        $names = ($pu | Select-Object -First 10 | ForEach-Object { $_.SamAccountName }) -join ', '
        Add-Check 'Protected Users group is in use' $st "$ct member(s)$(if ($names) { ': ' + $names })" $rk 'Add all Domain Admins, privileged service accounts, and executive accounts to Protected Users. Test each account for compatibility before bulk-adding.'
    }
    catch { Add-Check 'Protected Users group is in use' 'Warning' "Query failed: $($_.Exception.Message)" '' '' }

    # ── 12. LAPS deployment ────────────────────────────────────────────
    try {
        $schemaDN   = (Get-ADRootDSE @sp).schemaNamingContext
        $lapsOld    = Get-ADObject -SearchBase $schemaDN -Filter { lDAPDisplayName -eq 'ms-Mcs-AdmPwd' }    @sp -ErrorAction SilentlyContinue
        $lapsNew    = Get-ADObject -SearchBase $schemaDN -Filter { lDAPDisplayName -eq 'msLAPS-Password' }   @sp -ErrorAction SilentlyContinue
        $lapsNewEnc = Get-ADObject -SearchBase $schemaDN -Filter { lDAPDisplayName -eq 'msLAPS-EncryptedPassword' } @sp -ErrorAction SilentlyContinue

        $lapsType = if ($lapsNewEnc -or $lapsNew) { 'Windows LAPS (built-in)' }
                    elseif ($lapsOld)               { 'Legacy LAPS (Microsoft LAPS)' }
                    else                            { 'Not deployed' }
        $st = if ($lapsNewEnc -or $lapsNew) { 'Pass' }
              elseif ($lapsOld)              { 'Warning' }
              else                           { 'Critical' }
        $rk = switch ($lapsType) {
            'Not deployed'              { 'Local administrator passwords are unmanaged — every workstation likely shares the same local admin password. Compromise of one machine allows lateral movement to all others via Pass-the-Hash.' }
            'Legacy LAPS (Microsoft LAPS)' { 'Legacy LAPS stores passwords in plaintext in AD (ms-Mcs-AdmPwd). Any user who can read this attribute can retrieve local admin passwords. Migrate to Windows LAPS which supports encrypted storage.' }
            default                     { '' }
        }
        Add-Check 'Local Administrator Password Solution (LAPS) deployed' $st $lapsType $rk 'Deploy Windows LAPS (built into Server 2025/Windows 11 22H2). Configure with encrypted storage and rotation policy.'
    }
    catch { Add-Check 'LAPS deployment' 'Warning' "Schema query failed: $($_.Exception.Message)" '' '' }

    # ── 13. Fine-Grained Password Policies ─────────────────────────────
    try {
        $fgpp = @(Get-ADFineGrainedPasswordPolicy -Filter * @sp -ErrorAction SilentlyContinue)
        $ct   = $fgpp.Count
        $st   = if ($ct -gt 0) { 'Info' } else { 'Warning' }
        $rk   = if ($ct -eq 0) { 'No Fine-Grained Password Policies (FGPPs). All accounts use the Default Domain Policy — service accounts requiring non-expiring passwords cannot be separated from user accounts without policy-level exceptions.' } else { '' }
        $names = ($fgpp | ForEach-Object { "$($_.Name) (precedence $($_.Precedence))" }) -join '; '
        Add-Check 'Fine-Grained Password Policies configured' $st (if ($ct -gt 0) { "$ct FGPP(s): $names" } else { 'None configured' }) $rk 'Create FGPPs to enforce stricter policies for privileged accounts and service accounts independently of the default domain policy.'
    }
    catch { Add-Check 'Fine-Grained Password Policies configured' 'Info' "Query failed: $($_.Exception.Message)" '' '' }

    # ── Summary progress ──────────────────────────────────────────────
    $critical = @($rows | Where-Object { $_.Status -eq 'Critical' }).Count
    $warning  = @($rows | Where-Object { $_.Status -eq 'Warning'  }).Count
    $summaryColor = if ($critical -gt 0) { 'Red' } elseif ($warning -gt 0) { 'Yellow' } else { 'Green' }
    Write-Host "Security posture analysis complete: $($rows.Count) checks | Critical: $critical | Warning: $warning" `
        -ForegroundColor $summaryColor

    $result = @($rows)
    if (-not [string]::IsNullOrWhiteSpace($ExportPath)) {
        Export-M365ReportData -InputObject $result -ExportPath $ExportPath | Out-Null
    }
    return $result
}
