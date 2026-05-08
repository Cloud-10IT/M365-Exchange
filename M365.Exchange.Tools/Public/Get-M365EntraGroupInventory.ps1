function Get-M365EntraGroupInventory {
    [CmdletBinding()]
    param(
        [Parameter()]
        [switch]$IncludeMembers,

        [Parameter()]
        [string]$ExportPath
    )

    Assert-ExchangeOnlineConnected

    # ── 1. Groups (add renewedDateTime for last-renewed signal) ─────────────────
    Write-Host 'Retrieving Entra ID groups...' -ForegroundColor Cyan
    $groups = @(
        Get-M365GraphCollection -Uri 'https://graph.microsoft.com/v1.0/groups?$top=999&$select=id,displayName,description,mail,mailEnabled,securityEnabled,groupTypes,visibility,createdDateTime,renewedDateTime,membershipRule' |
            Sort-Object displayName
    )
    Write-Host "Retrieved $($groups.Count) group(s)." -ForegroundColor DarkCyan

    # ── 2. M365 Groups activity report (Unified groups only) ────────────────────
    # Requires Reports.Read.All. Falls back gracefully if scope missing.
    # Returns last 30-day activity: last email, SharePoint, Teams activity dates.
    $activityByGroupId  = @{}
    $activityStatus = 'Available'
    try {
        # beta endpoint returns JSON directly when Accept header is set
        $activityRows = Get-M365GraphCollection -Uri 'https://graph.microsoft.com/v1.0/reports/getOffice365GroupsActivityDetail(period=''D30'')?$format=application/json' -ErrorAction Stop
        foreach ($row in $activityRows) {
            $gid = [string]$row.groupId
            if (-not [string]::IsNullOrWhiteSpace($gid)) {
                $activityByGroupId[$gid] = $row
            }
        }
        Write-Host "Retrieved activity data for $($activityByGroupId.Count) M365 group(s)." -ForegroundColor DarkCyan
    }
    catch {
        $activityStatus = if ($_.Exception.Message -match 'Forbidden|insufficient privileges|AuthorizationRequestDenied') {
            'Unavailable (Reports.Read.All scope required)'
        } else {
            "Unavailable ($($_.Exception.Message))"
        }
        Write-Host "M365 group activity report: $activityStatus" -ForegroundColor DarkYellow
    }

    $total   = $groups.Count
    $counter = 0

    $results = foreach ($group in $groups) {
        $counter++
        Write-Progress -Activity 'Building group inventory' `
                       -Status "$counter / $total — $($group.displayName)" `
                       -PercentComplete ([int]($counter / [Math]::Max($total,1) * 100))

        # ── Members ──────────────────────────────────────────────────────────────
        $members    = @()
        $memberCount = 0
        if ($IncludeMembers) {
            $groupMembers = Get-M365GraphCollection -Uri "https://graph.microsoft.com/v1.0/groups/$($group.id)/members?`$top=999&`$select=mail,userPrincipalName"
            $members = @($groupMembers | ForEach-Object {
                if (-not [string]::IsNullOrWhiteSpace($_.mail)) { $_.mail }
                elseif (-not [string]::IsNullOrWhiteSpace($_.userPrincipalName)) { $_.userPrincipalName }
            } | Where-Object { $_ })
            $memberCount = $members.Count
        }
        else {
            # Lightweight count only — no member objects fetched
            try {
                $countResponse = Invoke-MgGraphRequest -Method GET `
                    -Uri "https://graph.microsoft.com/v1.0/groups/$($group.id)/members/`$count" `
                    -Headers @{ 'ConsistencyLevel' = 'eventual' } `
                    -OutputType PSObject -ErrorAction Stop
                $memberCount = [int]$countResponse
            }
            catch {
                $memberCount = -1
            }
        }

        # ── App role assignments (Enterprise Apps this group is assigned to) ─────
        $appAssignments = @()
        try {
            $assignments = Get-M365GraphCollection -Uri "https://graph.microsoft.com/v1.0/groups/$($group.id)/appRoleAssignments?`$select=resourceDisplayName,principalDisplayName" -ErrorAction Stop
            $appAssignments = @($assignments | ForEach-Object {
                [string]$_.resourceDisplayName
            } | Where-Object { $_ } | Sort-Object -Unique)
        }
        catch {
            # Scope or permission issue — skip silently
        }

        # ── Activity (M365/Unified groups only) ──────────────────────────────────
        $isUnified          = @($group.groupTypes) -contains 'Unified'
        $lastActivity       = $null
        $lastEmailDate      = $null
        $lastSharePointDate = $null
        $lastTeamsDate      = $null

        if ($isUnified -and $activityByGroupId.ContainsKey([string]$group.id)) {
            $act = $activityByGroupId[[string]$group.id]
            $lastEmailDate      = $act.exchangeLastActivityDate
            $lastSharePointDate = $act.sharePointLastActivityDate
            $lastTeamsDate      = $act.teamsLastActivityDate

            # Derive the most recent activity across all channels
            $actDates = @($lastEmailDate, $lastSharePointDate, $lastTeamsDate) |
                Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } |
                ForEach-Object { try { [datetime]$_ } catch { $null } } |
                Where-Object { $_ }
            if ($actDates.Count -gt 0) {
                $lastActivity = ($actDates | Sort-Object -Descending | Select-Object -First 1).ToString('yyyy-MM-ddTHH:mm:ssZ')
            }
        }

        [pscustomobject]@{
            ObjectId                    = $group.id
            DisplayName                 = $group.displayName
            Description                 = $group.description
            Mail                    = $group.mail
            MailEnabled             = $group.mailEnabled
            SecurityEnabled         = $group.securityEnabled
            GroupTypes              = (@($group.groupTypes) -join '; ')
            Visibility              = $group.visibility
            WhenCreated             = $group.createdDateTime
            RenewedDateTime         = $group.renewedDateTime
            MemberCount             = $memberCount
            Members                 = if ($IncludeMembers) { ($members -join '; ') } else { "$memberCount member(s)" }
            AppAssignmentCount         = $appAssignments.Count
            AssignedToApplications     = ($appAssignments -join '; ')
            LastActivityDate           = $lastActivity
            LastEmailActivityDate      = $lastEmailDate
            LastSharePointActivityDate = $lastSharePointDate
            LastTeamsActivityDate      = $lastTeamsDate
            ActivityDataStatus         = if ($isUnified) { $activityStatus } else { 'N/A (non-Unified group)' }
        }
    }

    Write-Progress -Activity 'Building group inventory' -Completed

    Export-M365ReportData -InputObject $results -ExportPath $ExportPath
}
