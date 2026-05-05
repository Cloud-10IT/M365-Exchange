function Get-M365EntraDuplicateGroupUsageReport {
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$DisplayName,

        [Parameter()]
        [switch]$IncludeAzureRbac,

        [Parameter()]
        [string]$ExportPath
    )

    Assert-ExchangeOnlineConnected

    $allGroups = Get-M365GraphCollection -Uri 'https://graph.microsoft.com/v1.0/groups?$top=999&$select=id,displayName,mail,groupTypes,createdDateTime,assignedLicenses'

    $targetGroups = if ([string]::IsNullOrWhiteSpace($DisplayName)) {
        $duplicateNameKeys = $allGroups |
            Group-Object { ([string]$_.displayName).ToLowerInvariant() } |
            Where-Object { $_.Count -gt 1 } |
            Select-Object -ExpandProperty Name

        $allGroups | Where-Object { $duplicateNameKeys -contains ([string]$_.displayName).ToLowerInvariant() }
    }
    else {
        $allGroups | Where-Object { [string]$_.displayName -eq $DisplayName }
    }

    if (-not $targetGroups) {
        return Export-M365ReportData -InputObject @() -ExportPath $ExportPath
    }

    $duplicateCounts = @{}
    foreach ($nameGroup in ($targetGroups | Group-Object displayName)) {
        $duplicateCounts[$nameGroup.Name] = $nameGroup.Count
    }

    $conditionalAccessPolicies = @()
    $conditionalAccessState = 'Not requested'
    try {
        $conditionalAccessPolicies = Get-M365GraphCollection -Uri 'https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies?$top=999'
        $conditionalAccessState = 'Available'
    }
    catch {
        $conditionalAccessState = if ($_.Exception.Message -match 'Forbidden|insufficient privileges') { 'Permission denied (Policy.Read.All)' } else { 'Unavailable' }
    }

    $results = foreach ($group in $targetGroups | Sort-Object displayName, id) {
        $evidence = @()
        $lastSeenCandidates = @()
        $warningNotes = @()

        $licenseCount = @($group.assignedLicenses).Count
        if ($licenseCount -gt 0) {
            $evidence += "Group-based licenses ($licenseCount)"
        }

        $appAssignmentCount = 0
        try {
            $appAssignments = Get-M365GraphCollection -Uri "https://graph.microsoft.com/v1.0/groups/$($group.id)/appRoleAssignments?`$top=999"
            $appAssignmentCount = @($appAssignments).Count
            if ($appAssignmentCount -gt 0) {
                $appNames = @($appAssignments | Select-Object -ExpandProperty resourceDisplayName -ErrorAction SilentlyContinue | Where-Object { $_ } | Select-Object -Unique)
                $evidence += if ($appNames.Count -gt 0) {
                    "Enterprise app assignments ($appAssignmentCount): $($appNames -join ', ')"
                }
                else {
                    "Enterprise app assignments ($appAssignmentCount)"
                }

                $assignmentDates = @($appAssignments | Select-Object -ExpandProperty createdDateTime -ErrorAction SilentlyContinue | Where-Object { $_ })
                $lastSeenCandidates += $assignmentDates
            }
        }
        catch {
            if ($_.Exception.Message -match 'Forbidden|insufficient privileges') {
                $warningNotes += 'Enterprise app assignments require AppRoleAssignment.Read.All'
            }
            else {
                $warningNotes += 'Enterprise app assignment lookup failed'
            }
        }

        $directoryRoleAssignmentCount = 0
        try {
            $directoryRoleAssignments = Get-M365GraphCollection -Uri "https://graph.microsoft.com/v1.0/roleManagement/directory/roleAssignments?`$filter=principalId eq '$($group.id)'&`$top=999"
            $directoryRoleAssignmentCount = @($directoryRoleAssignments).Count
            if ($directoryRoleAssignmentCount -gt 0) {
                $evidence += "Directory role assignments ($directoryRoleAssignmentCount)"
                $directoryRoleDates = @($directoryRoleAssignments | Select-Object -ExpandProperty createdDateTime -ErrorAction SilentlyContinue | Where-Object { $_ })
                $lastSeenCandidates += $directoryRoleDates
            }
        }
        catch {
            if ($_.Exception.Message -match 'Forbidden|insufficient privileges') {
                $warningNotes += 'Directory role assignments require RoleManagement.Read.Directory'
            }
            else {
                $warningNotes += 'Directory role assignment lookup failed'
            }
        }

        $parentGroupCount = 0
        try {
            $parentGroups = Get-M365GraphCollection -Uri "https://graph.microsoft.com/v1.0/groups/$($group.id)/memberOf/microsoft.graph.group?`$top=999&`$select=id,displayName"
            $parentGroupCount = @($parentGroups).Count
            if ($parentGroupCount -gt 0) {
                $parentNames = @($parentGroups | Select-Object -ExpandProperty displayName -ErrorAction SilentlyContinue | Where-Object { $_ } | Select-Object -Unique)
                $evidence += if ($parentNames.Count -gt 0) {
                    "Nested in groups ($parentGroupCount): $($parentNames -join ', ')"
                }
                else {
                    "Nested in groups ($parentGroupCount)"
                }
            }
        }
        catch {
            if ($_.Exception.Message -match 'Forbidden|insufficient privileges') {
                $warningNotes += 'Nested group lookup requires Group.Read.All'
            }
            else {
                $warningNotes += 'Nested group lookup failed'
            }
        }

        $conditionalAccessMatches = @()
        if ($conditionalAccessState -eq 'Available') {
            $conditionalAccessMatches = @($conditionalAccessPolicies | Where-Object {
                @($_.conditions.users.includeGroups) -contains $group.id -or @($_.conditions.users.excludeGroups) -contains $group.id
            })

            if ($conditionalAccessMatches.Count -gt 0) {
                $policyNames = @($conditionalAccessMatches | Select-Object -ExpandProperty displayName | Where-Object { $_ })
                $evidence += "Conditional Access policies ($($conditionalAccessMatches.Count)): $($policyNames -join ', ')"
                $policyDates = @($conditionalAccessMatches | Select-Object -ExpandProperty modifiedDateTime -ErrorAction SilentlyContinue | Where-Object { $_ })
                $lastSeenCandidates += $policyDates
            }
        }
        elseif ($conditionalAccessState -ne 'Not requested') {
            $warningNotes += $conditionalAccessState
        }

        $azureRbacCount = 0
        $azureRbacStatus = 'Not requested'
        if ($IncludeAzureRbac) {
            $azureRbacStatus = 'Unavailable'
            $getAzRoleAssignmentCommand = Get-Command -Name Get-AzRoleAssignment -ErrorAction SilentlyContinue
            if (-not $getAzRoleAssignmentCommand) {
                $azureRbacStatus = 'Az.Resources not installed/imported'
            }
            else {
                try {
                    $azureAssignments = @(Get-AzRoleAssignment -ObjectId $group.id -ErrorAction Stop)
                    $azureRbacCount = $azureAssignments.Count
                    $azureRbacStatus = 'Available'
                    if ($azureRbacCount -gt 0) {
                        $evidence += "Azure RBAC assignments ($azureRbacCount)"
                    }
                }
                catch {
                    $azureRbacStatus = if ($_.Exception.Message -match 'Run Connect-AzAccount|No subscription|not logged in') { 'Connect-AzAccount required' } else { 'Query failed or permission denied' }
                }
            }
        }

        $lastSeen = $null
        if (@($lastSeenCandidates).Count -gt 0) {
            $parsedDates = @($lastSeenCandidates | ForEach-Object {
                try {
                    [datetime]$_
                }
                catch {
                    $null
                }
            } | Where-Object { $_ })

            if ($parsedDates.Count -gt 0) {
                $lastSeen = ($parsedDates | Sort-Object -Descending | Select-Object -First 1)
            }
        }

        $evidenceCount = @($evidence).Count
        $confidence = if ($evidenceCount -ge 2 -or $appAssignmentCount -gt 0 -or $directoryRoleAssignmentCount -gt 0 -or $azureRbacCount -gt 0) {
            'High'
        }
        elseif ($evidenceCount -eq 1) {
            'Medium'
        }
        else {
            'Low'
        }

        [pscustomobject]@{
            DisplayName         = $group.displayName
            DuplicateCount      = $duplicateCounts[$group.displayName]
            GroupId             = $group.id
            Mail                = $group.mail
            GroupTypes          = (@($group.groupTypes) -join '; ')
            CreatedDateTime     = $group.createdDateTime
            EvidenceCount       = $evidenceCount
            WhereUsed           = if ($evidenceCount -gt 0) { $evidence -join ' | ' } else { 'No usage references found in queried sources' }
            LastSeen            = $lastSeen
            Confidence          = $confidence
            PermissionNotes     = if (@($warningNotes).Count -gt 0) { ($warningNotes | Select-Object -Unique) -join ' | ' } else { '' }
            AzureRbacStatus     = $azureRbacStatus
            AzureRbacCount      = $azureRbacCount
        }
    }

    Export-M365ReportData -InputObject $results -ExportPath $ExportPath
}
