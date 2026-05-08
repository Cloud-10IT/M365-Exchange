function Get-M365PrivilegedRoleMembers {
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$ExportPath
    )

    Assert-ExchangeOnlineConnected

    Write-Host 'Fetching active directory roles...' -ForegroundColor Cyan
    $roles = @(Get-M365GraphCollection -Uri '/v1.0/directoryRoles?$select=id,displayName,roleTemplateId')

    if ($roles.Count -eq 0) {
        Write-Warning 'No active directory roles returned.'
        return @()
    }

    $results = [System.Collections.Generic.List[object]]::new()
    $total = $roles.Count
    $i = 0

    foreach ($role in $roles) {
        $i++
        Write-Progress -Activity 'Fetching role members' -Status $role.displayName -PercentComplete (($i / $total) * 100)

        try {
            $members = @(Get-M365GraphCollection -Uri "/v1.0/directoryRoles/$($role.id)/members?`$select=id,displayName,userPrincipalName,accountEnabled,userType")
        }
        catch {
            $members = @()
        }

        foreach ($member in $members) {
            $results.Add([pscustomobject]@{
                RoleName          = [string]$role.displayName
                MemberDisplayName = [string]$member.displayName
                UserPrincipalName = [string]$member.userPrincipalName
                AccountEnabled    = $member.accountEnabled
                UserType          = [string]$member.userType
                RoleId            = [string]$role.id
                MemberId          = [string]$member.id
            })
        }
    }

    Write-Progress -Activity 'Fetching role members' -Completed

    $rows = @($results)

    if (-not [string]::IsNullOrWhiteSpace($ExportPath)) {
        Export-M365ReportData -InputObject $rows -ExportPath $ExportPath | Out-Null
    }

    return $rows
}
