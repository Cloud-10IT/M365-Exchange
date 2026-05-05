function Get-M365UnifiedGroupInventory {
    [CmdletBinding()]
    param(
        [Parameter()]
        [switch]$IncludeMembers,

        [Parameter()]
        [string]$ExportPath
    )

    Assert-ExchangeOnlineConnected

    $groups = Get-M365GraphCollection -Uri 'https://graph.microsoft.com/v1.0/groups?$top=999&$select=id,displayName,mail,groupTypes,visibility,createdDateTime' |
        Where-Object { $_.groupTypes -contains 'Unified' } |
        Sort-Object displayName

    $results = foreach ($group in $groups) {
        $members = if ($IncludeMembers) {
            $groupMembers = Get-M365GraphCollection -Uri "https://graph.microsoft.com/v1.0/groups/$($group.id)/members?`$top=999&`$select=mail,userPrincipalName"
            @($groupMembers | ForEach-Object {
                if (-not [string]::IsNullOrWhiteSpace($_.mail)) {
                    $_.mail
                }
                elseif (-not [string]::IsNullOrWhiteSpace($_.userPrincipalName)) {
                    $_.userPrincipalName
                }
            } | Where-Object { $_ })
        }
        else {
            @()
        }

        [pscustomobject]@{
            DisplayName        = $group.displayName
            PrimarySmtpAddress = $group.mail
            AccessType         = $group.visibility
            ManagedBy          = ''
            WhenCreated        = $group.createdDateTime
            MemberCount        = $members.Count
            Members            = ($members -join '; ')
        }
    }

    Export-M365ReportData -InputObject $results -ExportPath $ExportPath
}