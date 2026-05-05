function Get-M365MailboxInventory {
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$ExportPath
    )

    Assert-ExchangeOnlineConnected

    $users = Get-M365GraphCollection -Uri 'https://graph.microsoft.com/v1.0/users?$top=999&$select=id,displayName,userPrincipalName,mail,department,createdDateTime,userType'
    $results = $users |
        Where-Object {
            $_.userType -eq 'Member' -and
            -not [string]::IsNullOrWhiteSpace($_.mail)
        } |
        Sort-Object displayName |
        Select-Object @{
            Name       = 'DisplayName'
            Expression = { $_.displayName }
        }, @{
            Name       = 'UserPrincipalName'
            Expression = { $_.userPrincipalName }
        }, @{
            Name       = 'PrimarySmtpAddress'
            Expression = { $_.mail }
        }, @{
            Name       = 'Department'
            Expression = { $_.department }
        }, @{
            Name       = 'WhenCreated'
            Expression = { $_.createdDateTime }
        }, @{
            Name       = 'MailboxKind'
            Expression = { 'User' }
        }

    Export-M365ReportData -InputObject $results -ExportPath $ExportPath
}