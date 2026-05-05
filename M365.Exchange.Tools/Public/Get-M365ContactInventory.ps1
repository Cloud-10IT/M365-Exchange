function Get-M365ContactInventory {
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$ExportPath
    )

    Assert-ExchangeOnlineConnected

    $contacts = Get-M365GraphCollection -Uri 'https://graph.microsoft.com/v1.0/contacts?$top=999&$select=id,displayName,mail,companyName,createdDateTime'
    $results = $contacts |
        Sort-Object displayName |
        Select-Object @{
            Name       = 'DisplayName'
            Expression = { $_.displayName }
        }, @{
            Name       = 'PrimarySmtpAddress'
            Expression = { $_.mail }
        }, @{
            Name       = 'ExternalEmailAddress'
            Expression = { $_.mail }
        }, @{
            Name       = 'Alias'
            Expression = { $_.id }
        }, @{
            Name       = 'WhenCreated'
            Expression = { $_.createdDateTime }
        }

    Export-M365ReportData -InputObject $results -ExportPath $ExportPath
}
