function Assert-M365GraphConnected {
    [CmdletBinding()]
    param(
        [Parameter()]
        [string[]]$RequiredScopes = @(
            'User.Read.All',
            'Group.Read.All',
            'Directory.Read.All',
            'Organization.Read.All',
            'Policy.Read.All',
            'OrgContact.Read.All',
            'MailboxSettings.Read',
            'AuditLog.Read.All'
        )
    )

    if (-not (Test-M365GraphConnection -RequiredScopes $RequiredScopes -RequireAllScopes)) {
        throw 'Not connected to Microsoft Graph with the required scopes. Run Connect-M365ExchangeTools first.'
    }
}