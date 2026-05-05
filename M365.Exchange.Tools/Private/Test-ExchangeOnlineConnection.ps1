function Test-ExchangeOnlineConnection {
    $requiredScopes = @(
        'User.Read.All',
        'Group.Read.All',
        'Directory.Read.All',
        'OrgContact.Read.All',
        'MailboxSettings.Read',
        'AuditLog.Read.All'
    )

    $getContextCommand = Get-Command -Name Get-MgContext -ErrorAction SilentlyContinue
    if (-not $getContextCommand) {
        return $false
    }

    try {
        $context = Get-MgContext
    }
    catch {
        return $false
    }

    if (-not $context) {
        return $false
    }

    if (-not ($context.Account -and $context.TenantId)) {
        return $false
    }

    $grantedScopes = @($context.Scopes)
    foreach ($scope in $requiredScopes) {
        if ($grantedScopes -notcontains $scope) {
            return $false
        }
    }

    return $true
}