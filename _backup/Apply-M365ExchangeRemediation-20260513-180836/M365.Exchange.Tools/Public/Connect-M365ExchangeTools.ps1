function Connect-M365ExchangeTools {
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$UserPrincipalName
    )

    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
        throw "Microsoft.Graph.Authentication is not installed. Run 'Install-Module Microsoft.Graph -Scope CurrentUser'."
    }

    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop | Out-Null

    if (Test-ExchangeOnlineConnection) {
        Write-Host 'Microsoft Graph connection already active.' -ForegroundColor Green
        return
    }

    $requiredScopes = @(
        'User.Read.All',
        'Group.Read.All',
        'Directory.Read.All',
        'Organization.Read.All',
        'Policy.Read.All',
        'OrgContact.Read.All',
        'MailboxSettings.Read',
        'AuditLog.Read.All'
    )

    $connectParams = @{
        Scopes   = $requiredScopes
        NoWelcome = $true
    }

    if ($UserPrincipalName) {
        Write-Host 'UserPrincipalName is not used by Microsoft Graph interactive sign-in and will be ignored.' -ForegroundColor DarkYellow
    }

    Connect-MgGraph @connectParams

    $selectProfileCommand = Get-Command -Name Select-MgProfile -ErrorAction SilentlyContinue
    if ($selectProfileCommand) {
        Select-MgProfile -Name 'v1.0' | Out-Null
    }

    if (-not (Test-ExchangeOnlineConnection)) {
        throw 'Sign-in completed but a Microsoft Graph context was not detected. Try reconnecting, then run Check prerequisites to verify module and session state.'
    }

    Write-Host 'Connected to Microsoft Graph.' -ForegroundColor Green
}