function Get-M365DeviceInventory {
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$ExportPath,

        [Parameter()]
        [int]$StaleDaysThreshold = 90,

        [Parameter()]
        [switch]$ProblemDevicesOnly
    )

    Assert-ExchangeOnlineConnected

    Write-Host 'Fetching device inventory from Entra ID...' -ForegroundColor Cyan

    $selectFields = 'id,deviceId,displayName,operatingSystem,operatingSystemVersion,accountEnabled,approximateLastSignInDateTime,trustType,managementType,isManaged,registrationDateTime,isCompliant'
    $devices = @(Get-M365GraphCollection -Uri "/v1.0/devices?`$select=$selectFields")

    $cutoff = (Get-Date).AddDays(-$StaleDaysThreshold)

    $rows = @(
        $devices | ForEach-Object {
            $lastSignIn = if ($_.approximateLastSignInDateTime) { [datetime]$_.approximateLastSignInDateTime } else { $null }
            $regDate    = if ($_.registrationDateTime)          { [datetime]$_.registrationDateTime }          else { $null }
            $daysSince  = if ($null -ne $lastSignIn)            { [int]((Get-Date) - $lastSignIn).TotalDays }  else { $null }
            $stale      = if ($null -ne $lastSignIn)            { $lastSignIn -lt $cutoff }                    else { $null }

            $joinType = switch ([string]$_.trustType) {
                'AzureAD'   { 'Azure AD Joined' }
                'Workplace' { 'Azure AD Registered' }
                'ServerAD'  { 'Hybrid Azure AD Joined' }
                default     { [string]$_.trustType }
            }

            [pscustomobject]@{
                DisplayName      = [string]$_.displayName
                OperatingSystem  = [string]$_.operatingSystem
                OSVersion        = [string]$_.operatingSystemVersion
                JoinType         = $joinType
                IsManaged        = $_.isManaged
                IsCompliant      = $_.isCompliant
                ManagementType   = [string]$_.managementType
                AccountEnabled   = $_.accountEnabled
                LastSignInDate   = if ($null -ne $lastSignIn) { $lastSignIn.ToString('yyyy-MM-dd') } else { '' }
                DaysSinceSignIn  = $daysSince
                StaleDevice      = $stale
                RegistrationDate = if ($null -ne $regDate) { $regDate.ToString('yyyy-MM-dd') } else { '' }
                DeviceId         = [string]$_.deviceId
                EntraObjectId    = [string]$_.id
            }
        }
    )

    if ($ProblemDevicesOnly) {
        $rows = @($rows | Where-Object { $_.StaleDevice -eq $true -or $_.IsManaged -eq $false })
    }

    if (-not [string]::IsNullOrWhiteSpace($ExportPath)) {
        Export-M365ReportData -InputObject $rows -ExportPath $ExportPath | Out-Null
    }

    return $rows
}
