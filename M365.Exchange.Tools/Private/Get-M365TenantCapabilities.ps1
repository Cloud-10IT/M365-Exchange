function Get-M365TenantCapabilities {
    [CmdletBinding()]
    param()

    $isGraphConnected = Test-ExchangeOnlineConnection
    $isExchangeConnected = Test-M365ExchangePowerShellConnection

    $hasSearchUnifiedAuditLog = [bool](Get-Command -Name Search-UnifiedAuditLog -ErrorAction SilentlyContinue)
    $hasSearchMailboxAuditLog = [bool](Get-Command -Name Search-MailboxAuditLog -ErrorAction SilentlyContinue)
    $canRunMailboxDeletionAuditByCmd = ($hasSearchUnifiedAuditLog -or $hasSearchMailboxAuditLog)

    $licenseStatus = 'Unknown'
    $skuPartNumbers = @()
    $hasExchangePlan2OrBetter = $false
    $hasPurviewAuditPremium = $false

    if ($isGraphConnected) {
        $getSubscribedSkuCmd = Get-Command -Name Get-MgSubscribedSku -ErrorAction SilentlyContinue
        if ($getSubscribedSkuCmd) {
            try {
                $skus = @(Get-MgSubscribedSku -All -ErrorAction Stop)
                $skuPartNumbers = @($skus | ForEach-Object { [string]$_.SkuPartNumber } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique)

                $plan2Indicators = @(
                    'ENTERPRISEPACK',            # Office 365 E3
                    'ENTERPRISEPREMIUM',         # Office 365 E5
                    'M365_E3',                   # Microsoft 365 E3
                    'M365_E5',                   # Microsoft 365 E5
                    'EXCHANGEENTERPRISE',        # Exchange Online Plan 2
                    'SPE_E3',                    # Microsoft 365 E3 (new family)
                    'SPE_E5'                     # Microsoft 365 E5 (new family)
                )

                $purviewAuditIndicators = @(
                    'ENTERPRISEPREMIUM',         # Office 365 E5
                    'M365_E5',                   # Microsoft 365 E5
                    'SPE_E5',                    # Microsoft 365 E5 (new family)
                    'MICROSOFT_365_E5_COMPLIANCE',
                    'IDENTITY_THREAT_PROTECTION'
                )

                foreach ($sku in $skuPartNumbers) {
                    if ($plan2Indicators -contains $sku) {
                        $hasExchangePlan2OrBetter = $true
                    }
                    if ($purviewAuditIndicators -contains $sku) {
                        $hasPurviewAuditPremium = $true
                    }
                }

                $licenseStatus = if ($skuPartNumbers.Count -gt 0) { 'Detected' } else { 'NoneFound' }
            }
            catch {
                $licenseStatus = 'Error'
            }
        }
        else {
            $licenseStatus = 'SkuCmdletMissing'
        }
    }

    [pscustomobject]@{
        IsGraphConnected            = $isGraphConnected
        IsExchangeConnected         = $isExchangeConnected
        LicenseStatus               = $licenseStatus
        SkuPartNumbers              = $skuPartNumbers
        HasExchangePlan2OrBetter    = $hasExchangePlan2OrBetter
        HasPurviewAuditPremium      = $hasPurviewAuditPremium
        HasSearchUnifiedAuditLog    = $hasSearchUnifiedAuditLog
        HasSearchMailboxAuditLog    = $hasSearchMailboxAuditLog
        CanRunMailboxDeletionAudit  = ($canRunMailboxDeletionAuditByCmd -and ($isExchangeConnected -or $isGraphConnected))
    }
}
