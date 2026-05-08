function Get-M365TenantCapabilities {
    [CmdletBinding()]
    param()

    function ConvertTo-M365FriendlyLabel {
        [CmdletBinding()]
        param(
            [Parameter()]
            [string]$Value
        )

        if ([string]::IsNullOrWhiteSpace($Value)) {
            return ''
        }

        $friendly = $Value -replace '[_\-]+', ' '
        $friendly = ($friendly -replace '\s+', ' ').Trim()
        $friendly = (Get-Culture).TextInfo.ToTitleCase($friendly.ToLowerInvariant())

        # Restore common product acronyms after title-casing.
        $friendly = $friendly -replace '\bM365\b', 'M365'
        $friendly = $friendly -replace '\bO365\b', 'O365'
        $friendly = $friendly -replace '\bId\b', 'ID'
        $friendly = $friendly -replace '\bE[135]\b', { param($m) $m.Value.ToUpperInvariant() }
        $friendly = $friendly -replace '\bF[13]\b', { param($m) $m.Value.ToUpperInvariant() }
        $friendly = $friendly -replace '\bP[125]\b', { param($m) $m.Value.ToUpperInvariant() }

        return $friendly
    }

    $isGraphConnected = Test-ExchangeOnlineConnection
    $isExchangeConnected = Test-M365ExchangePowerShellConnection

    $hasSearchUnifiedAuditLog = [bool](Get-Command -Name Search-UnifiedAuditLog -ErrorAction SilentlyContinue)
    $hasSearchMailboxAuditLog = [bool](Get-Command -Name Search-MailboxAuditLog -ErrorAction SilentlyContinue)
    $canRunMailboxDeletionAuditByCmd = ($hasSearchUnifiedAuditLog -or $hasSearchMailboxAuditLog)

    $licenseStatus = 'Unknown'
    $skuPartNumbers = @()
    $skuCatalog = @()
    $skuServicePlans = @()
    $hasExchangePlan2OrBetter = $false
    $hasPurviewAuditPremium = $false

    if ($isGraphConnected) {
        $getSubscribedSkuCmd = Get-Command -Name Get-MgSubscribedSku -ErrorAction SilentlyContinue
        if ($getSubscribedSkuCmd) {
            try {
                $skus = @(Get-MgSubscribedSku -All -ErrorAction Stop)

                $skuCatalog = @(
                    $skus |
                        ForEach-Object {
                            $skuPartNumber = [string]$_.SkuPartNumber
                            if ([string]::IsNullOrWhiteSpace($skuPartNumber)) {
                                continue
                            }

                            [pscustomobject]@{
                                SkuFriendlyName = ConvertTo-M365FriendlyLabel -Value $skuPartNumber
                                SkuPartNumber   = $skuPartNumber
                                SkuId           = [string]$_.SkuId
                            }
                        } |
                        Sort-Object SkuFriendlyName, SkuPartNumber -Unique
                )

                $skuPartNumbers = @($skuCatalog | ForEach-Object { $_.SkuPartNumber })

                $skuServicePlans = @(
                    $skus |
                        ForEach-Object {
                            $skuPartNumber = [string]$_.SkuPartNumber
                            foreach ($servicePlan in @($_.ServicePlans)) {
                                $servicePlanId = [string]$servicePlan.ServicePlanId
                                if ([string]::IsNullOrWhiteSpace($servicePlanId)) {
                                    continue
                                }

                                [pscustomobject]@{
                                    SkuFriendlyName    = ConvertTo-M365FriendlyLabel -Value $skuPartNumber
                                    SkuPartNumber      = $skuPartNumber
                                    ServicePlanFriendlyName = ConvertTo-M365FriendlyLabel -Value ([string]$servicePlan.ServicePlanName)
                                    ServicePlanName    = [string]$servicePlan.ServicePlanName
                                    ServicePlanId      = $servicePlanId
                                    ProvisioningStatus = [string]$servicePlan.ProvisioningStatus
                                }
                            }
                        } |
                        Sort-Object SkuFriendlyName, SkuPartNumber, ServicePlanFriendlyName, ServicePlanName, ServicePlanId, ProvisioningStatus -Unique
                )

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
        SkuCatalog                  = $skuCatalog
        SkuServicePlans             = $skuServicePlans
        HasExchangePlan2OrBetter    = $hasExchangePlan2OrBetter
        HasPurviewAuditPremium      = $hasPurviewAuditPremium
        HasSearchUnifiedAuditLog    = $hasSearchUnifiedAuditLog
        HasSearchMailboxAuditLog    = $hasSearchMailboxAuditLog
        CanRunMailboxDeletionAudit  = ($canRunMailboxDeletionAuditByCmd -and ($isExchangeConnected -or $isGraphConnected))
    }
}
