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
        if ($null -eq $friendly) {
            return ''
        }

        $friendly = ($friendly -replace '\s+', ' ').Trim()
        if ([string]::IsNullOrWhiteSpace($friendly)) {
            return ''
        }

        $textInfo = $null
        try {
            $textInfo = (Get-Culture).TextInfo
        }
        catch {
        }

        if ($textInfo -and -not [string]::IsNullOrWhiteSpace($friendly)) {
            try {
                $friendly = $textInfo.ToTitleCase($friendly.ToLowerInvariant())
            }
            catch {
            }
        }

        # Restore common product acronyms after title-casing.
        # Use [regex]::Replace (not -replace with scriptblocks) for Windows PowerShell 5.1 compatibility.
        $friendly = [regex]::Replace($friendly, '\bM365\b', 'M365')
        $friendly = [regex]::Replace($friendly, '\bO365\b', 'O365')
        $friendly = [regex]::Replace($friendly, '\bId\b', 'ID')
        $friendly = [regex]::Replace($friendly, '\bE[135]\b', {
                param([System.Text.RegularExpressions.Match]$m)
                if ($null -eq $m) { return '' }
                return $m.Value.ToUpperInvariant()
            })
        $friendly = [regex]::Replace($friendly, '\bF[13]\b', {
                param([System.Text.RegularExpressions.Match]$m)
                if ($null -eq $m) { return '' }
                return $m.Value.ToUpperInvariant()
            })
        $friendly = [regex]::Replace($friendly, '\bP[125]\b', {
                param([System.Text.RegularExpressions.Match]$m)
                if ($null -eq $m) { return '' }
                return $m.Value.ToUpperInvariant()
            })

        return $friendly
    }

    $isGraphConnected = Test-M365GraphConnection
    $isExchangeConnected = Test-M365ExchangePowerShellConnection

    $hasSearchUnifiedAuditLog = [bool](Get-Command -Name Search-UnifiedAuditLog -ErrorAction SilentlyContinue)
    $hasSearchMailboxAuditLog = [bool](Get-Command -Name Search-MailboxAuditLog -ErrorAction SilentlyContinue)
    $canRunMailboxDeletionAuditByCmd = ($hasSearchUnifiedAuditLog -or $hasSearchMailboxAuditLog)

    $licenseStatus = 'Unknown'
    $licenseStatusDetail = ''
    $skuPartNumbers = @()
    $skuCatalog = @()
    $skuServicePlans = @()
    $hasExchangePlan2OrBetter = $false
    $hasPurviewAuditPremium = $false
    $skuProcessingErrors = [System.Collections.Generic.List[string]]::new()

    if ($isGraphConnected) {
        $getSubscribedSkuCmd = Get-Command -Name Get-MgSubscribedSku -ErrorAction SilentlyContinue
        if ($getSubscribedSkuCmd) {
            # Check whether the current Graph token includes the required scope.
            $currentScopes = @()
            try {
                $mgCtx = Get-MgContext -ErrorAction SilentlyContinue
                if ($mgCtx -and $mgCtx.Scopes) {
                    $currentScopes = @($mgCtx.Scopes)
                }
            }
            catch {
            }

            $requiredScope = 'Organization.Read.All'
            $hasScopeAccess = ($currentScopes -contains $requiredScope) -or
                              ($currentScopes -contains 'Organization.ReadWrite.All') -or
                              ($currentScopes | Where-Object { $_ -match '^Directory\.' })

            if (-not $hasScopeAccess) {
                $licenseStatus = 'ScopeMissing'
                $licenseStatusDetail = "Graph token does not include '$requiredScope'. Re-connect with that scope to view SKU data."
            }
            else {
                try {
                    $skus = @(Get-MgSubscribedSku -All -ErrorAction Stop)

                    $catalogRows = [System.Collections.Generic.List[object]]::new()
                    $servicePlanRows = [System.Collections.Generic.List[object]]::new()

                    foreach ($sku in $skus) {
                        if ($null -eq $sku) {
                            continue
                        }

                        try {
                            $skuPartNumber = [string]$sku.SkuPartNumber
                            if ([string]::IsNullOrWhiteSpace($skuPartNumber)) {
                                continue
                            }

                            $catalogRows.Add([pscustomobject]@{
                                SkuFriendlyName = ConvertTo-M365FriendlyLabel -Value $skuPartNumber
                                SkuPartNumber   = $skuPartNumber
                                SkuId           = [string]$sku.SkuId
                            })

                            $servicePlans = @()
                            if ($null -ne $sku.ServicePlans) {
                                $servicePlans = @($sku.ServicePlans)
                            }

                            foreach ($servicePlan in $servicePlans) {
                                if ($null -eq $servicePlan) {
                                    continue
                                }

                                try {
                                    $servicePlanName = ''
                                    $servicePlanId = ''
                                    $provisioningStatus = ''

                                    if ($servicePlan.PSObject.Properties.Name -contains 'ServicePlanName' -and $null -ne $servicePlan.ServicePlanName) {
                                        $servicePlanName = [string]$servicePlan.ServicePlanName
                                    }

                                    if ($servicePlan.PSObject.Properties.Name -contains 'ServicePlanId' -and $null -ne $servicePlan.ServicePlanId) {
                                        $servicePlanId = [string]$servicePlan.ServicePlanId
                                    }

                                    if ($servicePlan.PSObject.Properties.Name -contains 'ProvisioningStatus' -and $null -ne $servicePlan.ProvisioningStatus) {
                                        $provisioningStatus = [string]$servicePlan.ProvisioningStatus
                                    }

                                    if ([string]::IsNullOrWhiteSpace($servicePlanName) -and [string]::IsNullOrWhiteSpace($servicePlanId)) {
                                        continue
                                    }

                                    $servicePlanFriendlyName = ''
                                    if (-not [string]::IsNullOrWhiteSpace($servicePlanName)) {
                                        $servicePlanFriendlyName = ConvertTo-M365FriendlyLabel -Value $servicePlanName
                                    }

                                    $servicePlanRows.Add([pscustomobject]@{
                                        SkuFriendlyName         = ConvertTo-M365FriendlyLabel -Value $skuPartNumber
                                        SkuPartNumber           = $skuPartNumber
                                        ServicePlanFriendlyName = $servicePlanFriendlyName
                                        ServicePlanName         = $servicePlanName
                                        ServicePlanId           = $servicePlanId
                                        ProvisioningStatus      = $provisioningStatus
                                    })
                                }
                                catch {
                                    $skuProcessingErrors.Add("Skipped service plan under SKU '$skuPartNumber': $($_.Exception.Message)") | Out-Null
                                }
                            }
                        }
                        catch {
                            $skuLabel = ''
                            try { $skuLabel = [string]$sku.SkuPartNumber } catch {}
                            if ([string]::IsNullOrWhiteSpace($skuLabel)) { $skuLabel = '<unknown sku>' }
                            $skuProcessingErrors.Add("Skipped SKU '$skuLabel': $($_.Exception.Message)") | Out-Null
                        }
                    }

                    $skuCatalog = @($catalogRows) | Sort-Object SkuFriendlyName, SkuPartNumber -Unique
                    $skuPartNumbers = @($skuCatalog | ForEach-Object { $_.SkuPartNumber })
                    $skuServicePlans = @($servicePlanRows) | Sort-Object SkuFriendlyName, SkuPartNumber, ServicePlanFriendlyName, ServicePlanName, ServicePlanId, ProvisioningStatus -Unique

                    $plan2Indicators = @(
                        'ENTERPRISEPACK',        # Office 365 E3
                        'ENTERPRISEPREMIUM',     # Office 365 E5
                        'M365_E3',               # Microsoft 365 E3
                        'M365_E5',               # Microsoft 365 E5
                        'EXCHANGEENTERPRISE',    # Exchange Online Plan 2
                        'SPE_E3',                # Microsoft 365 E3 (new family)
                        'SPE_E5'                 # Microsoft 365 E5 (new family)
                    )

                    $purviewAuditIndicators = @(
                        'ENTERPRISEPREMIUM',                 # Office 365 E5
                        'M365_E5',                          # Microsoft 365 E5
                        'SPE_E5',                           # Microsoft 365 E5 (new family)
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
                    if ($skuProcessingErrors.Count -gt 0) {
                        if ($skuPartNumbers.Count -gt 0 -or $skuServicePlans.Count -gt 0) {
                            $licenseStatus = 'PartialData'
                        }
                        $licenseStatusDetail = ($skuProcessingErrors | Select-Object -Unique | Select-Object -First 3) -join ' | '
                    }
                }
                catch {
                    $licenseStatus = 'Error'
                    $licenseStatusDetail = $_.Exception.Message
                }
            }
        }
        else {
            $licenseStatus = 'SkuCmdletMissing'
            $licenseStatusDetail = 'Get-MgSubscribedSku cmdlet not found. Ensure Microsoft.Graph.Identity.DirectoryManagement or Microsoft.Graph is installed.'
        }
    }

    [pscustomobject]@{
        IsGraphConnected           = $isGraphConnected
        IsExchangeConnected        = $isExchangeConnected
        LicenseStatus              = $licenseStatus
        LicenseStatusDetail        = $licenseStatusDetail
        SkuPartNumbers             = $skuPartNumbers
        SkuCatalog                 = $skuCatalog
        SkuServicePlans            = $skuServicePlans
        HasExchangePlan2OrBetter   = $hasExchangePlan2OrBetter
        HasPurviewAuditPremium     = $hasPurviewAuditPremium
        HasSearchUnifiedAuditLog   = $hasSearchUnifiedAuditLog
        HasSearchMailboxAuditLog   = $hasSearchMailboxAuditLog
        CanRunMailboxDeletionAudit = ($canRunMailboxDeletionAuditByCmd -and ($isExchangeConnected -or $isGraphConnected))
    }
}