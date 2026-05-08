function Get-M365FeatureCapabilityMatrix {
    <#
    .SYNOPSIS
        Returns a feature capability matrix for the tenant based on its licensed service plans.
    .DESCRIPTION
        Cross-references the tenant's provisioned service plans (from Get-MgSubscribedSku) against
        a curated catalog of well-known M365 features. Each row shows whether the feature is
        available, which service plan provides it, and which SKU that plan belongs to.
        Inspired by the m365maps.com feature matrix concept but driven by live tenant data.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [object[]]$ServicePlans
    )

    function ConvertTo-UpperSafe {
        param(
            [Parameter()]
            [AllowNull()]
            [object]$Value
        )

        if ($null -eq $Value) {
            return ''
        }

        $text = [string]$Value
        if ([string]::IsNullOrWhiteSpace($text)) {
            return ''
        }

        return $text.ToUpperInvariant()
    }

    # ---------------------------------------------------------------------------
    # Feature catalog: maps well-known service plan names/IDs to feature metadata
    # Add new entries here as Microsoft releases new plans.
    # ServicePlanNames is an array - the FIRST matching plan found wins.
    # ---------------------------------------------------------------------------
    $featureCatalog = @(

        # ── Exchange & Mail ─────────────────────────────────────────────────────
        [pscustomobject]@{ Category = 'Exchange & Mail'; Feature = 'Exchange Online Plan 2 (100 GB)';       ServicePlanNames = @('EXCHANGE_S_ENTERPRISE','EXCHANGEENTERPRISE') }
        [pscustomobject]@{ Category = 'Exchange & Mail'; Feature = 'Exchange Online Plan 1 (50 GB)';        ServicePlanNames = @('EXCHANGE_S_STANDARD','EXCHANGE_B_STANDARD','EXCHANGE_S_ESSENTIALS') }
        [pscustomobject]@{ Category = 'Exchange & Mail'; Feature = 'Exchange Online Kiosk (2 GB)';          ServicePlanNames = @('EXCHANGE_S_DESKLESS') }
        [pscustomobject]@{ Category = 'Exchange & Mail'; Feature = 'Exchange Online Archiving';             ServicePlanNames = @('EXCHANGE_S_ARCHIVE_ADDON','EXCHANGE_S_ARCHIVE','EXCHANGE_ONLINE_ARCHIVING') }
        [pscustomobject]@{ Category = 'Exchange & Mail'; Feature = 'Exchange Online Protection (EOP)';      ServicePlanNames = @('EOP_ENTERPRISE_PREMIUM','EOP_ENTERPRISE','EXCHANGE_S_FOUNDATION') }
        [pscustomobject]@{ Category = 'Exchange & Mail'; Feature = 'Hosted Voicemail (Exchange UM)';        ServicePlanNames = @('MCOVOICECONF','MCOSTANDARD') }

        # ── Teams & Collaboration ────────────────────────────────────────────────
        [pscustomobject]@{ Category = 'Teams & Collaboration'; Feature = 'Microsoft Teams';                 ServicePlanNames = @('TEAMS1','TEAMS_FREE','MCO_TEAMS_IW') }
        [pscustomobject]@{ Category = 'Teams & Collaboration'; Feature = 'Audio Conferencing';              ServicePlanNames = @('MCOMEETADV','MCOEV_VIRTUALUSER') }
        [pscustomobject]@{ Category = 'Teams & Collaboration'; Feature = 'Teams Phone (Voice)';             ServicePlanNames = @('MCOEV','MCOEV_TELSTRA') }
        [pscustomobject]@{ Category = 'Teams & Collaboration'; Feature = 'Teams Webinars';                  ServicePlanNames = @('TEAMS_PREMIUM_WEBINAR','TEAMS_ADVANCED_MEETINGS') }
        [pscustomobject]@{ Category = 'Teams & Collaboration'; Feature = 'Microsoft Viva Engage';           ServicePlanNames = @('YAMMER_ENTERPRISE','VIVAENGAGE_CORE','YAMMER_EDU') }

        # ── SharePoint & OneDrive ────────────────────────────────────────────────
        [pscustomobject]@{ Category = 'SharePoint & OneDrive'; Feature = 'SharePoint Online Plan 2';        ServicePlanNames = @('SHAREPOINTENTERPRISE','SHAREPOINTENTERPRISE_EDU') }
        [pscustomobject]@{ Category = 'SharePoint & OneDrive'; Feature = 'SharePoint Online Plan 1';        ServicePlanNames = @('SHAREPOINTSTANDARD','SHAREPOINTDESKLESS','SHAREPOINTSTANDARD_EDU') }
        [pscustomobject]@{ Category = 'SharePoint & OneDrive'; Feature = 'OneDrive for Business Plan 2';    ServicePlanNames = @('ONEDRIVE_BASIC','ONEDRIVE_ENTERPRISE') }
        [pscustomobject]@{ Category = 'SharePoint & OneDrive'; Feature = 'Microsoft Lists';                 ServicePlanNames = @('SHAREPOINTENTERPRISE','SHAREPOINTSTANDARD') }

        # ── Office Apps ──────────────────────────────────────────────────────────
        [pscustomobject]@{ Category = 'Office Apps';    Feature = 'Microsoft 365 Apps (Desktop)';           ServicePlanNames = @('OFFICESUBSCRIPTION','OFFICESUBSCRIPTION_EDU','OFFICE_BUSINESS') }
        [pscustomobject]@{ Category = 'Office Apps';    Feature = 'Office for the Web';                     ServicePlanNames = @('SHAREPOINTWAC','SHAREPOINTWAC_EDU') }
        [pscustomobject]@{ Category = 'Office Apps';    Feature = 'Clipchamp';                              ServicePlanNames = @('CLIPCHAMP','CLIPCHAMP_BASIC') }

        # ── Identity & Access ────────────────────────────────────────────────────
        [pscustomobject]@{ Category = 'Identity & Access'; Feature = 'Entra ID Plan 2 (P2)';               ServicePlanNames = @('AAD_PREMIUM_P2') }
        [pscustomobject]@{ Category = 'Identity & Access'; Feature = 'Entra ID Plan 1 (P1)';               ServicePlanNames = @('AAD_PREMIUM','AAD_PREMIUM_GOVERNMENT') }
        [pscustomobject]@{ Category = 'Identity & Access'; Feature = 'Privileged Identity Management';     ServicePlanNames = @('PRIVILEGED_IDENTITY_MANAGEMENT') }
        [pscustomobject]@{ Category = 'Identity & Access'; Feature = 'Entra ID Protection (Risk Policies)'; ServicePlanNames = @('ADALLOM_FOR_AATP') }
        [pscustomobject]@{ Category = 'Identity & Access'; Feature = 'Conditional Access (P1)';             ServicePlanNames = @('AAD_PREMIUM','AAD_PREMIUM_P2') }
        [pscustomobject]@{ Category = 'Identity & Access'; Feature = 'Dynamic Groups';                     ServicePlanNames = @('AAD_PREMIUM','AAD_PREMIUM_P2') }
        [pscustomobject]@{ Category = 'Identity & Access'; Feature = 'Entra ID Connect (Sync)';            ServicePlanNames = @('AAD_PREMIUM','AAD_PREMIUM_P2') }

        # ── Security ─────────────────────────────────────────────────────────────
        [pscustomobject]@{ Category = 'Security'; Feature = 'Defender for Office 365 Plan 2';              ServicePlanNames = @('THREAT_INTELLIGENCE') }
        [pscustomobject]@{ Category = 'Security'; Feature = 'Defender for Office 365 Plan 1';              ServicePlanNames = @('ATP_ENTERPRISE','ATP_ENTERPRISE_GOV') }
        [pscustomobject]@{ Category = 'Security'; Feature = 'Defender for Endpoint Plan 2';                ServicePlanNames = @('WINDEFATP') }
        [pscustomobject]@{ Category = 'Security'; Feature = 'Defender for Endpoint Plan 1';                ServicePlanNames = @('DEFENDER_ENDPOINT_P1') }
        [pscustomobject]@{ Category = 'Security'; Feature = 'Defender for Identity';                       ServicePlanNames = @('ATA','AATP_SERVICE') }
        [pscustomobject]@{ Category = 'Security'; Feature = 'Defender for Cloud Apps';                     ServicePlanNames = @('ADALLOM_S_O365','ADALLOM_S_DISCOVERY','MCAS_FOUNDATION') }
        [pscustomobject]@{ Category = 'Security'; Feature = 'Microsoft Secure Score';                      ServicePlanNames = @('EXCHANGE_S_ENTERPRISE','TEAMS1','AAD_PREMIUM') }
        [pscustomobject]@{ Category = 'Security'; Feature = 'Safe Links & Safe Attachments';               ServicePlanNames = @('ATP_ENTERPRISE','THREAT_INTELLIGENCE') }

        # ── Compliance ───────────────────────────────────────────────────────────
        [pscustomobject]@{ Category = 'Compliance'; Feature = 'Audit (Premium / Advanced)';                ServicePlanNames = @('M365_ADVANCED_AUDITING','PURVIEW_AUDIT_PREMIUM') }
        [pscustomobject]@{ Category = 'Compliance'; Feature = 'Audit (Standard)';                          ServicePlanNames = @('EXCHANGE_S_ENTERPRISE','EXCHANGE_S_STANDARD') }
        [pscustomobject]@{ Category = 'Compliance'; Feature = 'eDiscovery (Premium)';                      ServicePlanNames = @('EQUIVIO_ANALYTICS','ML_CLASSIFICATION') }
        [pscustomobject]@{ Category = 'Compliance'; Feature = 'eDiscovery (Standard / Content Search)';    ServicePlanNames = @('EXCHANGE_S_ENTERPRISE','EXCHANGE_S_STANDARD') }
        [pscustomobject]@{ Category = 'Compliance'; Feature = 'Customer Lockbox';                          ServicePlanNames = @('LOCKBOX_ENTERPRISE') }
        [pscustomobject]@{ Category = 'Compliance'; Feature = 'Communication Compliance';                  ServicePlanNames = @('COMMUNICATIONS_DLP') }
        [pscustomobject]@{ Category = 'Compliance'; Feature = 'Data Loss Prevention (DLP)';               ServicePlanNames = @('DLP','EXCHANGE_S_ENTERPRISE','SHAREPOINTENTERPRISE') }
        [pscustomobject]@{ Category = 'Compliance'; Feature = 'Information Protection Plan 2 (MIP P2)';   ServicePlanNames = @('MIP_S_CLP2','RMS_S_PREMIUM2') }
        [pscustomobject]@{ Category = 'Compliance'; Feature = 'Information Protection Plan 1 (MIP P1)';   ServicePlanNames = @('MIP_S_CLP1','RMS_S_PREMIUM') }
        [pscustomobject]@{ Category = 'Compliance'; Feature = 'Azure Rights Management (RMS)';             ServicePlanNames = @('RMS_S_ENTERPRISE','RMS_S_ENTERPRISE_GOV') }
        [pscustomobject]@{ Category = 'Compliance'; Feature = 'Retention Policies & Labels';               ServicePlanNames = @('EXCHANGE_S_ENTERPRISE','SHAREPOINTENTERPRISE') }
        [pscustomobject]@{ Category = 'Compliance'; Feature = 'Records Management';                        ServicePlanNames = @('RECORDS_MANAGEMENT') }
        [pscustomobject]@{ Category = 'Compliance'; Feature = 'Privileged Access Management';              ServicePlanNames = @('PAM_ENTERPRISE') }
        [pscustomobject]@{ Category = 'Compliance'; Feature = 'Insider Risk Management';                   ServicePlanNames = @('INSIDER_RISK_MANAGEMENT') }
        [pscustomobject]@{ Category = 'Compliance'; Feature = 'Compliance Manager';                        ServicePlanNames = @('COMPLIANCE_MANAGER_PREMIUM_ASSESSMENT_ADDON','EXCHANGE_S_ENTERPRISE') }

        # ── Device Management ────────────────────────────────────────────────────
        [pscustomobject]@{ Category = 'Device Management'; Feature = 'Microsoft Intune Plan 1';             ServicePlanNames = @('INTUNE_A','INTUNE_O365','INTUNE_A_VL') }
        [pscustomobject]@{ Category = 'Device Management'; Feature = 'Intune Plan 2 / Suite';               ServicePlanNames = @('INTUNE_A_D','INTUNE_PLAN2') }
        [pscustomobject]@{ Category = 'Device Management'; Feature = 'Windows Autopilot';                   ServicePlanNames = @('INTUNE_A','AAD_PREMIUM') }
        [pscustomobject]@{ Category = 'Device Management'; Feature = 'Endpoint Analytics';                  ServicePlanNames = @('Endpoint_Analytics','INTUNE_A') }

        # ── Productivity & Apps ──────────────────────────────────────────────────
        [pscustomobject]@{ Category = 'Productivity'; Feature = 'Microsoft Forms';                          ServicePlanNames = @('FORMS_PLAN_E5','FORMS_PLAN_E3','FORMS_PLAN_E1','FORMS_PLAN_K1') }
        [pscustomobject]@{ Category = 'Productivity'; Feature = 'Microsoft Planner';                        ServicePlanNames = @('PROJECTWORKMANAGEMENT','PROJECT_O365_P2','PROJECT_O365_P1') }
        [pscustomobject]@{ Category = 'Productivity'; Feature = 'Microsoft To Do';                          ServicePlanNames = @('BPOS_S_TODO_3','BPOS_S_TODO_2','BPOS_S_TODO_1') }
        [pscustomobject]@{ Category = 'Productivity'; Feature = 'Microsoft Whiteboard';                     ServicePlanNames = @('WHITEBOARD_PLAN3','WHITEBOARD_PLAN2','WHITEBOARD_PLAN1') }
        [pscustomobject]@{ Category = 'Productivity'; Feature = 'Microsoft Loop';                           ServicePlanNames = @('LOOP_INTELLIGENCE','MICROSOFT_LOOP') }
        [pscustomobject]@{ Category = 'Productivity'; Feature = 'Microsoft Bookings';                       ServicePlanNames = @('MICROSOFTBOOKINGS','BOOKINGS') }
        [pscustomobject]@{ Category = 'Productivity'; Feature = 'Microsoft Sway';                           ServicePlanNames = @('SWAY') }
        [pscustomobject]@{ Category = 'Productivity'; Feature = 'Power Automate for M365';                  ServicePlanNames = @('FLOW_O365_P3','FLOW_O365_P2','FLOW_O365_P1','FLOW_O365_S1') }
        [pscustomobject]@{ Category = 'Productivity'; Feature = 'Power Apps for M365';                      ServicePlanNames = @('POWERAPPS_O365_P3','POWERAPPS_O365_P2','POWERAPPS_O365_P1','POWERAPPS_O365_S1') }
        [pscustomobject]@{ Category = 'Productivity'; Feature = 'Power BI Pro';                             ServicePlanNames = @('BI_AZURE_P2','POWER_BI_PRO') }
        [pscustomobject]@{ Category = 'Productivity'; Feature = 'Power BI (Free/Standard)';                 ServicePlanNames = @('BI_AZURE_P0','BI_AZURE_P1') }
        [pscustomobject]@{ Category = 'Productivity'; Feature = 'Viva Insights (Personal)';                ServicePlanNames = @('MYANALYTICS_P2','EXCHANGE_ANALYTICS') }
        [pscustomobject]@{ Category = 'Productivity'; Feature = 'Viva Learning (Basic)';                   ServicePlanNames = @('VIVA_LEARNING_SEEDED','MICROSOFT_SEARCH') }

        # ── Copilot & AI ─────────────────────────────────────────────────────────
        [pscustomobject]@{ Category = 'Copilot & AI'; Feature = 'Microsoft 365 Copilot (Premium)';          ServicePlanNames = @('M365_COPILOT','COPILOT_FOR_MICROSOFT365') }
        [pscustomobject]@{ Category = 'Copilot & AI'; Feature = 'Copilot Chat (Basic/Web)';                 ServicePlanNames = @('COPILOT_CHAT_FOUNDATION','TEAMS1','EXCHANGE_S_ENTERPRISE') }
    )

    # Build a fast lookup: servicePlanName (upper) → matching plan from tenant
    $planLookup = @{}
    foreach ($sp in @($ServicePlans)) {
        if ($null -eq $sp) {
            continue
        }

        $key = ConvertTo-UpperSafe -Value $sp.ServicePlanName
        if ([string]::IsNullOrWhiteSpace($key)) {
            continue
        }

        if (-not $planLookup.ContainsKey($key)) {
            $planLookup[$key] = $sp
        }
    }

    # Build result rows
    $rows = foreach ($entry in $featureCatalog) {
        $matchedPlan = $null
        $matchedSkuFriendly = ''
        $matchedStatus = ''

        foreach ($planName in $entry.ServicePlanNames) {
            $key = ConvertTo-UpperSafe -Value $planName
            if ([string]::IsNullOrWhiteSpace($key)) {
                continue
            }

            if ($planLookup.ContainsKey($key)) {
                $matchedPlan = $planLookup[$key]
                $matchedSkuFriendly = [string]$matchedPlan.SkuFriendlyName
                $matchedStatus = [string]$matchedPlan.ProvisioningStatus
                break
            }
        }

        $available = $null -ne $matchedPlan

        [pscustomobject]@{
            Category           = $entry.Category
            Feature            = $entry.Feature
            Available          = if ($available) { 'Yes' } else { 'No' }
            ProvisioningStatus = if ($available) { $matchedStatus } else { '' }
            IncludedInSKU      = if ($available) { $matchedSkuFriendly } else { '' }
            ServicePlan        = if ($available) { [string]$matchedPlan.ServicePlanName } else { '' }
        }
    }

    return $rows
}
