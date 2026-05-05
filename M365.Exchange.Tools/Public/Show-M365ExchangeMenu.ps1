function Show-M365ExchangeMenu {
    [CmdletBinding()]
    param()

    function Invoke-M365DelegationReport {
        [CmdletBinding()]
        param(
            [Parameter()]
            [string]$ExportPath
        )

        $identity = Read-Host 'Optional mailbox identity filter (comma-separated, press Enter for all mailboxes)'
        $includeFolderPermissions = Read-M365YesNo -Prompt 'Include Calendar and Inbox permissions?' -Default $true
        $includeSelf = Read-M365YesNo -Prompt 'Include NT AUTHORITY\\SELF entries?' -Default $false

        if ([string]::IsNullOrWhiteSpace($identity)) {
            $reportData = Get-M365MailboxDelegationReport -IncludeFolderPermissions:$includeFolderPermissions -IncludeSelf:$includeSelf -ExportPath $ExportPath
            Show-M365ReportData -InputObject $reportData -Title 'Mailbox Delegation Report'
            return
        }

        $identities = $identity.Split(',').Trim() | Where-Object { $_ }
        $reportData = Get-M365MailboxDelegationReport -Identity $identities -IncludeFolderPermissions:$includeFolderPermissions -IncludeSelf:$includeSelf -ExportPath $ExportPath
        Show-M365ReportData -InputObject $reportData -Title 'Mailbox Delegation Report'
    }

    function Show-M365FeatureAvailability {
        [CmdletBinding()]
        param()

        $cap = Get-M365TenantCapabilities

        $rows = @(
            [pscustomobject]@{ Feature = 'Microsoft Graph connected'; Available = [bool]$cap.IsGraphConnected; Details = 'Required for Entra and Graph-based Exchange reports' }
            [pscustomobject]@{ Feature = 'Exchange PowerShell connected'; Available = [bool]$cap.IsExchangeConnected; Details = 'Required for Exchange PowerShell report features' }
            [pscustomobject]@{ Feature = 'SKU/license discovery status'; Available = ($cap.LicenseStatus -eq 'Detected'); Details = $cap.LicenseStatus }
            [pscustomobject]@{ Feature = 'Exchange Plan 2 or better (heuristic)'; Available = [bool]$cap.HasExchangePlan2OrBetter; Details = 'Used to estimate archive/audit-heavy feature coverage' }
            [pscustomobject]@{ Feature = 'Purview Audit premium (heuristic)'; Available = [bool]$cap.HasPurviewAuditPremium; Details = 'Indicates likely premium audit retention/capabilities' }
            [pscustomobject]@{ Feature = 'Advanced Entra duplicate analysis (heuristic)'; Available = [bool]$cap.HasPurviewAuditPremium; Details = 'Shown when premium audit/compliance coverage is detected' }
        )

        Show-M365ReportData -InputObject $rows -Title 'Tenant Feature Availability'

        if ($cap.SkuPartNumbers -and @($cap.SkuPartNumbers).Count -gt 0) {
            $skuRows = @($cap.SkuPartNumbers | ForEach-Object { [pscustomobject]@{ SkuPartNumber = $_ } })
            Show-M365ReportData -InputObject $skuRows -Title 'Detected Tenant SKU Part Numbers'
        }
    }

    function Show-M365ExchangeExportMenu {
        [CmdletBinding()]
        param()

        do {
            Clear-Host
            $isGraphConnected = Test-ExchangeOnlineConnection
            $isExchangeConnected = Test-M365ExchangePowerShellConnection

            Write-Host 'Export Exchange Reports to CSV' -ForegroundColor Cyan
            Write-Host '1. Export user mailbox inventory'
            Write-Host '2. Export shared mailbox inventory'
            Write-Host '3. Export resource mailbox inventory'
            Write-Host '4. Export contact inventory'
            Write-Host '5. Export distribution group inventory'
            Write-Host '6. Export M365 group inventory'
            Write-Host '7. Export mailbox delegation report (Exchange Online PowerShell)'
            Write-Host '8. Export mailbox size and archive report (Exchange Online PowerShell)'
            Write-Host 'B. Back'

            $exportSelection = Read-Host 'Select an export option'
            if ([string]::IsNullOrWhiteSpace($exportSelection)) {
                continue
            }

            $normalizedExportSelection = $exportSelection.ToUpperInvariant()

            if ((-not $isGraphConnected) -and ($normalizedExportSelection -in @('1', '2', '3', '4', '5', '6'))) {
                Write-Host 'Connect to Microsoft Graph first.' -ForegroundColor Yellow
                Read-Host 'Press Enter to continue' | Out-Null
                continue
            }

            if ((-not $isExchangeConnected) -and ($normalizedExportSelection -in @('7', '8'))) {
                Write-Host 'Connect to Exchange Online PowerShell first.' -ForegroundColor Yellow
                Read-Host 'Press Enter to continue' | Out-Null
                continue
            }

            try {
                switch ($normalizedExportSelection) {
                    '1' {
                        $exportPath = Read-M365ExportPath -ReportName 'user mailbox inventory'
                        $reportData = Get-M365MailboxInventory -ExportPath $exportPath
                        Show-M365ReportData -InputObject $reportData -Title 'User Mailbox Inventory'
                    }
                    '2' {
                        $exportPath = Read-M365ExportPath -ReportName 'shared mailbox inventory'
                        $reportData = Get-M365SharedMailboxInventory -ExportPath $exportPath
                        Show-M365ReportData -InputObject $reportData -Title 'Shared Mailbox Inventory'
                    }
                    '3' {
                        $exportPath = Read-M365ExportPath -ReportName 'resource mailbox inventory'
                        $reportData = Get-M365ResourceMailboxInventory -ExportPath $exportPath
                        Show-M365ReportData -InputObject $reportData -Title 'Resource Mailbox Inventory'
                    }
                    '4' {
                        $exportPath = Read-M365ExportPath -ReportName 'contact inventory'
                        $reportData = Get-M365ContactInventory -ExportPath $exportPath
                        Show-M365ReportData -InputObject $reportData -Title 'Contact Inventory'
                    }
                    '5' {
                        $includeMembers = Read-M365YesNo -Prompt 'Include distribution group members?' -Default $true
                        $exportPath = Read-M365ExportPath -ReportName 'distribution group inventory'
                        $reportData = Get-M365DistributionGroupInventory -IncludeMembers:$includeMembers -ExportPath $exportPath
                        Show-M365ReportData -InputObject $reportData -Title 'Distribution Group Inventory' -ExpandColumn 'Members'
                    }
                    '6' {
                        $includeMembers = Read-M365YesNo -Prompt 'Include Microsoft 365 group members?' -Default $true
                        $exportPath = Read-M365ExportPath -ReportName 'M365 group inventory'
                        $reportData = Get-M365UnifiedGroupInventory -IncludeMembers:$includeMembers -ExportPath $exportPath
                        Show-M365ReportData -InputObject $reportData -Title 'M365 Group Inventory' -ExpandColumn 'Members'
                    }
                    '7' {
                        $exportPath = Read-M365ExportPath -ReportName 'mailbox delegation report'
                        Invoke-M365DelegationReport -ExportPath $exportPath
                    }
                    '8' {
                        $exportPath = Read-M365ExportPath -ReportName 'mailbox size and archive report'
                        $includeLastEmailReceived = Read-M365YesNo -Prompt 'Include last email received date? (adds runtime)' -Default $false
                        $reportData = Get-M365MailboxSizeReport -ExportPath $exportPath -IncludeLastEmailReceived:$includeLastEmailReceived
                        Show-M365ReportData -InputObject $reportData -Title 'Mailbox Size and Archive Report' -ChartColumn 'TotalItemSizeMB'
                    }
                    'B' {
                        return
                    }
                    Default {
                        Write-Warning 'Unknown selection.'
                    }
                }
            }
            catch {
                Write-Host $_.Exception.Message -ForegroundColor Red
            }

            if ($normalizedExportSelection -ne 'B') {
                Read-Host 'Press Enter to continue' | Out-Null
            }
        }
        while ($true)
    }

    function Show-M365ExchangeReportsMenu {
        [CmdletBinding()]
        param()

        do {
            Clear-Host
            $isGraphConnected = Test-ExchangeOnlineConnection
            $isExchangeConnected = Test-M365ExchangePowerShellConnection
            $browserPopout = (Get-M365UiSettings).BrowserPopout

            Write-Host 'Exchange Reports' -ForegroundColor Cyan
            Write-Host "Graph status: $(if ($isGraphConnected) { 'Connected' } else { 'Not connected' })" -ForegroundColor ($(if ($isGraphConnected) { 'Green' } else { 'Yellow' }))
            Write-Host "Exchange PowerShell status: $(if ($isExchangeConnected) { 'Connected' } else { 'Not connected' })" -ForegroundColor ($(if ($isExchangeConnected) { 'Green' } else { 'Yellow' }))

            if (-not $isGraphConnected) {
                Write-Host '1. Connect Microsoft Graph' -ForegroundColor Green
            }
            if (-not $isExchangeConnected) {
                Write-Host '2. Connect Exchange Online PowerShell (delegation report)' -ForegroundColor Green
            }
            Write-Host '3. List user mailboxes' -ForegroundColor ($(if ($isGraphConnected) { 'White' } else { 'Gray' }))
            Write-Host '4. List shared mailboxes' -ForegroundColor ($(if ($isGraphConnected) { 'White' } else { 'Gray' }))
            Write-Host '5. List resource mailboxes' -ForegroundColor ($(if ($isGraphConnected) { 'White' } else { 'Gray' }))
            if ($browserPopout -ne 'None') {
                Write-Host '6. List contacts' -ForegroundColor ($(if ($isGraphConnected) { 'White' } else { 'Gray' }))
                Write-Host '7. List distribution groups' -ForegroundColor ($(if ($isGraphConnected) { 'White' } else { 'Gray' }))
                Write-Host '8. List M365 groups' -ForegroundColor ($(if ($isGraphConnected) { 'White' } else { 'Gray' }))
                Write-Host '9. Mailbox access and delegation report' -ForegroundColor ($(if ($isExchangeConnected) { 'White' } else { 'Gray' }))
                Write-Host '10. Mailbox size and archive report' -ForegroundColor ($(if ($isExchangeConnected) { 'White' } else { 'Gray' }))
                Write-Host '11. Export exchange reports to CSV' -ForegroundColor ($(if ($isGraphConnected -or $isExchangeConnected) { 'White' } else { 'Gray' }))
            }
            Write-Host 'B. Back'

            $selection = Read-Host 'Select an option'
            if ([string]::IsNullOrWhiteSpace($selection)) {
                continue
            }

            $normalizedSelection = $selection.ToUpperInvariant()

            if ($isGraphConnected -and ($normalizedSelection -eq '1')) {
                Write-Warning 'Unknown selection.'
                Read-Host 'Press Enter to continue' | Out-Null
                continue
            }

            if ($isExchangeConnected -and ($normalizedSelection -eq '2')) {
                Write-Warning 'Unknown selection.'
                Read-Host 'Press Enter to continue' | Out-Null
                continue
            }

            if ($browserPopout -eq 'None' -and ($normalizedSelection -in @('6', '7', '8', '9', '10', '11'))) {
                Write-Warning 'Unknown selection.'
                Read-Host 'Press Enter to continue' | Out-Null
                continue
            }

            if ((-not $isGraphConnected) -and ($normalizedSelection -in @('3', '4', '5', '6', '7', '8'))) {
                Write-Host 'Please connect to Microsoft Graph first (option 1).' -ForegroundColor Yellow
                Read-Host 'Press Enter to continue' | Out-Null
                continue
            }

            if ((-not $isExchangeConnected) -and ($normalizedSelection -in @('9', '10'))) {
                Write-Host 'Please connect to Exchange Online PowerShell first (option 2).' -ForegroundColor Yellow
                Read-Host 'Press Enter to continue' | Out-Null
                continue
            }

            try {
                switch ($normalizedSelection) {
                    '1' {
                        Connect-M365ExchangeTools
                    }
                    '2' {
                        $upn = Read-Host 'Optional UPN for Connect-ExchangeOnline (press Enter to use default sign-in)'
                        if ([string]::IsNullOrWhiteSpace($upn)) {
                            Connect-M365ExchangePowerShell
                        }
                        else {
                            Connect-M365ExchangePowerShell -UserPrincipalName $upn
                        }
                    }
                    '3' {
                        $reportData = Get-M365MailboxInventory
                        Show-M365ReportData -InputObject $reportData -Title 'User Mailbox Inventory'
                    }
                    '4' {
                        $reportData = Get-M365SharedMailboxInventory
                        Show-M365ReportData -InputObject $reportData -Title 'Shared Mailbox Inventory'
                    }
                    '5' {
                        $reportData = Get-M365ResourceMailboxInventory
                        Show-M365ReportData -InputObject $reportData -Title 'Resource Mailbox Inventory'
                    }
                    '6' {
                        $reportData = Get-M365ContactInventory
                        Show-M365ReportData -InputObject $reportData -Title 'Contact Inventory'
                    }
                    '7' {
                        $includeMembers = Read-M365YesNo -Prompt 'Include distribution group members?' -Default $true
                        $reportData = Get-M365DistributionGroupInventory -IncludeMembers:$includeMembers
                        Show-M365ReportData -InputObject $reportData -Title 'Distribution Group Inventory' -ExpandColumn 'Members'
                    }
                    '8' {
                        $includeMembers = Read-M365YesNo -Prompt 'Include Microsoft 365 group members?' -Default $true
                        $reportData = Get-M365UnifiedGroupInventory -IncludeMembers:$includeMembers
                        Show-M365ReportData -InputObject $reportData -Title 'M365 Group Inventory' -ExpandColumn 'Members'
                    }
                    '9' {
                        Invoke-M365DelegationReport
                    }
                    '10' {
                        $includeLastEmailReceived = Read-M365YesNo -Prompt 'Include last email received date? (adds runtime)' -Default $false
                        $reportData = Get-M365MailboxSizeReport -IncludeLastEmailReceived:$includeLastEmailReceived
                        Show-M365ReportData -InputObject $reportData -Title 'Mailbox Size and Archive Report' -ChartColumn 'TotalItemSizeMB'
                    }
                    '11' {
                        if ($browserPopout -eq 'None') { Write-Warning 'Unknown selection.' } else { Show-M365ExchangeExportMenu }
                    }
                    'B' {
                        return
                    }
                    Default {
                        Write-Warning 'Unknown selection.'
                    }
                }
            }
            catch {
                Write-Host $_.Exception.Message -ForegroundColor Red
            }

            if ($normalizedSelection -ne 'B') {
                Read-Host 'Press Enter to continue' | Out-Null
            }
        }
        while ($true)
    }

    function Show-M365EntraReportsMenu {
        [CmdletBinding()]
        param()

        do {
            Clear-Host
            $isGraphConnected = Test-ExchangeOnlineConnection
            $browserPopout = (Get-M365UiSettings).BrowserPopout
            $capabilities = Get-M365TenantCapabilities
            $canEntraDuplicateAnalysis = $true
            if (($capabilities.LicenseStatus -eq 'Detected') -and (-not $capabilities.HasPurviewAuditPremium)) {
                $canEntraDuplicateAnalysis = $false
            }

            Write-Host 'Entra ID Reports' -ForegroundColor Cyan
            Write-Host "Graph status: $(if ($isGraphConnected) { 'Connected' } else { 'Not connected' })" -ForegroundColor ($(if ($isGraphConnected) { 'Green' } else { 'Yellow' }))

            if (-not $isGraphConnected) {
                Write-Host '1. Connect Microsoft Graph' -ForegroundColor Green
            }
            Write-Host '2. List Entra ID users (sign-in + properties)' -ForegroundColor ($(if ($isGraphConnected) { 'White' } else { 'Gray' }))
            Write-Host '3. List Entra ID groups' -ForegroundColor ($(if ($isGraphConnected) { 'White' } else { 'Gray' }))
            if ($browserPopout -ne 'None') {
                Write-Host '4. Export Entra ID users (sign-in + properties) to CSV' -ForegroundColor ($(if ($isGraphConnected) { 'White' } else { 'Gray' }))
                Write-Host '5. Export Entra ID groups to CSV' -ForegroundColor ($(if ($isGraphConnected) { 'White' } else { 'Gray' }))
                if ($canEntraDuplicateAnalysis) {
                    Write-Host '6. Analyze duplicate group usage' -ForegroundColor ($(if ($isGraphConnected) { 'White' } else { 'Gray' }))
                    Write-Host '7. Export duplicate group usage to CSV' -ForegroundColor ($(if ($isGraphConnected) { 'White' } else { 'Gray' }))
                }
            }
            Write-Host 'B. Back'

            $selection = Read-Host 'Select an option'
            if ([string]::IsNullOrWhiteSpace($selection)) {
                continue
            }

            $normalizedSelection = $selection.ToUpperInvariant()

            if ($isGraphConnected -and ($normalizedSelection -eq '1')) {
                Write-Warning 'Unknown selection.'
                Read-Host 'Press Enter to continue' | Out-Null
                continue
            }

            if ((-not $isGraphConnected) -and ($normalizedSelection -in @('2', '3', '4', '5', '6', '7'))) {
                Write-Host 'Please connect to Microsoft Graph first (option 1).' -ForegroundColor Yellow
                Read-Host 'Press Enter to continue' | Out-Null
                continue
            }

            if ($browserPopout -eq 'None' -and ($normalizedSelection -in @('4', '5', '6', '7'))) {
                Write-Warning 'Unknown selection.'
                Read-Host 'Press Enter to continue' | Out-Null
                continue
            }

            if ((-not $canEntraDuplicateAnalysis) -and ($normalizedSelection -in @('6', '7'))) {
                Write-Host 'Duplicate group usage analysis is not available for this tenant licensing profile.' -ForegroundColor Yellow
                Read-Host 'Press Enter to continue' | Out-Null
                continue
            }

            try {
                switch ($normalizedSelection) {
                    '1' {
                        Connect-M365ExchangeTools
                    }
                    '2' {
                        $userScopeSelection = Read-Host 'User scope: A=All, M=Members, G=Guests (default: A)'
                        $userScope = switch (([string]$userScopeSelection).ToUpperInvariant()) {
                            'M' { 'Member' }
                            'G' { 'Guest' }
                            Default { 'All' }
                        }

                        $reportData = Get-M365EntraUserInventory -UserScope $userScope
                        Show-M365ReportData -InputObject $reportData -Title 'Entra ID Users'
                    }
                    '3' {
                        $includeMembers = Read-M365YesNo -Prompt 'Include group members?' -Default $false
                        $reportData = Get-M365EntraGroupInventory -IncludeMembers:$includeMembers
                        Show-M365ReportData -InputObject $reportData -Title 'Entra ID Groups' -ExpandColumn 'Members'
                    }
                    '4' {
                        $exportPath = Read-M365ExportPath -ReportName 'Entra ID users'
                        $userScopeSelection = Read-Host 'User scope: A=All, M=Members, G=Guests (default: A)'
                        $userScope = switch (([string]$userScopeSelection).ToUpperInvariant()) {
                            'M' { 'Member' }
                            'G' { 'Guest' }
                            Default { 'All' }
                        }

                        $reportData = Get-M365EntraUserInventory -UserScope $userScope -ExportPath $exportPath
                        Show-M365ReportData -InputObject $reportData -Title 'Entra ID Users'
                    }
                    '5' {
                        $includeMembers = Read-M365YesNo -Prompt 'Include group members?' -Default $false
                        $exportPath = Read-M365ExportPath -ReportName 'Entra ID groups'
                        $reportData = Get-M365EntraGroupInventory -IncludeMembers:$includeMembers -ExportPath $exportPath
                        Show-M365ReportData -InputObject $reportData -Title 'Entra ID Groups' -ExpandColumn 'Members'
                    }
                    '6' {
                        $duplicateName = Read-Host 'Optional duplicate group display name filter (press Enter to analyze all duplicate names)'
                        $includeAzureRbac = Read-M365YesNo -Prompt 'Include Azure RBAC assignment checks? (requires Az.Resources + Connect-AzAccount)' -Default $false
                        if ([string]::IsNullOrWhiteSpace($duplicateName)) {
                            $reportData = Get-M365EntraDuplicateGroupUsageReport -IncludeAzureRbac:$includeAzureRbac
                        }
                        else {
                            $reportData = Get-M365EntraDuplicateGroupUsageReport -DisplayName $duplicateName -IncludeAzureRbac:$includeAzureRbac
                        }

                        Show-M365ReportData -InputObject $reportData -Title 'Entra Duplicate Group Usage'
                    }
                    '7' {
                        $duplicateName = Read-Host 'Optional duplicate group display name filter (press Enter to analyze all duplicate names)'
                        $includeAzureRbac = Read-M365YesNo -Prompt 'Include Azure RBAC assignment checks? (requires Az.Resources + Connect-AzAccount)' -Default $false
                        $exportPath = Read-M365ExportPath -ReportName 'Entra ID duplicate group usage'
                        if ([string]::IsNullOrWhiteSpace($duplicateName)) {
                            $reportData = Get-M365EntraDuplicateGroupUsageReport -IncludeAzureRbac:$includeAzureRbac -ExportPath $exportPath
                        }
                        else {
                            $reportData = Get-M365EntraDuplicateGroupUsageReport -DisplayName $duplicateName -IncludeAzureRbac:$includeAzureRbac -ExportPath $exportPath
                        }

                        Show-M365ReportData -InputObject $reportData -Title 'Entra Duplicate Group Usage'
                    }
                    'B' {
                        return
                    }
                    Default {
                        Write-Warning 'Unknown selection.'
                    }
                }
            }
            catch {
                Write-Host $_.Exception.Message -ForegroundColor Red
            }

            if ($normalizedSelection -ne 'B') {
                Read-Host 'Press Enter to continue' | Out-Null
            }
        }
        while ($true)
    }

    function Show-M365ConfigurationMenu {
        [CmdletBinding()]
        param()

        do {
            Clear-Host
            $settings = Get-M365UiSettings

            Write-Host 'Configuration' -ForegroundColor Cyan
            Write-Host "1. Browser for reports: $($settings.BrowserPopout)"
            Write-Host '2. Select browser'
            Write-Host "3. Company name: $(if ([string]::IsNullOrWhiteSpace($settings.CompanyName)) { '[Not set]' } else { $settings.CompanyName })"
            Write-Host "4. Logo path: $(if ([string]::IsNullOrWhiteSpace($settings.LogoPath)) { '[Not set]' } else { $settings.LogoPath })"
            Write-Host '5. Set company name'
            Write-Host '6. Set logo path'
            Write-Host '7. Clear logo path'
            Write-Host '8. Preview logo'
            Write-Host 'B. Back'

            $selection = Read-Host 'Select an option'
            if ([string]::IsNullOrWhiteSpace($selection)) {
                continue
            }

            $normalizedSelection = $selection.ToUpperInvariant()
            switch ($normalizedSelection) {
                '1' {
                    Write-Host "Browser is currently set to: $($settings.BrowserPopout)" -ForegroundColor Green
                }
                '2' {
                    do {
                        Clear-Host
                        Write-Host 'Select Browser for Reports' -ForegroundColor Cyan
                        Write-Host '1. Edge'
                        Write-Host '2. Firefox'
                        Write-Host '3. Chrome'
                        Write-Host '4. Brave'
                        Write-Host '5. Default (system default browser)'
                        Write-Host '6. None (console table only)'
                        Write-Host 'B. Back'

                        $browserSelection = Read-Host 'Select a browser'
                        if ([string]::IsNullOrWhiteSpace($browserSelection)) {
                            continue
                        }

                        $normalizedBrowserSelection = $browserSelection.ToUpperInvariant()
                        $selectedBrowser = $null

                        switch ($normalizedBrowserSelection) {
                            '1' { $selectedBrowser = 'Edge' }
                            '2' { $selectedBrowser = 'Firefox' }
                            '3' { $selectedBrowser = 'Chrome' }
                            '4' { $selectedBrowser = 'Brave' }
                            '5' { $selectedBrowser = 'Default' }
                            '6' { $selectedBrowser = 'None' }
                            'B' { break }
                            Default { 
                                Write-Warning 'Unknown selection.'
                                Read-Host 'Press Enter to continue' | Out-Null
                                continue
                            }
                        }

                        if ($selectedBrowser) {
                            Set-M365UiSettings -BrowserPopout $selectedBrowser | Out-Null
                            Write-Host "Browser set to: $selectedBrowser" -ForegroundColor Green
                            Read-Host 'Press Enter to continue' | Out-Null
                            break
                        }
                    }
                    while ($true)
                }
                '3' {
                    Write-Host "Company name is currently: $(if ([string]::IsNullOrWhiteSpace($settings.CompanyName)) { '[Not set]' } else { $settings.CompanyName })" -ForegroundColor Green
                }
                '4' {
                    Write-Host "Logo path is currently: $(if ([string]::IsNullOrWhiteSpace($settings.LogoPath)) { '[Not set]' } else { $settings.LogoPath })" -ForegroundColor Green
                }
                '5' {
                    $companyName = Read-Host 'Enter company name (leave blank to clear)'
                    Set-M365UiSettings -CompanyName $companyName | Out-Null
                    Write-Host 'Company name saved.' -ForegroundColor Green
                }
                '6' {
                    $logoPathInput = Read-Host 'Enter full logo path (png/jpg/svg). Leave blank to cancel'
                    if ([string]::IsNullOrWhiteSpace($logoPathInput)) {
                        Write-Host 'Logo path update canceled.' -ForegroundColor Yellow
                    }
                    elseif (-not (Test-Path -Path $logoPathInput)) {
                        Write-Host 'Logo path not found. No changes were made.' -ForegroundColor Red
                    }
                    else {
                        $resolvedLogoPath = Resolve-Path -Path $logoPathInput | Select-Object -ExpandProperty Path -First 1
                        Set-M365UiSettings -LogoPath $resolvedLogoPath | Out-Null
                        Write-Host "Logo path saved: $resolvedLogoPath" -ForegroundColor Green
                    }
                }
                '7' {
                    Set-M365UiSettings -LogoPath '' | Out-Null
                    Write-Host 'Logo path cleared.' -ForegroundColor Green
                }
                '8' {
                    if ([string]::IsNullOrWhiteSpace($settings.LogoPath)) {
                        Write-Host 'Logo path is not set.' -ForegroundColor Yellow
                    }
                    elseif (-not (Test-Path -Path $settings.LogoPath)) {
                        Write-Host "Logo path not found: $($settings.LogoPath)" -ForegroundColor Red
                    }
                    else {
                        Start-Process -FilePath $settings.LogoPath
                        Write-Host 'Opened logo preview.' -ForegroundColor Green
                    }
                }
                'B' {
                    return
                }
                Default {
                    Write-Warning 'Unknown selection.'
                }
            }

            if ($normalizedSelection -ne 'B') {
                Read-Host 'Press Enter to continue' | Out-Null
            }
        }
        while ($true)
    }

    do {
        Clear-Host
        $isGraphConnected = Test-ExchangeOnlineConnection
        $isExchangeConnected = Test-M365ExchangePowerShellConnection

        Write-Host 'M365 Reporting Tools' -ForegroundColor Cyan
        Write-Host "Microsoft Graph: $(if ($isGraphConnected) { 'Connected' } else { 'Not connected' })" -ForegroundColor ($(if ($isGraphConnected) { 'Green' } else { 'Yellow' }))
        Write-Host "Exchange PowerShell: $(if ($isExchangeConnected) { 'Connected' } else { 'Not connected' })" -ForegroundColor ($(if ($isExchangeConnected) { 'Green' } else { 'Yellow' }))
        Write-Host ''
        Write-Host '1. Exchange' -ForegroundColor Green
        Write-Host '2. Entra ID' -ForegroundColor Green
        Write-Host '3. Check prerequisites' -ForegroundColor Green
        Write-Host '4. Configuration' -ForegroundColor Green
        Write-Host '5. Feature availability' -ForegroundColor Green
        Write-Host 'Q. Quit'

        $selection = Read-Host 'Select an option'
        if ([string]::IsNullOrWhiteSpace($selection)) {
            continue
        }

        $normalizedSelection = $selection.ToUpperInvariant()
        $skipContinuePrompt = $false

        try {
            switch ($normalizedSelection) {
                '1' {
                    $skipContinuePrompt = $true
                    Show-M365ExchangeReportsMenu
                }
                '2' {
                    $skipContinuePrompt = $true
                    Show-M365EntraReportsMenu
                }
                '3' {
                    Show-M365Prerequisites -PromptForActions
                }
                '4' {
                    $skipContinuePrompt = $true
                    Show-M365ConfigurationMenu
                }
                '5' {
                    Show-M365FeatureAvailability
                }
                'Q' {
                    return
                }
                Default {
                    Write-Warning 'Unknown selection.'
                }
            }
        }
        catch {
            Write-Host $_.Exception.Message -ForegroundColor Red
        }

        if (($normalizedSelection -ne 'Q') -and (-not $skipContinuePrompt)) {
            Read-Host 'Press Enter to continue' | Out-Null
        }
    }
    while ($true)
}