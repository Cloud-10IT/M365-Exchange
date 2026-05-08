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

    function Get-M365ConfiguredSavePath {
        [CmdletBinding()]
        param()

        $settings = Get-M365UiSettings
        $configuredPath = [string]$settings.ReportSavePath
        if ([string]::IsNullOrWhiteSpace($configuredPath)) {
            return (Join-Path -Path ([Environment]::GetFolderPath('MyDocuments')) -ChildPath 'M365-Exchange-Exports')
        }

        return $configuredPath
    }

    function New-M365ConfiguredFileStem {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)]
            [string]$Title
        )

        $settings = Get-M365UiSettings
        $template = if ([string]::IsNullOrWhiteSpace([string]$settings.FileNameTemplate)) { '{Title}-{Timestamp}' } else { [string]$settings.FileNameTemplate }
        $timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
        $dateToken = Get-Date -Format 'yyyyMMdd'
        $timeToken = Get-Date -Format 'HHmmss'

        $titleToken = ($Title -replace '[^A-Za-z0-9\-_ ]', '' -replace ' +', '-').Trim('-')
        if ([string]::IsNullOrWhiteSpace($titleToken)) {
            $titleToken = 'M365-Report'
        }

        $companyNameToken = ([string]$settings.CompanyName -replace '[^A-Za-z0-9\-_ ]', '' -replace ' +', '-').Trim('-')
        $fileStem = $template
        $fileStem = $fileStem.Replace('{Title}', $titleToken)
        $fileStem = $fileStem.Replace('{Timestamp}', $timestamp)
        $fileStem = $fileStem.Replace('{Date}', $dateToken)
        $fileStem = $fileStem.Replace('{Time}', $timeToken)
        $fileStem = $fileStem.Replace('{CompanyName}', $companyNameToken)
        $fileStem = ($fileStem -replace '[^A-Za-z0-9\-_ ]', '' -replace ' +', '-').Trim('-')

        if ([string]::IsNullOrWhiteSpace($fileStem)) {
            return "M365-Report-$timestamp"
        }

        return $fileStem
    }

    function Convert-M365HtmlToPdf {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)]
            [string]$HtmlPath,

            [Parameter(Mandatory)]
            [string]$PdfPath
        )

        if (-not (Test-Path -Path $HtmlPath)) {
            return $false
        }

        $edgePaths = @(
            'msedge.exe',
            'C:\Program Files\Microsoft\Edge\Application\msedge.exe',
            'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe'
        )

        $edgeExecutable = $null
        foreach ($edgePath in $edgePaths) {
            $edgeCommand = Get-Command -Name $edgePath -ErrorAction SilentlyContinue
            if ($edgeCommand) {
                $edgeExecutable = $edgeCommand.Source
                break
            }

            if (Test-Path -Path $edgePath) {
                $edgeExecutable = $edgePath
                break
            }
        }

        if (-not $edgeExecutable) {
            return $false
        }

        $pdfDirectory = Split-Path -Path $PdfPath -Parent
        if ($pdfDirectory -and -not (Test-Path -Path $pdfDirectory)) {
            New-Item -Path $pdfDirectory -ItemType Directory -Force | Out-Null
        }

        $htmlUri = ([System.Uri]$HtmlPath).AbsoluteUri
        $args = @(
            '--headless',
            '--disable-gpu',
            '--print-to-pdf=' + $PdfPath,
            '--no-first-run',
            '--no-default-browser-check',
            $htmlUri
        )

        try {
            $process = Start-Process -FilePath $edgeExecutable -ArgumentList $args -Wait -PassThru
            return (($process.ExitCode -eq 0) -and (Test-Path -Path $PdfPath))
        }
        catch {
            return $false
        }
    }

    function Export-M365FeatureAvailabilityBundle {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)]
            [AllowEmptyCollection()]
            [object[]]$FeatureRows,

            [Parameter()]
            [AllowEmptyCollection()]
            [object[]]$SkuRows,

            [Parameter()]
            [AllowEmptyCollection()]
            [object[]]$ServicePlanRows
        )

        $savePath = Get-M365ConfiguredSavePath
        if (-not (Test-Path -Path $savePath)) {
            New-Item -Path $savePath -ItemType Directory -Force | Out-Null
        }

        $exports = @(
            @{ Title = 'Tenant Feature Availability'; Rows = @($FeatureRows) }
            @{ Title = 'Detected Tenant SKUs'; Rows = @($SkuRows) }
            @{ Title = 'Detected Tenant SKU Service Plans'; Rows = @($ServicePlanRows) }
        )

        foreach ($export in $exports) {
            $rows = @($export.Rows)
            if ($rows.Count -eq 0) {
                continue
            }

            $stem = New-M365ConfiguredFileStem -Title $export.Title
            $csvPath = Join-Path -Path $savePath -ChildPath ($stem + '.csv')
            Export-M365ReportData -InputObject $rows -ExportPath $csvPath | Out-Null

            $htmlReport = Show-M365ReportData -InputObject $rows -Title $export.Title -ForcePopout -NoOpenBrowser -PassThru
            if ($htmlReport -and -not [string]::IsNullOrWhiteSpace([string]$htmlReport.ReportPath)) {
                $pdfPath = Join-Path -Path $savePath -ChildPath ($stem + '.pdf')
                if (Convert-M365HtmlToPdf -HtmlPath $htmlReport.ReportPath -PdfPath $pdfPath) {
                    Write-Host "Exported PDF: $pdfPath" -ForegroundColor Green
                }
                else {
                    Write-Host "PDF export skipped for '$($export.Title)' (Edge headless print unavailable)." -ForegroundColor Yellow
                }
            }
        }

        Write-Host "Feature availability bundle exported to: $savePath" -ForegroundColor Green
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

        $featureTitle = 'Tenant Feature Availability'
        $skuTitle = 'Detected Tenant SKUs (Sorted + Friendly Names)'
        $servicePlanTitle = 'Detected Tenant SKU Service Plans (Sorted + Friendly Names)'

        $skuRows = @()
        $servicePlanRows = @()

        if ($cap.SkuCatalog -and @($cap.SkuCatalog).Count -gt 0) {
            $skuRows = @(
                $cap.SkuCatalog |
                    Select-Object SkuFriendlyName, SkuPartNumber, SkuId
            )
        }
        elseif ($cap.SkuPartNumbers -and @($cap.SkuPartNumbers).Count -gt 0) {
            $skuRows = @($cap.SkuPartNumbers | ForEach-Object { [pscustomobject]@{ SkuPartNumber = $_ } })
        }

        if ($cap.SkuServicePlans -and @($cap.SkuServicePlans).Count -gt 0) {
            $servicePlanRows = @(
                $cap.SkuServicePlans |
                    Select-Object SkuFriendlyName, SkuPartNumber, ServicePlanFriendlyName, ServicePlanName, ServicePlanId, ProvisioningStatus
            )
        }

        do {
            Clear-Host
            Write-Host 'Feature Availability' -ForegroundColor Cyan
            Write-Host "License discovery status: $($cap.LicenseStatus)" -ForegroundColor ($(if ($cap.LicenseStatus -eq 'Detected') { 'Green' } else { 'Yellow' }))
            Write-Host "SKU rows: $(@($skuRows).Count)"
            Write-Host "Service plan rows: $(@($servicePlanRows).Count)"
            Write-Host ''
            Write-Host '1. View tenant feature availability report'
            Write-Host '2. View detected SKUs'
            Write-Host '3. View detected SKU service plans'
            Write-Host '4. Export feature availability bundle (CSV + PDF)'
            Write-Host 'B. Back'

            $selection = Read-Host 'Select an option'
            if ([string]::IsNullOrWhiteSpace($selection)) {
                continue
            }

            $normalizedSelection = $selection.ToUpperInvariant()

            switch ($normalizedSelection) {
                '1' {
                    Show-M365ReportData -InputObject $rows -Title $featureTitle -ForcePopout
                }
                '2' {
                    if (@($skuRows).Count -eq 0) {
                        Write-Host 'No SKU rows available to display.' -ForegroundColor Yellow
                    }
                    else {
                        Show-M365ReportData -InputObject $skuRows -Title $skuTitle -ForcePopout
                    }
                }
                '3' {
                    if (@($servicePlanRows).Count -eq 0) {
                        Write-Host 'No service plan rows available to display.' -ForegroundColor Yellow
                    }
                    else {
                        Show-M365ReportData -InputObject $servicePlanRows -Title $servicePlanTitle -ForcePopout
                    }
                }
                '4' {
                    Export-M365FeatureAvailabilityBundle -FeatureRows $rows -SkuRows $skuRows -ServicePlanRows $servicePlanRows
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
                        Connect-M365ExchangePowerShell
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

    function Show-M365WindowsConfigurationForm {
        [CmdletBinding()]
        param()

        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing

        $settings = Get-M365UiSettings

        $form = New-Object System.Windows.Forms.Form
        $form.Text = 'M365 Exchange Tools - Configuration'
        $form.Size = New-Object System.Drawing.Size(760, 640)
        $form.StartPosition = 'CenterScreen'
        $form.FormBorderStyle = 'FixedDialog'
        $form.MaximizeBox = $false
        $form.MinimizeBox = $false

        $font = New-Object System.Drawing.Font('Segoe UI', 9)

        $labels = @(
            @{ Text = 'Company Name'; Y = 20 },
            @{ Text = 'Logo Path'; Y = 60 },
            @{ Text = 'Save Path'; Y = 100 },
            @{ Text = 'File Name Template'; Y = 140 },
            @{ Text = 'Primary Color (#RRGGBB)'; Y = 180 },
            @{ Text = 'Secondary Color (#RRGGBB)'; Y = 220 },
            @{ Text = 'Font Family'; Y = 260 }
        )

        foreach ($item in $labels) {
            $label = New-Object System.Windows.Forms.Label
            $label.Text = $item.Text
            $label.Location = New-Object System.Drawing.Point(20, $item.Y)
            $label.Size = New-Object System.Drawing.Size(200, 22)
            $label.Font = $font
            $form.Controls.Add($label)
        }

        $txtCompany = New-Object System.Windows.Forms.TextBox
        $txtCompany.Location = New-Object System.Drawing.Point(230, 18)
        $txtCompany.Size = New-Object System.Drawing.Size(410, 24)
        $txtCompany.Font = $font
        $txtCompany.Text = [string]$settings.CompanyName
        $form.Controls.Add($txtCompany)

        $lblCompanyHint = New-Object System.Windows.Forms.Label
        $lblCompanyHint.Text = 'Example: Contoso Ltd'
        $lblCompanyHint.Location = New-Object System.Drawing.Point(230, 42)
        $lblCompanyHint.Size = New-Object System.Drawing.Size(280, 16)
        $lblCompanyHint.Font = New-Object System.Drawing.Font('Segoe UI', 8)
        $form.Controls.Add($lblCompanyHint)

        $txtLogo = New-Object System.Windows.Forms.TextBox
        $txtLogo.Location = New-Object System.Drawing.Point(230, 58)
        $txtLogo.Size = New-Object System.Drawing.Size(330, 24)
        $txtLogo.Font = $font
        $txtLogo.Text = [string]$settings.LogoPath
        $form.Controls.Add($txtLogo)

        $lblLogoHint = New-Object System.Windows.Forms.Label
        $lblLogoHint.Text = 'Example: C:\Branding\logo.png'
        $lblLogoHint.Location = New-Object System.Drawing.Point(230, 82)
        $lblLogoHint.Size = New-Object System.Drawing.Size(280, 16)
        $lblLogoHint.Font = New-Object System.Drawing.Font('Segoe UI', 8)
        $form.Controls.Add($lblLogoHint)

        $btnLogoBrowse = New-Object System.Windows.Forms.Button
        $btnLogoBrowse.Text = 'Browse...'
        $btnLogoBrowse.Location = New-Object System.Drawing.Point(570, 56)
        $btnLogoBrowse.Size = New-Object System.Drawing.Size(70, 28)
        $btnLogoBrowse.Font = $font
        $btnLogoBrowse.Add_Click({
            $dialog = New-Object System.Windows.Forms.OpenFileDialog
            $dialog.Filter = 'Image Files|*.png;*.jpg;*.jpeg;*.svg;*.gif;*.webp|All Files|*.*'
            if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                $txtLogo.Text = $dialog.FileName
            }
        })
        $form.Controls.Add($btnLogoBrowse)

        $txtSavePath = New-Object System.Windows.Forms.TextBox
        $txtSavePath.Location = New-Object System.Drawing.Point(230, 98)
        $txtSavePath.Size = New-Object System.Drawing.Size(330, 24)
        $txtSavePath.Font = $font
        $txtSavePath.Text = [string]$settings.ReportSavePath
        $form.Controls.Add($txtSavePath)

        $lblSavePathHint = New-Object System.Windows.Forms.Label
        $lblSavePathHint.Text = 'Example: C:\Reports\M365'
        $lblSavePathHint.Location = New-Object System.Drawing.Point(230, 122)
        $lblSavePathHint.Size = New-Object System.Drawing.Size(280, 16)
        $lblSavePathHint.Font = New-Object System.Drawing.Font('Segoe UI', 8)
        $form.Controls.Add($lblSavePathHint)

        $btnSaveBrowse = New-Object System.Windows.Forms.Button
        $btnSaveBrowse.Text = 'Browse...'
        $btnSaveBrowse.Location = New-Object System.Drawing.Point(570, 96)
        $btnSaveBrowse.Size = New-Object System.Drawing.Size(70, 28)
        $btnSaveBrowse.Font = $font
        $btnSaveBrowse.Add_Click({
            $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
            if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                $txtSavePath.Text = $dialog.SelectedPath
            }
        })
        $form.Controls.Add($btnSaveBrowse)

        $txtTemplate = New-Object System.Windows.Forms.TextBox
        $txtTemplate.Location = New-Object System.Drawing.Point(230, 138)
        $txtTemplate.Size = New-Object System.Drawing.Size(410, 24)
        $txtTemplate.Font = $font
        $txtTemplate.Text = [string]$settings.FileNameTemplate
        $form.Controls.Add($txtTemplate)

        $lblTemplateHint = New-Object System.Windows.Forms.Label
        $lblTemplateHint.Text = 'Example: {CompanyName}-{Title}-{Date}-{Time}'
        $lblTemplateHint.Location = New-Object System.Drawing.Point(230, 162)
        $lblTemplateHint.Size = New-Object System.Drawing.Size(410, 16)
        $lblTemplateHint.Font = New-Object System.Drawing.Font('Segoe UI', 8)
        $form.Controls.Add($lblTemplateHint)

        $txtPrimaryColor = New-Object System.Windows.Forms.TextBox
        $txtPrimaryColor.Location = New-Object System.Drawing.Point(230, 178)
        $txtPrimaryColor.Size = New-Object System.Drawing.Size(200, 24)
        $txtPrimaryColor.Font = $font
        $txtPrimaryColor.Text = [string]$settings.ThemePrimaryColor
        $form.Controls.Add($txtPrimaryColor)

        $lblPrimaryHint = New-Object System.Windows.Forms.Label
        $lblPrimaryHint.Text = 'Example: #0f766e'
        $lblPrimaryHint.Location = New-Object System.Drawing.Point(230, 202)
        $lblPrimaryHint.Size = New-Object System.Drawing.Size(220, 16)
        $lblPrimaryHint.Font = New-Object System.Drawing.Font('Segoe UI', 8)
        $form.Controls.Add($lblPrimaryHint)

        $btnPrimaryColor = New-Object System.Windows.Forms.Button
        $btnPrimaryColor.Text = 'Pick...'
        $btnPrimaryColor.Location = New-Object System.Drawing.Point(440, 176)
        $btnPrimaryColor.Size = New-Object System.Drawing.Size(70, 28)
        $btnPrimaryColor.Font = $font
        $btnPrimaryColor.Add_Click({
            $dialog = New-Object System.Windows.Forms.ColorDialog
            if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                $txtPrimaryColor.Text = ('#{0:X2}{1:X2}{2:X2}' -f $dialog.Color.R, $dialog.Color.G, $dialog.Color.B)
            }
        })
        $form.Controls.Add($btnPrimaryColor)

        $txtSecondaryColor = New-Object System.Windows.Forms.TextBox
        $txtSecondaryColor.Location = New-Object System.Drawing.Point(230, 218)
        $txtSecondaryColor.Size = New-Object System.Drawing.Size(200, 24)
        $txtSecondaryColor.Font = $font
        $txtSecondaryColor.Text = [string]$settings.ThemeSecondaryColor
        $form.Controls.Add($txtSecondaryColor)

        $lblSecondaryHint = New-Object System.Windows.Forms.Label
        $lblSecondaryHint.Text = 'Example: #1e293b'
        $lblSecondaryHint.Location = New-Object System.Drawing.Point(230, 242)
        $lblSecondaryHint.Size = New-Object System.Drawing.Size(220, 16)
        $lblSecondaryHint.Font = New-Object System.Drawing.Font('Segoe UI', 8)
        $form.Controls.Add($lblSecondaryHint)

        $btnSecondaryColor = New-Object System.Windows.Forms.Button
        $btnSecondaryColor.Text = 'Pick...'
        $btnSecondaryColor.Location = New-Object System.Drawing.Point(440, 216)
        $btnSecondaryColor.Size = New-Object System.Drawing.Size(70, 28)
        $btnSecondaryColor.Font = $font
        $btnSecondaryColor.Add_Click({
            $dialog = New-Object System.Windows.Forms.ColorDialog
            if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                $txtSecondaryColor.Text = ('#{0:X2}{1:X2}{2:X2}' -f $dialog.Color.R, $dialog.Color.G, $dialog.Color.B)
            }
        })
        $form.Controls.Add($btnSecondaryColor)

        $cmbFont = New-Object System.Windows.Forms.ComboBox
        $cmbFont.Location = New-Object System.Drawing.Point(230, 258)
        $cmbFont.Size = New-Object System.Drawing.Size(300, 24)
        $cmbFont.Font = $font
        $cmbFont.DropDownStyle = 'DropDown'
        $fontChoices = @('Segoe UI', 'Verdana', 'Calibri', 'Tahoma', 'Arial', 'Cambria', 'Trebuchet MS', 'Consolas')
        [void]$cmbFont.Items.AddRange($fontChoices)
        $cmbFont.Text = [string]$settings.ReportFontFamily
        $form.Controls.Add($cmbFont)

        $lblFontHint = New-Object System.Windows.Forms.Label
        $lblFontHint.Text = 'Example: Verdana'
        $lblFontHint.Location = New-Object System.Drawing.Point(230, 282)
        $lblFontHint.Size = New-Object System.Drawing.Size(220, 16)
        $lblFontHint.Font = New-Object System.Drawing.Font('Segoe UI', 8)
        $form.Controls.Add($lblFontHint)

        $chkBranding = New-Object System.Windows.Forms.CheckBox
        $chkBranding.Text = 'Enable HTML branding'
        $chkBranding.Location = New-Object System.Drawing.Point(230, 320)
        $chkBranding.Size = New-Object System.Drawing.Size(220, 24)
        $chkBranding.Font = $font
        $chkBranding.Checked = [bool]$settings.HtmlBrandingEnabled
        $form.Controls.Add($chkBranding)

        $chkShowName = New-Object System.Windows.Forms.CheckBox
        $chkShowName.Text = 'Show company name in HTML'
        $chkShowName.Location = New-Object System.Drawing.Point(230, 350)
        $chkShowName.Size = New-Object System.Drawing.Size(260, 24)
        $chkShowName.Font = $font
        $chkShowName.Checked = [bool]$settings.HtmlShowCompanyName
        $form.Controls.Add($chkShowName)

        $chkShowLogo = New-Object System.Windows.Forms.CheckBox
        $chkShowLogo.Text = 'Show company logo in HTML'
        $chkShowLogo.Location = New-Object System.Drawing.Point(230, 380)
        $chkShowLogo.Size = New-Object System.Drawing.Size(260, 24)
        $chkShowLogo.Font = $font
        $chkShowLogo.Checked = [bool]$settings.HtmlShowCompanyLogo
        $form.Controls.Add($chkShowLogo)

        $lblTokens = New-Object System.Windows.Forms.Label
        $lblTokens.Text = 'Template tokens: {Title} {Timestamp} {Date} {Time} {CompanyName}'
        $lblTokens.Location = New-Object System.Drawing.Point(230, 410)
        $lblTokens.Size = New-Object System.Drawing.Size(500, 24)
        $lblTokens.Font = $font
        $form.Controls.Add($lblTokens)

        $btnSave = New-Object System.Windows.Forms.Button
        $btnSave.Text = 'Save'
        $btnSave.Location = New-Object System.Drawing.Point(470, 510)
        $btnSave.Size = New-Object System.Drawing.Size(80, 30)
        $btnSave.Font = $font
        $btnSave.Add_Click({
            $primary = $txtPrimaryColor.Text.Trim()
            $secondary = $txtSecondaryColor.Text.Trim()
            if ($primary -notmatch '^#?[0-9A-Fa-f]{6}$') {
                [System.Windows.Forms.MessageBox]::Show('Primary color must be in #RRGGBB format.', 'Validation', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
                return
            }
            if ($secondary -notmatch '^#?[0-9A-Fa-f]{6}$') {
                [System.Windows.Forms.MessageBox]::Show('Secondary color must be in #RRGGBB format.', 'Validation', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
                return
            }

            $primary = if ($primary.StartsWith('#')) { $primary } else { "#$primary" }
            $secondary = if ($secondary.StartsWith('#')) { $secondary } else { "#$secondary" }

            Set-M365UiSettings \
                -CompanyName $txtCompany.Text \
                -LogoPath $txtLogo.Text \
                -ReportSavePath $txtSavePath.Text \
                -FileNameTemplate $txtTemplate.Text \
                -ThemePrimaryColor $primary \
                -ThemeSecondaryColor $secondary \
                -ReportFontFamily $cmbFont.Text \
                -HtmlBrandingEnabled $chkBranding.Checked \
                -HtmlShowCompanyName $chkShowName.Checked \
                -HtmlShowCompanyLogo $chkShowLogo.Checked | Out-Null

            $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $form.Close()
        })
        $form.Controls.Add($btnSave)

        $btnReset = New-Object System.Windows.Forms.Button
        $btnReset.Text = 'Reset Defaults'
        $btnReset.Location = New-Object System.Drawing.Point(350, 510)
        $btnReset.Size = New-Object System.Drawing.Size(110, 30)
        $btnReset.Font = $font
        $btnReset.Add_Click({
            $txtCompany.Text = 'Contoso'
            $txtLogo.Text = ''
            $txtSavePath.Text = ''
            $txtTemplate.Text = '{Title}-{Timestamp}'
            $txtPrimaryColor.Text = '#0f766e'
            $txtSecondaryColor.Text = '#1e293b'
            $cmbFont.Text = 'Segoe UI'
            $chkBranding.Checked = $true
            $chkShowName.Checked = $true
            $chkShowLogo.Checked = $true
        })
        $form.Controls.Add($btnReset)

        $btnCancel = New-Object System.Windows.Forms.Button
        $btnCancel.Text = 'Cancel'
        $btnCancel.Location = New-Object System.Drawing.Point(560, 510)
        $btnCancel.Size = New-Object System.Drawing.Size(80, 30)
        $btnCancel.Font = $font
        $btnCancel.Add_Click({ $form.Close() })
        $form.Controls.Add($btnCancel)

        [void]$form.ShowDialog()
    }

    function Invoke-M365SignOutOnExit {
        [CmdletBinding()]
        param()

        Write-Host 'Signing out of active module sessions...' -ForegroundColor Cyan

        try {
            $disconnectMgGraphCommand = Get-Command -Name Disconnect-MgGraph -ErrorAction SilentlyContinue
            if ($disconnectMgGraphCommand) {
                Disconnect-MgGraph -ErrorAction SilentlyContinue
            }
        }
        catch {
        }

        try {
            $disconnectExchangeCommand = Get-Command -Name Disconnect-ExchangeOnline -ErrorAction SilentlyContinue
            if ($disconnectExchangeCommand) {
                Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
            }
        }
        catch {
        }

        try {
            $disconnectAzCommand = Get-Command -Name Disconnect-AzAccount -ErrorAction SilentlyContinue
            if ($disconnectAzCommand) {
                Disconnect-AzAccount -Scope Process -ErrorAction SilentlyContinue | Out-Null
            }
        }
        catch {
        }

        try {
            Clear-History -ErrorAction SilentlyContinue
        }
        catch {
        }

        Write-Host 'Module sessions disconnected. Clearing host view...' -ForegroundColor DarkCyan
        Clear-Host
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
            Write-Host "9. Report save path: $(if ([string]::IsNullOrWhiteSpace($settings.ReportSavePath)) { '[Default: Documents\\M365-Exchange-Exports]' } else { $settings.ReportSavePath })"
            Write-Host '10. Set report save path'
            Write-Host "11. File name template: $($settings.FileNameTemplate)"
            Write-Host '12. Set file name template'
            Write-Host "13. HTML branding enabled: $($settings.HtmlBrandingEnabled)"
            Write-Host '14. Toggle HTML branding'
            Write-Host "15. Show company name in HTML: $($settings.HtmlShowCompanyName)"
            Write-Host '16. Toggle company name in HTML'
            Write-Host "17. Show company logo in HTML: $($settings.HtmlShowCompanyLogo)"
            Write-Host '18. Toggle company logo in HTML'
            Write-Host "19. Primary color: $(if ([string]::IsNullOrWhiteSpace($settings.ThemePrimaryColor)) { '#0f766e' } else { $settings.ThemePrimaryColor })"
            Write-Host "20. Secondary color: $(if ([string]::IsNullOrWhiteSpace($settings.ThemeSecondaryColor)) { '#1e293b' } else { $settings.ThemeSecondaryColor })"
            Write-Host "21. Report font: $(if ([string]::IsNullOrWhiteSpace($settings.ReportFontFamily)) { 'Segoe UI' } else { $settings.ReportFontFamily })"
            Write-Host '22. Open native Windows settings form'
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
                '9' {
                    Write-Host "Report save path is currently: $(if ([string]::IsNullOrWhiteSpace($settings.ReportSavePath)) { '[Default: Documents\\M365-Exchange-Exports]' } else { $settings.ReportSavePath })" -ForegroundColor Green
                }
                '10' {
                    $savePathInput = Read-Host 'Enter report save path (example: C:\Reports\M365). Leave blank to use default in Documents'
                    if ([string]::IsNullOrWhiteSpace($savePathInput)) {
                        Set-M365UiSettings -ReportSavePath '' | Out-Null
                        Write-Host 'Report save path reset to default.' -ForegroundColor Green
                    }
                    else {
                        $resolvedSavePath = [System.IO.Path]::GetFullPath($savePathInput)
                        Set-M365UiSettings -ReportSavePath $resolvedSavePath | Out-Null
                        Write-Host "Report save path saved: $resolvedSavePath" -ForegroundColor Green
                    }
                }
                '11' {
                    Write-Host "File name template is currently: $($settings.FileNameTemplate)" -ForegroundColor Green
                    Write-Host 'Supported tokens: {Title} {Timestamp} {Date} {Time} {CompanyName}' -ForegroundColor DarkCyan
                }
                '12' {
                    $templateInput = Read-Host 'Enter file name template (example: {CompanyName}-{Title}-{Date}-{Time}; tokens: {Title} {Timestamp} {Date} {Time} {CompanyName})'
                    if ([string]::IsNullOrWhiteSpace($templateInput)) {
                        Set-M365UiSettings -FileNameTemplate '{Title}-{Timestamp}' | Out-Null
                        Write-Host 'File name template reset to default.' -ForegroundColor Green
                    }
                    else {
                        Set-M365UiSettings -FileNameTemplate $templateInput | Out-Null
                        Write-Host "File name template saved: $templateInput" -ForegroundColor Green
                    }
                }
                '13' {
                    Write-Host "HTML branding enabled is currently: $($settings.HtmlBrandingEnabled)" -ForegroundColor Green
                }
                '14' {
                    $newValue = -not [bool]$settings.HtmlBrandingEnabled
                    Set-M365UiSettings -HtmlBrandingEnabled $newValue | Out-Null
                    Write-Host "HTML branding enabled set to: $newValue" -ForegroundColor Green
                }
                '15' {
                    Write-Host "Show company name in HTML is currently: $($settings.HtmlShowCompanyName)" -ForegroundColor Green
                }
                '16' {
                    $newValue = -not [bool]$settings.HtmlShowCompanyName
                    Set-M365UiSettings -HtmlShowCompanyName $newValue | Out-Null
                    Write-Host "Show company name in HTML set to: $newValue" -ForegroundColor Green
                }
                '17' {
                    Write-Host "Show company logo in HTML is currently: $($settings.HtmlShowCompanyLogo)" -ForegroundColor Green
                }
                '18' {
                    $newValue = -not [bool]$settings.HtmlShowCompanyLogo
                    Set-M365UiSettings -HtmlShowCompanyLogo $newValue | Out-Null
                    Write-Host "Show company logo in HTML set to: $newValue" -ForegroundColor Green
                }
                '19' {
                    Write-Host "Primary color is currently: $(if ([string]::IsNullOrWhiteSpace($settings.ThemePrimaryColor)) { '#0f766e' } else { $settings.ThemePrimaryColor })" -ForegroundColor Green
                }
                '20' {
                    Write-Host "Secondary color is currently: $(if ([string]::IsNullOrWhiteSpace($settings.ThemeSecondaryColor)) { '#1e293b' } else { $settings.ThemeSecondaryColor })" -ForegroundColor Green
                }
                '21' {
                    Write-Host "Report font is currently: $(if ([string]::IsNullOrWhiteSpace($settings.ReportFontFamily)) { 'Segoe UI' } else { $settings.ReportFontFamily })" -ForegroundColor Green
                }
                '22' {
                    Show-M365WindowsConfigurationForm
                    Write-Host 'Native Windows configuration form closed.' -ForegroundColor Green
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
                    Invoke-M365SignOutOnExit
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