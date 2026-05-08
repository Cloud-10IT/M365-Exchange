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

        $featureTitle = 'M365 Tenant Feature Capability Matrix'
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

        $matrixRows = @()
        if (@($servicePlanRows).Count -gt 0) {
            $matrixRows = @(Get-M365FeatureCapabilityMatrix -ServicePlans $cap.SkuServicePlans)
        }

        do {
            Clear-Host
            Write-Host 'Feature Availability' -ForegroundColor Cyan
            $statusColor = switch ($cap.LicenseStatus) {
                'Detected'         { 'Green' }
                'PartialData'      { 'Yellow' }
                'NoneFound'        { 'Yellow' }
                default            { 'Red' }
            }
            Write-Host "License discovery status: $($cap.LicenseStatus)" -ForegroundColor $statusColor
            if (-not [string]::IsNullOrWhiteSpace([string]$cap.LicenseStatusDetail)) {
                Write-Host "  Detail: $($cap.LicenseStatusDetail)" -ForegroundColor DarkYellow
            }
            $matrixAvailable = @($matrixRows).Count -gt 0
            $matrixAvailableCount  = @($matrixRows | Where-Object { $_.Available -eq 'Yes' }).Count
            $matrixTotalCount      = @($matrixRows).Count
            Write-Host "SKU rows: $(@($skuRows).Count)  |  Service plan rows: $(@($servicePlanRows).Count)  |  Feature matrix: $matrixAvailableCount/$matrixTotalCount features available"
            Write-Host ''
            Write-Host '1. View feature capability matrix (what this tenant has access to)'
            Write-Host '2. View detected SKUs'
            Write-Host '3. View detected SKU service plans'
            Write-Host '4. Export feature availability bundle (CSV + PDF)'
            if ($cap.LicenseStatus -in @('Error', 'ScopeMissing', 'Unknown', 'SkuCmdletMissing')) {
                Write-Host 'T. Re-connect Graph with Organization.Read.All scope and retry' -ForegroundColor Yellow
            }
            Write-Host 'B. Back'

            $selection = Read-Host 'Select an option'
            if ([string]::IsNullOrWhiteSpace($selection)) {
                continue
            }

            $normalizedSelection = $selection.ToUpperInvariant()

            switch ($normalizedSelection) {
                '1' {
                    if (-not $matrixAvailable) {
                        Write-Host 'Feature matrix is unavailable — SKU/license data could not be retrieved.' -ForegroundColor Yellow
                        Write-Host 'Ensure Graph is connected with Organization.Read.All scope and use T to re-connect.' -ForegroundColor Yellow
                    }
                    else {
                        Invoke-M365ConsoleReport -Data $matrixRows -Title $featureTitle
                    }
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
                    Export-M365FeatureAvailabilityBundle -FeatureRows $matrixRows -SkuRows $skuRows -ServicePlanRows $servicePlanRows
                }
                'T' {
                    Write-Host 'Reconnecting to Microsoft Graph with Organization.Read.All scope...' -ForegroundColor Cyan
                    try {
                        Connect-MgGraph -Scopes 'Organization.Read.All','User.Read.All','Group.Read.All','Mail.ReadBasic.All','MailboxSettings.Read','Directory.Read.All' -ErrorAction Stop | Out-Null
                        Write-Host 'Reconnected. Re-fetching capability data...' -ForegroundColor Green
                    }
                    catch {
                        Write-Host "Reconnect failed: $($_.Exception.Message)" -ForegroundColor Red
                        Read-Host 'Press Enter to continue' | Out-Null
                        continue
                    }

                    $cap = Get-M365TenantCapabilities

                    $skuRows = @()
                    $servicePlanRows = @()
                    if ($cap.SkuCatalog -and @($cap.SkuCatalog).Count -gt 0) {
                        $skuRows = @($cap.SkuCatalog | Select-Object SkuFriendlyName, SkuPartNumber, SkuId)
                    }
                    elseif ($cap.SkuPartNumbers -and @($cap.SkuPartNumbers).Count -gt 0) {
                        $skuRows = @($cap.SkuPartNumbers | ForEach-Object { [pscustomobject]@{ SkuPartNumber = $_ } })
                    }
                    if ($cap.SkuServicePlans -and @($cap.SkuServicePlans).Count -gt 0) {
                        $servicePlanRows = @($cap.SkuServicePlans | Select-Object SkuFriendlyName, SkuPartNumber, ServicePlanFriendlyName, ServicePlanName, ServicePlanId, ProvisioningStatus)
                    }
                    $matrixRows = @()
                    if (@($servicePlanRows).Count -gt 0) {
                        $matrixRows = @(Get-M365FeatureCapabilityMatrix -ServicePlans $cap.SkuServicePlans)
                    }
                    continue
                }
                'B' {
                    return
                }
                Default {
                    Write-Warning 'Unknown selection.'
                }
            }

            if ($normalizedSelection -notin @('B', 'T', '1')) {
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
            Write-Host '3. List user mailboxes' -ForegroundColor ($(if ($isGraphConnected -and $isExchangeConnected) { 'White' } elseif ($isGraphConnected -or $isExchangeConnected) { 'Yellow' } else { 'Gray' }))
            Write-Host '4. List shared mailboxes' -ForegroundColor ($(if ($isExchangeConnected) { 'White' } else { 'Gray' }))
            Write-Host '5. List resource mailboxes' -ForegroundColor ($(if ($isExchangeConnected) { 'White' } else { 'Gray' }))
            if ($browserPopout -ne 'None') {
                Write-Host '6. List contacts' -ForegroundColor ($(if ($isGraphConnected) { 'White' } else { 'Gray' }))
                Write-Host '7. List distribution groups' -ForegroundColor ($(if ($isGraphConnected) { 'White' } else { 'Gray' }))
                Write-Host '8. List M365 groups' -ForegroundColor ($(if ($isGraphConnected) { 'White' } else { 'Gray' }))
                Write-Host '9. Mailbox access and delegation report' -ForegroundColor ($(if ($isExchangeConnected) { 'White' } else { 'Gray' }))
                Write-Host '10. Mailbox size and archive report' -ForegroundColor ($(if ($isExchangeConnected) { 'White' } else { 'Gray' }))
                Write-Host '11. Export exchange reports to CSV' -ForegroundColor ($(if ($isGraphConnected -or $isExchangeConnected) { 'White' } else { 'Gray' }))
            }
            $recentCount = if (Test-Path variable:script:M365ReportHistory) { $script:M365ReportHistory.Count } else { 0 }
            Write-Host "R. Recent reports$(if ($recentCount -gt 0) { " ($recentCount this session)" } else { '' })" -ForegroundColor ($(if ($recentCount -gt 0) { 'Cyan' } else { 'Gray' }))
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

            if ((-not $isGraphConnected) -and ($normalizedSelection -in @('3', '6', '7', '8'))) {
                Write-Host 'Please connect to Microsoft Graph first (option 1).' -ForegroundColor Yellow
                Read-Host 'Press Enter to continue' | Out-Null
                continue
            }

            if ((-not $isExchangeConnected) -and ($normalizedSelection -eq '3')) {
                Write-Host 'The user mailbox inventory requires Exchange Online PowerShell (option 2) for mailbox schema data.' -ForegroundColor Yellow
                Read-Host 'Press Enter to continue' | Out-Null
                continue
            }

            if ((-not $isExchangeConnected) -and ($normalizedSelection -in @('4', '5', '9', '10'))) {
                Write-Host 'Please connect to Exchange Online PowerShell first (option 2).' -ForegroundColor Yellow
                Read-Host 'Press Enter to continue' | Out-Null
                continue
            }

            $skipContinuePrompt = $false
            try {
                switch ($normalizedSelection) {
                    '1' {
                        Connect-M365ExchangeTools
                    }
                    '2' {
                        Connect-M365ExchangePowerShell
                    }
                    '3' {
                        $skipContinuePrompt = $true
                        $reportData = Get-M365MailboxInventory
                        Invoke-M365ConsoleReport -Data $reportData -Title 'User Mailbox Inventory'
                    }
                    '4' {
                        $skipContinuePrompt = $true
                        $reportData = Get-M365SharedMailboxInventory
                        Invoke-M365ConsoleReport -Data $reportData -Title 'Shared Mailbox Inventory'
                    }
                    '5' {
                        $skipContinuePrompt = $true
                        $reportData = Get-M365ResourceMailboxInventory
                        Invoke-M365ConsoleReport -Data $reportData -Title 'Resource Mailbox Inventory'
                    }
                    '6' {
                        $skipContinuePrompt = $true
                        $reportData = Get-M365ContactInventory
                        Invoke-M365ConsoleReport -Data $reportData -Title 'Contact Inventory'
                    }
                    '7' {
                        $skipContinuePrompt = $true
                        $includeMembers = Read-M365YesNo -Prompt 'Include distribution group members?' -Default $true
                        $reportData = Get-M365DistributionGroupInventory -IncludeMembers:$includeMembers
                        Invoke-M365ConsoleReport -Data $reportData -Title 'Distribution Group Inventory' -ExpandColumn 'Members'
                    }
                    '8' {
                        $skipContinuePrompt = $true
                        $includeMembers = Read-M365YesNo -Prompt 'Include Microsoft 365 group members?' -Default $true
                        $reportData = Get-M365UnifiedGroupInventory -IncludeMembers:$includeMembers
                        Invoke-M365ConsoleReport -Data $reportData -Title 'M365 Group Inventory' -ExpandColumn 'Members'
                    }
                    '9' {
                        $skipContinuePrompt = $true
                        Invoke-M365DelegationReport
                    }
                    '10' {
                        $skipContinuePrompt = $true
                        $includeLastEmailReceived = Read-M365YesNo -Prompt 'Include last email received date? (adds runtime)' -Default $false
                        $reportData = Get-M365MailboxSizeReport -IncludeLastEmailReceived:$includeLastEmailReceived
                        Invoke-M365ConsoleReport -Data $reportData -Title 'Mailbox Size and Archive Report' -ChartColumn 'TotalItemSizeMB'
                    }
                    '11' {
                        $skipContinuePrompt = $true
                        if ($browserPopout -eq 'None') { Write-Warning 'Unknown selection.' } else { Show-M365ExchangeExportMenu }
                    }
                    'R' {
                        $skipContinuePrompt = $true
                        Show-M365RecentReports
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

            if (-not $skipContinuePrompt -and $normalizedSelection -ne 'B') {
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
            Write-Host '8. Analyze AI application usage (Copilot, ChatGPT, Claude, Gemini, Grok)' -ForegroundColor ($(if ($isGraphConnected) { 'White' } else { 'Gray' }))
            Write-Host '9. Export AI application usage to CSV' -ForegroundColor ($(if ($isGraphConnected) { 'White' } else { 'Gray' }))
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

            if ((-not $isGraphConnected) -and ($normalizedSelection -in @('2', '3', '4', '5', '6', '7', '8', '9'))) {
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
                    '8' {
                        $daysInput = Read-Host 'Days to review for AI app sign-in activity [default: 30]'
                        $days = 30
                        if (-not [string]::IsNullOrWhiteSpace($daysInput)) {
                            $parsedDays = 0
                            if ([int]::TryParse($daysInput, [ref]$parsedDays) -and $parsedDays -ge 1 -and $parsedDays -le 180) {
                                $days = $parsedDays
                            }
                        }

                        $reportData = Get-M365AIApplicationUsageReport -Days $days
                        Invoke-M365ConsoleReport -Data $reportData -Title 'AI Application Usage'
                    }
                    '9' {
                        $exportPath = Read-M365ExportPath -ReportName 'AI application usage'
                        $daysInput = Read-Host 'Days to review for AI app sign-in activity [default: 30]'
                        $days = 30
                        if (-not [string]::IsNullOrWhiteSpace($daysInput)) {
                            $parsedDays = 0
                            if ([int]::TryParse($daysInput, [ref]$parsedDays) -and $parsedDays -ge 1 -and $parsedDays -le 180) {
                                $days = $parsedDays
                            }
                        }

                        $reportData = Get-M365AIApplicationUsageReport -Days $days -ExportPath $exportPath
                        Invoke-M365ConsoleReport -Data $reportData -Title 'AI Application Usage'
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
        $form.Size = New-Object System.Drawing.Size(760, 690)
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
            @{ Text = 'Font Family'; Y = 260 },
            @{ Text = 'Exchange Auth Mode'; Y = 300 }
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

        $cmbAuthMode = New-Object System.Windows.Forms.ComboBox
        $cmbAuthMode.Location = New-Object System.Drawing.Point(230, 298)
        $cmbAuthMode.Size = New-Object System.Drawing.Size(220, 24)
        $cmbAuthMode.Font = $font
        $cmbAuthMode.DropDownStyle = 'DropDownList'
        [void]$cmbAuthMode.Items.AddRange(@('Auto', 'Interactive', 'DisableWAM', 'Device'))
        $cmbAuthMode.Text = [string]$settings.ExchangeAuthMode
        $form.Controls.Add($cmbAuthMode)

        $lblAuthHint = New-Object System.Windows.Forms.Label
        $lblAuthHint.Text = 'Example: DisableWAM to avoid broker/WAM issues'
        $lblAuthHint.Location = New-Object System.Drawing.Point(230, 322)
        $lblAuthHint.Size = New-Object System.Drawing.Size(340, 16)
        $lblAuthHint.Font = New-Object System.Drawing.Font('Segoe UI', 8)
        $form.Controls.Add($lblAuthHint)

        $chkBranding = New-Object System.Windows.Forms.CheckBox
        $chkBranding.Text = 'Enable HTML branding'
        $chkBranding.Location = New-Object System.Drawing.Point(230, 360)
        $chkBranding.Size = New-Object System.Drawing.Size(220, 24)
        $chkBranding.Font = $font
        $chkBranding.Checked = [bool]$settings.HtmlBrandingEnabled
        $form.Controls.Add($chkBranding)

        $chkShowName = New-Object System.Windows.Forms.CheckBox
        $chkShowName.Text = 'Show company name in HTML'
        $chkShowName.Location = New-Object System.Drawing.Point(230, 390)
        $chkShowName.Size = New-Object System.Drawing.Size(260, 24)
        $chkShowName.Font = $font
        $chkShowName.Checked = [bool]$settings.HtmlShowCompanyName
        $form.Controls.Add($chkShowName)

        $chkShowLogo = New-Object System.Windows.Forms.CheckBox
        $chkShowLogo.Text = 'Show company logo in HTML'
        $chkShowLogo.Location = New-Object System.Drawing.Point(230, 420)
        $chkShowLogo.Size = New-Object System.Drawing.Size(260, 24)
        $chkShowLogo.Font = $font
        $chkShowLogo.Checked = [bool]$settings.HtmlShowCompanyLogo
        $form.Controls.Add($chkShowLogo)

        $lblTokens = New-Object System.Windows.Forms.Label
        $lblTokens.Text = 'Template tokens: {Title} {Timestamp} {Date} {Time} {CompanyName}'
        $lblTokens.Location = New-Object System.Drawing.Point(230, 450)
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

            $saveParams = @{
                CompanyName         = $txtCompany.Text
                LogoPath            = $txtLogo.Text
                ReportSavePath      = $txtSavePath.Text
                FileNameTemplate    = $txtTemplate.Text
                ThemePrimaryColor   = $primary
                ThemeSecondaryColor = $secondary
                ReportFontFamily    = $cmbFont.Text
                ExchangeAuthMode    = $cmbAuthMode.Text
                HtmlBrandingEnabled = $chkBranding.Checked
                HtmlShowCompanyName = $chkShowName.Checked
                HtmlShowCompanyLogo = $chkShowLogo.Checked
            }

            Set-M365UiSettings @saveParams | Out-Null

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
            $cmbAuthMode.Text = 'Auto'
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

    function Add-M365ReportHistoryEntry {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)]
            [string]$Title,

            [Parameter(Mandatory)]
            [string]$Path
        )

        if (-not (Test-Path variable:script:M365ReportHistory)) {
            $script:M365ReportHistory = [System.Collections.Generic.List[object]]::new()
        }

        $script:M365ReportHistory.Add([pscustomobject]@{
            Title       = $Title
            Path        = $Path
            GeneratedAt = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        })
    }

    function Show-M365RecentReports {
        [CmdletBinding()]
        param()

        if (-not (Test-Path variable:script:M365ReportHistory) -or $script:M365ReportHistory.Count -eq 0) {
            Write-Host 'No reports generated in this session yet.' -ForegroundColor Yellow
            Read-Host 'Press Enter to continue' | Out-Null
            return
        }

        do {
            Clear-Host
            Write-Host 'Recent Reports  (this session)' -ForegroundColor Cyan
            Write-Host ''

            $index = 1
            foreach ($entry in $script:M365ReportHistory) {
                $exists = Test-Path $entry.Path
                $existsLabel = if ($exists) { '' } else { '  [file removed]' }
                $color = if ($exists) { 'White' } else { 'DarkGray' }
                Write-Host "$index. [$($entry.GeneratedAt)]  $($entry.Title)$existsLabel" -ForegroundColor $color
                Write-Host "   $($entry.Path)" -ForegroundColor DarkGray
                $index++
            }

            Write-Host ''
            Write-Host 'Enter number to re-open   C  Clear history   B  Back'
            $sel = (Read-Host 'Select').Trim()
            $selUpper = $sel.ToUpperInvariant()

            switch ($selUpper) {
                'B' { return }
                'C' {
                    $script:M365ReportHistory = [System.Collections.Generic.List[object]]::new()
                    Write-Host 'History cleared.' -ForegroundColor DarkCyan
                    Read-Host 'Press Enter to continue' | Out-Null
                    return
                }
                default {
                    $idx = 0
                    if ([int]::TryParse($sel, [ref]$idx) -and $idx -ge 1 -and $idx -le $script:M365ReportHistory.Count) {
                        $entry = $script:M365ReportHistory[$idx - 1]
                        if (Test-Path $entry.Path) {
                            Start-Process $entry.Path
                        }
                        else {
                            Write-Host 'File no longer exists on disk.' -ForegroundColor Red
                            Read-Host 'Press Enter to continue' | Out-Null
                        }
                    }
                    else {
                        Write-Host 'Invalid selection.' -ForegroundColor Yellow
                        Read-Host 'Press Enter to continue' | Out-Null
                    }
                }
            }
        }
        while ($true)
    }

    function Invoke-M365ConsoleReport {
        [CmdletBinding()]
        param(
            [Parameter()]
            [AllowNull()]
            [AllowEmptyCollection()]
            [object[]]$Data,

            [Parameter(Mandatory)]
            [string]$Title,

            [Parameter()]
            [string]$ChartColumn,

            [Parameter()]
            [string]$ExpandColumn
        )

        $rows = if ($null -eq $Data) { @() } else { @($Data) }

        if ($rows.Count -eq 0) {
            Write-Host "No records found for '$Title'." -ForegroundColor DarkYellow
            Read-Host 'Press Enter to continue' | Out-Null
            return
        }

        Clear-Host
        Write-Host "=== $Title ===" -ForegroundColor Cyan
        Write-Host "Records: $($rows.Count)" -ForegroundColor DarkCyan
        Write-Host ''
        $rows | Format-Table -AutoSize | Out-Host
        Write-Host ''
        Write-Host 'H  Open HTML report    E  Export CSV    B  Back' -ForegroundColor DarkGray

        do {
            $action = (Read-Host 'Action').Trim().ToUpperInvariant()
            switch ($action) {
                'H' {
                    $htmlParams = @{
                        InputObject = $rows
                        Title       = $Title
                        ForcePopout = $true
                        PassThru    = $true
                    }
                    if (-not [string]::IsNullOrWhiteSpace($ChartColumn))  { $htmlParams['ChartColumn']  = $ChartColumn }
                    if (-not [string]::IsNullOrWhiteSpace($ExpandColumn)) { $htmlParams['ExpandColumn'] = $ExpandColumn }
                    $report = Show-M365ReportData @htmlParams
                    if ($report -and -not [string]::IsNullOrWhiteSpace([string]$report.ReportPath)) {
                        Add-M365ReportHistoryEntry -Title $Title -Path $report.ReportPath
                    }
                    return
                }
                'E' {
                    $savePath = Get-M365ConfiguredSavePath
                    if (-not (Test-Path -Path $savePath)) {
                        New-Item -Path $savePath -ItemType Directory -Force | Out-Null
                    }
                    $stem = New-M365ConfiguredFileStem -Title $Title
                    $csvPath = Join-Path -Path $savePath -ChildPath "$stem.csv"
                    $rows | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
                    Write-Host "Exported: $csvPath" -ForegroundColor Green
                    Add-M365ReportHistoryEntry -Title $Title -Path $csvPath
                    Read-Host 'Press Enter to continue' | Out-Null
                    return
                }
                'B' { return }
                default { Write-Host 'Enter H, E, or B.' -ForegroundColor Yellow }
            }
        }
        while ($true)
    }

    function Show-M365TenantAssessmentMenu {
        [CmdletBinding()]
        param()

        do {
            Clear-Host
            $isGraphConnected    = Test-ExchangeOnlineConnection
            $isExchangeConnected = Test-M365ExchangePowerShellConnection

            Write-Host 'Tenant Assessment' -ForegroundColor Cyan
            Write-Host "Microsoft Graph: $(if ($isGraphConnected) { 'Connected' } else { 'Not connected' })" -ForegroundColor ($(if ($isGraphConnected) { 'Green' } else { 'Yellow' }))
            Write-Host "Exchange PowerShell: $(if ($isExchangeConnected) { 'Connected' } else { 'Not connected' })" -ForegroundColor ($(if ($isExchangeConnected) { 'Green' } else { 'Yellow' }))
            Write-Host ''
            Write-Host '  Identity & Access'
            Write-Host '  1. Privileged role members (Global Admin sprawl check)' -ForegroundColor ($(if ($isGraphConnected) { 'White' } else { 'Gray' }))
            Write-Host ''
            Write-Host '  Messaging Risk'
            Write-Host '  2. Mailbox forwarding rules (silent exfiltration check)' -ForegroundColor ($(if ($isExchangeConnected) { 'White' } else { 'Gray' }))
            Write-Host ''
            Write-Host '  Devices'
            Write-Host '  3. Device inventory (all devices with stale/unmanaged flags)' -ForegroundColor ($(if ($isGraphConnected) { 'White' } else { 'Gray' }))
            Write-Host ''
            Write-Host '  Security Configuration'
            Write-Host '  4. Conditional Access policy check (10 key security policies)' -ForegroundColor ($(if ($isGraphConnected) { 'White' } else { 'Gray' }))
            Write-Host ''
            Write-Host '  AI Governance'
            Write-Host '  5. AI application usage (Copilot, ChatGPT, Claude, Gemini, Grok, DeepSeek, Qwen, ERNIE, Kimi, Doubao...)' -ForegroundColor ($(if ($isGraphConnected) { 'White' } else { 'Gray' }))
            Write-Host ''
            Write-Host '  Licensing'
            Write-Host '  6. Feature availability' -ForegroundColor Green
            Write-Host ''
            $runAllColor = if ($isGraphConnected) { 'Cyan' } else { 'Gray' }
            Write-Host 'A. Run All — Generate full assessment HTML report' -ForegroundColor $runAllColor
            Write-Host 'B. Back'

            $selection = Read-Host 'Select an option'
            if ([string]::IsNullOrWhiteSpace($selection)) {
                continue
            }

            $normalizedSelection = $selection.ToUpperInvariant()

            if ((-not $isGraphConnected) -and ($normalizedSelection -in @('1', '3', '4', '5'))) {
                Write-Host 'Connect to Microsoft Graph first (option 1 on the main menu).' -ForegroundColor Yellow
                Read-Host 'Press Enter to continue' | Out-Null
                continue
            }

            if ((-not $isExchangeConnected) -and ($normalizedSelection -eq '2')) {
                Write-Host 'Connect to Exchange Online PowerShell first (option 2 on the Exchange menu).' -ForegroundColor Yellow
                Read-Host 'Press Enter to continue' | Out-Null
                continue
            }

            if ((-not $isGraphConnected) -and ($normalizedSelection -eq 'A')) {
                Write-Host 'Connect to Microsoft Graph first to run the full assessment.' -ForegroundColor Yellow
                Read-Host 'Press Enter to continue' | Out-Null
                continue
            }

            $skipContinuePrompt = $false

            try {
                switch ($normalizedSelection) {
                    '1' {
                        $skipContinuePrompt = $true
                        $reportData = Get-M365PrivilegedRoleMembers
                        Invoke-M365ConsoleReport -Data $reportData -Title 'Privileged Role Members'
                    }
                    '2' {
                        $skipContinuePrompt = $true
                        $reportData = Get-M365MailboxForwardingReport
                        Invoke-M365ConsoleReport -Data $reportData -Title 'Mailbox Forwarding Rules'
                    }
                    '3' {
                        $skipContinuePrompt = $true
                        $scopeSelection = Read-Host 'Scope: A=All devices, P=Problems only (stale or unmanaged) [default: P]'
                        $problemOnly = ([string]$scopeSelection).ToUpperInvariant() -ne 'A'
                        $reportTitle = if ($problemOnly) { 'Device Inventory — Problems Only' } else { 'Device Inventory — All Devices' }
                        $reportData = Get-M365DeviceInventory -ProblemDevicesOnly:$problemOnly
                        Invoke-M365ConsoleReport -Data $reportData -Title $reportTitle
                    }
                    '4' {
                        $skipContinuePrompt = $true
                        $reportData = Get-M365ConditionalAccessAnalysis
                        Invoke-M365ConsoleReport -Data $reportData -Title 'Conditional Access Policy Analysis'
                    }
                    '5' {
                        $skipContinuePrompt = $true
                        $daysInput = Read-Host 'Days to review for AI app sign-in activity [default: 30]'
                        $days = 30
                        if (-not [string]::IsNullOrWhiteSpace($daysInput)) {
                            $parsedDays = 0
                            if ([int]::TryParse($daysInput, [ref]$parsedDays) -and $parsedDays -ge 1 -and $parsedDays -le 180) {
                                $days = $parsedDays
                            }
                        }
                        $reportData = Get-M365AIApplicationUsageReport -Days $days
                        Invoke-M365ConsoleReport -Data $reportData -Title 'AI Application Usage'
                    }
                    '6' {
                        $skipContinuePrompt = $true
                        Show-M365FeatureAvailability
                    }
                    'A' {
                        $skipContinuePrompt = $true

                        $sections = [System.Collections.Generic.List[object]]::new()

                        # 1. Privileged roles
                        Write-Host '[1/6] Collecting privileged role members...' -ForegroundColor Cyan
                        $rolesData = @(); $rolesNote = ''; $rolesAvail = $isGraphConnected
                        if ($isGraphConnected) {
                            try {
                                $rolesData = @(Get-M365PrivilegedRoleMembers)
                                $adminCount = @($rolesData | Where-Object { [string]$_.RoleName -like '*Global Admin*' }).Count
                                $rolesNote = if ($adminCount -gt 4) { "Risk: $adminCount Global Administrator accounts detected — review for sprawl." }
                                             elseif ($adminCount -eq 0) { 'No Global Administrator role members found.' }
                                             else { "$adminCount Global Administrator account(s) detected." }
                            }
                            catch { $rolesNote = "Error collecting role members: $($_.Exception.Message)"; $rolesAvail = $false }
                        } else { $rolesNote = 'Microsoft Graph connection not available.' }
                        $sections.Add(@{
                            Id          = 'roles'
                            Title       = 'Identity & Access — Privileged Role Members'
                            Description = 'All active Entra ID directory role members. Focus on Global Administrator count — each additional Global Admin widens the blast radius of any account compromise.'
                            Note        = $rolesNote
                            Available   = $rolesAvail
                            Rows        = $rolesData
                        })

                        # 2. Forwarding rules
                        Write-Host '[2/6] Collecting mailbox forwarding rules...' -ForegroundColor Cyan
                        $fwdData = @(); $fwdNote = ''; $fwdAvail = $isExchangeConnected
                        if ($isExchangeConnected) {
                            try {
                                $fwdData = @(Get-M365MailboxForwardingReport)
                                $extCount = @($fwdData | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.ForwardingSmtpAddress) }).Count
                                $fwdNote = if ($extCount -gt 0) { "Risk: $extCount mailbox(es) forwarding externally via ForwardingSmtpAddress — review immediately." }
                                           elseif ($fwdData.Count -gt 0) { "$($fwdData.Count) mailbox(es) with internal forwarding configured." }
                                           else { 'No mailbox forwarding rules found — clean.' }
                            }
                            catch { $fwdNote = "Error collecting forwarding data: $($_.Exception.Message)"; $fwdAvail = $false }
                        } else { $fwdNote = 'Exchange Online PowerShell connection not available.' }
                        $sections.Add(@{
                            Id          = 'forwarding'
                            Title       = 'Messaging Risk — Mailbox Forwarding Rules'
                            Description = 'Mailboxes with active forwarding rules configured. External forwarding (ForwardingSmtpAddress) is a primary indicator of data exfiltration risk and should be reviewed as a priority.'
                            Note        = $fwdNote
                            Available   = $fwdAvail
                            Rows        = $fwdData
                        })

                        # 3. Devices
                        Write-Host '[3/6] Collecting device inventory...' -ForegroundColor Cyan
                        $devData = @(); $devNote = ''; $devAvail = $isGraphConnected
                        if ($isGraphConnected) {
                            try {
                                $devData = @(Get-M365DeviceInventory)
                                $staleCount   = @($devData | Where-Object { $_.StaleDevice -eq $true }).Count
                                $unmanagedCount = @($devData | Where-Object { $_.IsManaged -eq $false }).Count
                                $devNote = if ($staleCount -gt 0) { "$staleCount stale device(s) (no sign-in within 90 days). $unmanagedCount unmanaged device(s) detected." }
                                           else { "$($devData.Count) devices — none stale within 90 days." }
                            }
                            catch { $devNote = "Error collecting device data: $($_.Exception.Message)"; $devAvail = $false }
                        } else { $devNote = 'Microsoft Graph connection not available.' }
                        $sections.Add(@{
                            Id          = 'devices'
                            Title       = 'Devices — Entra ID Device Inventory'
                            Description = 'All devices registered in Entra ID with join type, management state, compliance state, and last sign-in activity. Stale or unmanaged devices represent uncontrolled credential exposure.'
                            Note        = $devNote
                            Available   = $devAvail
                            Rows        = $devData
                        })

                        # 4. Conditional Access analysis
                        Write-Host '[4/6] Analysing Conditional Access policies...' -ForegroundColor Cyan
                        $caData = @(); $caNote = ''; $caAvail = $isGraphConnected
                        if ($isGraphConnected) {
                            try {
                                $caData = @(Get-M365ConditionalAccessAnalysis)
                                $missingCount  = @($caData | Where-Object { $_.Status -eq 'Missing' }).Count
                                $disabledCount = @($caData | Where-Object { $_.Status -eq 'Present' -and $_.PolicyState -ne 'Enabled' }).Count
                                $caNote = if ($missingCount -gt 5) { "Risk: $missingCount of 10 key Conditional Access policies are missing from this tenant." }
                                          elseif ($missingCount -gt 0) { "Warning: $missingCount of 10 recommended CA policies not found. $disabledCount present but not enabled." }
                                          else { 'All 10 recommended Conditional Access policies are present.' }
                            }
                            catch { $caNote = "Error running CA analysis: $($_.Exception.Message)"; $caAvail = $false }
                        } else { $caNote = 'Microsoft Graph connection not available.' }
                        $sections.Add(@{
                            Id          = 'ca-analysis'
                            Title       = 'Security Configuration — Conditional Access Policy Check'
                            Description = '10-point check against the recommended Conditional Access baseline: MFA enforcement, legacy auth block, risk-based policies, session controls, and more.'
                            Note        = $caNote
                            Available   = $caAvail
                            Rows        = $caData
                        })

                        # 5. AI application usage
                        Write-Host '[5/6] Collecting AI application usage...' -ForegroundColor Cyan
                        $aiData = @(); $aiNote = ''; $aiAvail = $isGraphConnected
                        if ($isGraphConnected) {
                            try {
                                $aiData = @(Get-M365AIApplicationUsageReport -Days 30)
                                $detectedCount = @($aiData | Where-Object { $_.Status -eq 'Detected' }).Count
                                $chinaCount    = @($aiData | Where-Object { $_.Status -eq 'Detected' -and [string]$_.AIApplication -match 'China' }).Count
                                $aiNote = if ($detectedCount -gt 0 -and $chinaCount -gt 0) {
                                    "Warning: $detectedCount AI platform(s) detected, including $chinaCount Chinese-origin platform(s). Review vendor risk, data residency, and approved use immediately."
                                }
                                elseif ($detectedCount -gt 0) {
                                    "Warning: $detectedCount AI platform(s) detected in Entra telemetry. Validate approval status, OAuth exposure, and data handling boundaries."
                                }
                                else {
                                    'No matched AI application activity found in Entra sign-ins or enterprise apps for the sampled period.'
                                }
                            }
                            catch { $aiNote = "Error collecting AI app usage: $($_.Exception.Message)"; $aiAvail = $false }
                        } else { $aiNote = 'Microsoft Graph connection not available.' }
                        $sections.Add(@{
                            Id          = 'ai-apps'
                            Title       = 'AI Governance — AI Application Usage'
                            Description = 'Catalog-based detection of major AI platforms from Entra enterprise apps and sign-in logs, including Microsoft Copilot, OpenAI/ChatGPT, Anthropic Claude, Google Gemini, xAI Grok, Perplexity, and major Chinese-origin platforms such as DeepSeek, Qwen, ERNIE, Kimi, Doubao, and Yuanbao.'
                            Note        = $aiNote
                            Available   = $aiAvail
                            Rows        = $aiData
                        })

                        # 6. Feature availability
                        Write-Host '[6/6] Collecting feature availability / licensing...' -ForegroundColor Cyan
                        $featData = @(); $featNote = ''; $featAvail = $isGraphConnected
                        if ($isGraphConnected) {
                            try {
                                $cap = Get-M365TenantCapabilities
                                if ($cap.SkuServicePlans -and @($cap.SkuServicePlans).Count -gt 0) {
                                    $featData = @(Get-M365FeatureCapabilityMatrix -ServicePlans $cap.SkuServicePlans)
                                    $availCount = @($featData | Where-Object { $_.Available -eq 'Yes' }).Count
                                    $featNote = "$availCount of $($featData.Count) catalogued features available in this tenant."
                                    if ([string]$cap.LicenseStatus -eq 'PartialData' -and -not [string]::IsNullOrWhiteSpace([string]$cap.LicenseStatusDetail)) {
                                        $featNote += " Warning: Partial SKU/service plan data returned by Graph. $([string]$cap.LicenseStatusDetail)"
                                    }
                                } else {
                                    switch ([string]$cap.LicenseStatus) {
                                        'PartialData' {
                                            $featNote = if ([string]::IsNullOrWhiteSpace([string]$cap.LicenseStatusDetail)) {
                                                'Partial SKU/service plan data was returned by Graph, but no usable service plans remained for the feature matrix.'
                                            } else {
                                                "Partial SKU/service plan data returned by Graph, but no usable service plans remained for the feature matrix. $([string]$cap.LicenseStatusDetail)"
                                            }
                                        }
                                        'ScopeMissing' {
                                            $featNote = if ([string]::IsNullOrWhiteSpace([string]$cap.LicenseStatusDetail)) {
                                                "Graph token missing Organization.Read.All scope. Reconnect Graph with -Scopes 'Organization.Read.All'."
                                            } else {
                                                [string]$cap.LicenseStatusDetail
                                            }
                                        }
                                        'SkuCmdletMissing' {
                                            $featNote = if ([string]::IsNullOrWhiteSpace([string]$cap.LicenseStatusDetail)) {
                                                'Get-MgSubscribedSku cmdlet not found. Install/update Microsoft.Graph modules.'
                                            } else {
                                                [string]$cap.LicenseStatusDetail
                                            }
                                        }
                                        'Error' {
                                            $featNote = if ([string]::IsNullOrWhiteSpace([string]$cap.LicenseStatusDetail)) {
                                                'Error retrieving SKU data from Microsoft Graph.'
                                            } else {
                                                "Error retrieving SKU data: $([string]$cap.LicenseStatusDetail)"
                                            }
                                        }
                                        'NoneFound' {
                                            $featNote = 'No subscribed SKUs returned by Graph for this tenant.'
                                        }
                                        default {
                                            $featNote = 'License/SKU data could not be retrieved from Microsoft Graph.'
                                        }
                                    }
                                    $featAvail = $false
                                }
                            }
                            catch { $featNote = "Error collecting feature data: $($_.Exception.Message)"; $featAvail = $false }
                        } else { $featNote = 'Microsoft Graph connection not available.' }
                        $sections.Add(@{
                            Id          = 'features'
                            Title       = 'Licensing — Feature Capability Matrix'
                            Description = 'Cross-reference of this tenant''s active service plans against a curated catalog of M365 features, organized by category.'
                            Note        = $featNote
                            Available   = $featAvail
                            Rows        = $featData
                        })

                        Write-Host 'Generating assessment report...' -ForegroundColor Cyan
                        $savePath = Get-M365ConfiguredSavePath
                        $settings = Get-M365UiSettings
                        $companyName = [string]$settings.CompanyName
                        $reportTitle = if (-not [string]::IsNullOrWhiteSpace($companyName)) { "$companyName — Tenant Assessment" } else { 'M365 Tenant Assessment' }

                        $reportPath = New-M365TenantAssessmentReport -Sections $sections -ReportTitle $reportTitle
                        if (Test-Path -Path $reportPath) {
                            Add-M365ReportHistoryEntry -Title $reportTitle -Path $reportPath
                            Write-Host "Assessment report saved: $reportPath" -ForegroundColor Green

                            $fileUri = ([System.Uri]$reportPath).AbsoluteUri
                            $browserSetting = (Get-M365UiSettings).BrowserPopout
                            switch ($browserSetting) {
                                'Edge'    { $cmd = Get-Command -Name msedge.exe  -ErrorAction SilentlyContinue; if ($cmd) { Start-Process -FilePath $cmd.Source -ArgumentList "--app=$fileUri" } else { Start-Process $reportPath } }
                                'Firefox' { $cmd = Get-Command -Name firefox.exe -ErrorAction SilentlyContinue; if ($cmd) { Start-Process -FilePath $cmd.Source -ArgumentList $fileUri } else { Start-Process $reportPath } }
                                'Chrome'  { $cmd = Get-Command -Name chrome.exe  -ErrorAction SilentlyContinue; if ($cmd) { Start-Process -FilePath $cmd.Source -ArgumentList $fileUri } else { Start-Process $reportPath } }
                                default   { Start-Process $reportPath }
                            }
                        }
                        Read-Host 'Press Enter to continue' | Out-Null
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

            if (-not $skipContinuePrompt -and $normalizedSelection -ne 'B') {
                Read-Host 'Press Enter to continue' | Out-Null
            }
        }
        while ($true)
    }

    function Show-ADAssessmentMenu {
        [CmdletBinding()]
        param()

        do {
            Clear-Host
            $adModuleAvail = [bool](Get-Module ActiveDirectory -ErrorAction SilentlyContinue)
            if (-not $adModuleAvail) {
                $adModuleAvail = [bool](Get-Module ActiveDirectory -ListAvailable -ErrorAction SilentlyContinue)
            }

            Write-Host 'Active Directory Assessment' -ForegroundColor Cyan
            $modColor = if ($adModuleAvail) { 'Green' } else { 'Yellow' }
            $modLabel = if ($adModuleAvail) { 'Available' } else { 'Not available (install RSAT: AD DS Tools)' }
            Write-Host "ActiveDirectory module: $modLabel" -ForegroundColor $modColor
            Write-Host ''
            Write-Host '  Forest & Domain' -ForegroundColor DarkGray
            Write-Host '  1. Domain summary (forest, password policy, inventory)' -ForegroundColor $(if ($adModuleAvail) { 'White' } else { 'Gray' })
            Write-Host '  2. Domain controller inventory (OS, roles, EOS status)' -ForegroundColor $(if ($adModuleAvail) { 'White' } else { 'Gray' })
            Write-Host ''
            Write-Host '  Health & Topology' -ForegroundColor DarkGray
            Write-Host '  3. Replication health (per-DC partnership status)' -ForegroundColor $(if ($adModuleAvail) { 'White' } else { 'Gray' })
            Write-Host '  4. DNS health (zones, scavenging, forwarders)' -ForegroundColor $(if ($adModuleAvail) { 'White' } else { 'Gray' })
            Write-Host '  5. Sites and Services (sites, subnets, site links)' -ForegroundColor $(if ($adModuleAvail) { 'White' } else { 'Gray' })
            Write-Host ''
            Write-Host '  Security & Risk' -ForegroundColor DarkGray
            Write-Host '  6. Security posture (delegation, LAPS, AS-REP, stale accounts...)' -ForegroundColor $(if ($adModuleAvail) { 'White' } else { 'Gray' })
            Write-Host '  7. Operational risk (SYSVOL, GPO health, tombstone, schema...)' -ForegroundColor $(if ($adModuleAvail) { 'White' } else { 'Gray' })
            Write-Host ''
            Write-Host 'A. Run All — Generate full AD assessment HTML report' -ForegroundColor $(if ($adModuleAvail) { 'Cyan' } else { 'Gray' })
            Write-Host 'B. Back'

            $selection = Read-Host 'Select an option'
            if ([string]::IsNullOrWhiteSpace($selection)) { continue }

            $normalizedSelection = $selection.ToUpperInvariant()

            if (-not $adModuleAvail -and $normalizedSelection -ne 'B') {
                Write-Host 'ActiveDirectory module is not available. Install RSAT: Active Directory Domain Services and Lightweight Directory Services Tools.' -ForegroundColor Yellow
                Read-Host 'Press Enter to continue' | Out-Null
                continue
            }

            $serverInput = $null
            if ($normalizedSelection -match '^[1-7A]$' -and $normalizedSelection -ne 'B') {
                $srv = Read-Host 'Optional: target DC or domain FQDN (press Enter for default)'
                $serverInput = $srv.Trim()
            }
            $spArgs = if (-not [string]::IsNullOrWhiteSpace($serverInput)) { @{ Server = $serverInput } } else { @{} }

            $skipContinuePrompt = $false

            try {
                switch ($normalizedSelection) {
                    '1' {
                        $skipContinuePrompt = $true
                        $reportData = Get-ADDomainSummary @spArgs
                        Invoke-M365ConsoleReport -Data $reportData -Title 'AD Domain Summary'
                    }
                    '2' {
                        $skipContinuePrompt = $true
                        $reportData = Get-ADDomainControllerInventory @spArgs
                        Invoke-M365ConsoleReport -Data $reportData -Title 'Domain Controller Inventory'
                    }
                    '3' {
                        $skipContinuePrompt = $true
                        $reportData = Get-ADReplicationHealth @spArgs
                        Invoke-M365ConsoleReport -Data $reportData -Title 'AD Replication Health'
                    }
                    '4' {
                        $skipContinuePrompt = $true
                        $reportData = Get-ADDNSHealth @spArgs
                        Invoke-M365ConsoleReport -Data $reportData -Title 'DNS Health'
                    }
                    '5' {
                        $skipContinuePrompt = $true
                        $reportData = Get-ADSitesAndServicesReport @spArgs
                        Invoke-M365ConsoleReport -Data $reportData -Title 'Sites and Services'
                    }
                    '6' {
                        $skipContinuePrompt = $true
                        $reportData = Get-ADSecurityPosture @spArgs
                        Invoke-M365ConsoleReport -Data $reportData -Title 'AD Security Posture'
                    }
                    '7' {
                        $skipContinuePrompt = $true
                        $reportData = Get-ADOperationalRiskReport @spArgs
                        Invoke-M365ConsoleReport -Data $reportData -Title 'AD Operational Risk'
                    }
                    'A' {
                        $skipContinuePrompt = $true
                        $sections = [System.Collections.Generic.List[object]]::new()

                        # 1. Domain Summary
                        Write-Host '[1/7] Collecting domain summary...' -ForegroundColor Cyan
                        $data = @(); $note = ''; $avail = $true
                        try {
                            $data = @(Get-ADDomainSummary @spArgs)
                            $crit = @($data | Where-Object { $_.Status -eq 'Critical' }).Count
                            $warn = @($data | Where-Object { $_.Status -eq 'Warning'  }).Count
                            $note = if ($crit -gt 0) { "Critical findings: $crit item(s) require immediate attention." } elseif ($warn -gt 0) { "$warn warning(s) detected." } else { 'Domain configuration within expected parameters.' }
                        }
                        catch { $note = "Error: $($_.Exception.Message)"; $avail = $false }
                        $sections.Add(@{ Id = 'domain-summary'; Title = 'Domain Summary'; Description = 'Forest/domain configuration, tombstone lifetime, KRBTGT password age, AD Recycle Bin, and default password policy.'; Note = $note; Available = $avail; Rows = $data })

                        # 2. DC Inventory
                        Write-Host '[2/7] Inventorying domain controllers...' -ForegroundColor Cyan
                        $data = @(); $note = ''; $avail = $true
                        try {
                            $data = @(Get-ADDomainControllerInventory @spArgs)
                            $eol  = @($data | Where-Object { $_.OSStatus -eq 'Critical' }).Count
                            $note = if ($eol -gt 0) { "Critical: $eol DC(s) running end-of-life operating system — no security patches being applied." } else { "$($data.Count) DC(s) found — none on end-of-life OS." }
                        }
                        catch { $note = "Error: $($_.Exception.Message)"; $avail = $false }
                        $sections.Add(@{ Id = 'dc-inventory'; Title = 'Domain Controller Inventory'; Description = 'Per-DC inventory with OS version, FSMO roles, Global Catalog status, RODC flag, SYSVOL replication mode, and end-of-support status.'; Note = $note; Available = $avail; Rows = $data })

                        # 3. Replication
                        Write-Host '[3/7] Collecting replication health...' -ForegroundColor Cyan
                        $data = @(); $note = ''; $avail = $true
                        try {
                            $data = @(Get-ADReplicationHealth @spArgs)
                            $fail = @($data | Where-Object { $_.Status -eq 'Critical' }).Count
                            $note = if ($fail -gt 0) { "Critical: $fail replication partnership(s) failing. Domain consistency is at risk." } elseif ($data.Count -eq 0) { 'No replication partnerships found.' } else { "All $($data.Count) replication partnership(s) healthy." }
                        }
                        catch { $note = "Error: $($_.Exception.Message)"; $avail = $false }
                        $sections.Add(@{ Id = 'replication'; Title = 'Replication Health'; Description = 'Per-DC replication partnership status, consecutive failure counts, days since last successful sync, and sync result codes.'; Note = $note; Available = $avail; Rows = $data })

                        # 4. DNS Health
                        Write-Host '[4/7] Collecting DNS health...' -ForegroundColor Cyan
                        $data = @(); $note = ''; $avail = $true
                        try {
                            $data = @(Get-ADDNSHealth @spArgs)
                            $issues = @($data | Where-Object { $_.Status -ne 'OK' -and $_.Status -ne 'Info' }).Count
                            $note = if ($issues -gt 0) { "$issues DNS zone(s) with configuration issues (scavenging, dynamic update)." } else { "$($data.Count) zone(s) checked — no issues detected." }
                        }
                        catch { $note = "Error: $($_.Exception.Message)"; $avail = $false }
                        $sections.Add(@{ Id = 'dns-health'; Title = 'DNS Health'; Description = 'DNS zone inventory with scavenging, dynamic update, and AD integration status. Uses DnsServer module when available; falls back to AD zone enumeration.'; Note = $note; Available = $avail; Rows = $data })

                        # 5. Sites & Services
                        Write-Host '[5/7] Collecting Sites and Services...' -ForegroundColor Cyan
                        $data = @(); $note = ''; $avail = $true
                        try {
                            $data = @(Get-ADSitesAndServicesReport @spArgs)
                            $issues = @($data | Where-Object { $_.Status -eq 'Warning' }).Count
                            $note = if ($issues -gt 0) { "$issues site/subnet/link item(s) with topology issues." } else { "Sites and Services topology has no detected issues." }
                        }
                        catch { $note = "Error: $($_.Exception.Message)"; $avail = $false }
                        $sections.Add(@{ Id = 'sites-services'; Title = 'Sites and Services'; Description = 'AD replication topology: sites with DC and subnet counts, unassigned subnets, and site link replication intervals.'; Note = $note; Available = $avail; Rows = $data })

                        # 6. Security Posture
                        Write-Host '[6/7] Running security posture analysis...' -ForegroundColor Cyan
                        $data = @(); $note = ''; $avail = $true
                        try {
                            $data = @(Get-ADSecurityPosture @spArgs)
                            $crit = @($data | Where-Object { $_.Status -eq 'Critical' }).Count
                            $warn = @($data | Where-Object { $_.Status -eq 'Warning'  }).Count
                            $note = if ($crit -gt 0) { "Critical: $crit security check(s) failed — immediate remediation required. $warn warning(s) also detected." } elseif ($warn -gt 0) { "$warn security warning(s) detected." } else { 'All security checks passed.' }
                        }
                        catch { $note = "Error: $($_.Exception.Message)"; $avail = $false }
                        $sections.Add(@{ Id = 'security-posture'; Title = 'Security Posture'; Description = 'Structured security checks: privileged group sprawl, delegation misconfigurations, AS-REP/Kerberoastable accounts, LAPS deployment, Protected Users, and stale accounts.'; Note = $note; Available = $avail; Rows = $data })

                        # 7. Operational Risk
                        Write-Host '[7/7] Running operational risk analysis...' -ForegroundColor Cyan
                        $data = @(); $note = ''; $avail = $true
                        try {
                            $data = @(Get-ADOperationalRiskReport @spArgs)
                            $high = @($data | Where-Object { $_.Severity -in 'Critical','High' }).Count
                            $note = if ($high -gt 0) { "$high Critical/High risk finding(s): SYSVOL replication, EOL OS, tombstone lifetime, or GPO health issues detected." } else { 'No Critical or High operational risks detected.' }
                        }
                        catch { $note = "Error: $($_.Exception.Message)"; $avail = $false }
                        $sections.Add(@{ Id = 'operational-risk'; Title = 'Operational Risk'; Description = 'Infrastructure risk indicators: SYSVOL replication mode, tombstone lifetime, AD Recycle Bin, functional levels, GPO health, DC count, and end-of-life OS.'; Note = $note; Available = $avail; Rows = $data })

                        Write-Host 'Generating AD assessment report...' -ForegroundColor Cyan
                        $settings    = Get-M365UiSettings
                        $companyName = [string]$settings.CompanyName
                        $domain      = if (-not [string]::IsNullOrWhiteSpace($serverInput)) { $serverInput } else {
                            try { (Get-ADDomain -ErrorAction SilentlyContinue).DNSRoot } catch { 'Active Directory' }
                        }
                        $reportTitle = if (-not [string]::IsNullOrWhiteSpace($companyName)) { "$companyName — AD Assessment ($domain)" } else { "AD Assessment — $domain" }
                        $reportPath  = New-ADAssessmentReport -Sections $sections -ReportTitle $reportTitle

                        if (Test-Path -Path $reportPath) {
                            Add-M365ReportHistoryEntry -Title $reportTitle -Path $reportPath
                            $fileUri = ([System.Uri]$reportPath).AbsoluteUri
                            $browserSetting = (Get-M365UiSettings).BrowserPopout
                            switch ($browserSetting) {
                                'Edge'    { $cmd = Get-Command -Name msedge.exe  -ErrorAction SilentlyContinue; if ($cmd) { Start-Process -FilePath $cmd.Source -ArgumentList "--app=$fileUri" } else { Start-Process $reportPath } }
                                'Firefox' { $cmd = Get-Command -Name firefox.exe -ErrorAction SilentlyContinue; if ($cmd) { Start-Process -FilePath $cmd.Source -ArgumentList $fileUri } else { Start-Process $reportPath } }
                                'Chrome'  { $cmd = Get-Command -Name chrome.exe  -ErrorAction SilentlyContinue; if ($cmd) { Start-Process -FilePath $cmd.Source -ArgumentList $fileUri } else { Start-Process $reportPath } }
                                default   { Start-Process $reportPath }
                            }
                        }
                        Read-Host 'Press Enter to continue' | Out-Null
                    }
                    'B' { return }
                    Default { Write-Warning 'Unknown selection.' }
                }
            }
            catch {
                Write-Host $_.Exception.Message -ForegroundColor Red
            }

            if (-not $skipContinuePrompt -and $normalizedSelection -ne 'B') {
                Read-Host 'Press Enter to continue' | Out-Null
            }
        }
        while ($true)
    }

    function Invoke-M365SignOutOnExit {
        [CmdletBinding()]
        param()

        function Remove-M365TemporaryReportFiles {
            [CmdletBinding()]
            param()

            $tempReportDirectory = Join-Path -Path $env:TEMP -ChildPath 'M365-Exchange-Reports'
            if (-not (Test-Path -Path $tempReportDirectory)) {
                return
            }

            try {
                Remove-Item -Path $tempReportDirectory -Recurse -Force -ErrorAction Stop
                Write-Host "Removed temporary report files: $tempReportDirectory" -ForegroundColor DarkCyan
            }
            catch {
                Write-Host "Could not remove temporary report files: $tempReportDirectory" -ForegroundColor Yellow
            }
        }

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

        Remove-M365TemporaryReportFiles

        Write-Host 'Module sessions disconnected. Clearing host view...' -ForegroundColor DarkCyan
        Clear-Host
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
        Write-Host '6. Tenant Assessment' -ForegroundColor Green
        Write-Host '7. Active Directory Assessment' -ForegroundColor Green
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
                    Show-M365WindowsConfigurationForm
                }
                '5' {
                    Show-M365FeatureAvailability
                }
                '6' {
                    $skipContinuePrompt = $true
                    Show-M365TenantAssessmentMenu
                }
                '7' {
                    $skipContinuePrompt = $true
                    Show-ADAssessmentMenu
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