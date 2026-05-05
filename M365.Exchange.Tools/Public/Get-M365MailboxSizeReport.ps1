function Get-M365MailboxSizeReport {
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$ExportPath,

        [Parameter()]
        [switch]$IncludeLastEmailReceived
    )

    Assert-M365ExchangePowerShellConnected

    Write-Host 'Retrieving mailbox list from Exchange Online...' -ForegroundColor Cyan
    $mailboxes = Get-EXOMailbox -ResultSize Unlimited -PropertySets StatisticsSeed -Properties DisplayName, PrimarySmtpAddress, RecipientTypeDetails, ArchiveStatus, UserPrincipalName, ExternalDirectoryObjectId -ErrorAction Stop

    $total = @($mailboxes).Count
    Write-Host "Retrieving statistics for $total mailboxes (this may take a few minutes)..." -ForegroundColor Cyan

    $counter = 0
    $results = foreach ($mbx in ($mailboxes | Sort-Object DisplayName)) {
        $counter++
        if ($counter % 25 -eq 0) {
            Write-Host "  Processed $counter / $total..." -ForegroundColor DarkCyan
        }

        $mailboxKind = switch ($mbx.RecipientTypeDetails) {
            'SharedMailbox'    { 'Shared' }
            'RoomMailbox'      { 'Resource' }
            'EquipmentMailbox' { 'Resource' }
            default            { 'User' }
        }

        $totalSizeMB        = $null
        $totalDeletedSizeMB = $null
        $itemCount          = $null
        $deletedItemCount   = $null
        $lastLogonTime      = $null
        $lastEmailReceivedTime = $null

        try {
            $stats = Get-EXOMailboxStatistics -Identity $mbx.UserPrincipalName -ErrorAction Stop
            if ($stats.TotalItemSize) {
                $totalSizeMB = [Math]::Round($stats.TotalItemSize.Value.ToBytes() / 1MB, 2)
            }
            if ($stats.TotalDeletedItemSize) {
                $totalDeletedSizeMB = [Math]::Round($stats.TotalDeletedItemSize.Value.ToBytes() / 1MB, 2)
            }
            $itemCount        = $stats.ItemCount
            $deletedItemCount = $stats.DeletedItemCount
            $lastLogonTime    = $stats.LastLogonTime
        }
        catch { }

        $archiveSizeMB   = $null
        $archiveItemCount = $null

        if ($mbx.ArchiveStatus -eq 'Active') {
            try {
                $archiveStats = Get-EXOMailboxStatistics -Identity $mbx.UserPrincipalName -Archive -ErrorAction Stop
                if ($archiveStats.TotalItemSize) {
                    $archiveSizeMB = [Math]::Round($archiveStats.TotalItemSize.Value.ToBytes() / 1MB, 2)
                }
                $archiveItemCount = $archiveStats.ItemCount
            }
            catch { }
        }

        if ($IncludeLastEmailReceived) {
            try {
                $folderStats = Get-EXOMailboxFolderStatistics -Identity $mbx.UserPrincipalName -FolderScope Inbox -ErrorAction Stop
                $dateValues = foreach ($folder in @($folderStats)) {
                    foreach ($propName in @('NewestItemReceivedDate', 'NewestItemLastModifiedDate')) {
                        $prop = $folder.PSObject.Properties[$propName]
                        if ($prop -and $null -ne $prop.Value -and -not [string]::IsNullOrWhiteSpace([string]$prop.Value)) {
                            try {
                                [datetime]$prop.Value
                                break
                            }
                            catch { }
                        }
                    }
                }

                if ($dateValues) {
                    $lastEmailReceivedTime = ($dateValues | Sort-Object -Descending | Select-Object -First 1)
                }
            }
            catch { }
        }

        $row = [ordered]@{
            DisplayName            = $mbx.DisplayName
            PrimarySmtpAddress     = $mbx.PrimarySmtpAddress
            MailboxKind            = $mailboxKind
            RecipientTypeDetails   = $mbx.RecipientTypeDetails
            TotalItemSizeMB        = $totalSizeMB
            TotalDeletedItemSizeMB = $totalDeletedSizeMB
            ItemCount              = $itemCount
            DeletedItemCount       = $deletedItemCount
            ArchiveStatus          = $mbx.ArchiveStatus
            ArchiveSizeMB          = $archiveSizeMB
            ArchiveItemCount       = $archiveItemCount
            LastLogonTime          = $lastLogonTime
        }

        if ($IncludeLastEmailReceived) {
            $row.LastEmailReceivedTime = $lastEmailReceivedTime
        }

        [pscustomobject]$row
    }

    Export-M365ReportData -InputObject $results -ExportPath $ExportPath
}
