function Get-M365MailboxDeletionAuditReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Identity,

        [Parameter()]
        [ValidateRange(1, 365)]
        [int]$Days = 7,

        [Parameter()]
        [ValidateRange(1, 50000)]
        [int]$ResultSize = 5000,

        [Parameter()]
        [string]$ExportPath
    )

    Assert-M365ExchangePowerShellConnected

    $targetMailbox = Get-EXOMailbox -Identity $Identity -Properties DisplayName, UserPrincipalName, PrimarySmtpAddress -ErrorAction Stop
    $targetValues = @(
        [string]$targetMailbox.UserPrincipalName,
        [string]$targetMailbox.PrimarySmtpAddress,
        [string]$targetMailbox.DisplayName,
        [string]$Identity
    ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { $_.ToLowerInvariant() } | Select-Object -Unique

    $startDate = (Get-Date).AddDays(-1 * [Math]::Abs($Days))
    $endDate = Get-Date
    $operations = @('HardDelete', 'SoftDelete', 'MoveToDeletedItems')

    Write-Host "Searching audit logs for mailbox '$($targetMailbox.PrimarySmtpAddress)' over last $Days day(s)..." -ForegroundColor Cyan

    $results = @()
    $searchUnifiedCmd = Get-Command -Name Search-UnifiedAuditLog -ErrorAction SilentlyContinue
    $searchMailboxCmd = Get-Command -Name Search-MailboxAuditLog -ErrorAction SilentlyContinue

    if ($searchUnifiedCmd) {
        $records = Search-UnifiedAuditLog -StartDate $startDate -EndDate $endDate -Operations $operations -ResultSize $ResultSize -ErrorAction Stop

        foreach ($record in @($records)) {
            $auditData = $null
            if ($record.AuditData) {
                try {
                    $auditData = $record.AuditData | ConvertFrom-Json -ErrorAction Stop
                }
                catch { }
            }

            $ownerValues = @(
                [string]$auditData.MailboxOwnerUPN,
                [string]$auditData.MailboxOwnerSid,
                [string]$auditData.ObjectId,
                [string]$auditData.MailboxGuid
            ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { $_.ToLowerInvariant() }

            $isTargetMailbox = $false
            foreach ($value in $ownerValues) {
                if ($targetValues -contains $value) {
                    $isTargetMailbox = $true
                    break
                }
            }

            if (-not $isTargetMailbox) {
                continue
            }

            $results += [pscustomobject]@{
                TimeStamp            = $record.CreationDate
                Mailbox              = if ($auditData.MailboxOwnerUPN) { $auditData.MailboxOwnerUPN } else { $targetMailbox.PrimarySmtpAddress }
                DeletedBy            = if ($auditData.UserId) { $auditData.UserId } else { $record.UserIds }
                Operation            = if ($auditData.Operation) { $auditData.Operation } else { $record.Operations }
                Subject              = [string]$auditData.ItemSubject
                FolderPath           = [string]$auditData.FolderPathName
                InternetMessageId    = [string]$auditData.InternetMessageId
                ClientIPAddress      = [string]$auditData.ClientIPAddress
                ClientInfoString     = [string]$auditData.ClientInfoString
                ResultStatus         = [string]$auditData.ResultStatus
            }
        }
    }
    elseif ($searchMailboxCmd) {
        $records = Search-MailboxAuditLog -Identity $targetMailbox.UserPrincipalName -LogonTypes Owner, Delegate, Admin -Operations $operations -StartDate $startDate -EndDate $endDate -ShowDetails -ResultSize $ResultSize -ErrorAction Stop

        foreach ($record in @($records)) {
            $results += [pscustomobject]@{
                TimeStamp            = $record.LastAccessed
                Mailbox              = $targetMailbox.PrimarySmtpAddress
                DeletedBy            = if ($record.LogonUserDisplayName) { $record.LogonUserDisplayName } else { $record.LogonUserSid }
                Operation            = $record.Operation
                Subject              = [string]$record.ItemSubject
                FolderPath           = [string]$record.FolderPathName
                InternetMessageId    = [string]$record.InternetMessageId
                ClientIPAddress      = [string]$record.ClientIPAddress
                ClientInfoString     = [string]$record.ClientInfoString
                ResultStatus         = [string]$record.OperationResult
            }
        }
    }
    else {
        throw 'No supported mailbox audit command found. Connect Exchange Online with compliance/audit cmdlets available and try again.'
    }

    if (-not $results) {
        Write-Host 'No deletion audit records found for the selected mailbox in the requested time window.' -ForegroundColor Yellow
    }

    $orderedResults = $results | Sort-Object TimeStamp -Descending
    Export-M365ReportData -InputObject $orderedResults -ExportPath $ExportPath
}
