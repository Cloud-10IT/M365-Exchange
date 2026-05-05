function Export-M365ReportData {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowNull()]
        [AllowEmptyCollection()]
        [object[]]$InputObject,

        [Parameter()]
        [string]$ExportPath
    )

    $reportData = if ($null -eq $InputObject) { @() } else { @($InputObject) }

    if ($reportData.Count -eq 0) {
        Write-Host 'No records found.' -ForegroundColor DarkYellow
    }

    if ($ExportPath) {
        $directoryPath = Split-Path -Path $ExportPath -Parent
        if ($directoryPath -and -not (Test-Path -Path $directoryPath)) {
            New-Item -Path $directoryPath -ItemType Directory -Force | Out-Null
        }

        if ($reportData.Count -gt 0) {
            $reportData | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8
            Write-Host "Exported report to $ExportPath" -ForegroundColor Green
        }
        else {
            Write-Host "Skipped CSV export because no records were returned for this report." -ForegroundColor DarkYellow
        }
    }

    return $reportData
}