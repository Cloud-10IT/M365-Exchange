function Read-M365ExportPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ReportName
    )

    Read-Host "Optional CSV export path for $ReportName (press Enter to skip)"
}