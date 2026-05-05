function Get-InstalledModuleVersion {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Name
    )

    $installedModule = Get-Module -ListAvailable -Name $Name |
        Sort-Object Version -Descending |
        Select-Object -First 1

    if (-not $installedModule) {
        return $null
    }

    return [version]$installedModule.Version
}