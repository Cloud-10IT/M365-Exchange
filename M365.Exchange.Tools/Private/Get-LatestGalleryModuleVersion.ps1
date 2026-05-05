function Get-LatestGalleryModuleVersion {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Name
    )

    $findModuleCommand = Get-Command -Name Find-Module -ErrorAction SilentlyContinue
    if (-not $findModuleCommand) {
        return $null
    }

    try {
        $galleryModule = Find-Module -Name $Name -ErrorAction Stop
        return [version]$galleryModule.Version
    }
    catch {
        Write-Host "Unable to query the PowerShell Gallery for $Name. $($_.Exception.Message)" -ForegroundColor DarkYellow
        return $null
    }
}