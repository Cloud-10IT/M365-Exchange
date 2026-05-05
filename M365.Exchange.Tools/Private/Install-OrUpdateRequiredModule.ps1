function Install-OrUpdateRequiredModule {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Name,

        [Parameter(Mandatory)]
        [ValidateSet('Install', 'Update')]
        [string]$Action
    )

    $commandName = if ($Action -eq 'Install') { 'Install-Module' } else { 'Update-Module' }
    $command = Get-Command -Name $commandName -ErrorAction SilentlyContinue
    if (-not $command) {
        Write-Host "$commandName is not available in this PowerShell session." -ForegroundColor Red
        return $false
    }

    try {
        if ($Action -eq 'Install') {
            Install-Module -Name $Name -Scope CurrentUser -Force -AllowClobber -Confirm:$false -ErrorAction Stop
        }
        else {
            Update-Module -Name $Name -Scope CurrentUser -Force -Confirm:$false -ErrorAction Stop
        }

        Write-Host "$Action completed for $Name." -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "$Action failed for $Name. $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}