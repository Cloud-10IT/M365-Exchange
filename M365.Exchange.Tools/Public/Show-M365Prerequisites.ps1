function Show-M365Prerequisites {
    [CmdletBinding()]
    param(
        [Parameter()]
        [switch]$PromptForActions
    )

    $requiredModules = @(
        'Microsoft.Graph',
        'ExchangeOnlineManagement'
    )

    $isAdministrator = Test-IsAdministrator
    $hasIssues = $false

    Write-Host 'Checking required PowerShell modules...' -ForegroundColor Cyan

    if (-not $isAdministrator) {
        Write-Host 'This session is not elevated. Module install and update actions will be skipped.' -ForegroundColor DarkYellow
    }

    foreach ($requiredModule in $requiredModules) {
        $installedVersion = Get-InstalledModuleVersion -Name $requiredModule
        $latestVersion = Get-LatestGalleryModuleVersion -Name $requiredModule

        if (-not $installedVersion) {
            $hasIssues = $true

            if ($latestVersion) {
                Write-Host "[Missing] $requiredModule (Missing modules) (latest available: $latestVersion)" -ForegroundColor Red
            }
            else {
                Write-Host "[Missing] $requiredModule (Missing modules)" -ForegroundColor Red
            }

            if ($isAdministrator -and $PromptForActions) {
                if (Read-M365YesNo -Prompt "Install $requiredModule now?" -Default $true) {
                    Install-OrUpdateRequiredModule -Name $requiredModule -Action Install | Out-Null
                }
            }
            else {
                Write-Host "Run an elevated PowerShell session and execute: Install-Module $requiredModule -Scope CurrentUser" -ForegroundColor Yellow
            }

            continue
        }

        if ($latestVersion -and $installedVersion -lt $latestVersion) {
            $hasIssues = $true
            Write-Host "[Installed] $requiredModule $installedVersion (Module installed)" -ForegroundColor Gray
            Write-Host "[Update Available] $requiredModule $installedVersion -> $latestVersion" -ForegroundColor Yellow

            if ($isAdministrator -and $PromptForActions) {
                if (Read-M365YesNo -Prompt "Update $requiredModule now?" -Default $true) {
                    Install-OrUpdateRequiredModule -Name $requiredModule -Action Update | Out-Null
                }
            }
            else {
                Write-Host "Run an elevated PowerShell session and execute: Update-Module $requiredModule -Scope CurrentUser" -ForegroundColor Yellow
            }
        }
        elseif ($latestVersion) {
            Write-Host "[Installed] $requiredModule $installedVersion (Module installed)" -ForegroundColor Gray
        }
        else {
            Write-Host "[Installed] $requiredModule $installedVersion (Module installed)" -ForegroundColor Gray
        }
    }

    if (-not $hasIssues) {
        Write-Host ''
        Write-Host '[READY] All prerequisites are installed and current.' -ForegroundColor Green
    }

    Write-Host ''
}