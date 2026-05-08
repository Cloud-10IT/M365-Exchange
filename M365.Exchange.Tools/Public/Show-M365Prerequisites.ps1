function Show-M365Prerequisites {
    [CmdletBinding()]
    param(
        [Parameter()]
        [switch]$PromptForActions
    )

    function Invoke-M365ElevatedModuleAction {
        [CmdletBinding()]
        param(
            [Parameter(Mandatory)]
            [string]$Name,

            [Parameter(Mandatory)]
            [ValidateSet('Install', 'Update')]
            [string]$Action
        )

        $hostExecutable = if (Get-Command -Name 'pwsh' -ErrorAction SilentlyContinue) {
            'pwsh'
        }
        elseif (Get-Command -Name 'powershell' -ErrorAction SilentlyContinue) {
            'powershell'
        }
        else {
            $null
        }

        if (-not $hostExecutable) {
            Write-Host 'Could not locate pwsh or powershell to launch an elevated install session.' -ForegroundColor Red
            return $false
        }

        $commandText = if ($Action -eq 'Install') {
            "Install-Module -Name '$Name' -Scope CurrentUser -Force -AllowClobber -Confirm:`$false -ErrorAction Stop"
        }
        else {
            "Update-Module -Name '$Name' -Scope CurrentUser -Force -Confirm:`$false -ErrorAction Stop"
        }

        $elevatedScript = @"
`$ErrorActionPreference = 'Stop'
try {
    $commandText
    Write-Host '$Action completed for $Name.' -ForegroundColor Green
}
catch {
    Write-Host '$Action failed for $Name. ' -NoNewline -ForegroundColor Red
    Write-Host `$_.Exception.Message -ForegroundColor Red
    exit 1
}
"@

        $encodedCommand = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($elevatedScript))

        try {
            $process = Start-Process -FilePath $hostExecutable -Verb RunAs -ArgumentList @('-NoLogo', '-NoProfile', '-EncodedCommand', $encodedCommand) -Wait -PassThru
            if ($process.ExitCode -eq 0) {
                return $true
            }

            Write-Host "$Action did not complete successfully for $Name in the elevated session." -ForegroundColor Red
            return $false
        }
        catch {
            Write-Host "Unable to start elevated session for $Action $Name. $($_.Exception.Message)" -ForegroundColor Red
            return $false
        }
    }

    $requiredModules = @(
        'Microsoft.Graph',
        'ExchangeOnlineManagement'
    )

    $isAdministrator = Test-IsAdministrator
    $hasIssues = $false

    Write-Host 'Checking required PowerShell modules...' -ForegroundColor Cyan

    if (-not $isAdministrator) {
        Write-Host 'This session is not elevated. Install/update can still run by launching an elevated prompt from here.' -ForegroundColor DarkYellow
    }

    foreach ($requiredModule in $requiredModules) {
        $moduleHasOutstandingIssue = $false
        $installedVersion = Get-InstalledModuleVersion -Name $requiredModule
        $latestVersion = Get-LatestGalleryModuleVersion -Name $requiredModule

        if (-not $installedVersion) {
            $moduleHasOutstandingIssue = $true

            if ($latestVersion) {
                Write-Host "[Missing] $requiredModule (Missing modules) (latest available: $latestVersion)" -ForegroundColor Red
            }
            else {
                Write-Host "[Missing] $requiredModule (Missing modules)" -ForegroundColor Red
            }

            if ($isAdministrator -and $PromptForActions) {
                if (Read-M365YesNo -Prompt "Install $requiredModule now?" -Default $true) {
                    if (Install-OrUpdateRequiredModule -Name $requiredModule -Action Install) {
                        $moduleHasOutstandingIssue = $false
                    }
                }
            }
            elseif ($PromptForActions) {
                if (Read-M365YesNo -Prompt "Install $requiredModule now by launching elevated PowerShell?" -Default $true) {
                    if (Invoke-M365ElevatedModuleAction -Name $requiredModule -Action Install) {
                        $moduleHasOutstandingIssue = $false
                    }
                }
            }
            else {
                Write-Host "Run an elevated PowerShell session and execute: Install-Module $requiredModule -Scope CurrentUser" -ForegroundColor Yellow
            }

            if ($moduleHasOutstandingIssue) {
                $hasIssues = $true
            }

            continue
        }

        if ($latestVersion -and $installedVersion -lt $latestVersion) {
            $moduleHasOutstandingIssue = $true
            Write-Host "[Installed] $requiredModule $installedVersion (Module installed)" -ForegroundColor Gray
            Write-Host "[Update Available] $requiredModule $installedVersion -> $latestVersion" -ForegroundColor Yellow

            if ($isAdministrator -and $PromptForActions) {
                if (Read-M365YesNo -Prompt "Update $requiredModule now?" -Default $true) {
                    if (Install-OrUpdateRequiredModule -Name $requiredModule -Action Update) {
                        $moduleHasOutstandingIssue = $false
                    }
                }
            }
            elseif ($PromptForActions) {
                if (Read-M365YesNo -Prompt "Update $requiredModule now by launching elevated PowerShell?" -Default $true) {
                    if (Invoke-M365ElevatedModuleAction -Name $requiredModule -Action Update) {
                        $moduleHasOutstandingIssue = $false
                    }
                }
            }
            else {
                Write-Host "Run an elevated PowerShell session and execute: Update-Module $requiredModule -Scope CurrentUser" -ForegroundColor Yellow
            }

            if ($moduleHasOutstandingIssue) {
                $hasIssues = $true
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