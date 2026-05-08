@{
    RootModule        = 'M365.Exchange.Tools.psm1'
    ModuleVersion     = '0.3.0'
    GUID              = '9c8f9094-0ee5-4c18-b5fc-c5548cf9867c'
    Author            = 'GitHub Copilot'
    CompanyName       = 'Local'
    Copyright         = '(c) 2026'
    Description       = 'Menu-driven Microsoft 365 Exchange reporting helpers.'
    PowerShellVersion = '5.1'
    RequiredModules   = @()
    FunctionsToExport = @(
        'Connect-M365ExchangePowerShell',
        'Connect-M365ExchangeTools',
        'Get-M365AIApplicationUsageReport',
        'Get-ADDNSHealth',
        'Get-ADDomainControllerInventory',
        'Get-ADDomainSummary',
        'Get-ADOperationalRiskReport',
        'Get-ADReplicationHealth',
        'Get-ADSecurityPosture',
        'Get-ADSitesAndServicesReport',
        'Get-M365ConditionalAccessAnalysis',
        'Get-M365ConditionalAccessReport',
        'Get-M365ContactInventory',
        'Get-M365DeviceInventory',
        'Get-M365DistributionGroupInventory',
        'Get-M365EntraDuplicateGroupUsageReport',
        'Get-M365EntraGroupInventory',
        'Get-M365EntraUserInventory',
        'Get-M365MailboxDelegationReport',
        'Get-M365MailboxForwardingReport',
        'Get-M365MailboxInventory',
        'Get-M365MailboxSizeReport',
        'Get-M365PrivilegedRoleMembers',
        'Get-M365ResourceMailboxInventory',
        'Show-M365Prerequisites',
        'Get-M365SharedMailboxInventory',
        'Get-M365UnifiedGroupInventory',
        'Show-M365ExchangeMenu'
    )
    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @()
}