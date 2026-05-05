# M365 Exchange Tools

Menu-driven PowerShell module for Microsoft 365 reporting with top-level Exchange and Entra ID menus:

- Exchange reports: user/shared/resource mailboxes, contacts, distribution groups, M365 groups
- Exchange mailbox delegation and access report (Exchange Online PowerShell hybrid path)
- Entra ID reports: users and groups
- Entra duplicate-group usage analysis with derived LastSeen/Confidence
- Entra user reporting includes member/guest filtering, last sign-in activity, password expiry indicators, and extended user properties
- Export each report to CSV

## Prerequisites

1. PowerShell 5.1 or PowerShell 7
2. Microsoft Graph PowerShell SDK
3. Exchange Online Management module (for delegation report)

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
Install-Module ExchangeOnlineManagement -Scope CurrentUser
```

## Launch

```powershell
.\Start-M365ExchangeMenu.ps1
```

On startup, the launcher checks whether the required external module is installed, compares the installed version to the newest version available in the PowerShell Gallery, and prompts you to install or update before the menu opens when the shell is elevated.

If the shell is not elevated, the launcher still reports missing modules or available updates, but it skips the install or update prompt and shows the command to run from an elevated PowerShell session.

The main menu also includes a green `Check prerequisites` option so you can re-run the same prerequisite check on demand.

Or import the module and call functions directly:

```powershell
Import-Module .\M365.Exchange.Tools\M365.Exchange.Tools.psd1
Connect-M365ExchangeTools
Connect-M365ExchangePowerShell
Get-M365MailboxInventory
Get-M365ResourceMailboxInventory
Get-M365ContactInventory
Get-M365DistributionGroupInventory -IncludeMembers -ExportPath .\distribution-groups.csv
Get-M365UnifiedGroupInventory -IncludeMembers -ExportPath .\m365-groups.csv
Get-M365MailboxDelegationReport -IncludeFolderPermissions -ExportPath .\mailbox-delegation.csv
Get-M365EntraUserInventory
Get-M365EntraGroupInventory -IncludeMembers
Get-M365EntraDuplicateGroupUsageReport
```

## Notes

- The main menu lists report data without prompting for export.
- A dedicated Export menu provides CSV export choices for each report.
- Exchange reports use Microsoft Graph delegated authentication for inventory/group/contact reporting.
- Entra user last sign-in reporting requires Microsoft Graph `AuditLog.Read.All` consent.
- Delegation report uses Exchange Online PowerShell and requires a separate Exchange connection.
- Group inventory reports can include expanded member addresses.
- Main menu now has top-level Exchange and Entra ID sections.
- Duplicate-group usage report can optionally include Azure RBAC evidence when Az.Resources is available and you are connected with `Connect-AzAccount`.
- Configuration menu includes browser selection for report windows: Edge, Firefox, Chrome, Brave, Default (system browser), or None (console table only).
- Configuration menu also lets you set Company Name and a logo path used in popout report headers.
- Configuration menu includes a logo preview option.
- Popout report view supports profile filters for Users, Shared Mailboxes, Guests, and On-Prem Synced users, plus text search.