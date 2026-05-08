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

If the shell is not elevated, the launcher still reports missing modules or available updates. You can approve install or update actions directly from `Check prerequisites`, and the tool will launch an elevated PowerShell prompt to complete the action.

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
- Exchange PowerShell connect now reuses the signed-in Microsoft Graph account when available (no manual UPN prompt required in the Exchange menu).
- Entra user last sign-in reporting requires Microsoft Graph `AuditLog.Read.All` consent.
- Delegation report uses Exchange Online PowerShell and requires a separate Exchange connection.
- Group inventory reports can include expanded member addresses.
- Main menu now has top-level Exchange and Entra ID sections.
- Feature availability now lists detected SKUs and service plans in sorted order, with friendly names plus raw identifiers (`SkuPartNumber`, `ServicePlanName`, `ServicePlanId`) for Entra dynamic group rules.
- Main menu option `5. Feature availability` always opens a modern HTML popout with search, column filter, and export actions (CSV and PDF).
- Option `5. Feature availability` can also export a full bundle to the configured save path (CSV + generated PDF files) without browser save dialogs.
- Duplicate-group usage report can optionally include Azure RBAC evidence when Az.Resources is available and you are connected with `Connect-AzAccount`.
- Configuration menu includes browser selection for report windows: Edge, Firefox, Chrome, Brave, Default (system browser), or None (console table only).
- Configuration menu also lets you set Company Name and a logo path used in popout report headers.
- Configuration menu includes report save path and file name template settings for generated files.
- File name template supports tokens: `{Title}`, `{Timestamp}`, `{Date}`, `{Time}`, `{CompanyName}`.
- Configuration menu includes HTML branding toggles: enable/disable branding, show/hide company name, and show/hide logo.
- Configuration menu includes a logo preview option.
- Popout report view supports profile filters for Users, Shared Mailboxes, Guests, and On-Prem Synced users, plus text search.