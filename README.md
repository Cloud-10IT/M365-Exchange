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
- Main menu option `4. Configuration` opens a native Windows settings form directly.
- Configuration form includes browser selection for report windows: Edge, Firefox, Chrome, Brave, Default (system browser), or None (console table only).
- Configuration form also lets you set Company Name and a logo path used in popout report headers.
- Configuration form includes report save path and file name template settings for generated files.
- File name template supports tokens: `{Title}`, `{Timestamp}`, `{Date}`, `{Time}`, `{CompanyName}`.
- Configuration form includes HTML branding toggles: enable/disable branding, show/hide company name, and show/hide logo.
- Configuration form includes Company Name, Logo, Save Path, File Name Template, theme colors (Primary/Secondary), report font family (for example, `Verdana`), and preferred Exchange auth mode (`Auto`, `Interactive`, `DisableWAM`, `Device`).
- Native Windows settings form includes a `Reset Defaults` button to restore recommended generic values quickly.
- Configuration form includes a logo preview option.
- Popout report view supports profile filters for Users, Shared Mailboxes, Guests, and On-Prem Synced users, plus text search.
- Choosing `Q. Quit` now signs out of Microsoft Graph, Exchange Online PowerShell, and Az (if loaded), clears PowerShell command history for the session, and clears the host view.

## Git Hygiene

- The tracked settings file `M365.Exchange.Tools/Config/M365.Exchange.Tools.Settings.json` is intentionally sanitized with generic values.
- A pre-commit hook is included at `.githooks/pre-commit` to auto-sanitize the settings file before each commit.
- Enable hooks once per clone:

```powershell
git config core.hooksPath .githooks
```

- You can run sanitization manually any time:

```powershell
.\scripts\Sanitize-M365Settings.ps1
```

## Configuration Guide

Use `Configuration` from the main menu to open the native Windows settings form.

### Configuration Fields and Examples

- `CompanyName`: Name shown in report branding.
Example: `Contoso Ltd`
- `LogoPath`: Full local path to logo image.
Example: `C:\Branding\logo.png`
- `ReportSavePath`: Folder path for generated CSV/PDF bundles.
Example: `C:\Reports\M365`
- `FileNameTemplate`: Output naming pattern.
Example: `{CompanyName}-{Title}-{Date}-{Time}`
- `ThemePrimaryColor`: Primary accent color in `#RRGGBB` format.
Example: `#0f766e`
- `ThemeSecondaryColor`: Secondary/header color in `#RRGGBB` format.
Example: `#1e293b`
- `ReportFontFamily`: Font used in HTML report body.
Example: `Verdana`
- `ExchangeAuthMode`: Preferred Exchange Online sign-in mode.
Example: `DisableWAM`
- `HtmlBrandingEnabled`: Enables/disables report branding section.
Example: `true`
- `HtmlShowCompanyName`: Shows/hides company name in report header.
Example: `true`
- `HtmlShowCompanyLogo`: Shows/hides logo image in report header.
Example: `true`

### Token Reference for FileNameTemplate

- `{Title}`: Report title text
- `{Timestamp}`: `yyyyMMdd-HHmmss`
- `{Date}`: `yyyyMMdd`
- `{Time}`: `HHmmss`
- `{CompanyName}`: Sanitized company name token

---

## MSP Microsoft 365 Tenant Baseline Assessment

### Purpose

A practical, repeatable baseline data set for collecting meaningful insight from a customer Microsoft 365 tenant. Covers:

- Security posture
- Identity risk
- Licensing efficiency
- Collaboration sprawl
- Operational and governance gaps

> **Guiding principle:** Pull data that changes decisions.

### Scope & Assumptions

- Customer uses Microsoft 365 / Entra ID
- PowerShell access is approved and admin read permissions are granted
- Tools: Microsoft Graph PowerShell (identity, licensing, groups, devices) and Exchange Online PowerShell (mailboxes)

### Authentication

```powershell
# Entra ID / Microsoft Graph
Connect-MgGraph -Scopes "Directory.Read.All","User.Read.All","Group.Read.All","Policy.Read.All"

# Exchange Online
Connect-ExchangeOnline
```

Or use the built-in connect options from the main menu which request all required scopes in one step.

---

### 1. Identity & Access Overview

**Why this matters:** Identifies dormant accounts, guest sprawl, and security exposure.

```powershell
Get-MgUser -All | Select DisplayName, UserPrincipalName, UserType, AccountEnabled, CreatedDateTime
```

Key signals:
- Disabled but still-licensed users
- Guest accounts older than expected
- High account creation volume

**Privileged Roles** — Global Admin sprawl is the #1 real-world tenant risk.

```powershell
Get-MgDirectoryRole | ForEach-Object {
    Get-MgDirectoryRoleMember -DirectoryRoleId $_.Id |
    Select @{n="Role"; e={$_.AdditionalProperties.displayName}}, DisplayName, UserPrincipalName
}
```

Key signals:
- Too many Global Administrators
- No dedicated break-glass account
- Privileged access assigned to regular users

---

### 2. Licensing & Cost Efficiency

**Why this matters:** This is where MSPs often create immediate savings.

```powershell
# Tenant license summary
Get-MgSubscribedSku | Select SkuPartNumber, ConsumedUnits, PrepaidUnits

# Per-user license assignment
Get-MgUser -All | ForEach-Object {
    [PSCustomObject]@{
        UserPrincipalName = $_.UserPrincipalName
        LicenseCount      = ($_.AssignedLicenses | Measure-Object).Count
    }
}
```

Key signals:
- Licenses assigned to disabled users
- Over-licensed tenants
- Inefficient license tier usage

The **Feature Availability** report (option 5 in the main menu) cross-references the tenant's active service plans against a built-in feature catalog to show which M365 features are actually available in the tenant.

---

### 3. Mailboxes & Messaging Risk

> **Important:** Use `Get-Mailbox` (Exchange Online PowerShell) for insight and audit work. `Get-EXOMailbox` is optimized for bulk operations but intentionally limits available properties.

```powershell
# Full mailbox inventory
Get-Mailbox -ResultSize Unlimited | Select DisplayName, RecipientTypeDetails, PrimarySmtpAddress

# Forwarding rules — critical security signal
Get-Mailbox -ResultSize Unlimited |
    Where-Object { $_.ForwardingSmtpAddress -or $_.ForwardingAddress } |
    Select DisplayName, ForwardingSmtpAddress

# Shared mailboxes
Get-Mailbox -RecipientTypeDetails SharedMailbox | Select DisplayName, PrimarySmtpAddress
```

Key signals:
- Silent data exfiltration via forwarding rules
- Forgotten or orphaned shared mailboxes
- Shared mailboxes being used as user accounts

The **User Mailbox Inventory** report merges Exchange schema data with Entra ID identity fields (`AccountEnabled`, `UsageLocation`, `AssignedLicenseCount`, `PasswordLastSet`) and sign-in activity from Graph beta (`LastSignInDateTime`), which replaces the deprecated Exchange `LastLogonTime` field.

---

### 4. Groups & Collaboration Sprawl

```powershell
# All groups overview
Get-MgGroup -All | Select DisplayName, GroupTypes, SecurityEnabled, MailEnabled, CreatedDateTime

# M365 Unified groups only
Get-MgGroup -All | Where-Object { $_.GroupTypes -contains "Unified" }
```

> **Note on Group Activity fields:**
> - `LastActivityDate`, `LastEmailActivityDate`, `LastSharePointActivityDate`, and `LastTeamsActivityDate` only populate for **Microsoft 365 (Unified) groups**.
> - They do **not** populate for Security Groups or Distribution Lists.
> - Requires `Reports.Read.All` scope. The `ActivityDataStatus` column explains the result per row.

The **Entra Group Inventory** report also includes:
- `AppAssignmentCount` and `AssignedToApplications` — which Enterprise Applications the group is assigned to (via `/appRoleAssignments`)
- `RenewedDateTime` — last renewal date, useful for identifying stale security groups
- `MemberCount` populated even when full member list is not fetched (uses a lightweight `$count` call)

Key signals:
- Group sprawl with no clear ownership
- Unused M365 Groups / Teams with no recent activity
- Security groups with no app assignments and no members

---

### 5. Devices & Endpoint Hygiene

```powershell
Get-MgDevice -All |
    Select DisplayName, OperatingSystem, AccountEnabled, ApproximateLastSignInDateTime
```

Key signals:
- Stale devices not signing in
- Azure-joined but unmanaged endpoints
- Credential risk from shared machines

---

### 6. Security Configuration Signals

```powershell
# Conditional Access policies
Get-MgConditionalAccessPolicy
```

Key questions:
- Is MFA enforced for all users?
- Is legacy authentication blocked?
- Are administrator accounts protected with a dedicated CA policy?

---

### 7. Operational Risk Indicators

Non-technical but critical checks that no script can fully automate:

| Check | Risk if missing |
|---|---|
| Number of Global Admins | Sprawl increases blast radius of account compromise |
| Documented break-glass account | Lockout risk during incident response |
| Shared resources have clear owners | Orphaned resources accumulate silently |
| License assignment is centralized | Ad-hoc assignment leads to cost waste and gaps |
| Forwarding rules reviewed periodically | Silent exfiltration goes undetected |