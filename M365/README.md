# M365 Toolkit

A collection of PowerShell scripts for Microsoft 365 management and reporting. Designed for MSPs and organizations managing multiple M365 tenants.

## Scripts

### Get-CustomerReport.ps1
Comprehensive M365 customer audit report with:
- **Key Findings** — License usage, inactive users, MFA coverage, security posture
- **Tenant Overview** — Subscriptions, user counts, license breakdowns
- **License Details** — Renewal dates, assignments, allocation % by license type
- **User Security** — MFA enrollment, last sign-in age, password age
- **Secure Score** — Overall posture and trending
- **Devices** — Enrollment count, OS distribution, compliance status
- **Exchange Online** — Mailbox quota usage, forwarding rules (if available)

Output is a single HTML report with charts, tables, and key metrics.

**Usage:**
```powershell
# Using a customer profile (single-tenant customers.json)
.\Get-CustomerReport.ps1 -CustomersFile ".\Customer.se\contoso.com.json"

# Using a named customer profile
.\Get-CustomerReport.ps1 -Customer contoso.com

# Using ad-hoc credentials
.\Get-CustomerReport.ps1 -TenantId "contoso.com" -AdminUPN "admin@contoso.com" -TenantName "Contoso Inc"

# Skip Exchange Online connection (faster if not needed)
.\Get-CustomerReport.ps1 -Customer contoso.com -SkipExchangeOnline

# Custom output path and inactive threshold
.\Get-CustomerReport.ps1 -Customer contoso.com -OutputPath "C:\Reports" -InactiveThresholdDays 120
```

### Get-TenantOverview.ps1
Quick snapshot of a single M365 tenant:
- Tenant ID and name
- User and group counts
- License summary
- Device enrollment status

**Usage:**
```powershell
.\Get-TenantOverview.ps1 -TenantId "contoso.com"
```

### Get-LicensedUsers.ps1
Export detailed licensed user information with:
- UPN, Display Name, Department
- License SKUs and assignment date
- Last sign-in activity
- MFA status

Outputs to HTML and CSV.

**Usage:**
```powershell
.\Get-LicensedUsers.ps1 -TenantId "contoso.com" -OutputPath "C:\Reports"
```

### Get-CADeletionAudit.ps1
Audit Conditional Access deletion history and configuration changes. Useful for compliance and security reviews.

**Usage:**
```powershell
.\Get-CADeletionAudit.ps1 -TenantId "contoso.com"
```

### move365user-shared.ps1
Migrate a user's mailbox to shared status and transfer ownership in M365. Useful for offboarding scenarios.

**Usage:**
```powershell
.\move365user-shared.ps1 -UserUPN "user@contoso.com" -NewOwnerUPN "manager@contoso.com"
```

### Org-wide calendar sharing (without exposing history)

Two scripts that solve a common request: *"make everyone able to see everyone's calendar — but
don't expose the sensitive historical meetings, only going forward."*

Exchange Online has **no** native way to share "future events only" — calendar folder permissions
always cover the whole folder (past + future). The supported workaround is to mask the history as
**Private** first, then open the calendars. A Private item is hidden from anyone with normal
sharing permissions (they see a busy block labelled *Private*, no subject/body), unless they hold
Full Access or Delegate + ViewPrivateItems — which these scripts do not grant.

> **Caveat:** "Private" is an honour-based Outlook/EXO flag, **not encryption**. For genuinely
> confidential meetings use Microsoft Purview sensitivity labels with encryption instead.

Run them in order:

**1. `Set-HistoricalCalendarPrivate.ps1`** — bulk-sets `sensitivity = private` (via Microsoft
Graph) on every event ending before a cutoff date, across all or selected mailboxes.

```powershell
# Preview only (no changes), one pilot mailbox:
.\Set-HistoricalCalendarPrivate.ps1 -CutoffUtc (Get-Date '2026-06-14T00:00:00Z') -Mailbox pilot@contoso.com -WhatIf

# Live, all mailboxes (recurring series masked too):
.\Set-HistoricalCalendarPrivate.ps1 -CutoffUtc (Get-Date '2026-06-14T00:00:00Z')

# Skip recurring series that may still have future occurrences:
.\Set-HistoricalCalendarPrivate.ps1 -CutoffUtc (Get-Date '2026-06-14T00:00:00Z') -SkipRecurringMasters
```

**2. `Enable-OrgCalendarSharing.ps1`** — sets the `Default` (everyone-in-org) permission on each
Calendar folder via `Set-MailboxFolderPermission`, without granting ViewPrivateItems.

```powershell
# Preview, one mailbox:
.\Enable-OrgCalendarSharing.ps1 -AccessRight LimitedDetails -Mailbox pilot@contoso.com -WhatIf

# Live, all mailboxes (subject + location, no body):
.\Enable-OrgCalendarSharing.ps1 -AccessRight LimitedDetails

# Other levels: AvailabilityOnly (free/busy only) | Reviewer (full details, read-only)
```

**Requirements:** `Set-HistoricalCalendarPrivate.ps1` needs the `Microsoft.Graph` module and the
`Calendars.ReadWrite` scope; `Enable-OrgCalendarSharing.ps1` needs `ExchangeOnlineManagement` and
an Exchange admin role. Localised tenants: the Calendar folder may be named in the mailbox language
(e.g. "Kalender") — see the note at the foot of the sharing script.

**Sources:**
[Set-MailboxFolderPermission](https://learn.microsoft.com/powershell/module/exchangepowershell/set-mailboxfolderpermission?view=exchange-ps) ·
[Graph event resource (sensitivity)](https://learn.microsoft.com/graph/api/resources/event?view=graph-rest-1.0) ·
[KB 4021947 — private appointments on shared calendars](https://learn.microsoft.com/troubleshoot/outlook/calendaring/outlook-or-owa-not-shows-full-details-of-private-appointments) ·
[Purview sensitivity labels for meetings](https://learn.microsoft.com/purview/sensitivity-labels-meetings)

## Configuration

### customers.json

Create a `customers.json` file next to the scripts (or in a subfolder) to manage multiple customer tenants:

```json
[
  {
    "ShortName":         "example",
    "DisplayName":       "Example Company AB",
    "AdminUPN":          "admin@example.com",
    "TenantId":          "00000000-0000-0000-0000-000000000000",
    "DefaultLicenseSku": "O365_BUSINESS_PREMIUM"
  }
]
```

**Fields:**
- `ShortName` — Short identifier used in reports and output folders
- `DisplayName` — Friendly name for reports
- `AdminUPN` — Service account UPN (used for Exchange Online auth)
- `TenantId` — Entra ID Tenant ID or primary domain name
- `DefaultLicenseSku` — Default license for new users (e.g., `O365_BUSINESS_PREMIUM`, `SPB`)

Start with `customers.template.json` and customize for your environment.

## Microsoft Graph Permissions

All scripts require the following scopes to be granted by an admin in your M365 tenant(s):

| Scope | Purpose |
|-------|---------|
| `User.Read.All` | Read user profiles and sign-in activity |
| `Directory.Read.All` | Read directory data (groups, org structure) |
| `Policy.Read.All` | Read security policies (MFA, Conditional Access) |
| `AuditLog.Read.All` | Read audit logs and activity history |
| `UserAuthenticationMethod.Read.All` | Read MFA enrollment status |
| `SecurityEvents.Read.All` | Read security event data and Secure Score |
| `DeviceManagementManagedDevices.Read.All` | Read Intune device inventory |

For **Exchange Online** operations (Get-CustomerReport, move365user-shared):
- Service account must have `Exchange Online` admin role or `Mail Recipients Administrator`

### Granting Permissions

Run any script with first execution. You'll be prompted to grant permissions in your browser:

```powershell
.\Get-CustomerReport.ps1 -Customer example
# Browser opens → Grant permissions → Done
```

Permissions are cached per account; subsequent runs do not require re-authorization.

## Requirements

- **PowerShell 7.0+** (cross-platform; tested on Windows Server 2019+)
- **Microsoft.Graph modules** — automatically installed on first run
- **Exchange Online PowerShell Module** (for Get-CustomerReport, move365user-shared)
- **Admin access** to M365 tenant(s)

## Installation

1. Clone this repository or download the scripts
2. Copy to a safe location (e.g., `C:\Scripts\M365\`)
3. Create `customers.json` from `customers.template.json`
4. Run any script; modules install automatically on first execution

## Privacy & Security

**IMPORTANT:**
- Never commit real `customers.json` to version control. It contains service account credentials and tenant IDs.
- Use `customers.template.json` as a starting point.
- Store `customers.json` in a secure location outside the repository.
- Treat `customers.json` as you would treat API keys or credentials.
- Reports contain sensitive user and device data; store securely and rotate as needed.

## Output

- **Get-CustomerReport** → `.\<shortname>\Report-<date>.html`
- **Get-LicensedUsers** → `.\<outputpath>\LicensedUsers-<date>.html` and `.csv`
- Other scripts → Console output and optional file export

## Troubleshooting

### "NullReferenceException" or "RuntimeBroker"
Occurs in non-interactive environments (headless servers). Solution:
- Scripts automatically retry with `-NoWAM` flag (Windows Account Manager disabled)
- If still failing, ensure you have internet access and modern TLS support

### "Permission denied" on scopes
Solution:
- Re-run with first execution to update permissions
- Verify you have admin consent in your M365 tenant
- Check if conditional access policies are blocking delegated auth

### Exchange Online connection fails
Solution:
- Verify service account has `Mail.ReadWrite` or `Mail.Send` permissions
- Check if MFA is required (scripts support MFA)
- Try `-SkipExchangeOnline` to skip EXO data if not needed

## License

These scripts are provided as-is for M365 management and reporting.

## Authors

- Rosvall & Claude
