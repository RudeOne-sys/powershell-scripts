# M365 Tenant Assessment Tool v3.0

A PowerShell script that performs a comprehensive read-only assessment of a Microsoft 365 tenant. Outputs a set of CSV data files and a self-contained HTML report. Designed for use with both existing customers and potential new customers.

---

## How to Run

### Prerequisites

**PowerShell version:** 5.1 or later (Windows PowerShell on Windows 10/11 is sufficient — PowerShell 7 is not required)

**Modules:** The script installs all required modules automatically on first run. To pre-install manually:
```powershell
Install-Module Microsoft.Graph -Scope CurrentUser -Force
Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force  # optional but recommended
```

**Execution policy** — run this once if you haven't already:
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

---

### Account & Permissions

#### Minimum required role: Global Reader

Global Reader covers the vast majority of the script. Assign it to the account that will run the assessment in the target tenant.

Some sections need additional roles for full data. These degrade gracefully — they show N/A rather than failing the whole script:

| Section | Additional Role Needed | Without It |
|---|---|---|
| Secure Score | Security Reader | Score shows but some controls may be limited |
| MFA Registration % | Reports Reader | Shows N/A |
| Risky Users | Security Reader + Entra ID P2 | Shows 0 |
| Conditional Access | Global Reader is sufficient | — |
| Usage Reports (SharePoint, OneDrive, Teams) | Reports Reader | Sections skipped |
| Audit Log | Global Reader is sufficient | — |

**Recommended role combination for a complete assessment:**
> Global Reader + Security Reader + Reports Reader

All three are read-only. A customer's admin can grant these in **Entra admin center → Roles and administrators** in under 2 minutes.

---

### Running the Script

#### Option A — Your own tenant (existing customer)

```powershell
# Full assessment
.\M365TenantAssessment.ps1 -CustomerName "Contoso"

# Quick scan only (Tenant Overview, Licensing, Secure Score)
.\M365TenantAssessment.ps1 -CustomerName "Contoso" -QuickMode
```

A browser window will open asking you to sign in. Once authenticated, the script runs unattended.

---

#### Option B — External or potential customer tenant

There are three approaches depending on your access situation:

**1. Customer creates a guest account for you (simplest)**

The customer:
1. Goes to **Entra admin center → Users → Invite external user**
2. Invites your email address
3. Assigns the guest account: Global Reader + Security Reader + Reports Reader

You then run:
```powershell
.\M365TenantAssessment.ps1 -CustomerName "Potential Customer Ltd"
```
Sign in with your own email when the browser prompt appears — Microsoft will authenticate you into their tenant via the guest invite.

---

**2. Customer runs the script themselves**

The safest option for a prospect who is not ready to grant you access. Hand them the script and they run it in their own tenant, then send you the output folder from their Downloads.

```powershell
# They run this — no parameters needed
.\M365TenantAssessment.ps1
```

---

**3. Partner / CSP relationship (using -TenantId)**

If you have a Microsoft Partner relationship (CSP or GDAP) with the customer:
```powershell
.\M365TenantAssessment.ps1 -CustomerName "Fabrikam" -TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
```
Sign in with your partner account when prompted. The `-TenantId` parameter directs both Microsoft Graph and Exchange Online to authenticate into the customer's tenant rather than your own.

The customer's Tenant ID can be found in their **Entra admin center → Overview**, or via your Partner Center dashboard.

---

### What Happens When You Run It

1. Required modules are checked and installed if missing (first run only — takes a couple of minutes)
2. A browser sign-in prompt appears
3. You will be asked to consent to the requested permissions — click **Accept**
4. The script works through all 12 sections (or 3 in QuickMode), printing progress to the console
5. When complete, the output folder and HTML report open automatically in Explorer and your browser
6. All output is saved to: `C:\Users\<you>\Downloads\M365_Assessment_<CustomerName>_<timestamp>\`

---

### Consent Note for New Tenants

The first time you run against a tenant, Microsoft will display a consent prompt listing all the permissions the script is requesting. If the account being used is not a tenant admin, an admin may need to grant consent first. Ask the customer's IT admin to either:

- Run the script once themselves so they can approve consent as an admin, or
- Grant admin consent manually at: **Entra admin center → App registrations → Microsoft Graph PowerShell → API permissions → Grant admin consent**

---

## What the Script Does

The script connects to Microsoft Graph (and optionally Exchange Online) using delegated permissions and collects data across 12 assessment areas. Everything is read-only — no changes are made to the tenant.

### Sections

#### 1. Tenant Overview
Collects the organisation name, Tenant ID, all verified domains, primary domain, country, technical contact, and tenant creation date. This forms the header of the HTML report and gives a baseline identity record for the assessment.

#### 2. Licensing
Queries all subscribed SKUs and maps internal Microsoft part numbers to friendly product names (e.g. `SPB` → Microsoft 365 Business Premium). Reports total purchased, consumed, and available licences per SKU, plus warning and suspended unit counts. Highlights SKUs where more than 20% of licences are unused, which is useful for identifying cost-saving opportunities.

#### 3. Secure Score
Retrieves the tenant's current Microsoft Secure Score and the full list of security control recommendations. Each recommendation is ranked by point value, mapped to a product area (Entra ID, Defender for Endpoint, Exchange Online, etc.), and exported. The HTML report highlights the top 10 highest-value gaps that have not yet been implemented.

#### 4. User Health
Provides a full breakdown of the user population: total users, enabled members, guest users, licensed vs unlicensed active accounts, newly created accounts in the last 30 days, and accounts with no sign-in activity in the last 90 days. Also enumerates all directory role assignments and specifically calls out Global Administrator accounts, flagging if there are more than 5 (Microsoft's recommended maximum for most organisations).

#### 5. Entra ID Security
Reads all Conditional Access policies and reports how many are enabled, in report-only mode, or disabled. A tenant with zero enabled policies is flagged as a critical security gap. Also queries MFA and Self-Service Password Reset (SSPR) registration statistics to show what percentage of users have enrolled. Attempts to retrieve risky user counts if Entra ID P2 is licensed.

#### 6. Exchange Online
Reports user mailboxes, shared mailboxes, Microsoft 365 groups, distribution groups, mail-enabled security groups, and plain security groups. If the ExchangeOnlineManagement module is installed and connected, shared mailboxes are identified accurately using `Get-EXOMailbox -RecipientTypeDetails SharedMailbox`. Without the module, a Graph-based heuristic is used (disabled accounts with no assigned licences) and the report notes that the count is an estimate.

#### 7. Email Security (DNS)
Performs live DNS lookups against the tenant's primary domain to check the four key email security records:

| Record | What it checks |
|---|---|
| SPF | Whether a Sender Policy Framework record exists and how strict it is (`-all`, `~all`, or `+all`) |
| DMARC | Whether a DMARC policy exists and whether it is enforced (`reject`, `quarantine`, or `none`) |
| DKIM | Whether a DKIM public key record is present on common selectors |
| MTA-STS | Whether MTA Strict Transport Security is configured |

Missing or unenforced records are flagged with explanations of the risk.

#### 8. SharePoint Online
Uses the Microsoft Graph reporting API to pull 30-day usage data for all SharePoint sites: URL, owner, storage used, file count, active file count, last activity date, and site type. Reports total storage consumption and the number of sites with no activity in the last 30 days.

#### 9. OneDrive for Business
Pulls 30-day OneDrive usage data for all accounts: storage used, allocated storage, file count, active file count, and last activity. Summarises total storage and identifies the top 10 storage consumers.

#### 10. Microsoft Teams
Lists all Teams with visibility (public/private) and archive status, and pulls 30-day activity data including active user count, channel message volume, and meetings organised per team. Active and archived team counts are separated. Tables in the report are capped at 25 rows with a note to refer to the CSV for the full list.

#### 11. Entra ID Applications
Enumerates all App Registrations and Enterprise Applications (Service Principals) in the tenant. Third-party applications (those not published by Microsoft) are separated out and listed with their publisher and verification status. Useful for identifying shadow IT integrations, overly permissive apps, or applications that are no longer in use.

#### 12. Compliance Posture
Checks for the presence of Sensitivity Labels (Microsoft Purview Information Protection) and tests whether the Audit Log is accessible. Notes that DLP policies and Retention policies require a Compliance Administrator role and cannot be read by Global Reader — these are flagged for manual review in the Microsoft Purview portal.

---

### Output Files

All files are saved to a timestamped folder in your Downloads directory.

| File | Contents |
|---|---|
| `M365_Assessment_Report.html` | Self-contained HTML executive summary — the main deliverable |
| `Tenant_Overview.csv` | Organisation and domain details |
| `Licenses.csv` | Full licence SKU breakdown |
| `SecureScore_Current.csv` | Current score snapshot |
| `SecureScore_Recommendations.csv` | All security control recommendations |
| `Users.csv` | Full user list with status and last sign-in |
| `Users_AdminRoles.csv` | All directory role assignments |
| `EntraID_ConditionalAccess.csv` | Conditional Access policy list |
| `Exchange_Mailboxes.csv` | User mailboxes |
| `Exchange_SharedMailboxes.csv` | Shared mailboxes |
| `Exchange_Groups.csv` | M365, distribution, and mail-enabled security groups |
| `Security_Groups.csv` | Non-mail security groups |
| `EmailSecurity_DNS.csv` | SPF, DMARC, DKIM, MTA-STS results |
| `SharePoint_Usage.csv` | Per-site storage and activity (30 days) |
| `OneDrive_Usage.csv` | Per-user OneDrive storage and activity (30 days) |
| `Teams_List.csv` | All Teams |
| `Teams_Usage.csv` | Per-team activity data (30 days) |
| `EntraID_AppRegistrations.csv` | App Registrations |
| `EntraID_ServicePrincipals.csv` | Enterprise Applications / Service Principals |
| `Compliance_SensitivityLabels.csv` | Sensitivity labels (if accessible) |

---

### Parameters

| Parameter | Type | Description |
|---|---|---|
| `-CustomerName` | String | Customer or organisation name. Used in the output folder name and HTML report header. |
| `-TenantId` | String | Target tenant ID for partner or delegated access scenarios. |
| `-QuickMode` | Switch | Runs only Tenant Overview, Licensing, and Secure Score. Useful for a fast pre-sales scan. |
