# Azure Assessment Toolkit

How to run a Microsoft Azure assessment for a customer or prospect, and how its output goes into the BMAD Assessment Toolkit.

Three tools, each answering a different question:

| Tool | Answers | Output | Run it |
|---|---|---|---|
| **Azure-Tenant-Assessment-v1.ps1** (Netsurit) | What is wrong, how serious, what it costs | CSV evidence, `Findings.csv`, `NotCollected.csv`, `Assessment_Summary.json`, branded HTML report | Every engagement |
| **ARI** (Azure Resource Inventory, Microsoft open source) | What is there, in detail | Excel inventory (one tab per resource type) + draw.io network and organisation diagrams | Every engagement |
| **WARA** (Well-Architected Reliability Assessment, Microsoft) | How resilient the production workloads are, against the Azure Proactive Resiliency Library | Collector JSON → Action Plan Excel → PowerPoint | When reliability or a production workload review is in scope |

All three are read-only. None of them change anything in the customer's tenant.

### Framework alignment

Every assessment activity, finding and recommendation is referenced to Microsoft's two Azure frameworks:

- **Cloud Adoption Framework (CAF)**: the platform view. Its seven phases are Strategy, Plan, Ready, Adopt, Govern, Secure and Manage, and its Azure landing zone has eight design areas.
- **Well-Architected Framework (WAF)**: the workload view. It has five pillars, each with a design review checklist (RE:01-10, SE:01-12, CO:01-14, OE:01-11, PE:01-12).

| Activity | CAF | WAF |
|---|---|---|
| Netsurit script | Ready, Adopt, Govern, Secure, Manage; seven of the eight landing zone design areas | All five pillars (Performance Efficiency through Advisor and VM size checks) |
| ARI | Ready: inventory and network diagrams as evidence for the design areas | Evidence only |
| WARA | Manage (protect and recover) | Reliability, against the Azure Proactive Resiliency Library |
| Customer workshop (step 4) | Azure Landing Zone Review: covers Platform automation and DevOps, and what APIs cannot show | Azure Well-Architected Review: covers targets, testing, practices, cost model |

Section 6 has the detail. In reports, findings are **referenced to** CAF and WAF. They are never "compliant" or "non-compliant".

---

## 1. Access the customer must grant

One assessment account (a Netsurit guest account, or a named account in their tenant) with:

| Role | Where | Why | Without it |
|---|---|---|---|
| **Reader** | Tenant Root Group (management group) | Every subscription and management group in one go; policy and RBAC at MG scope | Only subscriptions you are given are seen; MG hierarchy and MG-level policy/RBAC are missing |
| **Security Reader** | Tenant Root Group | Defender for Cloud plans, secure score, recommendations | Security section shows N/A |
| **Cost Management Reader** | Tenant Root Group | 90-day actual cost and budgets | Cost and budgets show N/A (Advisor and waste checks still run) |
| **Directory Readers** | Entra ID role | Turns object IDs into names; finds guests and deleted identities in RBAC | RBAC shows IDs only; guest/orphan checks show N/A |

All four are read-only. Reader at the root needs a Global Administrator to elevate access once:

1. Global Admin signs in to the Azure portal → **Microsoft Entra ID → Properties** → **Access management for Azure resources** → **Yes** → Save.
2. **Management groups** → **Tenant Root Group** → **Access control (IAM)** → **Add role assignment** → add the three Azure roles above for the assessment account.
3. Set **Access management for Azure resources** back to **No**.
4. **Entra ID → Roles and administrators → Directory Readers** → add the assessment account.

Or, after step 1, from PowerShell as the Global Admin:

```powershell
$root = "/providers/Microsoft.Management/managementGroups/<tenant-id>"
foreach ($r in 'Reader', 'Security Reader', 'Cost Management Reader') {
    New-AzRoleAssignment -SignInName "assessor@customer.com" -RoleDefinitionName $r -Scope $root
}
```

If root access is not possible, assign the same three roles on each subscription in scope and run with `-SubscriptionId`.

**Other ways in:**

- **Customer runs it themselves.** Send them the script (and ARI if they will run that too). They send back the output folder. Safest for prospects.
- **Azure Lighthouse.** If Netsurit manages the subscriptions through Lighthouse, run with `-TenantId <customer tenant>`. Lighthouse covers subscriptions only, so management groups, MG-scope policy and Directory Readers will be missing.
- **CSP (admin on behalf of).** CSP subscriptions give the partner Admin Agents group Owner. That works for the inventory, but Cost Management in a CSP subscription is often not visible to the customer side. Cost comes back as a NotCollected item; ask for a cost export or invoice.

---

## 2. Workstation setup (once)

| Need | For | Install |
|---|---|---|
| PowerShell 7.2+ | All three | `winget install Microsoft.PowerShell` |
| Az.Accounts, Az.ResourceGraph | Netsurit script | Installed automatically on first run, or `Install-Module Az.Accounts, Az.ResourceGraph -Scope CurrentUser` |
| AzureResourceInventory | ARI | `Install-Module AzureResourceInventory -Scope CurrentUser` (pulls ImportExcel and the Az modules it needs) |
| Az.CostManagement | ARI `-IncludeCosts` only | `Install-Module Az.CostManagement -Scope CurrentUser` |
| WARA | WARA | `Install-Module WARA -Scope CurrentUser` |
| Windows + Microsoft Excel + PowerPoint | WARA analyser and report | WARA's analyser and report steps automate Office; they do not run on Mac or in Cloud Shell |

Run once if scripts are blocked:

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

ARI does not upgrade modules that are already installed, and both ARI and WARA check the gallery for newer versions. Before an engagement:

```powershell
Update-Module AzureResourceInventory, WARA, Az.Accounts, Az.ResourceGraph
```

WARA stops with "Module is out of date" if a newer version exists. That is expected; update and open a new PowerShell window.

---

## 3. Run order on the day

Run everything from one PowerShell 7 window, signed in once. Create a working folder per customer first:

```powershell
$cust = "Contoso"
$tid  = "<tenant-id>"
$work = "C:\Assessments\$cust\Azure"
New-Item -ItemType Directory -Path $work -Force | Out-Null
```

### Step 1: Netsurit Azure assessment (10 to 30 minutes)

```powershell
.\Azure-Tenant-Assessment-v1.ps1 -CustomerName $cust -TenantId $tid -OutputRoot $work
```

A browser sign-in opens. After that it runs unattended and opens the report when it finishes.

Common variations:

```powershell
# Pre-sales quick look: Overview, Security, Cost & Waste only
.\Azure-Tenant-Assessment-v1.ps1 -CustomerName $cust -TenantId $tid -QuickMode

# Only some subscriptions
.\Azure-Tenant-Assessment-v1.ps1 -CustomerName $cust -TenantId $tid -SubscriptionId "<sub-1>","<sub-2>"

# Browser sign-in fails (VS Code terminal, RDP session, locked-down laptop)
.\Azure-Tenant-Assessment-v1.ps1 -CustomerName $cust -TenantId $tid -UseDeviceCode

# Large estate: check diagnostic settings on more resources (one REST call each)
.\Azure-Tenant-Assessment-v1.ps1 -CustomerName $cust -TenantId $tid -MaxDiagnosticChecks 1500
```

If a session is already signed in to the same tenant, the script reuses it. Run `Disconnect-AzAccount` first to sign in as someone else.

**Check before moving on:** open `NotCollected.csv`. Every row is either an access problem to fix now (re-run), or an item for the Information Request.

### Step 2: ARI (5 minutes to 1 hour, depending on size)

```powershell
Import-Module AzureResourceInventory
Invoke-ARI -TenantID $tid -SecurityCenter -IncludeTags -ReportName "${cust}_ARI" -ReportDir "$work\ARI"
```

| Option | Use when |
|---|---|
| `-SubscriptionID <id>` / `-ManagementGroup <id>` | Scope is limited |
| `-IncludeCosts` | Cost per subscription wanted in the inventory (needs Az.CostManagement) |
| `-DiagramFullEnvironment` | The network diagram should include every component, not only the core topology |
| `-SkipDiagram` | Very large estates where diagram generation is slow |
| `-SkipAdvisory`, `-SkipPolicy` | Speed; the Netsurit script already covers Advisor and Policy |
| `-Lite` | Faster Excel without charts |
| `-DeviceLogin` | Browser sign-in fails |
| `-NoAutoUpdate` | Stops ARI updating itself mid-engagement |

Output in `$work\ARI`: `<ReportName>.xlsx` and the draw.io diagram `.xml` files. Without `-ReportDir`, ARI writes to `C:\AzureResourceInventory\`.

Open the `.xml` diagrams at https://app.diagrams.net (or with the draw.io MCP plugin) and export PNGs for the report.

### Step 3: WARA (optional, 30 minutes plus your review time)

Only for production workloads where reliability is in scope. WARA is driven per workload, so scope it to the subscriptions or resource groups that make up the workload.

**3a. Collect** (any OS, PowerShell 7):

```powershell
Set-Location "$work\WARA"
Start-WARACollector -TenantID $tid -SubscriptionIds "/subscriptions/<sub-id>"
# Narrow to a workload:
Start-WARACollector -TenantID $tid -ResourceGroups "/subscriptions/<sub-id>/resourceGroups/RG-PROD-APP"
# Specialised workloads: add -AVD, -SAP or -HPC
```

Note the `/subscriptions/` prefix on subscription IDs. Output: `WARA_File_<date>.json` in the current folder.

**3b. Analyse** (Windows with Excel):

```powershell
Start-WARAAnalyzer -JSONFile "$work\WARA\WARA_File_<date>.json"
```

Output: the WARA Action Plan / Expert-Analysis Excel file in the current folder.

**3c. Review.** Open the Excel and work through it: confirm or reject each recommendation and fill in the validation columns. This is the consultant's judgement step; do not skip it.

**3d. Report** (Windows with PowerPoint), only if you want Microsoft's deck as well as the Netsurit report:

```powershell
Start-WARAReport -ExpertAnalysisFile "$work\WARA\<action plan>.xlsx" -CustomerName $cust -WorkloadName "Production App"
```

WARA's parameters change between releases. If a command is rejected, check `Get-Help Start-WARACollector -Full` (and the same for the other two).

### Step 4: Customer workshop (2 to 3 hours)

Some CAF and WAF recommendations are about decisions and practices that the Azure APIs cannot show: recovery targets, IaC, deployment pipelines, the cost model, DR testing. Cover them with the customer using Microsoft's two free review tools. Sign in to Microsoft Learn with the Netsurit account before you start, because assessments are tied to the profile that creates them.

| Review | Framework | Link |
|---|---|---|
| **Azure Landing Zone Review** | CAF Ready, all eight design areas | https://learn.microsoft.com/assessments/21765fea-dfe6-4bc4-8bb7-db9df5a6f6c0/ |
| **Azure Well-Architected Review** (choose *Core Well-Architected Review*) | WAF, all five pillars, about 60 questions | https://learn.microsoft.com/assessments/azure-architecture-review/ |

How to run it:

1. Open the script's report on the **Framework coverage** table in the CAF/WAF appendix. Anything marked *workshop* is on the agenda.
2. Answer the questions with the customer's platform and workload owners. Use the script findings as evidence, so you are not asking things already measured.
3. Name each review `<Customer> - <scope> - <date>` and add notes to the questions you want in the report.
4. Export the recommendations (CSV) from each review into the evidence folder (section 4).

For large estates, do the Well-Architected Review per production workload, or one pillar per session.

---

## 4. Putting it into the BMAD toolkit

### Evidence folders

Copy the outputs into the engagement:

```
engagements/<Client-Engagement>/evidence/
  azure/
    assessment/   ← the whole Azure_Assessment_<Customer>_<timestamp> folder (CSVs, JSON, HTML)
    ari/          ← ARI .xlsx and diagram .xml files
    wara/         ← WARA JSON and the reviewed Action Plan .xlsx
    workshop/     ← Azure Landing Zone Review and Well-Architected Review CSV exports, workshop notes
```

Catalogue entries for `config/evidence-catalogue.yaml` (adjust to the catalogue's existing field names):

```yaml
- id: azure-assessment
  name: Azure tenant assessment (Netsurit script)
  owner: Netsurit
  folder: azure/assessment
  patterns: ["Assessment_Summary.json", "Findings.csv", "NotCollected.csv", "Azure_*.csv"]
  parser: azure
- id: azure-ari
  name: Azure Resource Inventory (ARI)
  owner: Netsurit
  folder: azure/ari
  patterns: ["*.xlsx", "*.xml"]
  parser: tabular        # inventory appendix only until an ari parser exists
- id: azure-wara
  name: Well-Architected Reliability Assessment (WARA)
  owner: Netsurit
  folder: azure/wara
  patterns: ["*.xlsx", "WARA_File_*.json"]
  parser: tabular
- id: azure-workshop
  name: CAF Azure Landing Zone Review and WAF Well-Architected Review (workshop)
  owner: Netsurit + customer
  folder: azure/workshop
  patterns: ["*.csv", "*.md", "*.docx"]
  parser: tabular        # findings from here are confidence: reported
```

In `engagement.yaml`, list the frameworks so the engine allows the references:

```yaml
frameworks: [caf, waf]       # add cis-v8 / nist-csf-2 as well if the proposal commits to them
```

### Parser: `engine/parsers/azure.py`

Same four functions as the other parsers (`parse / validate / metrics / candidates`), registered in `engine/parsers/__init__.py` as `"azure": azure`.

**What the script guarantees, so the parser can rely on it:**

- Every CSV in the file list below is always written. A header-only file means **none exist** (same rule as the M365 script). Example: header-only `Azure_LogAnalytics.csv` means no workspaces.
- If something could not be read, it is **not** silently zero: it has a row in `NotCollected.csv`, and "absence" findings (no policy assignments, no workspace, no Service Health alerts) are suppressed. Rule for the parser: a zero is only a zero if there is no NotCollected row for that section and item.
- Numbers always use a decimal point (the script forces invariant culture), so no decimal-comma handling is needed, unlike the M365 output.
- CSVs are UTF-8 with BOM (`utf-8-sig` in Python).
- Column names and file names are fixed. Treat a change to either as a breaking change and bump the script version.

**metrics():** read `Assessment_Summary.json` → `metrics` and expose each key as `azure.<Key>`, e.g. `{{ m('azure.SecureScorePct') }}`, `{{ m('azure.RunningVMsWithoutBackup') }}`, `{{ m('azure.CostTotal') }}` with `{{ m('azure.CostCurrency') }}`. A `null` value means not collected (or, for cost, more than one currency); render it as `[TBC]`, never as 0.

**candidates():** one candidate per row in `Findings.csv`.

| Findings.csv column | Candidate field |
|---|---|
| `Id` | stable finding key (e.g. `AZ-SEC-004`), so a re-run updates rather than duplicates |
| `SuggestedSeverity` | starting point only; `config/severity.yaml` and Roan decide |
| `Title`, `Detail`, `Recommendation` | draft text |
| `Count` | metric for the finding |
| `EvidenceFile` | evidence reference (the detail rows live in that CSV) |
| `Confidence` | `verified` (API data), `reported` (Advisor/Microsoft's own assessment), `inferred` (our rule, e.g. Hybrid Benefit) |
| `CafPhase`, `CafDesignArea` | CAF reference (`Workload` = workload-level, WAF only) |
| `WafPillar`, `WafCodes` | WAF reference, e.g. `Security` / `SE:05; SE:06` |
| `ReferencedTo` | ready-made reference line for the report, e.g. "CAF Secure: Identity and access management · WAF SE:05 Identity and access management" |
| `References` | space-separated Microsoft Learn URLs (CAF design area or phase page, then each WAF recommendation page) |

`NotCollected.csv` rows (Section, Item, Reason, HowToFix) become Information Request items, in the same way as the M365 script's "not available" notes.

`Assessment_Summary.json` → `frameworks` holds finding counts by `wafPillar`, `cafDesignArea` and `cafPhase` (metric keys e.g. `azure.frameworks.wafPillar.Security`). It also has an `unmapped` list, which must be empty. `Framework_Coverage.csv` states which design areas and pillars are covered by the script, ARI, WARA or the workshop. Use it as the coverage statement in the report's approach section.

**ARI and WARA** are not parsed for metrics in v1. Use ARI as the inventory appendix and source of the diagrams. Use WARA's reviewed Action Plan as evidence for reliability findings, referenced to the Well-Architected Reliability pillar. A dedicated `ari` parser is worth building once real ARI exports from two customers are available (tab names vary by version and by which resource types exist).

### Which reports it feeds

- **Infrastructure Assessment**: the Azure section (estate, governance, identity, security, network, resiliency, cost, operations, lifecycle), grouped by CAF design area, with the WAF references on each finding.
- **Azure tool report** (new, optional, like the M365 Tenant Assessment tool report): the script's nine sections and the CAF/WAF appendix reproduced in Netsurit format. Engine-only, no thinking step, the same as the other vendor reports.
- **Information Request**: everything in `NotCollected.csv`.

### Parser status line for BMAD-Assessment-Toolkit.md

> Assumed until real exports: Azure-Tenant-Assessment-v1.ps1 (tested end-to-end against mocked Azure data only), ARI .xlsx, WARA Action Plan .xlsx.

---

## 5. What the script checks

| # | Section | Checks | Main evidence files |
|---|---|---|---|
| 1 | Estate Overview | Subscriptions (state, offer type, MG path), management groups, resource counts by type/region/subscription, resource groups, VMs, Arc servers | `Azure_Subscriptions.csv`, `Azure_ManagementGroups.csv`, `Azure_ResourceCounts_*.csv` |
| 2 | Governance | Subscriptions directly under Tenant Root, MG depth, policy assignments (scope, built-in/custom, enforcement), compliance %, top non-compliant policies, tag coverage and keys, resource locks | `Azure_Policy*.csv`, `Azure_Tag*.csv`, `Azure_Locks.csv` |
| 3 | Identity & RBAC | All role assignments with names, privileged roles (Owner, UAA, RBAC Admin) by scope, guests, deleted identities, direct-to-user assignments, service principals with privileged roles, custom roles with `*`, PIM eligible assignments | `Azure_RoleAssignments.csv`, `Azure_CustomRoles.csv`, `Azure_PIM_Eligible.csv` |
| 4 | Security | Defender plans per subscription, secure score, unhealthy recommendations by severity, NSG rules open to the internet on management/database ports, public IPs on VM NICs, storage (anonymous blob, TLS, shared key, network), Key Vault (purge protection, RBAC, network), PaaS public network access | `Azure_Defender_*.csv`, `Azure_SecureScore.csv`, `Azure_NSG_InternetExposed.csv`, `Azure_StorageAccounts.csv`, `Azure_KeyVaults.csv` |
| 5 | Network Topology | Topology (Virtual WAN / hub and spoke / isolated), hub components (Azure Firewall, gateways, NVAs), peering state, overlapping address spaces, in-use subnets without NSG, DDoS protection, Network Watcher per region, private endpoints and DNS zones | `Azure_VNets.csv`, `Azure_VNetPeerings.csv`, `Azure_Subnets.csv`, `Azure_VNet_Overlaps.csv` |
| 6 | Resiliency & Backup | Azure Backup coverage of running VMs, stale/failed backups, vault soft delete and immutability, ASR replicated items, VM placement (zone / availability set / scale set / single), storage and SQL redundancy | `Azure_BackupItems.csv`, `Azure_RecoveryVaults.csv`, `Azure_ASR_ReplicatedItems.csv`, `Azure_SQLDatabases.csv` |
| 7 | Cost & Waste | Actual cost by month/service/subscription, budgets, Advisor recommendations by WAF pillar, Advisor cost savings, reservation/savings plan opportunities, unattached disks, unused public IPs and NICs, old snapshots, empty App Service plans and resource groups, stopped-not-deallocated VMs, Hybrid Benefit opportunities | `Azure_Cost_*.csv`, `Azure_Budgets.csv`, `Azure_Advisor.csv`, `Azure_Advisor_ByPillar.csv`, `Azure_Waste.csv` |
| 8 | Operations | Log Analytics workspaces and Sentinel, Azure Monitor Agent vs legacy MMA/OMS, Defender for Endpoint extension, Update Manager assessments and missing patches, activity log export, diagnostic settings on key resources, alert rules and Service Health alerts, Arc servers | `Azure_LogAnalytics.csv`, `Azure_ActivityLogExport.csv`, `Azure_DiagnosticSettings.csv`, `Azure_Alerting.csv`, `Azure_ArcMachines.csv` |
| 9 | Lifecycle & Modernisation | OS lifecycle for VMs and Arc servers, previous-generation VM sizes, retired features (Basic public IP and load balancer, unmanaged disks, classic resources, App Gateway v1, classic App Insights, GPv1 storage, Cloud Services extended support), App Service runtimes, AKS versions vs supported | `Azure_VMs.csv`, `Azure_Lifecycle_Retirements.csv`, `Azure_AppServices.csv`, `Azure_AKS.csv` |

`Azure_VMs.csv` is the joined VM view: one row per VM with size generation, OS lifecycle, power state, placement, public IP, Azure Backup, ASR, AMA, legacy agent, MDE and patch status. It is the easiest file to filter for a customer conversation.

QuickMode runs sections 1, 4 and 7. Each section heading in the report shows the CAF design area and WAF recommendations it is referenced to.

### Finding IDs

| Prefix | Section |
|---|---|
| `AZ-OVR` | Estate Overview |
| `AZ-GOV` | Governance |
| `AZ-IAM` | Identity & RBAC |
| `AZ-SEC` | Security |
| `AZ-NET` | Network Topology |
| `AZ-ADV` | Azure Advisor reliability, performance and operational excellence recommendations (by WAF pillar) |
| `AZ-RES` | Resiliency & Backup |
| `AZ-CST` | Cost & Waste |
| `AZ-OPS` | Operations |
| `AZ-LCM` | Lifecycle & Modernisation (`AZ-LCM-R01`…`R08` are the retirement items, in the order of the `$script:Retirements` table) |

IDs are stable between runs. A finding only appears when the condition is met.

### Output files

| File | Contents |
|---|---|
| `Azure_Assessment_Report.html` | Branded report, A4 print-ready (Ctrl+P → Save as PDF) |
| `Assessment_Summary.json` | Run details and headline metrics for the engine |
| `Findings.csv` | All findings with suggested severity, count, evidence file and confidence |
| `NotCollected.csv` | What could not be read, why, and how to fix it |
| `Framework_Coverage.csv` | Which CAF design areas and WAF pillars are covered by the script, ARI, WARA or the workshop |
| `Azure_*.csv` | Evidence, one file per check (see the table above) |

### Parameters

| Parameter | Default | Description |
|---|---|---|
| `-CustomerName` | | Used in the folder name and report header |
| `-TenantId` | current context | Tenant to sign in to |
| `-SubscriptionId` | all readable | Limit to one or more subscriptions |
| `-QuickMode` | off | Overview, Security, Cost & Waste only |
| `-SkipCost` | off | Skip Cost Management (Advisor and waste still run) |
| `-CostDays` | 90 | Days of actual cost, 7 to 365 |
| `-MaxDiagnosticChecks` | 300 | Resources checked for diagnostic settings; 0 skips |
| `-UseDeviceCode` | off | Device-code sign-in |
| `-OutputRoot` | Downloads | Where the output folder is created |
| `-NoOpen` | off | Do not open the report at the end |

### How it collects

Almost everything is Azure Resource Graph (KQL), one query per check, across all readable subscriptions in a single call with paging. That is why it is fast and only needs Reader. REST calls are used only where Resource Graph has no data: Cost Management query, locks, PIM eligibility schedules, activity log and resource diagnostic settings, and AKS supported versions. Microsoft Graph is used once, to resolve role assignment principals to names.

It never reads data inside resources (no blob contents, no database queries, no Key Vault secrets).

---

## 6. CAF and WAF alignment in detail

### How each finding is referenced

Each finding ID maps to three things, kept in `$script:FrameworkMap` in the script:

1. **CAF phase**, which says where the fix belongs: *Ready* (landing zone design), *Govern* (policy and cost guardrails), *Secure*, *Manage* (operations, monitoring, backup) or *Adopt* (workload migration and modernisation).
2. **CAF Azure landing zone design area**, for platform findings. Workload-level findings (e.g. a single-instance VM) are marked `Workload`, because the landing zone design areas do not cover them and WAF does.
3. **WAF design review checklist codes**, one or more per finding. A few platform-only findings (overlapping address spaces, disconnected peerings, no resource locks) have CAF references only, because WAF does not cover them.

Every finding has at least one Microsoft Learn link. The mock test fails if any finding is unmapped.

### Mapping by finding

| Findings | CAF | WAF |
|---|---|---|
| `AZ-OVR-001` disabled subscriptions | Govern · Azure billing and Microsoft Entra tenant | CO:07 |
| `AZ-GOV-001` subscriptions under root | Ready · Resource organization | SE:04 |
| `AZ-GOV-002/003/004` policy baseline, enforcement, compliance | Govern · Governance | SE:01 (+ CO:04 for 002) |
| `AZ-GOV-005` tagging | Govern · Resource organization | CO:03 |
| `AZ-GOV-006` locks | Govern · Governance | (CAF only) |
| `AZ-IAM-001…007` owners, guests, orphans, groups, custom roles, PIM, service principals | Secure / Govern / Ready · Identity and access management | SE:05 |
| `AZ-SEC-001/003` secure score, CSPM | Secure · Security | SE:01 (+ SE:10) |
| `AZ-SEC-002`, `AZ-OPS-009` Defender for Servers, MDE | Secure · Security | SE:10 |
| `AZ-SEC-004/005/008/010` internet exposure, public IPs, open storage and PaaS | Secure · Network topology and connectivity | SE:06 (+ SE:04) |
| `AZ-SEC-006/007/009`, `AZ-LCM-005` blob access, TLS, Key Vault, HTTPS | Secure · Security | SE:08 / SE:07 / SE:09 / SE:07 |
| `AZ-NET-001/002/005` overlaps, peerings, flat topology | Ready / Manage · Network topology and connectivity | (CAF only) |
| `AZ-NET-003/004/006` subnets without NSG, no central inspection, DDoS | Secure / Ready · Network topology and connectivity | SE:06 (+ SE:04) |
| `AZ-NET-007` Network Watcher | Manage · Management | OE:07 |
| `AZ-RES-001/002/003` backup coverage, failures, vault immutability | Manage · Management | RE:09 (+ RE:10) |
| `AZ-RES-004/005`, `AZ-ADV-001` single instances, LRS, Advisor reliability | Adopt / Manage · Workload | RE:05 |
| `AZ-ADV-002` Advisor performance | Manage · Workload | PE:12 |
| `AZ-ADV-003` Advisor operational excellence | Manage · Management | OE:02 |
| `AZ-CST-001/003…007` Advisor savings, waste | Manage · Workload | CO:07 |
| `AZ-CST-002/008` reservations, Hybrid Benefit | Manage · Workload | CO:05 |
| `AZ-CST-009` no budgets | Govern · Governance | CO:03, CO:04 |
| `AZ-OPS-001/002/003/006/007` workspace, AMA, legacy agent, activity log, diagnostics | Manage · Management | OE:07 (+ SE:10 for 006) |
| `AZ-OPS-004/005` missing patches, no assessment | Manage · Management | SE:08 |
| `AZ-OPS-008` Service Health alerts | Manage · Management | RE:10, OE:08 |
| `AZ-LCM-001/002/004/006` OS, runtimes, AKS out of support | Adopt · Workload | SE:08 |
| `AZ-LCM-003` previous-generation sizes | Adopt · Workload | PE:03, CO:07 |
| `AZ-LCM-R01…R08` retired features | Adopt · Workload | per item: RE:05 (Basic IP/LB, unmanaged disks), SE:08 (classic), SE:06 (App Gateway v1), OE:07 (classic App Insights), PE:03 (GPv1) |

### Coverage, and what the workshop must answer

| Area | Covered by | Workshop questions (not visible through APIs) |
|---|---|---|
| CAF Azure billing and Microsoft Entra tenant | Script + workshop | EA / MCA / CSP structure, billing ownership, number of tenants and why |
| CAF Identity and access management | Script + M365 script | Break-glass accounts, access reviews, how platform admins get access |
| CAF Resource organization | Script | Subscription vending: how new subscriptions are requested and baselined |
| CAF Network topology and connectivity | Script + ARI | Hybrid connectivity design, DNS resolution path, IP address management |
| CAF Security | Script | Security operations ownership, SIEM integration, vulnerability management process |
| CAF Management | Script | Who owns patching, backup restore tests, operations runbooks |
| CAF Governance | Script | Who approves policy exemptions, cost review cadence |
| CAF Platform automation and DevOps | **Workshop only** | Is the landing zone deployed with IaC (Bicep/Terraform), pipelines, who can change it |
| WAF Reliability | Script + WARA + workshop | RE:04 availability/recovery targets, RE:08 resilience testing, RE:09 DR plan and last test |
| WAF Security | Script + workshop | SE:02 secure development, SE:03 data classification, SE:11 testing, SE:12 incident response |
| WAF Cost Optimization | Script + workshop | CO:01 accountability, CO:02 cost model and budget ownership |
| WAF Operational Excellence | Script + workshop | OE:01-OE:06 and OE:09-OE:11: practices, IaC, pipelines, safe deployment |
| WAF Performance Efficiency | Advisor + workshop | PE:01 targets, PE:02 capacity planning, PE:06 performance testing |

### Wording rules for reports

- Write "referenced to CAF … / WAF …". Never write "compliant with", "fails", or "certified".
- Put the CAF design area in the finding header and the WAF codes in the reference line (the `ReferencedTo` column).
- Use the framework only when `engagement.yaml` → `frameworks` lists `caf` / `waf`.
- Group the Azure section of the Infrastructure Assessment by CAF design area. Recommendations follow the WAF wording on the linked page.

### Keeping it current

CAF and WAF are updated by Microsoft regularly. The codes, titles and URLs are checked against Microsoft Learn in October 2026 and live in one place: the framework block (`$script:WafCodes`, `$script:CafDesignAreaUrl`, `$script:FrameworkMap`, `$script:FrameworkCoverage`). When a new finding is added, add its ID to `$script:FrameworkMap`; the mock test reports anything unmapped.

---

## 7. Troubleshooting

| Symptom | Cause | Fix |
|---|---|---|
| Sign-in window never appears or the sign-in fails | The Windows sign-in broker does not work in some terminals (VS Code, some RDP sessions) | `-UseDeviceCode` (script) / `-DeviceLogin` (ARI) |
| Fewer subscriptions than expected | Reader not on every subscription | Reader at Tenant Root Group, or list the subscriptions with `-SubscriptionId` |
| Management groups empty, policy only at subscription level | No Reader at the root management group | Elevate access and assign at Tenant Root Group (section 1) |
| Principal names blank, guest/orphan counts N/A | No Directory Readers | Add the Entra role |
| Cost N/A with "not available for this CSP subscription" | CSP cost visibility | Ask the partner/customer for a cost export; it is listed in NotCollected |
| Defender section all N/A | No Security Reader | Add Security Reader at Tenant Root Group (or on each subscription) |
| "Diagnostic settings (partial)" in NotCollected | Estate has more key resources than `-MaxDiagnosticChecks` | Re-run with a higher value |
| OS shows "Unknown - check manually" | VM agent not reporting and no marketplace image info (custom image) | Confirm with the customer or from ARI |
| ARI runs for a very long time | Large estate or diagram generation | `-SkipDiagram`, `-Lite`, or scope with `-SubscriptionID` |
| WARA "Module is out of date" | Newer WARA in the gallery | `Update-Module WARA`, new PowerShell window |
| WARA analyser fails on Mac or Cloud Shell | Needs Windows and Excel | Copy the collector JSON to a Windows laptop with Office |

---

## 8. Maintaining the script

| Change | Where |
|---|---|
| OS lifecycle dates, retired features, outdated runtime patterns, sensitive ports, Defender plan names, privileged roles, diagnostic-settings resource types | Configuration block at the top of the script |
| A new check | The section function: query with `Invoke-ArgQuery`, export with `Export-AssessmentCsv` (always pass the column list), raise with `Add-Finding` (new stable ID), add the ID to `$script:FrameworkMap`, add a stat card or table |
| CAF/WAF references | Framework block near the top of the script (section 6) |
| Report look | `New-HtmlReport` (CSS uses the Netsurit palette; print rules are in the `@media print` block) |

Lifecycle and retirement dates were checked in October 2026. Re-check them against Microsoft Learn each quarter; the dates in the configuration block are the only place they live.

**Regression test:** `tests/mock-harness.ps1` replaces the Az cmdlets with canned data and runs the full script offline:

```powershell
pwsh -File .\tests\mock-harness.ps1 -Script .\Azure-Tenant-Assessment-v1.ps1 -Out .\tests\out
pwsh -File .\tests\mock-harness.ps1 -Script .\Azure-Tenant-Assessment-v1.ps1 -Out .\tests\out -Quick
```

Reference results on the mock data:

| Run | Findings | NotCollected | Framework check |
|---|---|---|---|
| Full | 59 (10 high) | 2 | 0 unmapped |
| `-Quick` | 22 (3 high) | 1 | 0 unmapped |
| `$env:MOCK_FAIL = 1` (no access) | 0 | 61 | 0 unmapped |

The full run also shows: secure score 48.6%, policy compliance 65.8%, spend 103,451.24 ZAR, topology "Hub and spoke". A no-access run must never raise findings, because absence findings are suppressed when the query failed.

---

## 9. What to tell the customer

> We run three read-only tools with an account you grant Reader access to. They read configuration through the Azure management APIs: what resources exist, how they are configured, who has access, what Defender for Cloud and Azure Advisor recommend, and 90 days of billed cost. They do not read data inside your resources, make no changes, and install nothing in your environment. You can remove the account's access as soon as we are done. Every recommendation is referenced to Microsoft's Cloud Adoption Framework and Well-Architected Framework, so you can follow it straight to Microsoft's own guidance.

---

## Sources

- [Cloud Adoption Framework overview](https://learn.microsoft.com/azure/cloud-adoption-framework/overview)
- [Azure landing zone design areas](https://learn.microsoft.com/azure/cloud-adoption-framework/ready/landing-zone/design-areas)
- [WAF pillars](https://learn.microsoft.com/azure/well-architected/pillars) and checklists: [Reliability](https://learn.microsoft.com/azure/well-architected/reliability/checklist), [Security](https://learn.microsoft.com/azure/well-architected/security/checklist), [Cost Optimization](https://learn.microsoft.com/azure/well-architected/cost-optimization/checklist), [Operational Excellence](https://learn.microsoft.com/azure/well-architected/operational-excellence/checklist), [Performance Efficiency](https://learn.microsoft.com/azure/well-architected/performance-efficiency/checklist)
- [Complete an Azure Well-Architected Review assessment](https://learn.microsoft.com/azure/well-architected/design-guides/implementing-recommendations)
- [Azure landing zones: validate platform readiness (Azure Landing Zone Review)](https://learn.microsoft.com/azure/migration/migrate-from-on-premises-platform-landing-zone#validate-platform-readiness)

- [Azure Resource Inventory (GitHub)](https://github.com/microsoft/ARI)
- [Azure Proactive Resiliency Library v2 – WARA tools](https://azure.github.io/Azure-Proactive-Resiliency-Library-v2/tools-preview/analyzer)
- [WARA module (PowerShell Gallery)](https://www.powershellgallery.com/packages/wara)
- [Search-AzGraph](https://learn.microsoft.com/powershell/module/az.resourcegraph/search-azgraph)
- [Resource Graph sample queries for Defender for Cloud](https://learn.microsoft.com/azure/defender-for-cloud/resource-graph-samples)
- [Query backups using Azure Resource Graph](https://learn.microsoft.com/azure/backup/query-backups-using-azure-resource-graph)
