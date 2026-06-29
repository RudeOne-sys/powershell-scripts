# FileServer-MigrationAssessment.ps1

**Version:** 1.0  
**Purpose:** Read-only pre-migration assessment of a Windows file server to scope a SharePoint Online migration project.

---

## What the Script Does

Scans all SMB shares on a target server and collects data across nine categories, writing everything to a single filterable CSV file. It performs **no write, modify, delete, or move operations** on any files or shares and is safe to run against a production file server during business hours.

A companion error log captures any access-denied paths or unreadable items without crashing the script.

---

## What It Collects

| CSV Section | What is captured |
|---|---|
| `SHARE_SUMMARY` | Share names, paths, descriptions, share-level permissions, NTFS root permissions, HIGH RISK flags for overly broad access, and total data size/file/folder counts per share |
| `FILE_AGE` | File counts and sizes grouped into four age buckets (Under 1 yr / 1–3 yr / 3–5 yr / Over 5 yr), oldest and newest modified dates, and count of files not accessed in over 1 year |
| `LARGEST_FILES` | Top 20 largest files per share — full path, size in MB, last modified date, extension. Files over 250 MB flagged `OVER250MB`; files over 15 GB flagged `OVER15GB` |
| `FILE_TYPES` | Every file extension found, with total count and size. Flags: `BLOCKED by SPO` (e.g. .tmp, thumbs.db), `NEEDS ATTENTION` (e.g. .pst, .accdb), `EXECUTABLE` (e.g. .exe, .ps1) |
| `SPO_VIOLATIONS` | Files and folders incompatible with SharePoint Online naming rules — illegal characters, trailing spaces/periods, reserved names (CON, NUL, _vti_, etc.), and paths exceeding 400 characters |
| `PERMISSIONS` | Unique ACE count per share (complexity rating LOW/MEDIUM/HIGH), folders with broken inheritance, unique AD identities referenced, and any local (non-domain) accounts found in NTFS ACLs |
| `DUPLICATES` | Files sharing the same name and file size within a share (no hashing — lightweight detection) |
| `EMPTY_FOLDERS` | Count and percentage of empty folders per share |
| `RUN_SUMMARY` | Total runtime, shares scanned, error count, and output file paths |

---

## Minimum Permissions Required

| Task | Requirement |
|---|---|
| File/folder scanning, NTFS ACL reading | Standard domain user with Read access to the shares |
| Share enumeration (`Get-SmbShare`) | Local administrator on the target file server |
| Share-level permission reading (`Get-SmbShareAccess`) | Local administrator on the target file server |

**Recommended:** Create a dedicated service account (e.g. `DOMAIN\SVC_FSAssessment`), add it to the local Administrators group on the file server, and grant it at least Read & Execute on all share root paths.

---

## How to Run

### Basic — scan the local machine

```powershell
.\FileServer-MigrationAssessment.ps1
```

### Remote file server

```powershell
.\FileServer-MigrationAssessment.ps1 -ServerName FILESVR01
```

### Remote server with custom output folder and specific shares only

```powershell
.\FileServer-MigrationAssessment.ps1 `
    -ServerName FILESVR01 `
    -OutputPath D:\Reports `
    -ShareList @("Finance", "HR", "Projects")
```

### With a custom SharePoint Online base URL (for path-length calculations)

```powershell
.\FileServer-MigrationAssessment.ps1 `
    -ServerName FILESVR01 `
    -SPOBaseUrl "https://contoso.sharepoint.com/sites/migration"
```

### Run as a service account

```powershell
Start-Process powershell `
    -Credential (Get-Credential DOMAIN\SVC_FSAssessment) `
    -ArgumentList "-File `"C:\Scripts\FileServer-MigrationAssessment.ps1`" -ServerName FILESVR01 -OutputPath D:\Reports"
```

### Check execution policy before running

```powershell
Get-ExecutionPolicy
# If the result is "Restricted", allow scripts for this session only:
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
```

---

## Output Files

Both files are written to `-OutputPath` (default: `C:\`).

| File | Description |
|---|---|
| `FileServer_Assessment_<yyyyMMdd_HHmmss>.csv` | All assessment data in a single file, filterable by the `Section` column |
| `FileServer_Assessment_<yyyyMMdd_HHmmss>_ERRORS.log` | Timestamped log of any access-denied paths or read failures |

### Filtering the CSV in Excel

1. Open the CSV in Excel.
2. Select the header row and choose **Data → Filter**.
3. Use the dropdown on the **Section** column to isolate any category (e.g. `SPO_VIOLATIONS`, `LARGEST_FILES`).

---

## How to Generate a Migration Report Using Claude

Once the script has run and you have the `.csv` file, you can use Claude to analyse it and produce a migration scoping report.

### Step 1 — Open Claude

Go to [claude.ai](https://claude.ai) or open the Claude desktop app.

### Step 2 — Upload the CSV

Click the **paperclip / attachment icon** in the chat input and select your `FileServer_Assessment_<timestamp>.csv` file.

### Step 3 — Use one of the prompts below

Copy and paste the relevant prompt after uploading the file.

---

#### Executive Summary Report

```
I have attached a CSV output from a file server assessment tool scoping a SharePoint Online 
migration. The CSV has a "Section" column that groups rows into categories: SHARE_SUMMARY, 
FILE_AGE, LARGEST_FILES, FILE_TYPES, SPO_VIOLATIONS, PERMISSIONS, DUPLICATES, EMPTY_FOLDERS.

Please produce an executive summary report covering:
1. Total data volume and file count across all shares
2. Data age profile — how much data is stale (over 3 years old) vs. active
3. Top SharePoint Online compatibility risks (path violations, blocked file types, oversized files)
4. Permission complexity assessment
5. Overall migration complexity rating (Low / Medium / High) with justification
6. Recommended next steps

Write it in plain business language suitable for a non-technical stakeholder.
```

---

#### Technical Migration Scoping Report

```
I have attached a CSV output from a file server assessment. The Section column organises 
the data into: SHARE_SUMMARY, FILE_AGE, LARGEST_FILES, FILE_TYPES, SPO_VIOLATIONS, 
PERMISSIONS, DUPLICATES, EMPTY_FOLDERS.

Please produce a detailed technical migration scoping report with the following sections:

1. Share inventory — list each share, its size, and file count
2. Data age analysis — summarise by age bucket per share and identify candidates for 
   archiving rather than migration
3. SharePoint Online blockers — list all SPO_VIOLATIONS with counts, all BLOCKED file 
   types, and all files exceeding 250 MB or 15 GB
4. Permission complexity — flag shares with HIGH complexity or broken inheritance
5. Duplicate data — estimate storage savings if duplicates were removed before migration
6. Remediation checklist — ordered list of actions that should be completed before 
   migration begins
7. Estimated migration wave sizing — suggest grouping shares into migration waves based 
   on size and complexity
```

---

#### SharePoint Online Blockers Only

```
I have attached a file server assessment CSV. Please focus only on SharePoint Online 
compatibility issues. Using the SPO_VIOLATIONS and FILE_TYPES sections:

1. List every unique violation type and how many files/folders are affected
2. List all file extensions flagged as BLOCKED by SPO and their file counts
3. List all files flagged OVER15GB or OVER250MB from the LARGEST_FILES section
4. For each issue, provide the specific remediation action required before migration
5. Produce a summary count: total blockers that MUST be resolved vs. items that need 
   attention but are not hard blockers
```

---

#### Stale Data & Archiving Candidates Report

```
I have attached a file server assessment CSV. Using the FILE_AGE section, please:

1. Identify which shares have the highest proportion of files older than 3 years
2. Calculate the total storage that could be archived rather than migrated to SharePoint Online
3. Recommend a data retention policy approach for each share based on the age distribution
4. Estimate the reduction in migration scope (GB and file count) if files older than 5 years 
   are archived instead of migrated
```

---

### Tips for Getting the Best Results from Claude

- If the CSV is large (many shares or thousands of violations), ask Claude to focus on one section at a time.
- Ask Claude to output tables where possible — they are easier to paste into reports.
- Ask for a "RAG status" (Red / Amber / Green) per share to quickly communicate risk to stakeholders.
- You can follow up with specific questions after the initial report, for example: *"Which share should we migrate first based on the data?"* or *"Write a remediation email to the IT team listing the files that need to be renamed before migration."*
