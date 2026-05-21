# Classify-Downloads

An AI-powered PowerShell script that automatically sorts your Downloads folder into **Work**, **Personal**, and **Unclear** subfolders using the [Claude API](https://www.anthropic.com/api) to classify files by filename.

Rather than deleting files blindly, the script moves them into organised subfolders and logs every decision with the AI's reasoning — keeping you in control.

---

## How It Works

1. Scans your Downloads folder for files
2. Sends filenames to Claude (Anthropic's AI) in batches
3. Claude classifies each file as `work`, `personal`, or `unclear`
4. Files are moved into `_Work`, `_Personal`, or `_Unclear` subfolders
5. Every decision is logged to `_classifier_log.txt` with a reason

---

## Requirements

- Windows with PowerShell 5.1 or later
- An [Anthropic API key](https://console.anthropic.com/settings/keys)
- An Anthropic account with active credits ([billing](https://console.anthropic.com/settings/billing))

---

## Setup

### 1. Download the script

Save `Classify-Downloads.ps1` to a permanent location, e.g. `C:\Scripts\`.

### 2. Unblock the script

Windows may block scripts downloaded from the internet. Run once to unblock:

```powershell
Unblock-File -Path "C:\Scripts\Classify-Downloads.ps1"
```

Or allow your user account to run local scripts in general:

```powershell
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
```

### 3. Store your API key

Set your Anthropic API key as a persistent environment variable:

```powershell
[System.Environment]::SetEnvironmentVariable("ANTHROPIC_API_KEY", "sk-ant-...", "User")
```

Close and reopen PowerShell after setting this.

---

## Usage

### Dry run (preview only — no files are moved)

```powershell
.\Classify-Downloads.ps1 -DryRun
```

Always do a dry run first to review how files will be classified before committing.

### Normal run

```powershell
.\Classify-Downloads.ps1
```

### Custom Downloads folder

```powershell
.\Classify-Downloads.ps1 -DownloadsPath "D:\Downloads"
```

### Pass API key directly

```powershell
.\Classify-Downloads.ps1 -ApiKey "sk-ant-..."
```

---

## Parameters

| Parameter | Default | Description |
|---|---|---|
| `-ApiKey` | `$env:ANTHROPIC_API_KEY` | Your Anthropic API key. Falls back to environment variable if not provided. |
| `-DownloadsPath` | `$env:USERPROFILE\Downloads` | Path to the folder you want to classify. |
| `-DryRun` | `$false` | Preview classifications without moving any files. |

---

## Scheduling at Logon (Task Scheduler)

Run this once in PowerShell to create a scheduled task that fires every time you log in:

```powershell
$action   = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NonInteractive -WindowStyle Hidden -File `"C:\Scripts\Classify-Downloads.ps1`""
$trigger  = New-ScheduledTaskTrigger -AtLogOn
$settings = New-ScheduledTaskSettingsSet -ExecutionTimeLimit (New-TimeSpan -Minutes 10)

Register-ScheduledTask -TaskName "Classify Downloads" -Action $action -Trigger $trigger -Settings $settings -RunLevel Highest
```

Verify the task was created:

```powershell
Get-ScheduledTask -TaskName "Classify Downloads"
```

Test it manually without logging off:

```powershell
Start-ScheduledTask -TaskName "Classify Downloads"
```

---

## Output Folders

After running, your Downloads folder will contain:

```
Downloads\
├── _Work\          ← Work/business related files
├── _Personal\      ← Clearly personal files
├── _Unclear\       ← Ambiguous — review these manually
└── _classifier_log.txt
```

Files that can't be confidently classified land in `_Unclear` for you to handle manually. This is intentional — the script is conservative by design.

---

## Customisation

### Change the AI model

In the script, find this line and swap in a different [Claude model](https://docs.anthropic.com/en/docs/about-claude/models/overview):

```powershell
$Model = "claude-sonnet-4-20250514"
```

### Adjust batch size

Controls how many filenames are sent per API call (default: 50):

```powershell
$BatchSize = 50
```

Reduce if you hit API errors on large folders. Increase if you want fewer API calls.

### Change classification behaviour

The classification prompt is in the `Invoke-ClaudeClassify` function. Edit the prompt text to adjust how Claude makes decisions — for example, you could add domain-specific rules like company names, project codes, or file naming conventions relevant to your workflow.

### Keep the log file small

The log is trimmed to the last 2000 lines automatically. Adjust this with:

```powershell
$MaxLogLines = 2000
```

---

## Notes

- The script skips files whose names start with `_` (i.e. its own subfolders and log file)
- Hidden files are skipped
- If a destination filename already exists, a numeric suffix is appended automatically (e.g. `report_1.pdf`)
- The script never permanently deletes files — everything is moved, not removed

---

## License

MIT — free to use, modify, and distribute.
