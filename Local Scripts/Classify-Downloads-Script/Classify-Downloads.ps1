# =============================================================================
# Classify-Downloads.ps1
# AI-powered Downloads folder classifier using the Claude API
# Sorts files into _Work, _Personal, and _Unclear subfolders
# =============================================================================
#
# USAGE:
#   .\Classify-Downloads.ps1
#   .\Classify-Downloads.ps1 -DryRun                          # Preview only, no moves
#   .\Classify-Downloads.ps1 -ApiKey "sk-ant-..."             # Pass key directly
#   .\Classify-Downloads.ps1 -DownloadsPath "D:\Downloads"    # Custom folder
#
# SETUP:
#   Set your API key as a persistent environment variable (recommended):
#     [System.Environment]::SetEnvironmentVariable("ANTHROPIC_API_KEY", "sk-ant-...", "User")
#   Then restart PowerShell.
#
# SCHEDULING (Task Scheduler):
#   Action: powershell.exe
#   Arguments: -NonInteractive -WindowStyle Hidden -File "C:\Scripts\Classify-Downloads.ps1"
# =============================================================================

param(
    [string]$ApiKey       = $env:ANTHROPIC_API_KEY,
    [string]$DownloadsPath = "E:\Downloads",
    [switch]$DryRun
)

# --------------------------------------------------------------------------
# Configuration
# --------------------------------------------------------------------------
$WorkFolder     = Join-Path $DownloadsPath "_Work"
$PersonalFolder = Join-Path $DownloadsPath "_Personal"
$UnclearFolder  = Join-Path $DownloadsPath "_Unclear"
$LogFile        = Join-Path $DownloadsPath "_classifier_log.txt"
$Model          = "claude-sonnet-4-20250514"
$BatchSize      = 50      # Files per API call (keeps prompts manageable)
$MaxLogLines    = 2000    # Trim log file if it grows beyond this

# --------------------------------------------------------------------------
# Validate inputs
# --------------------------------------------------------------------------
if (-not $ApiKey) {
    Write-Error @"
No API key found.
Set it permanently with:
  [System.Environment]::SetEnvironmentVariable('ANTHROPIC_API_KEY', 'sk-ant-...', 'User')
Or pass it directly: .\Classify-Downloads.ps1 -ApiKey 'sk-ant-...'
"@
    exit 1
}

if (-not (Test-Path $DownloadsPath)) {
    Write-Error "Downloads folder not found: $DownloadsPath"
    exit 1
}

# --------------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------------
function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $entry = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$Level] $Message"
    Write-Host $entry
    Add-Content -Path $LogFile -Value $entry -Encoding UTF8
}

function Trim-Log {
    if (Test-Path $LogFile) {
        $lines = Get-Content $LogFile
        if ($lines.Count -gt $MaxLogLines) {
            $lines | Select-Object -Last $MaxLogLines | Set-Content $LogFile -Encoding UTF8
        }
    }
}

function Get-SafeDestination {
    param([string]$Folder, [string]$Filename)
    $dest = Join-Path $Folder $Filename
    if (-not (Test-Path $dest)) { return $dest }
    $base = [System.IO.Path]::GetFileNameWithoutExtension($Filename)
    $ext  = [System.IO.Path]::GetExtension($Filename)
    $i    = 1
    do {
        $dest = Join-Path $Folder "${base}_${i}${ext}"
        $i++
    } while (Test-Path $dest)
    return $dest
}

function Invoke-ClaudeClassify {
    param([string[]]$Filenames)

    $list   = ($Filenames | ForEach-Object { "- $_" }) -join "`n"
    $prompt = @"
You are a file classifier. Classify each filename from a Downloads folder as:
- "work"     : clearly work/business related — invoices, contracts, proposals, SOWs,
               reports, technical docs, IT tools, company/client names, CVs, etc.
- "personal" : clearly personal — holiday photos, music, personal hobby files,
               family documents, personal receipts, entertainment downloads, etc.
- "unclear"  : genuinely ambiguous from the filename alone

IMPORTANT: When in doubt, use "unclear". Only classify as "work" or "personal" when
you are confident. Do not guess.

Return ONLY a valid JSON array with no markdown, no explanation, no backticks.
Each element: {"filename": "...", "category": "work|personal|unclear", "reason": "brief one-line reason"}

Filenames to classify:
$list
"@

    $body = @{
        model      = $Model
        max_tokens = 2048
        messages   = @(@{ role = "user"; content = $prompt })
    } | ConvertTo-Json -Depth 5

    $headers = @{
        "x-api-key"         = $ApiKey
        "anthropic-version" = "2023-06-01"
        "content-type"      = "application/json"
    }

    try {
        $response = Invoke-RestMethod `
            -Uri     "https://api.anthropic.com/v1/messages" `
            -Method  POST `
            -Headers $headers `
            -Body    $body `
            -ErrorAction Stop

        $raw = $response.content[0].text.Trim()

        # Strip any accidental markdown fences
        $raw = $raw -replace '```json', '' -replace '```', ''

        return $raw | ConvertFrom-Json
    }
    catch {
        Write-Log "API call failed: $_" "ERROR"
        return $null
    }
}

# --------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------
Write-Log "=== Classifier run started$(if ($DryRun) { ' [DRY RUN]' }) ==="

# Create output folders
foreach ($folder in @($WorkFolder, $PersonalFolder, $UnclearFolder)) {
    if (-not (Test-Path $folder)) {
        if (-not $DryRun) { New-Item -ItemType Directory -Path $folder | Out-Null }
        Write-Log "Created folder: $folder"
    }
}

# Collect files — skip our own subfolders, log file, and hidden files
$files = Get-ChildItem -Path $DownloadsPath -File |
         Where-Object { $_.Name -notlike "_*" -and -not $_.Attributes.HasFlag([IO.FileAttributes]::Hidden) }

if ($files.Count -eq 0) {
    Write-Log "No files found to classify. Exiting."
    exit 0
}

Write-Log "Found $($files.Count) file(s) to classify."

# Counters
$moved   = @{ work = 0; personal = 0; unclear = 0 }
$skipped = 0

# Process in batches
$batches = [math]::Ceiling($files.Count / $BatchSize)

for ($b = 0; $b -lt $batches; $b++) {
    $batch     = $files | Select-Object -Skip ($b * $BatchSize) -First $BatchSize
    $batchNums = "$($b * $BatchSize + 1)-$([math]::Min(($b + 1) * $BatchSize, $files.Count))"
    Write-Log "Classifying batch $($b + 1)/$batches (files $batchNums)..."

    $results = Invoke-ClaudeClassify -Filenames ($batch | Select-Object -ExpandProperty Name)

    if (-not $results) {
        Write-Log "Batch $($b + 1) returned no results — skipping." "WARN"
        $skipped += $batch.Count
        continue
    }

    # Build a lookup by filename for quick matching
    $lookup = @{}
    foreach ($r in $results) { $lookup[$r.filename] = $r }

    foreach ($file in $batch) {
        $result = $lookup[$file.Name]

        if (-not $result) {
            Write-Log "No classification returned for: $($file.Name) — moving to _Unclear" "WARN"
            $result = [PSCustomObject]@{ filename = $file.Name; category = "unclear"; reason = "Not returned by API" }
        }

        $destFolder = switch ($result.category) {
            "work"     { $WorkFolder }
            "personal" { $PersonalFolder }
            default    { $UnclearFolder }
        }

        $destPath = Get-SafeDestination -Folder $destFolder -Filename $file.Name
        $tag      = $result.category.ToUpper().PadRight(8)
        Write-Log "$tag $($file.Name)  |  $($result.reason)"

        if (-not $DryRun) {
            try {
                Move-Item -Path $file.FullName -Destination $destPath -ErrorAction Stop
                $moved[$result.category]++
            }
            catch {
                Write-Log "Failed to move '$($file.Name)': $_" "ERROR"
                $skipped++
            }
        }
        else {
            $moved[$result.category]++  # Count even in dry run for the summary
        }
    }
}

# --------------------------------------------------------------------------
# Summary
# --------------------------------------------------------------------------
$dryLabel = if ($DryRun) { " (DRY RUN — no files were actually moved)" } else { "" }
Write-Log "=== Run complete$dryLabel ==="
Write-Log "  Work:     $($moved.work) file(s)"
Write-Log "  Personal: $($moved.personal) file(s)"
Write-Log "  Unclear:  $($moved.unclear) file(s)"
if ($skipped -gt 0) { Write-Log "  Skipped:  $skipped file(s)" "WARN" }
Write-Log "Log: $LogFile"

Trim-Log