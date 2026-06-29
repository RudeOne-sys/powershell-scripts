<#
.SYNOPSIS
    Read-only file server assessment for SharePoint Online migration scoping.

.DESCRIPTION
    Performs a non-destructive inventory of all SMB shares on a Windows file server,
    collecting data to scope a SharePoint Online migration. Output goes to a single
    sectioned CSV file plus an error log.

    THIS SCRIPT PERFORMS NO WRITE, MODIFY, DELETE, OR MOVE OPERATIONS.

.PARAMETER ServerName
    Target server to assess. Defaults to localhost.

.PARAMETER ShareList
    Optional array of specific share names to scan. If omitted, all non-admin shares are scanned.

.PARAMETER OutputPath
    Directory where the CSV and error log are written. Defaults to C:\.

.PARAMETER SPOBaseUrl
    Base SharePoint Online URL used for path-length calculations.
    Default: https://tenant.sharepoint.com/sites/sitename

.EXAMPLE
    .\FileServer-MigrationAssessment.ps1

.EXAMPLE
    .\FileServer-MigrationAssessment.ps1 -ServerName FILESVR01 -OutputPath D:\Reports

.EXAMPLE
    .\FileServer-MigrationAssessment.ps1 -ServerName FILESVR01 -ShareList @("Finance","HR","Projects")

.NOTES
    Script Name : FileServer-MigrationAssessment.ps1
    Version     : 1.0
    Author      : <Author Placeholder>
    Purpose     : Pre-migration assessment for SharePoint Online scoping

    =========================================================================
    MINIMUM PERMISSIONS REQUIRED
    =========================================================================
    Most sections: Standard domain user with Read permission on all shares.

    Sections that MAY require elevation or additional rights:
      - Share Inventory (Get-SmbShare): Requires local admin or membership in
        the local Administrators group on the target server to enumerate all
        shares via WMI/CIM. A domain user can only see shares they have access
        to via net view, which may miss some shares.
      - Share-Level Permissions (Get-SmbShareAccess): Requires local admin on
        the target server.
      - NTFS ACL reading: Works with Read access to the folder. Folders with
        restricted ACLs (e.g., hidden from non-owners) will generate Access
        Denied errors that are caught and logged.
      - LastAccessTime: May be disabled by default on Windows Server 2008+
        (NtfsDisableLastAccessUpdate = 1). If so, LAT data will not be
        meaningful. No elevated rights needed to READ it; only to re-enable it.

    RECOMMENDED SERVICE ACCOUNT SETUP:
      1. Create a dedicated domain service account (e.g., SVC_FSAssessment).
      2. Add it to the local Administrators group on the file server being assessed
         (required for share enumeration and share-level permission reading).
      3. Ensure it has at least Read & Execute NTFS rights on all share root paths.
      4. Run the script under that account:
           Start-Process powershell -Credential (Get-Credential DOMAIN\SVC_FSAssessment) `
               -ArgumentList "-File .\FileServer-MigrationAssessment.ps1 -ServerName FILESVR01"
    =========================================================================
#>

[CmdletBinding()]
param(
    [string]$ServerName  = 'localhost',
    [string[]]$ShareList = @(),
    [string]$OutputPath  = 'C:\',
    [string]$SPOBaseUrl  = 'https://tenant.sharepoint.com/sites/sitename'
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'SilentlyContinue'

#region -- Helpers & Constants ------------------------------------------------

$Timestamp     = Get-Date -Format 'yyyyMMdd_HHmmss'
$CsvPath       = Join-Path $OutputPath "FileServer_Assessment_$Timestamp.csv"
$ErrorLogPath  = Join-Path $OutputPath "FileServer_Assessment_${Timestamp}_ERRORS.log"
$ScriptStart   = Get-Date
$CsvRows       = [System.Collections.Generic.List[PSObject]]::new()
$TotalErrors   = 0

# Admin shares to always exclude
$ExcludedShares = @('C$','D$','E$','F$','ADMIN$','IPC$','SYSVOL','NETLOGON','PRINT$')

# SPO illegal characters in file/folder names
$SpoIllegalChars = [char[]]('~','"','#','%','&','*',':','<','>','?','/','\','{','|','}')

# SPO reserved names (case-insensitive)
$SpoReservedNames = @(
    '_vti_','.lock','CON','PRN','AUX','NUL',
    'COM1','COM2','COM3','COM4','COM5','COM6','COM7','COM8','COM9',
    'LPT1','LPT2','LPT3','LPT4','LPT5','LPT6','LPT7','LPT8','LPT9'
)

# SPO flagged extensions
$SpoBlocked    = @('.tmp','.ds_store','desktop.ini','thumbs.db','ehthumbs.db')
$SpoAttention  = @('.pst','.lnk','.url','.one','.accdb','.mdb')
$SpoExeTypes   = @('.exe','.dll','.bat','.ps1','.vbs','.cmd','.msi','.msp')

function Write-ErrorLog {
    param([string]$Message)
    $script:TotalErrors++
    $line = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $Message"
    Add-Content -Path $ErrorLogPath -Value $line -ErrorAction SilentlyContinue
}

function Add-CsvRow {
    param([hashtable]$Row)
    $script:CsvRows.Add([PSCustomObject]$Row)
}

function Format-GB { param([long]$Bytes) [math]::Round($Bytes / 1GB, 3) }
function Format-MB { param([long]$Bytes) [math]::Round($Bytes / 1MB, 2) }

function Get-SpoViolation {
    param([string]$Name, [string]$FullPath)
    $violations = [System.Collections.Generic.List[string]]::new()

    foreach ($ch in $SpoIllegalChars) {
        if ($Name.Contains($ch)) { $violations.Add("IllegalChar:[$ch]") }
    }
    if ($Name -match '[ .]$') { $violations.Add('TrailingSpaceOrPeriod') }
    foreach ($r in $SpoReservedNames) {
        if ($Name -ieq $r -or $Name.StartsWith($r, [System.StringComparison]::OrdinalIgnoreCase)) {
            $violations.Add("ReservedName:$r")
        }
    }
    if (($SPOBaseUrl.TrimEnd('/') + '/' + $FullPath).Length -gt 400) {
        $violations.Add('PathTooLong(>400)')
    }
    return $violations
}

function Is-EmptyFolder {
    param([string]$Path)
    $hasItems = Get-ChildItem -LiteralPath $Path -Recurse -File -ErrorAction SilentlyContinue |
                Select-Object -First 1
    return ($null -eq $hasItems)
}

#endregion

#region -- Share Enumeration --------------------------------------------------

Write-Host "`n=====================================================" -ForegroundColor Cyan
Write-Host "  File Server Migration Assessment" -ForegroundColor Cyan
Write-Host "  Target: $ServerName  |  Started: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Cyan
Write-Host "=====================================================" -ForegroundColor Cyan
Write-Host "  CSV  : $CsvPath" -ForegroundColor Gray
Write-Host "  Errors: $ErrorLogPath`n" -ForegroundColor Gray

Write-Progress -Activity 'File Server Assessment' -Status 'Enumerating shares...' -PercentComplete 1

$AllShares = @()
try {
    if ($ServerName -eq 'localhost' -or $ServerName -eq $env:COMPUTERNAME) {
        $AllShares = Get-SmbShare -ErrorAction Stop |
                     Where-Object { $_.Name -notin $ExcludedShares -and $_.Name -notmatch '\$$' }
    } else {
        $AllShares = Get-SmbShare -CimSession $ServerName -ErrorAction Stop |
                     Where-Object { $_.Name -notin $ExcludedShares -and $_.Name -notmatch '\$$' }
    }
} catch {
    Write-ErrorLog "Share enumeration failed: $($_.Exception.Message)"
    Write-Warning "Could not enumerate shares via Get-SmbShare. Attempting fallback via WMI..."
    try {
        $wmiShares = Get-WmiObject -Class Win32_Share -ComputerName $ServerName -ErrorAction Stop |
                     Where-Object { $_.Name -notin $ExcludedShares -and $_.Name -notmatch '\$$' -and $_.Type -eq 0 }
        $AllShares = $wmiShares | Select-Object @{N='Name';E={$_.Name}},
                                                @{N='Path';E={$_.Path}},
                                                @{N='Description';E={$_.Description}}
    } catch {
        Write-ErrorLog "WMI share enumeration also failed: $($_.Exception.Message)"
        Write-Error "Cannot enumerate shares on $ServerName. Check permissions and connectivity."
        exit 1
    }
}

if ($ShareList.Count -gt 0) {
    $AllShares = $AllShares | Where-Object { $ShareList -contains $_.Name }
}

$ShareCount = ($AllShares | Measure-Object).Count
Write-Host "Found $ShareCount shares to assess.`n" -ForegroundColor Green

#endregion

#region -- Per-Share Processing -----------------------------------------------

$ShareIndex = 0
foreach ($Share in $AllShares) {
    $ShareIndex++
    $ShareName = $Share.Name
    $SharePath = $Share.Path
    $ShareDesc = $Share.Description

    $ShareStart = Get-Date
    Write-Host "[$ShareIndex/$ShareCount] Processing share: $ShareName ($SharePath)" -ForegroundColor Yellow
    Write-Progress -Activity 'File Server Assessment' `
                   -Status "Share $ShareIndex of $ShareCount : $ShareName" `
                   -PercentComplete ([int](($ShareIndex - 1) / $ShareCount * 100))

    # Resolve UNC path for remote servers
    $ScanPath = if ($ServerName -notin @('localhost',$env:COMPUTERNAME)) {
        "\\$ServerName\$ShareName"
    } else {
        $SharePath
    }

    if (-not (Test-Path -LiteralPath $ScanPath -ErrorAction SilentlyContinue)) {
        Write-ErrorLog "Share path inaccessible: $ScanPath"
        Write-Warning "  Skipping $ShareName -- path not accessible."
        continue
    }

    #region -- 1. Share-Level Permissions -------------------------------------

    Write-Progress -Activity 'File Server Assessment' `
                   -Status "[$ShareName] Reading share permissions..." -PercentComplete 5

    $SharePerms = @()
    try {
        $SharePerms = Get-SmbShareAccess -Name $ShareName -ErrorAction Stop
    } catch {
        try {
            $SharePerms = Get-SmbShareAccess -Name $ShareName -CimSession $ServerName -ErrorAction Stop
        } catch {
            Write-ErrorLog "Cannot read share permissions for $ShareName : $($_.Exception.Message)"
        }
    }

    foreach ($perm in $SharePerms) {
        $isHighRisk = ($perm.AccountName -match 'Everyone|Authenticated Users') -and
                      ($perm.AccessRight -match 'Full|Change')
        Add-CsvRow @{
            Section             = 'SHARE_SUMMARY'
            ShareName           = $ShareName
            SharePath           = $SharePath
            ShareDescription    = $ShareDesc
            PermType            = 'Share'
            Identity            = $perm.AccountName
            Rights              = $perm.AccessRight
            AccessControlType   = $perm.AccessControlType
            IsInherited         = ''
            HighRiskFlag        = if ($isHighRisk) { 'HIGH RISK - Broad Write Access' } else { '' }
            Detail              = ''
        }
    }

    # NTFS root permissions
    try {
        $acl = Get-Acl -LiteralPath $ScanPath -ErrorAction Stop
        foreach ($ace in $acl.Access) {
            $isHighRisk = ($ace.IdentityReference -match 'Everyone|Authenticated Users') -and
                          ($ace.FileSystemRights  -match 'FullControl|Modify|Write')
            Add-CsvRow @{
                Section           = 'SHARE_SUMMARY'
                ShareName         = $ShareName
                SharePath         = $SharePath
                ShareDescription  = $ShareDesc
                PermType          = 'NTFS-Root'
                Identity          = $ace.IdentityReference
                Rights            = $ace.FileSystemRights
                AccessControlType = $ace.AccessControlType
                IsInherited       = $ace.IsInherited
                HighRiskFlag      = if ($isHighRisk) { 'HIGH RISK - Broad Write Access' } else { '' }
                Detail            = ''
            }
        }
    } catch {
        Write-ErrorLog "Cannot read NTFS ACL on $ScanPath : $($_.Exception.Message)"
    }

    #endregion

    #region -- 2 & 3. File Enumeration: Size, Count, Age ---------------------

    Write-Progress -Activity 'File Server Assessment' `
                   -Status "[$ShareName] Enumerating files (this may take a while)..." -PercentComplete 10

    $Now          = Get-Date
    $OneYearAgo   = $Now.AddYears(-1)
    $ThreeYearsAgo= $Now.AddYears(-3)
    $FiveYearsAgo = $Now.AddYears(-5)

    # Accumulators
    $TotalBytes   = [long]0
    $TotalFiles   = [long]0
    $TotalFolders = [long]0
    $PathViolCount= [long]0

    $Age_Under1yr_Count    = [long]0 ; $Age_Under1yr_Bytes    = [long]0
    $Age_1to3yr_Count      = [long]0 ; $Age_1to3yr_Bytes      = [long]0
    $Age_3to5yr_Count      = [long]0 ; $Age_3to5yr_Bytes      = [long]0
    $Age_Over5yr_Count     = [long]0 ; $Age_Over5yr_Bytes     = [long]0
    $NotAccessed1yr_Count  = [long]0
    $OldestDate            = $Now
    $NewestDate            = [datetime]::MinValue

    # Top 20 largest files
    $LargestFiles = [System.Collections.Generic.SortedList[long,PSObject]]::new()

    # File type buckets: extension -> [count, bytes]
    $ExtMap  = [System.Collections.Generic.Dictionary[string,long[]]]::new()

    # Duplicate detection: "name|size" -> list of paths
    $DupMap  = [System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[string]]]::new()

    # SPO violations list
    $SpoViolations = [System.Collections.Generic.List[PSObject]]::new()

    # Permission complexity
    $UniqueAces     = [System.Collections.Generic.HashSet[string]]::new()
    $BrokenInherit  = [System.Collections.Generic.List[string]]::new()
    $UniqueIdentities = [System.Collections.Generic.HashSet[string]]::new()
    $LocalAccounts  = [System.Collections.Generic.List[string]]::new()

    # Empty folder tracking
    $EmptyFolderCount = [long]0

    # -- Folder pass ----------------------------------------------------------
    $Folders = Get-ChildItem -LiteralPath $ScanPath -Recurse -Directory -ErrorAction SilentlyContinue
    foreach ($folder in $Folders) {
        $TotalFolders++

        # SPO violations on folder name
        $fViolations = Get-SpoViolation -Name $folder.Name -FullPath ($folder.FullName -replace [regex]::Escape($ScanPath),'')
        foreach ($v in $fViolations) {
            $PathViolCount++
            $SpoViolations.Add([PSCustomObject]@{
                ShareName = $ShareName
                Type      = 'Folder'
                Path      = $folder.FullName
                Violation = $v
            })
        }

        # NTFS ACL for permission complexity
        try {
            $facl = Get-Acl -LiteralPath $folder.FullName -ErrorAction Stop
            foreach ($ace in $facl.Access) {
                $key = "$($ace.IdentityReference)|$($ace.FileSystemRights)|$($ace.AccessControlType)"
                $UniqueAces.Add($key) | Out-Null
                $UniqueIdentities.Add($ace.IdentityReference.ToString()) | Out-Null

                # Detect local accounts (no domain prefix, no well-known SID prefixes)
                $id = $ace.IdentityReference.ToString()
                if ($id -notmatch '\\' -and $id -notmatch '^(NT |BUILTIN|Everyone|Creator|SELF)') {
                    $LocalAccounts.Add($id) | Out-Null
                }
            }
            if (-not $facl.AreAccessRulesProtected -eq $false) {
                # Inheritance is protected (broken) on this folder
            }
            if ($facl.AreAccessRulesProtected) {
                $BrokenInherit.Add($folder.FullName) | Out-Null
            }
        } catch {
            Write-ErrorLog "ACL read failed: $($folder.FullName) -- $($_.Exception.Message)"
        }

        # Empty folder check
        if (Is-EmptyFolder -Path $folder.FullName) {
            $EmptyFolderCount++
        }
    }

    # -- File pass ------------------------------------------------------------
    $Files = Get-ChildItem -LiteralPath $ScanPath -Recurse -File -ErrorAction SilentlyContinue
    foreach ($file in $Files) {
        $TotalFiles++
        $fBytes = $file.Length
        $TotalBytes += $fBytes

        # Age buckets
        $lastMod = $file.LastWriteTime
        if ($lastMod -gt $OneYearAgo)      { $Age_Under1yr_Count++; $Age_Under1yr_Bytes += $fBytes }
        elseif ($lastMod -gt $ThreeYearsAgo){ $Age_1to3yr_Count++;  $Age_1to3yr_Bytes  += $fBytes }
        elseif ($lastMod -gt $FiveYearsAgo) { $Age_3to5yr_Count++;  $Age_3to5yr_Bytes  += $fBytes }
        else                                { $Age_Over5yr_Count++;  $Age_Over5yr_Bytes += $fBytes }

        if ($lastMod -lt $OldestDate) { $OldestDate = $lastMod }
        if ($lastMod -gt $NewestDate) { $NewestDate = $lastMod }

        # Last access time
        if ($file.LastAccessTime -lt $OneYearAgo) { $NotAccessed1yr_Count++ }

        # Extension map
        $ext = $file.Extension.ToLower()
        if (-not $ExtMap.ContainsKey($ext)) { $ExtMap[$ext] = @([long]0,[long]0) }
        $ExtMap[$ext][0]++
        $ExtMap[$ext][1] += $fBytes

        # Top 20 largest -- SortedList key must be unique; use negative bytes + index trick
        $sortKey = $fBytes * 1000 + $TotalFiles   # ensures unique ascending key
        if ($LargestFiles.Count -lt 20) {
            $LargestFiles.Add($sortKey, [PSCustomObject]@{
                Path     = $file.FullName
                SizeMB   = Format-MB $fBytes
                SizeBytes= $fBytes
                Modified = $lastMod
                Ext      = $ext
            })
        } elseif ($sortKey -gt $LargestFiles.Keys[0]) {
            $LargestFiles.RemoveAt(0)
            $LargestFiles.Add($sortKey, [PSCustomObject]@{
                Path     = $file.FullName
                SizeMB   = Format-MB $fBytes
                SizeBytes= $fBytes
                Modified = $lastMod
                Ext      = $ext
            })
        }

        # Duplicate detection key: name|size
        $dupKey = "$($file.Name)|$fBytes"
        if (-not $DupMap.ContainsKey($dupKey)) {
            $DupMap[$dupKey] = [System.Collections.Generic.List[string]]::new()
        }
        $DupMap[$dupKey].Add($file.FullName)

        # SPO violations on file name
        $relPath = $file.FullName -replace [regex]::Escape($ScanPath),''
        $fViolations = Get-SpoViolation -Name $file.Name -FullPath $relPath
        foreach ($v in $fViolations) {
            $PathViolCount++
            $SpoViolations.Add([PSCustomObject]@{
                ShareName = $ShareName
                Type      = 'File'
                Path      = $file.FullName
                Violation = $v
            })
        }
    }

    $ShareElapsed = (Get-Date) - $ShareStart

    #endregion

    #region -- Write Section: SHARE_SUMMARY (size + age totals) --------------

    Add-CsvRow @{
        Section            = 'SHARE_SUMMARY'
        ShareName          = $ShareName
        SharePath          = $SharePath
        ShareDescription   = $ShareDesc
        PermType           = 'DataSummary'
        Identity           = ''
        Rights             = ''
        AccessControlType  = ''
        IsInherited        = ''
        HighRiskFlag       = ''
        Detail             = "TotalGB=$(Format-GB $TotalBytes)|Files=$TotalFiles|Folders=$TotalFolders|PathViolations=$PathViolCount|ScanTimeSec=$([int]$ShareElapsed.TotalSeconds)"
    }

    #endregion

    #region -- Write Section: FILE_AGE ----------------------------------------

    $fa_OldestStr = if ($Age_Over5yr_Count -gt 0) { $OldestDate.ToString('yyyy-MM-dd') } else { '' }
    $fa_NewestStr = if ($Age_Under1yr_Count -gt 0) { $NewestDate.ToString('yyyy-MM-dd') } else { '' }

    foreach ($row in @(
        [PSCustomObject]@{ Label='Under1Yr';       Count=$Age_Under1yr_Count; Bytes=$Age_Under1yr_Bytes; Oldest='';          Newest=$fa_NewestStr },
        [PSCustomObject]@{ Label='1to3Yr';         Count=$Age_1to3yr_Count;  Bytes=$Age_1to3yr_Bytes;  Oldest='';          Newest='' },
        [PSCustomObject]@{ Label='3to5Yr';         Count=$Age_3to5yr_Count;  Bytes=$Age_3to5yr_Bytes;  Oldest='';          Newest='' },
        [PSCustomObject]@{ Label='Over5Yr';        Count=$Age_Over5yr_Count; Bytes=$Age_Over5yr_Bytes; Oldest=$fa_OldestStr; Newest='' },
        [PSCustomObject]@{ Label='NotAccessed1Yr'; Count=$NotAccessed1yr_Count; Bytes=[long]0;         Oldest='';          Newest='' }
    )) {
        Add-CsvRow @{
            Section   = 'FILE_AGE'
            ShareName = $ShareName
            SharePath = $ScanPath
            PermType  = $row.Label
            Detail    = "Count=$($row.Count)|SizeGB=$(Format-GB $row.Bytes)|OldestFile=$($row.Oldest)|NewestFile=$($row.Newest)"
        }
    }

    #endregion

    #region -- Write Section: LARGEST_FILES -----------------------------------

    foreach ($entry in ($LargestFiles.Values | Sort-Object SizeBytes -Descending)) {
        $flag = ''
        if ($entry.SizeBytes -gt 15GB)      { $flag = 'OVER15GB' }
        elseif ($entry.SizeBytes -gt 250MB) { $flag = 'OVER250MB' }

        Add-CsvRow @{
            Section      = 'LARGEST_FILES'
            ShareName    = $ShareName
            SharePath    = $entry.Path
            PermType     = 'LargeFile'
            Detail       = "SizeMB=$($entry.SizeMB)|LastModified=$($entry.Modified.ToString('yyyy-MM-dd'))|Extension=$($entry.Ext)"
            HighRiskFlag = $flag
        }
    }

    #endregion

    #region -- Write Section: FILE_TYPES --------------------------------------

    foreach ($kvp in $ExtMap.GetEnumerator()) {
        $ext      = $kvp.Key
        $count    = $kvp.Value[0]
        $bytes    = $kvp.Value[1]
        $riskFlag = ''
        if ($SpoBlocked   -contains $ext) { $riskFlag = 'BLOCKED by SPO' }
        elseif ($SpoAttention -contains $ext) { $riskFlag = 'NEEDS ATTENTION' }
        elseif ($SpoExeTypes  -contains $ext) { $riskFlag = 'EXECUTABLE' }

        Add-CsvRow @{
            Section   = 'FILE_TYPES'
            ShareName = $ShareName
            SharePath = $ScanPath
            PermType  = if ($ext -eq '') { '(no extension)' } else { $ext }
            Detail    = "Count=$count|SizeGB=$(Format-GB $bytes)|Flag=$riskFlag"
        }
    }

    #endregion

    #region -- Write Section: SPO_VIOLATIONS ----------------------------------

    foreach ($v in $SpoViolations) {
        Add-CsvRow @{
            Section   = 'SPO_VIOLATIONS'
            ShareName = $v.ShareName
            SharePath = $v.Path
            PermType  = $v.Type
            Detail    = "Reason=$($v.Violation)|PathLength=$($v.Path.Length)"
        }
    }

    #endregion

    #region -- Write Section: PERMISSIONS -------------------------------------

    Add-CsvRow @{
        Section                = 'PERMISSIONS'
        ShareName              = $ShareName
        UniqueAceCount         = $UniqueAces.Count
        BrokenInheritanceCount = $BrokenInherit.Count
        UniqueIdentities       = ($UniqueIdentities | Sort-Object) -join '; '
        LocalAccountsFound     = ($LocalAccounts | Select-Object -Unique | Sort-Object) -join '; '
        ComplexityRating       = if     ($UniqueAces.Count -gt 500)  { 'HIGH' }
                                  elseif ($UniqueAces.Count -gt 100)  { 'MEDIUM' }
                                  else                                 { 'LOW' }
    }

    foreach ($bf in $BrokenInherit) {
        Add-CsvRow @{
            Section   = 'PERMISSIONS'
            ShareName = $ShareName
            Detail    = "BrokenInheritance: $bf"
        }
    }

    #endregion

    #region -- Write Section: DUPLICATES --------------------------------------

    foreach ($kvp in $DupMap.GetEnumerator()) {
        if ($kvp.Value.Count -gt 1) {
            $parts    = $kvp.Key -split '\|'
            $fileName = $parts[0]
            $fileSize = [long]$parts[1]

            Add-CsvRow @{
                Section   = 'DUPLICATES'
                ShareName = $ShareName
                SharePath = $ScanPath
                Identity  = $fileName
                Detail    = "SizeBytes=$fileSize|Count=$($kvp.Value.Count)|Paths=$($kvp.Value -join ' | ')"
            }
        }
    }

    #endregion

    #region -- Write Section: EMPTY_FOLDERS -----------------------------------

    Add-CsvRow @{
        Section          = 'EMPTY_FOLDERS'
        ShareName        = $ShareName
        EmptyFolderCount = $EmptyFolderCount
        TotalFolders     = $TotalFolders
        PercentEmpty     = if ($TotalFolders -gt 0) { [math]::Round($EmptyFolderCount / $TotalFolders * 100, 1) } else { 0 }
    }

    #endregion

    $elapsed = [int]$ShareElapsed.TotalSeconds
    Write-Host "  Done -- $TotalFiles files, $(Format-GB $TotalBytes) GB, $elapsed sec" -ForegroundColor Green
}

#endregion

#region -- Final Summary ------------------------------------------------------

$TotalElapsed = (Get-Date) - $ScriptStart

Write-Progress -Activity 'File Server Assessment' -Status 'Writing output...' -PercentComplete 95

# Export to CSV
$CsvRows | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8 -ErrorAction SilentlyContinue

$SummaryLines = @(
    '',
    '===================================================',
    "  ASSESSMENT COMPLETE",
    "  Server      : $ServerName",
    "  Shares Scanned: $ShareIndex",
    "  Total Runtime : $([int]$TotalElapsed.TotalMinutes)m $($TotalElapsed.Seconds)s",
    "  Total Errors  : $TotalErrors",
    "  CSV Output    : $CsvPath",
    "  Error Log     : $ErrorLogPath",
    '==================================================='
)

$SummaryLines | ForEach-Object { Write-Host $_ -ForegroundColor Cyan }

# Append summary to CSV as a SUMMARY section
Add-CsvRow @{
    Section       = 'RUN_SUMMARY'
    ServerName    = $ServerName
    SharesScanned = $ShareIndex
    TotalErrors   = $TotalErrors
    RuntimeSeconds= [int]$TotalElapsed.TotalSeconds
    CsvPath       = $CsvPath
    ErrorLogPath  = $ErrorLogPath
    RunAt         = $ScriptStart.ToString('yyyy-MM-dd HH:mm:ss')
}

$CsvRows | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8 -ErrorAction SilentlyContinue

Write-Progress -Activity 'File Server Assessment' -Completed

#endregion

#region -- Validation ---------------------------------------------------------

Write-Host "`n----- CSV Validation -----" -ForegroundColor Cyan

try {
    $valCsv = Import-Csv -Path $CsvPath -ErrorAction Stop
} catch {
    Write-Error "Validation: cannot read output CSV -- $($_.Exception.Message)"
    exit 1
}

# Derive total files scanned from the DataSummary rows already written to the CSV
$totalFilesScanned = (
    $valCsv |
    Where-Object { $_.Section -eq 'SHARE_SUMMARY' -and $_.PermType -eq 'DataSummary' } |
    ForEach-Object { if ($_.Detail -match 'Files=(\d+)') { [long]$Matches[1] } else { 0L } } |
    Measure-Object -Sum
).Sum

$genuinelyEmpty = ($totalFilesScanned -eq 0)

# Schema columns other than Section and ShareName
$dataCols = @('Identity','HighRiskFlag','IsInherited','SharePath','PermType','Detail',
              'ShareDescription','AccessControlType','Rights')

# Sections to validate.
# AllowEmpty=$true  -> 0 data rows is legitimate (no violations / no duplicates found)
# AllowEmpty=$false -> 0 data rows when files were scanned is a bug
$sectionsToValidate = @(
    [PSCustomObject]@{ Name='FILE_AGE';       AllowEmpty=$false },
    [PSCustomObject]@{ Name='LARGEST_FILES';  AllowEmpty=$genuinelyEmpty },
    [PSCustomObject]@{ Name='FILE_TYPES';     AllowEmpty=$genuinelyEmpty },
    [PSCustomObject]@{ Name='SPO_VIOLATIONS'; AllowEmpty=$true },
    [PSCustomObject]@{ Name='DUPLICATES';     AllowEmpty=$true }
)

$validationFailed = $false

foreach ($sec in $sectionsToValidate) {
    $sectionRows   = @($valCsv | Where-Object { $_.Section -eq $sec.Name })
    # A row "has data" when at least 3 schema columns beyond Section and ShareName are non-empty
    $populatedRows = @($sectionRows | Where-Object {
        $r = $_
        ($dataCols | Where-Object { -not [string]::IsNullOrEmpty($r.$_) } | Measure-Object).Count -ge 3
    })
    $count = $populatedRows.Count

    if ($count -eq 0 -and -not $sec.AllowEmpty) {
        Write-Host "  Section $($sec.Name): $count data rows written  [FAIL]" -ForegroundColor Red
        $validationFailed = $true
        Write-Warning "  Section '$($sec.Name)' has 0 populated rows but $totalFilesScanned files were scanned -- column mapping error."
    } else {
        Write-Host "  Section $($sec.Name): $count data rows written" -ForegroundColor Green
    }
}

if ($validationFailed) {
    Write-Error "CSV validation failed. Review output:`n  $CsvPath`n  $ErrorLogPath"
    exit 1
}

Write-Host "  Validation PASSED." -ForegroundColor Green

#endregion
