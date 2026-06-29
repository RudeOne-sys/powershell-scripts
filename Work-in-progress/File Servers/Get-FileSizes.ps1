#region ScriptInfo

<#

.SYNOPSIS
Gets file sizes from a specified directory and outputs results to a .csv file.

.DESCRIPTION
Gets file sizes from a specified directory and outputs results to a .csv file. Details gathered include File Path, Directory Name, File Name, Size (Bytes), Size (MB), Size (GB), Extension, Creation Time (UTC),
Last Write Time (UTC) and Read Only.

.PARAMETER outDir
This is the directory where the csv report will be saved to. If the directory doesn't exist it will be created.

.PARAMETER scanDir
This is the directory to scan to generate the report for file sizes.

.PARAMETER roundTo
[OPTIONAL] This is the number of decimal points the file sizes will be calculated to. The default is 0 and the choices are 0, 1, 2, 3 & 4 decimal places.

.PARAMETER filesBiggerThan
[OPTIONAL] This parameter is used to limit the scan to files bigger than the set size in GB. Cannot be used with the -filesLessThan parameter.

.PARAMETER filesLessThan
[OPTIONAL] This parameter is used to limit the scan to files less than the set size in GB. Cannot be used with the -filesBiggerThan parameter.

.PARAMETER sortBy
[OPTIONAL] This parameter is used to specify the sort order for the report, the default is File Patch. The choices are File Path, Big Files & Small Files.

.EXAMPLE
.\Get-FileSizes.ps1 -outDir "c:\temp" -scanDir "C:\Windows"
Scans C:\Windows and saves the .csv report to c:\temp

.EXAMPLE
.\Get-FileSizes.ps1 -outDir "c:\temp" -scanDir "C:\Windows" -fileBiggerThan 1 -roundTo 2
Scans C:\Windows for files larger than 1 GB and saves the .csv report to c:\temp. The file sizes will be rounded to 2 decimal places.

.EXAMPLE
.\Get-FileSizes.ps1 -outDir "c:\temp" -scanDir "C:\Windows" -fileLessThan 5
Scans C:\Windows for files less than 5 GB and saves the .csv report to c:\temp

.EXAMPLE
.\Get-FileSizes.ps1 -outDir "c:\temp" -scanDir "C:\Windows" -fileBiggerThan 1 -orderBy BigFiles
Scans C:\Windows for files larger than 1 GB and saves the .csv report to c:\temp. The report will be ordered by the biggest files.

.LINK
https://github.com/gordonrankine/get-filesizes

.NOTES
License:            MIT License
Compatibility:      Windows 10
Author:             Gordon Rankine
Date:               05/11/2021
Version:            1.0
PSScriptAnalyzer:   Pass.
Change Log:         Version  Date        Author          Comments
                    1.0      05/11/2021  Gordon Rankine  Initial script

#>

#endregion ScriptInfo

#region Bindings
[cmdletbinding()]

Param(

    [Parameter(Mandatory=$True, Position=1, HelpMessage="This is the directory where the .cvs report file will be saved.")]
    [string]$outDir,
    [Parameter(Mandatory=$True, Position=2, HelpMessage="This is the directory to scan for file sizes.")]
    [string]$scanDir,
    [Parameter(Mandatory=$False, Position=3, HelpMessage="This is the amount of decimal places the files sizes will be in. The default is 0 and you can choose between 0 & 4.")]
    [ValidateSet(0,1,2,3,4)]
    [int]$roundTo = 0,
    [Parameter(Mandatory=$False, Position=4, HelpMessage="Use this to scan for files bigger than x GB, where x is the number of GB.")]
    [long]$filesBiggerThan,
    [Parameter(Mandatory=$False, Position=5, HelpMessage="Use this to scan for files less than x GB, where x is the number of GB.")]
    [long]$filesLessThan,
    [Parameter(Mandatory=$False, Position=6, HelpMessage="Use this to sort the date either by File Path, Big Files or Small Files. The default is File Path")]
    [ValidateSet('FilePath','BigFiles', 'SmallFiles')]
    [string]$sortBy = 'FilePath'
)
#endregion Bindings

#region Functions
### START FUNCTIONS ###

function fnCreateDir {

<#

.SYNOPSIS
Creates a directory.

.DESCRIPTION
Creates a directory.

.PARAMETER outDir
This is the directory to be created.

.EXAMPLE
.\Create-Directory.ps1 -outDir "c:\test"
Creates a directory called "test" in c:\

.EXAMPLE
.\Create-Directory.ps1 -outDir "\\COMP01\c$\test"
Creates a directory called "test" in c:\ on COMP01

.LINK
https://github.com/gordonrankine/powershell

.NOTES
    License:            MIT License
    Compatibility:      Windows 7 or Server 2008 and higher
    Author:             Gordon Rankine
    Date:               13/01/2019
    Version:            1.1
    PSSscriptAnalyzer:  Pass

#>

    [CmdletBinding()]

        Param(

        # The directory to be created.
        [Parameter(Mandatory=$True, Position=0, HelpMessage='This is the directory to be created. E.g. C:\Temp')]
        [string]$outDir

        )

        # Create out directory if it doesnt exist
        if(!(Test-Path -path $outDir)){
            if(($outDir -notlike "*:\*") -and ($outDir -notlike "*\\*")){
            Write-Output "[ERROR] $outDir is not a valid path. Script terminated."
            break
            }
                try{
                New-Item $outDir -type directory -Force -ErrorAction Stop | Out-Null
                Write-Output "[INFO] Created output directory $outDir"
                }
                catch{
                Write-Output "[ERROR] There was an issue creating $outDir. Script terminated."
                Write-Output ($_.Exception.Message)
                Write-Output ""
                break
                }
        }
        # Directory already exists
        else{
        Write-Output "[INFO] $outDir already exists."
        }

} # end fnCreateDir

function fnCheckPSAdmin {

<#

.SYNOPSIS
Checks PowerShell is running as Administrator.

.DESCRIPTION
Checks PowerShell is running as Administrator.

.LINK
https://github.com/gordonrankine/powershell

.NOTES
    License:            MIT License
    Compatibility:      Windows 7 or Server 2008 and higher
    Author:             Gordon Rankine
    Date:               19/09/2019
    Version:            1.0
    PSSscriptAnalyzer:  Pass

#>

    try{
    $wIdCurrent = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $wPrinCurrent = New-Object System.Security.Principal.WindowsPrincipal($wIdCurrent)
    $wBdminRole = [System.Security.Principal.WindowsBuiltInRole]::Administrator

        if(!$wPrinCurrent.IsInRole($wBdminRole)){
        Write-Output "[ERROR] PowerShell is not running as administrator. Script terminated."
        Break
        }

    }

    catch{
    Write-Output "[ERROR] There was an unexpected error checking if PowerShell is running as administrator. Script terminated."
    Break
    }

} # end fnCheckPSAdmin
#endregion Functions

#region StartCollecting
Clear-Host

# Start stopwatch
$sw = [system.diagnostics.stopwatch]::StartNew()

fnCheckPSAdmin
fnCreateDir $outDir

# Date and Endpoint used to construct $outFile
$date = Get-Date -UFormat %Y%m%d%H%M
$endpoint = $env:COMPUTERNAME
$outFile = "$outDir\GetFileSizes_$($endpoint)_$($date).csv"

    if($filesBiggerThan -ne '' -and $filesLessThan -ne ''){
    Write-Output "[ERROR] -filesBiggerThan and -filesLessThan can't be run togther. Please choose one only and try again."
    Write-Output ""
    break
    }

# Create File array to store results then get files.
# Script will be running as admin but there may be some files even the admin cant see. These will be skipped.
$filesArray = @()

    try{
    Write-Output "[INFO] File size rounding set to $($roundTo) decimal places."
    Write-Output "[INFO] Ordering set to $($SortBy)."

        if($filesBiggerThan -ne ''){
        Write-Output "[INFO] Scanning for files bigger than $($filesBiggerThan) GB."
        Write-Output "[INFO] Scanning $($scanDir), please wait..."
        $filesBiggerThan = $filesBiggerThan * 1073741824
        $files = Get-ChildItem -Path $scanDir -File -Recurse -ErrorAction SilentlyContinue -Force | Where-Object {$_.Length -ge $filesBiggerThan}
        }

        elseif($filesLessThan -ne ''){
        Write-Output "[INFO] Scanning for files less than $($filesLessThan) GB."
        Write-Output "[INFO] Scanning $($scanDir), please wait..."
        $filesLessThan = $filesLessThan * 1073741824
        $files = Get-ChildItem -Path $scanDir -File -Recurse -ErrorAction SilentlyContinue -Force | Where-Object {$_.Length -le $filesLessThan}
        }

        else{
        Write-Output "[INFO] Scanning $($scanDir), please wait..."
        $files = Get-ChildItem -Path $scanDir -File -Recurse -ErrorAction SilentlyContinue -Force
        }

    }
    catch{
    Write-Output "[ERROR] There was an issue scanning $($scanDir). Script terminated."
    Write-Output ($_.Exception.Message)
    Write-Output ""
    break
    }

    # Get file properties for each file and do some calculations
    foreach($file in $files){
    $fileArray = New-Object System.Object
    $fileArray | Add-Member -type NoteProperty -name "File Path" -Value $file.FullName -Force
    $fileArray | Add-Member -type NoteProperty -name DirectoryName -Value $file.DirectoryName -Force
    $fileArray | Add-Member -type NoteProperty -name Name -Value $file.Name -Force
    $fileArray | Add-Member -type NoteProperty -name "Size (Bytes)" -Value $file.Length -Force
    $fileArray | Add-Member -type NoteProperty -name "Size (MB)" -Value ([math]::Round($file.Length/1048576,$roundTo)) -Force
    $fileArray | Add-Member -type NoteProperty -name "Size (GB)" -Value ([math]::Round($file.Length/1073741824,$roundTo)) -Force
    $fileArray | Add-Member -type NoteProperty -name Extension -Value $file.Extension -Force
    $fileArray | Add-Member -type NoteProperty -name "Creation Time (UTC)" -Value $file.CreationTimeUtc -Force
    $fileArray | Add-Member -type NoteProperty -name "Last Write Time (UTC)" -Value $file.LastWriteTimeUtc -Force
    $fileArray | Add-Member -type NoteProperty -name "Read Only" -Value $file.IsReadOnly -Force
    $filesArray += $fileArray
    }

#endregion StartCollecting

#region GenerateReports

    try{
    Write-Output "[INFO] Saving report file."
        if($SortBy -eq 'BigFiles'){
        $filesArray | Sort-Object "Size (Bytes)" -Descending | Export-Csv $outFile -Encoding ASCII -NoTypeInformation -Force
        }
        elseif($SortBy -eq 'SmallFiles'){
        $filesArray | Sort-Object "Size (Bytes)" | Export-Csv $outFile -Encoding ASCII -NoTypeInformation -Force
        }
        else{
        $filesArray | Export-Csv $outFile -Encoding ASCII -NoTypeInformation -Force
        }
    }
    catch{
    Write-Output "[ERROR] There was an issue saving $($outFile). Script terminated."
    Write-Output ($_.Exception.Message)
    Write-Output ""
    break
    }

# Display script execution time
Write-Output "[INFO] Script complete in $($sw.Elapsed.Hours) hours, $($sw.Elapsed.Minutes) minutes, $($sw.Elapsed.Seconds) seconds."
Write-Output "[INFO] Results can be found at $($outfile)"
Write-Output ""

 #endregion GenerateReportsg