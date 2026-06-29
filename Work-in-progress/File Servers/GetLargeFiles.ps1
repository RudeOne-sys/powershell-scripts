# Define the directory to search and the size threshold (in MB)
$directory = "C:\Users\LeeV\OneDrive - Netsurit\Documents\Service Requests\Amaro Foods"
$sizeThresholdMB = 100  # Size threshold in MB

# Convert size threshold to bytes
$sizeThresholdBytes = $sizeThresholdMB * 1MB

# Get the current date
$currentDate = Get-Date

# Function to get large files
function Get-LargeFiles {
    param (
        [string]$directoryPath,
        [int64]$sizeThreshold
    )
    Get-ChildItem -Path $directoryPath -Recurse -File | Where-Object { $_.Length -gt $sizeThreshold }
}

# Function to get folder sizes in the root directory
function Get-FolderSizes {
    param (
        [string]$directoryPath
    )
    Get-ChildItem -Path $directoryPath -Directory | ForEach-Object {
        $folderSize = (Get-ChildItem -Path $_.FullName -Recurse -File | Measure-Object -Property Length -Sum).Sum
        [PSCustomObject]@{
            FolderPath = $_.FullName
            FolderSizeMB = [math]::Round($folderSize / 1MB, 2)
        }
    }
}

# Get large files
$largeFiles = Get-LargeFiles -directoryPath $directory -sizeThreshold $sizeThresholdBytes

# Get folder sizes in the root directory
$folderSizes = Get-FolderSizes -directoryPath $directory

# Create an array to store file information
$fileInfo = @()

# Loop through each large file and store its information
foreach ($file in $largeFiles) {
    $fileInfo += [PSCustomObject]@{
        FilePath = $file.FullName
        FileSizeMB = [math]::Round($file.Length / 1MB, 2)
        CreationDate = $file.CreationTime
        LastAccessDate = $file.LastAccessTime
        LastWriteDate = $file.LastWriteTime
    }
}

# Output the file information to a CSV file
$fileInfo | Export-Csv -Path "C:\temp\LargeFilesInfo.csv" -NoTypeInformation

# Output the folder size information to a CSV file
$folderSizes | Export-Csv -Path "C:\temp\FolderSizesInfo.csv" -NoTypeInformation

Write-Host "Large file information has been exported to C:\temp\LargeFilesInfo.csv"
Write-Host "Folder size information has been exported to C:\temp\FolderSizesInfo.csv"