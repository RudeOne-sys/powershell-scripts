# Define the shared location
$sharedLocation = "C:\Users\LeeV\OneDrive - Netsurit\Documents\Service Requests\Amaro Foods"

# Get the current date
$currentDate = Get-Date

# Function to calculate item age
function Get-ItemAge {
    param (
        [string]$itemPath
    )
    $itemInfo = Get-Item $itemPath
    $itemAge = $currentDate - $itemInfo.CreationTime
    return $itemAge.Days
}

# Get all folders and files
$items = Get-ChildItem -Path $sharedLocation -Recurse -File -Directory

# Create an array to store item age information
$itemAgeInfo = @()

# Loop through each item and calculate its age
foreach ($item in $items) {
    $itemAge = Get-ItemAge -itemPath $item.FullName
    $itemAgeInfo += [PSCustomObject]@{
        ItemPath = $item.FullName
        ItemAgeDays = $itemAge
    }
}

# Output the item age information to a CSV file
$itemAgeInfo | Export-Csv -Path "C:\temp\ItemAgeInfo.csv" -NoTypeInformation

Write-Host "Item age information has been exported to C:\temp\ItemAgeInfo.csv"