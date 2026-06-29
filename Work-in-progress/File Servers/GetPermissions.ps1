# Define the directory to search
$directory = "C:\Users\LeeV\OneDrive - Netsurit\Documents\Service Requests\Amaro Foods"

# Function to get permissions for a given item
function Get-Permissions {
    param (
        [string]$itemPath
    )
    $acl = Get-Acl -Path $itemPath
    $permissions = @()
    foreach ($access in $acl.Access) {
        $permissions += [PSCustomObject]@{
            ItemPath = $itemPath
            IdentityReference = $access.IdentityReference
            AccessControlType = $access.AccessControlType
            IsInherited = $access.IsInherited
            FileSystemRights = $access.FileSystemRights
        }
    }
    return $permissions
}

# Get all folders and files
$items = Get-ChildItem -Path $directory -Recurse

# Create an array to store permissions information
$permissionsInfo = @()

# Loop through each item and get its permissions
foreach ($item in $items) {
    $permissionsInfo += Get-Permissions -itemPath $item.FullName
}

# Output the permissions information to a CSV file
$permissionsInfo | Export-Csv -Path "C:\temp\PermissionsInfo.csv" -NoTypeInformation

Write-Host "Permissions information has been exported to C:\temp\PermissionsInfo.csv"