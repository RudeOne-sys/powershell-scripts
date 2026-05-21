# Connect to SharePoint Online
Connect-PnPOnline -Url "https://yourtenant.sharepoint.com/sites/ARC" -Interactive

# Specify document library name
$libraryName = "Documents"
$folderPath = "/Shared Documents/ARCLtd_Underwriting"

# Get all files and folders in the specified location
$items = Get-PnPListItem -List $libraryName -Fields "FileRef", "FileDirRef", "FileLeafRef"

$permissionsReport = @()

foreach ($item in $items) {
    $fileUrl = $item["FileRef"]
    $fileName = $item["FileLeafRef"]

    # Get permissions for the file/folder
    $permissions = Get-PnPListItemPermission -List $libraryName -Identity $item.Id

    foreach ($perm in $permissions) {
        $permissionsReport += [PSCustomObject]@{
            FilePath   = $fileUrl
            FileName   = $fileName
            User       = $perm.PrincipalName
            Permission = $perm.Roles -join ", "
        }
    }
}

# Export to CSV
$permissionsReport | Export-Csv -Path "C:\SharePoint_Permissions.csv" -NoTypeInformation
