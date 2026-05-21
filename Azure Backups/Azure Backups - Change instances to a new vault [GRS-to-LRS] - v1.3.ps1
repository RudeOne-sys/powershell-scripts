# Connect to Azure (if not already connected)
# Connect-AzAccount

# Set variables
$subscriptionId = "YourSubscriptionId"
$grsResourceGroupName = "YourGRSResourceGroupName"
$grsVaultName = "YourGRSVaultName"
$lrsResourceGroupName = "YourLRSResourceGroupName"
$lrsVaultName = "YourLRSVaultName"
$vmName = "YourVMName"
$vmResourceGroupName = "YourVMResourceGroupName"

# Select the correct subscription
Select-AzSubscription -SubscriptionId $subscriptionId

# Function to check and enable soft delete
function Ensure-SoftDeleteEnabled($vault) {
    $softDeleteStatus = Get-AzRecoveryServicesVaultProperty -VaultId $vault.ID
    if (-not $softDeleteStatus.SoftDeleteFeatureState -eq 'Enabled') {
        Set-AzRecoveryServicesVaultProperty -VaultId $vault.ID -SoftDeleteFeatureState Enable
        Write-Output "Soft delete enabled for vault $($vault.Name)"
    } else {
        Write-Output "Soft delete is already enabled for vault $($vault.Name)"
    }
}

# Step 1: Check and enable soft delete in GRS vault
$grsVault = Get-AzRecoveryServicesVault -ResourceGroupName $grsResourceGroupName -Name $grsVaultName
Ensure-SoftDeleteEnabled $grsVault

# Step 2: Delete backup in GRS vault
Set-AzRecoveryServicesVaultContext -Vault $grsVault
$container = Get-AzRecoveryServicesBackupContainer -ContainerType AzureVM -FriendlyName $vmName
$backupItem = Get-AzRecoveryServicesBackupItem -Container $container -WorkloadType AzureVM
Disable-AzRecoveryServicesBackupProtection -Item $backupItem -RemoveRecoveryPoints -Force
Write-Output "Backup deleted for VM $vmName in GRS vault"

# Wait for the deletion to complete
Start-Sleep -Seconds 60

# Step 3: Perform VM backup in LRS vault
$lrsVault = Get-AzRecoveryServicesVault -ResourceGroupName $lrsResourceGroupName -Name $lrsVaultName
Set-AzRecoveryServicesVaultContext -Vault $lrsVault
$policy = Get-AzRecoveryServicesBackupProtectionPolicy -Name "DefaultPolicy" # Change this to your policy name if different
$vm = Get-AzVM -ResourceGroupName $vmResourceGroupName -Name $vmName
Enable-AzRecoveryServicesBackupProtection -Policy $policy -Name $vm.Name -ResourceGroupName $vm.ResourceGroupName
Write-Output "Backup enabled for VM $vmName in LRS vault"

# Step 4: Undelete the backup in GRS vault
Set-AzRecoveryServicesVaultContext -Vault $grsVault
$container = Get-AzRecoveryServicesBackupContainer -ContainerType AzureVM -FriendlyName $vmName -Status Deleted
$backupItem = Get-AzRecoveryServicesBackupItem -Container $container -WorkloadType AzureVM
Undo-AzRecoveryServicesBackupItemDeletion -Item $backupItem -Force
Write-Output "Backup undeleted for VM $vmName in GRS vault"

Write-Output "Process completed successfully"