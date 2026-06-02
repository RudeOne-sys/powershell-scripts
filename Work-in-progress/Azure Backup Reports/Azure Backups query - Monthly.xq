// Find backups that have completed in the last "31Day"
AddonAzureBackupJobs
| where JobOperation=="Backup"
| where TimeGenerated > ago(31d)
| where JobStatus=="Completed" or JobStatus == "Running"
| where JobFailureCode =="Success" or JobFailureCode == "Failed"
| join kind=inner
(
    CoreAzureBackup
    | where OperationName == "BackupItem"
    | where BackupManagementType=="IaaSVM" or BackupManagementType == "AzureStorage"
    | distinct BackupItemUniqueId, BackupItemFriendlyName
)
on BackupItemUniqueId
// Output the results based on the parameters below.
| project JobStartDateTime, TimeGenerated, JobDurationInSecs, JobOperation, BackupItemFriendlyName, JobFailureCode, JobStatus, BackupManagementType, DataTransferredInMB, BackupItemUniqueId1
| sort by JobStartDateTime
