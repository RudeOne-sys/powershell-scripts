<#
.SYNOPSIS
    Exports Hyper-V cluster VM information to a CSV file.

.DESCRIPTION
    Queries all nodes in a Hyper-V failover cluster and collects VM details
    including state, resource allocation, replication status, and host placement.

.PARAMETER ClusterName
    The name of the Hyper-V failover cluster. Defaults to the local cluster.

.PARAMETER OutputPath
    Full path for the output CSV file. Defaults to the script directory.

.EXAMPLE
    .\Export-HyperVClusterVMs.ps1 -ClusterName "HV-CLUSTER-01" -OutputPath "C:\Reports\VMs.csv"
#>

param (
    [string]$ClusterName = $env:COMPUTERNAME,
    [string]$OutputPath  = "$PSScriptRoot\HyperV_VM_Export_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
)

#Requires -Modules Hyper-V, FailoverClusters

# ── Helper: convert bytes to GB ──────────────────────────────────────────────
function ConvertTo-GB {
    param ([long]$Bytes)
    return [math]::Round($Bytes / 1GB, 2)
}

# ── Resolve cluster nodes ─────────────────────────────────────────────────────
Write-Host "Connecting to cluster: $ClusterName" -ForegroundColor Cyan

try {
    $clusterNodes = Get-ClusterNode -Cluster $ClusterName | Select-Object -ExpandProperty Name
    Write-Host "Found $($clusterNodes.Count) node(s): $($clusterNodes -join ', ')" -ForegroundColor Green
} catch {
    Write-Error "Failed to connect to cluster '$ClusterName'. Check the name and your permissions.`n$_"
    exit 1
}

# ── Collect VM data from all nodes ────────────────────────────────────────────
$vmData = foreach ($node in $clusterNodes) {
    Write-Host "Querying node: $node" -ForegroundColor Yellow

    try {
        $vms = Get-VM -ComputerName $node -ErrorAction Stop

        foreach ($vm in $vms) {

            # Network adapters (comma-separated IPs) — PS5.1 compatible
            $networkInfo = (Get-VMNetworkAdapter -VM $vm |
                Select-Object -ExpandProperty IPAddresses |
                Where-Object { $_ -match '\d+\.\d+\.\d+\.\d+' }) -join '; '

            # Disk info — per-disk detail and totals
            $vhdDrives  = Get-VMHardDiskDrive -VM $vm
            $vhdInfo    = $vhdDrives | ForEach-Object {
                            Get-VHD -Path $_.Path -ComputerName $node -ErrorAction SilentlyContinue
                          }
            $totalVHDGB = ($vhdInfo | Measure-Object -Property Size -Sum).Sum
            $usedVHDGB  = ($vhdInfo | Measure-Object -Property FileSize -Sum).Sum

            # Build per-disk summary string: "Filename.vhdx | 100GB provisioned / 45GB used"
            $vhdDetails = ($vhdInfo | ForEach-Object {
                $name = Split-Path $_.Path -Leaf
                $prov = [math]::Round($_.Size / 1GB, 2)
                $used = [math]::Round($_.FileSize / 1GB, 2)
                "$name | ${prov}GB provisioned / ${used}GB used"
            }) -join ' || '

            # Replication status
            $replState  = (Get-VMReplication -VM $vm -ErrorAction SilentlyContinue).State

            # Checkpoint count
            $checkpoints = (Get-VMSnapshot -VM $vm -ErrorAction SilentlyContinue | Measure-Object).Count

            [PSCustomObject]@{
                VMName               = $vm.Name
                HostNode             = $node
                State                = $vm.State
                Generation           = $vm.Generation
                CPUCount             = $vm.ProcessorCount
                MemoryAssignedGB     = ConvertTo-GB $vm.MemoryAssigned
                MemoryStartupGB      = ConvertTo-GB $vm.MemoryStartup
                DynamicMemoryEnabled = $vm.DynamicMemoryEnabled
                MemoryMinGB          = if ($vm.DynamicMemoryEnabled) { ConvertTo-GB $vm.MemoryMinimum } else { 'N/A' }
                MemoryMaxGB          = if ($vm.DynamicMemoryEnabled) { ConvertTo-GB $vm.MemoryMaximum } else { 'N/A' }
                TotalVHDSizeGB       = ConvertTo-GB $totalVHDGB
                UsedVHDSizeGB        = ConvertTo-GB $usedVHDGB
                VHDCount             = ($vhdInfo | Measure-Object).Count
                VHDDetails           = $vhdDetails
                IPAddresses          = $networkInfo
                IntegrationServices  = $vm.IntegrationServicesVersion
                Uptime               = $vm.Uptime
                ReplicationState     = if ($replState) { $replState } else { 'Not Configured' }
                CheckpointCount      = $checkpoints
                ConfigurationPath    = $vm.ConfigurationLocation
                CreationTime         = $vm.CreationTime
                Notes                = $vm.Notes
            }
        }
    } catch {
        Write-Warning "Could not query node '$node': $_"
    }
}

# ── Export ────────────────────────────────────────────────────────────────────
if ($vmData) {
    $vmData | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Host "`nExport complete: $OutputPath" -ForegroundColor Green
    Write-Host "Total VMs exported: $($vmData.Count)" -ForegroundColor Green
} else {
    Write-Warning "No VM data collected. Check cluster connectivity and permissions."
}
