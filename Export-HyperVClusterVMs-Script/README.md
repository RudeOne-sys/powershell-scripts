# Export-HyperVClusterVMs

A PowerShell script that queries all nodes in a Hyper-V failover cluster and exports comprehensive VM inventory data to a CSV file.

---

## Requirements

| Requirement | Detail |
|---|---|
| PowerShell | 5.1 or later |
| Modules | `Hyper-V`, `FailoverClusters` |
| Permissions | Domain Admin or Hyper-V Administrator on all cluster nodes |
| Execution context | Run from a cluster node or a management host with remote access to all nodes |

The `Hyper-V` and `FailoverClusters` modules are included with Windows Server. If running from a management workstation, install them via RSAT:

```powershell
Add-WindowsCapability -Online -Name Rsat.FailoverCluster.Management.Tools~~~~0.0.1.0
Add-WindowsCapability -Online -Name Rsat.Hyper-V.Tools~~~~0.0.1.0
```

---

## Usage

### Basic — query the local cluster

```powershell
.\Export-HyperVClusterVMs.ps1
```

Output file is created in the script directory with a timestamped filename:
`HyperV_VM_Export_20250506_143022.csv`

### Specify a remote cluster

```powershell
.\Export-HyperVClusterVMs.ps1 -ClusterName "HV-CLUSTER-01"
```

### Specify a custom output path

```powershell
.\Export-HyperVClusterVMs.ps1 -ClusterName "HV-CLUSTER-01" -OutputPath "C:\Reports\VMs.csv"
```

---

## Parameters

| Parameter | Type | Default | Description |
|---|---|---|---|
| `-ClusterName` | String | `$env:COMPUTERNAME` | Name of the Hyper-V failover cluster to query |
| `-OutputPath` | String | Script directory, timestamped | Full file path for the CSV output |

---

## CSV Output Fields

Each row represents one VM. All size values are in GB rounded to 2 decimal places.

| Column | Description |
|---|---|
| `VMName` | Display name of the virtual machine |
| `HostNode` | Cluster node the VM is currently running on |
| `State` | Power state (Running, Off, Saved, Paused) |
| `Generation` | VM generation (1 or 2) |
| `CPUCount` | Number of virtual processors assigned |
| `MemoryAssignedGB` | Memory currently assigned to the VM |
| `MemoryStartupGB` | Memory allocated at startup |
| `DynamicMemoryEnabled` | Whether Dynamic Memory is active (True/False) |
| `MemoryMinGB` | Dynamic Memory minimum (N/A if disabled) |
| `MemoryMaxGB` | Dynamic Memory maximum (N/A if disabled) |
| `TotalVHDSizeGB` | Sum of provisioned size across all attached VHDs |
| `UsedVHDSizeGB` | Sum of actual used/allocated size across all attached VHDs |
| `VHDCount` | Number of virtual hard disks attached |
| `VHDDetails` | Per-disk breakdown (see format below) |
| `IPAddresses` | IPv4 addresses across all virtual network adapters, semicolon-separated |
| `IntegrationServices` | Integration Services version installed in the guest |
| `Uptime` | Current uptime of the VM |
| `ReplicationState` | Hyper-V Replica state, or `Not Configured` |
| `CheckpointCount` | Number of checkpoints (snapshots) present |
| `ConfigurationPath` | Path to the VM configuration files on the host |
| `CreationTime` | Date and time the VM was created |
| `Notes` | VM notes field from Hyper-V Manager |

### VHDDetails format

Multiple disks are separated by ` || `. Each disk entry follows this format:

```
DiskName.vhdx | 127GB provisioned / 62.4GB used || DataDisk.vhdx | 500GB provisioned / 310.75GB used
```

The `||` separator was chosen to avoid conflicts with commas and semicolons in CSV parsing (e.g. when opening in Excel).

---

## Safety Notes

This script is **read-only**. It does not create, modify, start, stop, or delete any virtual machines or cluster resources.

The one consideration on a live production cluster is `Get-VHD`, which briefly inspects each VHD/VHDX file to read size metadata. On environments with a large number of dynamically expanding disks under heavy I/O, this can add query time. It will not cause data loss or VM interruption.

If you want to skip VHD file inspection entirely and only capture disk counts, replace the VHD block in the script with:

```powershell
$vhdDrives  = Get-VMHardDiskDrive -VM $vm
$totalVHDGB = 'Skipped'
$usedVHDGB  = 'Skipped'
$vhdDetails = 'Skipped'
```

---

## How It Works

1. Connects to the specified cluster and enumerates all nodes via `Get-ClusterNode`
2. Iterates through each node sequentially and calls `Get-VM` remotely
3. For each VM, collects network, storage, memory, replication, and checkpoint data
4. Assembles a `PSCustomObject` per VM and streams results into a collection
5. Exports the full collection to CSV via `Export-Csv`

Nodes that fail to respond are skipped with a warning rather than terminating the entire run, so a single unreachable node does not prevent data collection from the rest of the cluster.

---

## Troubleshooting

| Symptom | Likely cause | Fix |
|---|---|---|
| `Failed to connect to cluster` | Wrong cluster name or no network access | Verify the cluster name with `Get-Cluster` from the node |
| `Could not query node` | WinRM not enabled or firewall blocking remote PS | Run `Enable-PSRemoting` on the target node |
| `The term 'Get-VM' is not recognized` | Hyper-V module not installed | Install RSAT Hyper-V Tools (see Requirements) |
| Empty `IPAddresses` column | Integration Services not running or guest has no IP | Check VM integration services status in Hyper-V Manager |
| VHDDetails shows partial results | One or more VHD paths inaccessible from the querying host | Ensure the host running the script has read access to the VHD storage paths |

---

## License

MIT — free to use, modify, and distribute.
