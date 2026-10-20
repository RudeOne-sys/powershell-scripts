# Mock harness: replaces Az cmdlets with canned data so every section of the script runs offline.
param([string]$Script, [string]$Out, [switch]$Quick)

$global:T  = '11111111-1111-1111-1111-111111111111'
$global:S1 = 'aaaaaaaa-0000-0000-0000-000000000001'
$global:S2 = 'aaaaaaaa-0000-0000-0000-000000000002'
$global:S3 = 'aaaaaaaa-0000-0000-0000-000000000003'
$global:ArgCalls = [System.Collections.Generic.List[string]]::new()

function global:Get-Module { [pscustomobject]@{ Name = 'x'; Version = [version]'99.0' } }
function global:Install-Module { }
function global:Import-Module { }
function global:Update-AzConfig { }
function global:Connect-AzAccount { }
function global:Get-AzContext { [pscustomobject]@{ Account = [pscustomobject]@{ Id = 'roan@netsurit.com' }; Tenant = [pscustomobject]@{ Id = $global:T } } }
function global:Get-AzAccessToken { [pscustomobject]@{ Token = 'fake' } }

function global:J($o) { $o | ConvertTo-Json -Depth 10 -Compress }
function global:MkRow { param([hashtable]$h) [pscustomobject]$h }

function global:Search-AzGraph {
    param([string]$Query, [int]$First, [string]$SkipToken, [string[]]$Subscription, [switch]$UseTenantScope)
    if ($env:MOCK_FAIL) { throw 'AuthorizationFailed: The client does not have authorization to perform action Microsoft.ResourceGraph/resources/read' }
    $q = $Query
    $global:ArgCalls.Add(($q -split "`n")[0..2] -join ' ')
    $vmBase = "/subscriptions/$S1/resourcegroups/rg-prod/providers/microsoft.compute/virtualmachines"
    $rows = switch -Regex ($q) {
        'mv-expand p = properties\.virtualNetworkPeerings' {
            $hub = "/subscriptions/$S1/resourcegroups/rg-hub/providers/microsoft.network/virtualnetworks/hub-vnet"
            (MkRow @{ id = 'h'; vnet = 'hub-vnet'; subscriptionId = $S1; peering = 'hub-to-prod'; state = 'Connected'; remote = "/subscriptions/$S1/resourcegroups/rg-prod/providers/microsoft.network/virtualnetworks/spoke-prod"; gatewayTransit = 'true'; useRemoteGateways = 'false' }),
            (MkRow @{ id = 's'; vnet = 'spoke-prod'; subscriptionId = $S1; peering = 'prod-to-hub'; state = 'Connected'; remote = $hub; gatewayTransit = 'false'; useRemoteGateways = 'false' }),
            (MkRow @{ id = 'h'; vnet = 'hub-vnet'; subscriptionId = $S1; peering = 'hub-to-dev'; state = 'Disconnected'; remote = "/subscriptions/$S2/resourcegroups/rg-dev/providers/microsoft.network/virtualnetworks/spoke-dev"; gatewayTransit = 'true'; useRemoteGateways = 'false' })
            break
        }
        'mv-expand s = properties\.subnets' {
            (MkRow @{ id = 'h'; vnet = 'hub-vnet'; resourceGroup = 'rg-hub'; subscriptionId = $S1; subnet = 'AzureFirewallSubnet'; prefix = '10.0.0.0/26'; nsg = ''; routeTable = ''; natGateway = ''; ipConfigs = 1; privateEndpoints = $null }),
            (MkRow @{ id = 's'; vnet = 'spoke-prod'; resourceGroup = 'rg-prod'; subscriptionId = $S1; subnet = 'snet-app'; prefix = '10.1.1.0/24'; nsg = '/x/networkSecurityGroups/nsg-app'; routeTable = '/x/routeTables/rt-spoke'; natGateway = ''; ipConfigs = 3; privateEndpoints = $null }),
            (MkRow @{ id = 's'; vnet = 'spoke-prod'; resourceGroup = 'rg-prod'; subscriptionId = $S1; subnet = 'snet-data'; prefix = '10.1.2.0/24'; nsg = ''; routeTable = ''; natGateway = ''; ipConfigs = 2; privateEndpoints = 1 }),
            (MkRow @{ id = 'd'; vnet = 'spoke-dev'; resourceGroup = 'rg-dev'; subscriptionId = $S2; subnet = 'default'; prefix = '10.1.128.0/24'; nsg = ''; routeTable = ''; natGateway = ''; ipConfigs = $null; privateEndpoints = $null })
            break
        }
        "microsoft\.network/virtualnetworks'\s*\n\| project" {
            (MkRow @{ id = "/subscriptions/$S1/resourcegroups/rg-hub/providers/microsoft.network/virtualnetworks/hub-vnet"; name = 'hub-vnet'; resourceGroup = 'rg-hub'; subscriptionId = $S1; location = 'southafricanorth'; addressSpace = '["10.0.0.0/16"]'; ddos = 'false'; peerings = 2; subnetCount = 1 }),
            (MkRow @{ id = "/subscriptions/$S1/resourcegroups/rg-prod/providers/microsoft.network/virtualnetworks/spoke-prod"; name = 'spoke-prod'; resourceGroup = 'rg-prod'; subscriptionId = $S1; location = 'southafricanorth'; addressSpace = '["10.1.0.0/16"]'; ddos = 'false'; peerings = 1; subnetCount = 2 }),
            (MkRow @{ id = "/subscriptions/$S2/resourcegroups/rg-dev/providers/microsoft.network/virtualnetworks/spoke-dev"; name = 'spoke-dev'; resourceGroup = 'rg-dev'; subscriptionId = $S2; location = 'westeurope'; addressSpace = '["10.1.128.0/17"]'; ddos = 'false'; peerings = 0; subnetCount = 1 }),
            (MkRow @{ id = "/subscriptions/$S2/resourcegroups/rg-dev/providers/microsoft.network/virtualnetworks/isolated"; name = 'isolated'; resourceGroup = 'rg-dev'; subscriptionId = $S2; location = 'westeurope'; addressSpace = '["192.168.0.0/24"]'; ddos = 'false'; peerings = 0; subnetCount = 0 })
            break
        }
        "'microsoft\.network/azurefirewalls','microsoft\.network/virtualnetworkgateways'" {
            (MkRow @{ id = 'fw'; name = 'afw-hub'; type = 'microsoft.network/azurefirewalls'; resourceGroup = 'rg-hub'; subscriptionId = $S1; location = 'southafricanorth'; subnetId = "/subscriptions/$S1/resourcegroups/rg-hub/providers/microsoft.network/virtualnetworks/hub-vnet/subnets/azurefirewallsubnet"; gatewayType = ''; sku = 'Standard' }),
            (MkRow @{ id = 'nw'; name = 'NetworkWatcher_southafricanorth'; type = 'microsoft.network/networkwatchers'; resourceGroup = 'NetworkWatcherRG'; subscriptionId = $S1; location = 'southafricanorth'; subnetId = ''; gatewayType = ''; sku = '' }),
            (MkRow @{ id = 'pe'; name = 'pe-sql'; type = 'microsoft.network/privateendpoints'; resourceGroup = 'rg-prod'; subscriptionId = $S1; location = 'southafricanorth'; subnetId = ''; gatewayType = ''; sku = '' }),
            (MkRow @{ id = 'pip'; name = 'pip-afw'; type = 'microsoft.network/publicipaddresses'; resourceGroup = 'rg-hub'; subscriptionId = $S1; location = 'southafricanorth'; subnetId = ''; gatewayType = ''; sku = 'Standard' })
            break
        }
        'enableIPForwarding' { break }
        "microsoft\.resources/subscriptions'\s*\n\| project" {
            (MkRow @{ id = "/subscriptions/$S1"; subscriptionId = $S1; name = 'Contoso-Prod'; state = 'Enabled'; quotaId = 'CSP_2015-05-01'; spendingLimit = 'Off'; mgChain = (J @(@{ name = 'lz-prod'; displayName = 'Landing Zones' }, @{ name = $T; displayName = 'Tenant Root Group' })) }),
            (MkRow @{ id = "/subscriptions/$S2"; subscriptionId = $S2; name = 'Contoso-Dev'; state = 'Enabled'; quotaId = 'MSDN_2014-09-01'; spendingLimit = 'On'; mgChain = (J @(@{ name = $T; displayName = 'Tenant Root Group' })) }),
            (MkRow @{ id = "/subscriptions/$S3"; subscriptionId = $S3; name = 'Old-Sandbox'; state = 'Disabled'; quotaId = 'PayAsYouGo_2014-09-01'; spendingLimit = 'Off'; mgChain = '[]' })
            break
        }
        'microsoft\.management/managementgroups' { (MkRow @{ id = 'x'; name = $T; displayName = 'Tenant Root Group'; parent = '' }), (MkRow @{ id = 'y'; name = 'lz-prod'; displayName = 'Landing Zones'; parent = $T }); break }
        'count\(\) by type' {
            if (!$SkipToken) { return [pscustomobject]@{ Data = @((MkRow @{ type = 'microsoft.compute/virtualmachines'; Count = 6 }), (MkRow @{ type = 'microsoft.compute/disks'; Count = 9 })); SkipToken = 'page2' } }
            (MkRow @{ type = 'microsoft.network/networkinterfaces'; Count = 7 }), (MkRow @{ type = 'microsoft.hybridcompute/machines'; Count = 1 }); break
        }
        'count\(\) by location' { (MkRow @{ location = 'southafricanorth'; Count = 20 }), (MkRow @{ location = 'westeurope'; Count = 3 }); break }
        'count\(\) by subscriptionId' { (MkRow @{ subscriptionId = $S1; Count = 18 }), (MkRow @{ subscriptionId = $S2; Count = 5 }); break }
        "resourcegroups' \| summarize Count" { MkRow @{ Count = 12 }; break }
        "microsoft\.compute/virtualmachines'\s*\n\| extend size" {
            (MkRow @{ id = "$vmBase/app01"; name = 'app01'; resourceGroup = 'rg-prod'; subscriptionId = $S1; location = 'southafricanorth'; size = 'Standard_D4s_v5'; osType = 'Windows'; powerState = 'PowerState/running'; osName = 'Windows Server 2019 Datacenter'; osVersion = '10.0.17763'; imgPublisher = 'MicrosoftWindowsServer'; imgOffer = 'WindowsServer'; imgSku = '2019-Datacenter'; zones = '["1"]'; avset = ''; vmss = ''; managedOs = $true; nics = (J @(@{ id = "/subscriptions/$S1/resourceGroups/rg-prod/providers/Microsoft.Network/networkInterfaces/app01-nic" })); winPatch = 'AutomaticByPlatform'; linPatch = ''; licenseType = 'Windows_Server'; priority = '' }),
            (MkRow @{ id = "$vmBase/sql01"; name = 'sql01'; resourceGroup = 'rg-prod'; subscriptionId = $S1; location = 'southafricanorth'; size = 'Standard_DS3_v2'; osType = 'Windows'; powerState = 'PowerState/running'; osName = ''; osVersion = ''; imgPublisher = 'MicrosoftSQLServer'; imgOffer = 'sql2016sp2-ws2012r2'; imgSku = 'enterprise'; zones = ''; avset = 'x'; vmss = ''; managedOs = $true; nics = '[]'; winPatch = 'Manual'; linPatch = ''; licenseType = ''; priority = '' }),
            (MkRow @{ id = "$vmBase/dc01"; name = 'dc01'; resourceGroup = 'rg-prod'; subscriptionId = $S1; location = 'southafricanorth'; size = 'Standard_A2'; osType = 'Windows'; powerState = 'PowerState/running'; osName = 'Windows Server 2016 Datacenter'; osVersion = '10.0.14393'; imgPublisher = ''; imgOffer = ''; imgSku = ''; zones = ''; avset = ''; vmss = ''; managedOs = $true; nics = '[]'; winPatch = ''; linPatch = ''; licenseType = ''; priority = '' }),
            (MkRow @{ id = "$vmBase/web01"; name = 'web01'; resourceGroup = 'rg-prod'; subscriptionId = $S1; location = 'southafricanorth'; size = 'Standard_B2ms'; osType = 'Linux'; powerState = 'PowerState/running'; osName = 'ubuntu'; osVersion = '18.04'; imgPublisher = 'Canonical'; imgOffer = 'UbuntuServer'; imgSku = '18.04-LTS'; zones = ''; avset = ''; vmss = ''; managedOs = $true; nics = '[]'; winPatch = ''; linPatch = 'ImageDefault'; licenseType = ''; priority = '' }),
            (MkRow @{ id = "/subscriptions/$S2/resourcegroups/rg-dev/providers/microsoft.compute/virtualmachines/devbox"; name = 'devbox'; resourceGroup = 'rg-dev'; subscriptionId = $S2; location = 'westeurope'; size = 'Standard_D2s_v3'; osType = 'Linux'; powerState = 'PowerState/stopped'; osName = ''; osVersion = ''; imgPublisher = 'RedHat'; imgOffer = 'RHEL'; imgSku = '7.9'; zones = ''; avset = ''; vmss = ''; managedOs = $true; nics = '[]'; winPatch = ''; linPatch = ''; licenseType = ''; priority = '' }),
            (MkRow @{ id = "/subscriptions/$S2/resourcegroups/rg-dev/providers/microsoft.compute/virtualmachines/avd-0"; name = 'avd-0'; resourceGroup = 'rg-dev'; subscriptionId = $S2; location = 'westeurope'; size = 'Standard_D4as_v5'; osType = 'Windows'; powerState = 'PowerState/deallocated'; osName = ''; osVersion = ''; imgPublisher = 'MicrosoftWindowsDesktop'; imgOffer = 'windows-10'; imgSku = 'win10-22h2-avd'; zones = ''; avset = ''; vmss = ''; managedOs = $false; nics = '[]'; winPatch = ''; linPatch = ''; licenseType = 'Windows_Client'; priority = '' })
            break
        }
        'microsoft\.authorization/policyassignments' {
            (MkRow @{ id = "/subscriptions/$S1/providers/Microsoft.Authorization/policyAssignments/SecurityCenterBuiltIn"; name = 'SecurityCenterBuiltIn'; displayName = 'ASC Default'; scope = "/subscriptions/$S1"; definitionId = '/providers/Microsoft.Authorization/policySetDefinitions/1f3afdf9'; enforcementMode = 'Default' }),
            (MkRow @{ id = "/providers/Microsoft.Management/managementGroups/lz-prod/providers/Microsoft.Authorization/policyAssignments/allowed-locs"; name = 'allowed-locs'; displayName = 'Allowed locations'; scope = '/providers/Microsoft.Management/managementGroups/lz-prod'; definitionId = '/providers/Microsoft.Authorization/policyDefinitions/e56962a6'; enforcementMode = 'DoNotEnforce' }),
            (MkRow @{ id = "/subscriptions/$S1/providers/Microsoft.Authorization/policyAssignments/req-tags"; name = 'req-tags'; displayName = 'Require tags'; scope = "/subscriptions/$S1"; definitionId = "/subscriptions/$S1/providers/Microsoft.Authorization/policyDefinitions/custom1"; enforcementMode = '' })
            break
        }
        'summarize nc = max' { (MkRow @{ subscriptionId = $S1; Evaluated = 150; NonCompliant = 60 }), (MkRow @{ subscriptionId = $S2; Evaluated = 40; NonCompliant = 5 }); break }
        'dcount\(rid\) by assignmentId' { (MkRow @{ assignmentId = "/subscriptions/$S1/providers/microsoft.authorization/policyassignments/securitycenterbuiltin"; definitionName = 'abc-def'; Resources = 30 }), (MkRow @{ assignmentId = '/x/unknown'; definitionName = 'zzz'; Resources = 4 }); break }
        'microsoft\.authorization/policydefinitions' { MkRow @{ id = 'd'; name = 'abc-def'; displayName = 'Storage accounts should restrict network access' }; break }
        '^resources \| summarize Total' { MkRow @{ Total = 200; Untagged = 90 }; break }
        '^resourcecontainers .*summarize Total' { MkRow @{ Total = 12; Untagged = 3 }; break }
        'mvexpand tags' { (MkRow @{ tagKey = 'Environment'; Resources = 80 }), (MkRow @{ tagKey = 'Owner'; Resources = 40 }); break }
        'microsoft\.authorization/roledefinitions' {
            (MkRow @{ id = '1'; roleGuid = '8e3af657-a8ff-443c-a75c-2fe8c4bcb635'; roleName = 'Owner'; roleType = 'BuiltInRole'; actions = '["*"]'; scopes = '["/"]' }),
            (MkRow @{ id = '2'; roleGuid = 'b24988ac-6180-42a0-ab88-20f7382dd24c'; roleName = 'Contributor'; roleType = 'BuiltInRole'; actions = '["*"]'; scopes = '["/"]' }),
            (MkRow @{ id = '3'; roleGuid = 'cccccccc-0000-0000-0000-000000000001'; roleName = 'Contoso Ops Admin'; roleType = 'CustomRole'; actions = '["*"]'; scopes = (J @("/subscriptions/$S1")) })
            break
        }
        'microsoft\.authorization/roleassignments' {
            $o = '8e3af657-a8ff-443c-a75c-2fe8c4bcb635'; $c = 'b24988ac-6180-42a0-ab88-20f7382dd24c'; $u = '18d7d88d-d35e-4fb5-a5c3-7773c20a72d9'
            (MkRow @{ id = 'a1'; roleGuid = $o; principalId = 'u1'; principalType = 'User'; scope = "/subscriptions/$S1"; createdOn = '2023-01-01'; subscriptionId = $S1 }),
            (MkRow @{ id = 'a2'; roleGuid = $o; principalId = 'u2'; principalType = 'User'; scope = "/subscriptions/$S1"; createdOn = '2023-01-01'; subscriptionId = $S1 }),
            (MkRow @{ id = 'a3'; roleGuid = $o; principalId = 'u3'; principalType = 'User'; scope = "/subscriptions/$S1"; createdOn = '2023-01-01'; subscriptionId = $S1 }),
            (MkRow @{ id = 'a4'; roleGuid = $o; principalId = 'g1'; principalType = 'User'; scope = "/subscriptions/$S1"; createdOn = '2023-01-01'; subscriptionId = $S1 }),
            (MkRow @{ id = 'a5'; roleGuid = $u; principalId = 'gone'; principalType = 'User'; scope = '/providers/Microsoft.Management/managementGroups/lz-prod'; createdOn = '2022-01-01'; subscriptionId = '' }),
            (MkRow @{ id = 'a6'; roleGuid = $o; principalId = 'sp1'; principalType = 'ServicePrincipal'; scope = "/subscriptions/$S2"; createdOn = '2024-01-01'; subscriptionId = $S2 }),
            (MkRow @{ id = 'a7'; roleGuid = $c; principalId = 'grp1'; principalType = 'Group'; scope = "/subscriptions/$S2/resourceGroups/rg-dev"; createdOn = '2024-01-01'; subscriptionId = $S2 }),
            (MkRow @{ id = 'a8'; roleGuid = 'cccccccc-0000-0000-0000-000000000001'; principalId = 'u1'; principalType = 'User'; scope = "/subscriptions/$S1"; createdOn = '2024-01-01'; subscriptionId = $S1 })
            break
        }
        'microsoft\.security/pricings' {
            foreach ($s in $S1, $S2) {
                MkRow @{ id = 'p'; subscriptionId = $s; plan = 'VirtualMachines'; tier = $(if ($s -eq $S1) { 'Standard' } else { 'Free' }); subPlan = $(if ($s -eq $S1) { 'P1' } else { '' }); deprecated = '' }
                MkRow @{ id = 'p'; subscriptionId = $s; plan = 'CloudPosture'; tier = 'Free'; subPlan = ''; deprecated = '' }
                MkRow @{ id = 'p'; subscriptionId = $s; plan = 'KeyVaults'; tier = 'Standard'; subPlan = ''; deprecated = '' }
                MkRow @{ id = 'p'; subscriptionId = $s; plan = 'Dns'; tier = 'Free'; subPlan = ''; deprecated = 'true' }
            }
            break
        }
        'microsoft\.security/securescores' { (MkRow @{ id = 's'; subscriptionId = $S1; current = 22.5; max = 58; pct = 0.388 }), (MkRow @{ id = 's'; subscriptionId = $S2; current = 30; max = 50; pct = 0.6 }); break }
        'microsoft\.security/assessments' { (MkRow @{ displayName = 'Machines should have vulnerability findings resolved'; severity = 'High'; UnhealthyResources = 4 }), (MkRow @{ displayName = 'Storage accounts should restrict network access'; severity = 'Medium'; UnhealthyResources = 3 }), (MkRow @{ displayName = 'MFA should be enabled on accounts with owner permissions'; severity = 'High'; UnhealthyResources = 1 }); break }
        'microsoft\.network/networksecuritygroups' {
            (MkRow @{ id = 'n1'; nsgName = 'nsg-app'; resourceGroup = 'rg-prod'; subscriptionId = $S1; ruleName = 'Allow-RDP'; priority = 100; protocol = 'Tcp'; src = 'Internet'; srcs = '[]'; port = '3389'; ports = '[]'; attachedNics = 1; attachedSubnets = $null }),
            (MkRow @{ id = 'n1'; nsgName = 'nsg-app'; resourceGroup = 'rg-prod'; subscriptionId = $S1; ruleName = 'Allow-HTTPS'; priority = 110; protocol = 'Tcp'; src = '*'; srcs = '[]'; port = '443'; ports = '[]'; attachedNics = 1; attachedSubnets = $null }),
            (MkRow @{ id = 'n2'; nsgName = 'nsg-db'; resourceGroup = 'rg-prod'; subscriptionId = $S1; ruleName = 'Office-SSH'; priority = 100; protocol = 'Tcp'; src = '196.10.1.1/32'; srcs = '[]'; port = '22'; ports = '[]'; attachedNics = $null; attachedSubnets = 1 }),
            (MkRow @{ id = 'n2'; nsgName = 'nsg-db'; resourceGroup = 'rg-prod'; subscriptionId = $S1; ruleName = 'Wide-Range'; priority = 120; protocol = '*'; src = ''; srcs = '["0.0.0.0/0","10.0.0.0/8"]'; port = ''; ports = '["3000-4000","8080"]'; attachedNics = $null; attachedSubnets = 1 }),
            (MkRow @{ id = 'n3'; nsgName = 'nsg-orphan'; resourceGroup = 'rg-dev'; subscriptionId = $S2; ruleName = 'Any'; priority = 100; protocol = '*'; src = '*'; srcs = '[]'; port = '*'; ports = '[]'; attachedNics = $null; attachedSubnets = $null })
            break
        }
        "contains '/networkinterfaces/'" { MkRow @{ id = 'pip1'; name = 'app01-pip'; ip = '102.37.1.10'; sku = 'Standard'; nicId = "/subscriptions/$S1/resourcegroups/rg-prod/providers/microsoft.network/networkinterfaces/app01-nic"; resourceGroup = 'rg-prod'; subscriptionId = $S1 }; break }
        "microsoft\.storage/storageaccounts'\s*\n\| project" {
            (MkRow @{ id = 'st1'; name = 'stprod01'; resourceGroup = 'rg-prod'; subscriptionId = $S1; location = 'southafricanorth'; kind = 'StorageV2'; sku = 'Standard_LRS'; allowBlobPublicAccess = ''; minTls = 'TLS1_0'; httpsOnly = 'true'; sharedKey = ''; publicNetworkAccess = ''; defaultAction = 'Allow'; privateEndpoints = $null }),
            (MkRow @{ id = 'st2'; name = 'stprod02'; resourceGroup = 'rg-prod'; subscriptionId = $S1; location = 'southafricanorth'; kind = 'StorageV2'; sku = 'Standard_GRS'; allowBlobPublicAccess = 'false'; minTls = 'TLS1_2'; httpsOnly = 'true'; sharedKey = 'false'; publicNetworkAccess = 'Disabled'; defaultAction = 'Deny'; privateEndpoints = 1 }),
            (MkRow @{ id = 'st3'; name = 'stlegacy'; resourceGroup = 'rg-dev'; subscriptionId = $S2; location = 'westeurope'; kind = 'Storage'; sku = 'Standard_LRS'; allowBlobPublicAccess = 'true'; minTls = 'TLS1_2'; httpsOnly = 'true'; sharedKey = ''; publicNetworkAccess = ''; defaultAction = 'Allow'; privateEndpoints = $null })
            break
        }
        'microsoft\.keyvault/vaults' { (MkRow @{ id = 'kv1'; name = 'kv-prod'; resourceGroup = 'rg-prod'; subscriptionId = $S1; location = 'southafricanorth'; purgeProtection = ''; softDelete = 'true'; rbac = 'true'; publicNetworkAccess = ''; defaultAction = ''; privateEndpoints = $null }); break }
        'microsoft\.sql/servers'',''microsoft' { (MkRow @{ id = 'sql'; name = 'sql-prod'; type = 'microsoft.sql/servers'; resourceGroup = 'rg-prod'; subscriptionId = $S1; location = 'southafricanorth'; pna = ''; pe = $null }), (MkRow @{ id = 'w'; name = 'web-app'; type = 'microsoft.web/sites'; resourceGroup = 'rg-prod'; subscriptionId = $S1; location = 'southafricanorth'; pna = ''; pe = $null }); break }
        'backupfabrics/protectioncontainers/protecteditems' {
            (MkRow @{ id = 'b1'; src = "$vmBase/app01"; workload = 'VM'; friendlyName = 'app01'; vault = 'rsv-prod'; policy = 'Daily'; protectionState = 'Protected'; lastBackup = (Get-Date).ToUniversalTime().AddHours(-10).ToString('o'); lastStatus = 'Completed'; health = 'Passed'; subscriptionId = $S1 }),
            (MkRow @{ id = 'b2'; src = "$vmBase/dc01"; workload = 'VM'; friendlyName = 'dc01'; vault = 'rsv-prod'; policy = 'Daily'; protectionState = 'Protected'; lastBackup = '2026-09-01T01:00:00Z'; lastStatus = 'Failed'; health = 'ActionRequired'; subscriptionId = $S1 })
            break
        }
        'backupinstances' { MkRow @{ datasourceType = 'Microsoft.Compute/disks'; Count = 2 }; break }
        "microsoft\.recoveryservices/vaults'\s*\n\| project" { MkRow @{ id = 'v'; name = 'rsv-prod'; resourceGroup = 'rg-prod'; subscriptionId = $S1; location = 'southafricanorth'; sku = 'RS0'; redundancy = 'GeoRedundant'; crossRegionRestore = 'Disabled'; softDelete = 'Enabled'; immutability = ''; publicNetworkAccess = 'Enabled' }; break }
        'replicationprotecteditems' { MkRow @{ id = 'r'; name = 'x'; friendlyName = 'web01'; health = 'Normal'; state = 'Protected'; primary = 'South Africa North'; recovery = 'South Africa West'; subscriptionId = $S1 }; break }
        'microsoft\.sql/servers/databases'' and name' { MkRow @{ id = 'db'; name = 'appdb'; server = 'sql-prod'; resourceGroup = 'rg-prod'; subscriptionId = $S1; location = 'southafricanorth'; sku = 'GP_Gen5_2'; tier = 'GeneralPurpose'; zoneRedundant = 'false'; backupRedundancy = 'Geo' }; break }
        'microsoft\.advisor/recommendations' {
            (MkRow @{ id = 'a'; subscriptionId = $S1; category = 'Cost'; impact = 'High'; problem = 'Right-size or shutdown underutilized virtual machines'; solution = 'Right-size'; impactedType = 'Microsoft.Compute/virtualMachines'; impactedName = 'sql01'; annualSavings = 24000.5; monthlySavings = 2000; currency = 'ZAR'; resourceId = 'x' }),
            (MkRow @{ id = 'b'; subscriptionId = $S1; category = 'Cost'; impact = 'Medium'; problem = 'Consider virtual machine reserved instance to save over your on-demand costs'; solution = 'Buy RI'; impactedType = 'Subscriptions'; impactedName = $S1; annualSavings = $null; monthlySavings = 1500; currency = 'ZAR'; resourceId = 'y' }),
            (MkRow @{ id = 'd'; subscriptionId = $S1; category = 'Performance'; impact = 'High'; problem = 'Upgrade to Premium SSD for production disks'; solution = ''; impactedType = 'disk'; impactedName = 'sql01-data'; annualSavings = $null; monthlySavings = $null; currency = ''; resourceId = 'w' }),
            (MkRow @{ id = 'c'; subscriptionId = $S2; category = 'HighAvailability'; impact = 'Medium'; problem = 'Enable Backup'; solution = ''; impactedType = 'vm'; impactedName = 'devbox'; annualSavings = $null; monthlySavings = $null; currency = ''; resourceId = 'z' })
            break
        }
        "diskState\) =~ 'Unattached'" { (MkRow @{ id = 'd1'; name = 'old-data-disk'; resourceGroup = 'rg-prod'; subscriptionId = $S1; location = 'southafricanorth'; sizeGB = 512; sku = 'Premium_LRS'; timeCreated = '2024-01-01'; tags = '{}' }), (MkRow @{ id = 'd2'; name = 'web01-osdisk-ASRReplica'; resourceGroup = 'rg-asr'; subscriptionId = $S1; location = 'southafricawest'; sizeGB = 30; sku = 'Standard_LRS'; timeCreated = '2025-01-01'; tags = '{"ASR-ReplicaDisk":"x"}' }); break }
        'isempty\(tostring\(properties\.ipConfiguration\.id\)\)' { MkRow @{ id = 'p2'; name = 'pip-unused'; resourceGroup = 'rg-dev'; subscriptionId = $S2; location = 'westeurope'; sku = 'Standard'; ip = '20.1.1.1'; allocation = 'Static' }; break }
        "microsoft\.network/networkinterfaces'\s*\n\| where isempty" { MkRow @{ id = 'nic'; name = 'nic-orphan'; resourceGroup = 'rg-dev'; subscriptionId = $S2; location = 'westeurope' }; break }
        'microsoft\.compute/snapshots' { MkRow @{ id = 'sn'; name = 'snap-before-upgrade'; resourceGroup = 'rg-prod'; subscriptionId = $S1; location = 'southafricanorth'; sizeGB = 128; created = '2025-03-01T10:00:00Z' }; break }
        'microsoft\.web/serverfarms' { MkRow @{ id = 'asp'; name = 'asp-old'; resourceGroup = 'rg-dev'; subscriptionId = $S2; location = 'westeurope'; sku = 'S1'; tier = 'Standard' }; break }
        'join kind=leftouter' { MkRow @{ id = 'rg'; name = 'rg-empty'; resourceGroup = 'rg-empty'; subscriptionId = $S2; location = 'westeurope' }; break }
        'microsoft\.operationalinsights/workspaces' { MkRow @{ id = '/subscriptions/x/resourcegroups/rg-mgmt/providers/microsoft.operationalinsights/workspaces/law-prod'; name = 'law-prod'; resourceGroup = 'rg-mgmt'; subscriptionId = $S1; location = 'southafricanorth'; sku = 'PerGB2018'; retentionDays = 30; dailyCapGb = -1 }; break }
        'SecurityInsights' { MkRow @{ id = 's'; name = 'SecurityInsights(law-prod)'; workspace = '/subscriptions/x/resourcegroups/rg-mgmt/providers/microsoft.operationalinsights/workspaces/law-prod' }; break }
        "'/extensions/'" {
            (MkRow @{ id = 'e1'; machineId = "$vmBase/app01"; extType = 'AzureMonitorWindowsAgent'; state = 'Succeeded' }),
            (MkRow @{ id = 'e2'; machineId = "$vmBase/app01"; extType = 'MDE.Windows'; state = 'Succeeded' }),
            (MkRow @{ id = 'e3'; machineId = "$vmBase/sql01"; extType = 'MicrosoftMonitoringAgent'; state = 'Succeeded' }),
            (MkRow @{ id = 'e4'; machineId = "/subscriptions/$S1/resourcegroups/rg-arc/providers/microsoft.hybridcompute/machines/fs01"; extType = 'AzureMonitorWindowsAgent'; state = 'Succeeded' })
            break
        }
        "microsoft\.hybridcompute/machines'\s*\n\| project" { MkRow @{ id = "/subscriptions/$S1/resourcegroups/rg-arc/providers/microsoft.hybridcompute/machines/fs01"; name = 'fs01'; resourceGroup = 'rg-arc'; subscriptionId = $S1; location = 'southafricanorth'; status = 'Connected'; osType = 'windows'; osName = 'windows'; osSku = 'Windows Server 2012 R2 Standard'; osVersion = '6.3.9600'; agentVersion = '1.45'; lastStatusChange = '2026-10-19' }; break }
        'patchassessmentresources' { MkRow @{ id = 'pa'; machineId = "$vmBase/app01"; lastAssess = '2026-10-19T02:00:00Z'; critical = 2; security = 3 }; break }
        'diagnostic settings check|Resources for diagnostic' { break }
        'microsoft\.keyvault/vaults'', ''microsoft\.network/azurefirewalls' { break }
        'actiongroups' { (MkRow @{ type = 'microsoft.insights/actiongroups'; subscriptionId = $S1; Count = 1; ServiceHealth = 0 }), (MkRow @{ type = 'microsoft.insights/metricalerts'; subscriptionId = $S1; Count = 2; ServiceHealth = 0 }); break }
        "publicipaddresses' and tostring\(sku\.name\) =~ 'Basic'" { MkRow @{ id = 'bpip'; name = 'pip-basic'; type = 'microsoft.network/publicipaddresses'; resourceGroup = 'rg-dev'; subscriptionId = $S2; location = 'westeurope' }; break }
        "in~ \('Standard','WAF'\)" { MkRow @{ id = 'agw'; name = 'agw-v1'; type = 'microsoft.network/applicationgateways'; resourceGroup = 'rg-prod'; subscriptionId = $S1; location = 'southafricanorth' }; break }
        "kind =~ 'Storage'" { MkRow @{ id = 'st3'; name = 'stlegacy'; type = 'microsoft.storage/storageaccounts'; resourceGroup = 'rg-dev'; subscriptionId = $S2; location = 'westeurope' }; break }
        'mv-expand sp = properties\.siteProperties' { (MkRow @{ id = '/subscriptions/x/w1'; runtime = 'NODE|16-lts' }), (MkRow @{ id = '/subscriptions/x/w2'; runtime = 'DOTNETCORE|8.0' }); break }
        "microsoft\.web/sites'\s*\n\| project" { (MkRow @{ id = '/subscriptions/x/w1'; name = 'web-app'; kind = 'app,linux'; resourceGroup = 'rg-prod'; subscriptionId = $S1; location = 'southafricanorth'; state = 'Running'; httpsOnly = 'false'; plan = 'asp-prod' }), (MkRow @{ id = '/subscriptions/x/w2'; name = 'api-app'; kind = 'app,linux'; resourceGroup = 'rg-prod'; subscriptionId = $S1; location = 'southafricanorth'; state = 'Running'; httpsOnly = 'true'; plan = 'asp-prod' }); break }
        'microsoft\.containerservice/managedclusters''\s*\n\| project' { MkRow @{ id = 'aks'; name = 'aks-prod'; resourceGroup = 'rg-aks'; subscriptionId = $S1; location = 'southafricanorth'; version = '1.27.3'; tier = 'Free'; nodePools = 2; powerState = 'Running' }; break }
        default { @() }
    }
    if ($q -match 'type in~ \(.microsoft\.keyvault/vaults.,.microsoft\.network/azurefirewalls') {
        $rows = (MkRow @{ id = '/subscriptions/S1/rg/kv-prod'; name = 'kv-prod'; type = 'microsoft.keyvault/vaults'; resourceGroup = 'rg-prod'; subscriptionId = $S1 }),
                (MkRow @{ id = '/subscriptions/S1/rg/web-app'; name = 'web-app'; type = 'microsoft.web/sites'; resourceGroup = 'rg-prod'; subscriptionId = $S1 }),
                (MkRow @{ id = '/subscriptions/S1/rg/agw-v1'; name = 'agw-v1'; type = 'microsoft.network/applicationgateways'; resourceGroup = 'rg-prod'; subscriptionId = $S1 })
    }
    return [pscustomobject]@{ Data = @($rows); SkipToken = $null }
}

function global:Resp($code, $obj) { [pscustomobject]@{ StatusCode = $code; Content = $(if ($obj -is [string]) { $obj } else { J $obj }); Headers = $null } }

function global:Invoke-AzRestMethod {
    param([string]$Path, [string]$Uri, [string]$Method, [string]$Payload)
    $p = if ($Path) { $Path } else { $Uri }
    if ($env:MOCK_FAIL) { return Resp 403 @{ error = @{ message = 'AuthorizationFailed' } } }
    switch -Regex ($p) {
        'Authorization/locks' { if ($p -match $S1) { return Resp 200 @{ value = @(@{ id = "/subscriptions/$S1/resourceGroups/rg-mgmt/providers/Microsoft.Authorization/locks/no-delete"; name = 'no-delete'; properties = @{ level = 'CanNotDelete'; notes = '' } }) } } else { return Resp 200 @{ value = @() } } }
        'roleEligibilityScheduleInstances' { if ($p -match $S1) { return Resp 403 @{ error = @{ code = 'AuthorizationFailed'; message = 'No access to eligibility' } } } else { return Resp 200 @{ value = @() } } }
        'Consumption/budgets' { if ($p -match $S1) { return Resp 200 @{ value = @(@{ name = 'prod-monthly'; properties = @{ amount = 40000; timeGrain = 'Monthly'; notifications = @{ actual80 = @{ enabled = $true }; actual100 = @{ enabled = $true } } } }) } } else { return Resp 200 @{ value = @() } } }
        'CostManagement/query' {
            if ($p -match $S2) { return Resp 401 @{ error = @{ code = 'Unauthorized'; message = 'Cost Management data is not available for this CSP subscription.' } } }
            return Resp 200 @{ properties = @{ nextLink = $null; columns = @(@{ name = 'Cost'; type = 'Number' }, @{ name = 'BillingMonth'; type = 'Datetime' }, @{ name = 'ServiceName'; type = 'String' }, @{ name = 'Currency'; type = 'String' })
                    rows = @(@(41250.237, '2026-08-01T00:00:00', 'Virtual Machines', 'ZAR'), @(9100.1, '2026-08-01T00:00:00', 'Storage', 'ZAR'), @(43100.9, '2026-09-01T00:00:00', 'Virtual Machines', 'ZAR'), @(8800, '2026-09-01T00:00:00', 'Storage', 'ZAR'), @(1200, '2026-09-01T00:00:00', '', 'ZAR')) } }
        }
        "subscriptions/[^/]+/providers/Microsoft\.Insights/diagnosticSettings" { if ($p -match $S1) { return Resp 200 @{ value = @(@{ name = 'al'; properties = @{ workspaceId = '/x/workspaces/law-prod' } }) } } else { return Resp 200 @{ value = @() } } }
        'kv-prod/providers/Microsoft\.Insights/diagnosticSettings' { return Resp 200 @{ value = @(@{ name = 'd'; properties = @{ workspaceId = '/x/workspaces/law-prod' } }) } }
        'diagnosticSettings' { return Resp 200 @{ value = @() } }
        'kubernetesVersions' { return Resp 200 @{ values = @(@{ version = '1.31'; isPreview = $false }, @{ version = '1.32'; isPreview = $false }, @{ version = '1.33'; isPreview = $false }, @{ version = '1.34'; isPreview = $true }) } }
        default { return Resp 404 @{ error = @{ message = "mock: no route for $p" } } }
    }
}

function global:Invoke-RestMethod {
    param([string]$Method, [string]$Uri, $Headers, $Body, [string]$ContentType)
    $ids = ($Body | ConvertFrom-Json).ids
    $known = @{
        u1   = @{ '@odata.type' = '#microsoft.graph.user'; id = 'u1'; displayName = 'Roan Venter'; userPrincipalName = 'roan@contoso.com' }
        u2   = @{ '@odata.type' = '#microsoft.graph.user'; id = 'u2'; displayName = 'Admin Two'; userPrincipalName = 'admin2@contoso.com' }
        u3   = @{ '@odata.type' = '#microsoft.graph.user'; id = 'u3'; displayName = 'Admin Three'; userPrincipalName = 'admin3@contoso.com' }
        g1   = @{ '@odata.type' = '#microsoft.graph.user'; id = 'g1'; displayName = 'Vendor Guy'; userPrincipalName = 'vendor_partner.com#EXT#@contoso.onmicrosoft.com' }
        sp1  = @{ '@odata.type' = '#microsoft.graph.servicePrincipal'; id = 'sp1'; displayName = 'github-deploy'; appId = 'app-123' }
        grp1 = @{ '@odata.type' = '#microsoft.graph.group'; id = 'grp1'; displayName = 'Dev Team' }
    }
    return [pscustomobject]@{ value = @($ids | Where-Object { $known.ContainsKey($_) } | ForEach-Object { [pscustomobject]$known[$_] }) }
}

$p = @{ CustomerName = 'Contoso Mock'; OutputRoot = $Out; NoOpen = $true }
if ($Quick) { $p['QuickMode'] = $true }
& $Script @p
"`nARG queries issued: $($global:ArgCalls.Count)"
$f = Get-ChildItem $Out -Recurse -Filter Findings.csv | Sort-Object LastWriteTime | Select-Object -Last 1
$rows = Import-Csv $f.FullName
$bad = @($rows | Where-Object { $_.CafPhase -eq 'Unmapped' -or -not $_.References })
"Framework check: $($rows.Count) findings, $($bad.Count) unmapped or without references" + $(if ($bad.Count) { ": " + (($bad | ForEach-Object Id) -join ', ') } else { '' })
