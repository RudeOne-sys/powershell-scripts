#Requires -Version 7.0
<#
.SYNOPSIS
    Azure Tenant Assessment Tool v1.1 - Netsurit Professional Services
.DESCRIPTION
    Read-only assessment of a Microsoft Azure estate for customers and prospects.

    Most checks are Azure Resource Graph (KQL) queries that run across every subscription the
    signed-in account can read in a single call. A small number of ARM REST calls cover what
    Resource Graph does not hold: Cost Management, resource locks, PIM eligibility, diagnostic
    settings and AKS supported versions.

    Output (one timestamped folder):
      - Azure_*.csv              evidence files, fixed names and columns (always written, header-only when empty)
      - Findings.csv             findings with suggested severity, count and evidence file
      - NotCollected.csv         anything that could not be read, why, and how to fix it
      - Assessment_Summary.json  headline metrics for the BMAD engine
      - Azure_Assessment_Report.html  branded, A4 print-ready report

    Companion tools: ARI (full inventory + diagrams) and WARA (reliability).
    See Azure-Assessment-Toolkit.md for the full engagement runbook.
.PARAMETER CustomerName
    Customer name used for the output folder and report header.
.PARAMETER TenantId
    Tenant to sign in to (guest, Lighthouse/GDAP or partner access).
.PARAMETER SubscriptionId
    One or more subscription IDs to limit the assessment. Default: every subscription the account can read.
.PARAMETER QuickMode
    Runs only Estate Overview, Security and Cost & Waste.
.PARAMETER SkipCost
    Skips Cost Management queries (Advisor and waste checks still run).
.PARAMETER CostDays
    Days of actual cost to pull. Default 90.
.PARAMETER MaxDiagnosticChecks
    Maximum resources checked for diagnostic settings (one REST call each). Default 300. 0 skips the check.
.PARAMETER UseDeviceCode
    Sign in with a device code instead of a browser window.
.PARAMETER OutputRoot
    Folder to create the output folder in. Default: your Downloads folder.
.PARAMETER NoOpen
    Do not open the output folder and report when finished.
.EXAMPLE
    .\Azure-Tenant-Assessment-v1.ps1 -CustomerName "Contoso"
.EXAMPLE
    .\Azure-Tenant-Assessment-v1.ps1 -CustomerName "Fabrikam" -TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -QuickMode
.EXAMPLE
    .\Azure-Tenant-Assessment-v1.ps1 -CustomerName "Fabrikam" -SubscriptionId "sub-id-1","sub-id-2" -UseDeviceCode
#>
[CmdletBinding()]
param(
    [string]$CustomerName = "",
    [string]$TenantId = "",
    [string[]]$SubscriptionId = @(),
    [switch]$QuickMode,
    [switch]$SkipCost,
    [ValidateRange(7, 365)][int]$CostDays = 90,
    [ValidateRange(0, 5000)][int]$MaxDiagnosticChecks = 300,
    [switch]$UseDeviceCode,
    [string]$OutputRoot = "",
    [switch]$NoOpen
)

$ErrorActionPreference = 'Continue'

# Numbers in CSV/JSON must use a decimal point regardless of the machine's regional settings
$script:OriginalCulture = [System.Threading.Thread]::CurrentThread.CurrentCulture
[System.Threading.Thread]::CurrentThread.CurrentCulture = [System.Globalization.CultureInfo]::InvariantCulture

# ── Script-level state ────────────────────────────────────────────────────────
$script:ToolVersion    = '1.1'
$script:HtmlSections   = [System.Collections.Generic.List[string]]::new()
$script:TocEntries     = [System.Collections.Generic.List[string]]::new()
$script:Findings       = [System.Collections.Generic.List[object]]::new()
$script:NotCollected   = [System.Collections.Generic.List[object]]::new()
$script:Data           = [ordered]@{}
$script:SectionNum     = 0
$script:CurrentSection = ''
$script:SubNames       = @{}
$script:ScopeSubs      = @()
$script:TenantIdUsed   = ''
$script:Account        = ''
$script:Vms            = @()          # shared VM inventory, enriched by several sections
$script:VmById         = @{}
$script:OutputPath     = ''
$script:SubRows        = @()
$script:StorageRows    = @()
$script:ArcRows        = @()
$script:AdvisorRows    = @()
$script:FailedQueries  = @{}

# ── Configuration (edit here, not in the section code) ────────────────────────
$script:PrivilegedRoles = @('Owner', 'User Access Administrator', 'Role Based Access Control Administrator')

# Fallback names for the built-in roles that matter most, in case role definitions are not readable
$script:BuiltInRoleGuids = @{
    '8e3af657-a8ff-443c-a75c-2fe8c4bcb635' = 'Owner'
    'b24988ac-6180-42a0-ab88-20f7382dd24c' = 'Contributor'
    'acdd72a7-3385-48ef-bd42-f606fba81ae7' = 'Reader'
    '18d7d88d-d35e-4fb5-a5c3-7773c20a72d9' = 'User Access Administrator'
    'f58310d9-a9f6-439a-9e8d-f62e7b41a168' = 'Role Based Access Control Administrator'
}

$script:SensitivePorts = [ordered]@{
    22 = 'SSH'; 3389 = 'RDP'; 445 = 'SMB'; 1433 = 'SQL Server'; 3306 = 'MySQL'; 5432 = 'PostgreSQL'
    5985 = 'WinRM'; 5986 = 'WinRM (HTTPS)'; 23 = 'Telnet'; 21 = 'FTP'; 135 = 'RPC'; 139 = 'NetBIOS'
    27017 = 'MongoDB'; 6379 = 'Redis'
}

$script:DefenderPlanNames = [ordered]@{
    'VirtualMachines' = 'Servers'; 'CloudPosture' = 'CSPM'; 'StorageAccounts' = 'Storage'
    'SqlServers' = 'Azure SQL'; 'SqlServerVirtualMachines' = 'SQL on VMs'; 'OpenSourceRelationalDatabases' = 'OSS databases'
    'CosmosDbs' = 'Cosmos DB'; 'AppServices' = 'App Service'; 'KeyVaults' = 'Key Vault'; 'Arm' = 'Resource Manager'
    'Containers' = 'Containers'; 'Api' = 'APIs'; 'AI' = 'AI services'
}

# Windows Server lifecycle by release year
$script:WindowsServerLifecycle = @{
    2008 = @{ Status = 'End of support';                         Severity = 'High';   Note = 'Extended support ended Jan 2020; Azure ESU ended Jan 2024' }
    2012 = @{ Status = 'End of support';                         Severity = 'High';   Note = 'Extended support ended Oct 2023; ESU ended Oct 2026' }
    2016 = @{ Status = 'Extended support ends Jan 2027';         Severity = 'Medium'; Note = 'Plan upgrade now' }
    2019 = @{ Status = 'Supported (extended support to Jan 2029)'; Severity = 'Info'; Note = '' }
    2022 = @{ Status = 'Supported';                              Severity = 'Info';   Note = '' }
    2025 = @{ Status = 'Supported';                              Severity = 'Info';   Note = '' }
}

# App Service runtime strings (LinuxFxVersion / WindowsFxVersion) that are out of community/vendor support
$script:OutdatedRuntimePatterns = @(
    '(?i)^php\|(5|7|8\.0|8\.1)(\.|$)'
    '(?i)^node\|([0-9]|1[0-8])(\.|-|$)'
    '(?i)^python\|(2|3\.[0-9])$'
    '(?i)^dotnetcore\|[1-3]\.'
    '(?i)^dotnet\|[5-7]\.'
    '(?i)^ruby\|2'
)

# Retired / retiring platform features. Status 'Retired' = High, 'Retirement announced' = Medium.
$script:Retirements = @(
    [pscustomobject]@{ Label = 'Basic SKU public IP'; Waf = 'RE:05'; Status = 'Retired'; When = '30 Sep 2025'
        Query = "resources | where type =~ 'microsoft.network/publicipaddresses' and tostring(sku.name) =~ 'Basic' | project id, name, type, resourceGroup, subscriptionId, location" }
    [pscustomobject]@{ Label = 'Basic SKU load balancer'; Waf = 'RE:05'; Status = 'Retired'; When = '30 Sep 2025'
        Query = "resources | where type =~ 'microsoft.network/loadbalancers' and tostring(sku.name) =~ 'Basic' | project id, name, type, resourceGroup, subscriptionId, location" }
    [pscustomobject]@{ Label = 'VM on unmanaged disks'; Waf = 'RE:05'; Status = 'Retired'; When = '30 Sep 2025'
        Query = "resources | where type =~ 'microsoft.compute/virtualmachines' | where isnull(properties.storageProfile.osDisk.managedDisk) | project id, name, type, resourceGroup, subscriptionId, location" }
    [pscustomobject]@{ Label = 'Classic (ASM) resource'; Waf = 'SE:08'; Status = 'Retired'; When = 'Sep 2023 (VMs) / Aug 2024 (Cloud Services)'
        Query = "resources | where type startswith 'microsoft.classic' | project id, name, type, resourceGroup, subscriptionId, location" }
    [pscustomobject]@{ Label = 'Application Gateway v1'; Waf = 'SE:06'; Status = 'Retired'; When = '28 Apr 2026'
        Query = "resources | where type =~ 'microsoft.network/applicationgateways' | where tostring(properties.sku.tier) in~ ('Standard','WAF') | project id, name, type, resourceGroup, subscriptionId, location" }
    [pscustomobject]@{ Label = 'Classic Application Insights (no workspace)'; Waf = 'OE:07'; Status = 'Retired'; When = 'Feb 2024'
        Query = "resources | where type =~ 'microsoft.insights/components' | where isempty(tostring(properties.WorkspaceResourceId)) | project id, name, type, resourceGroup, subscriptionId, location" }
    [pscustomobject]@{ Label = 'General-purpose v1 storage account'; Waf = 'PE:03'; Status = 'Retirement announced'; When = 'Migrate to GPv2'
        Query = "resources | where type =~ 'microsoft.storage/storageaccounts' and kind =~ 'Storage' | project id, name, type, resourceGroup, subscriptionId, location" }
    [pscustomobject]@{ Label = 'Cloud Services (extended support)'; Waf = ''; Status = 'Retirement announced'; When = 'Mar 2027'
        Query = "resources | where type =~ 'microsoft.compute/cloudservices' | project id, name, type, resourceGroup, subscriptionId, location" }
)

$script:DiagnosticTypes = @(
    'microsoft.keyvault/vaults', 'microsoft.network/azurefirewalls', 'microsoft.network/applicationgateways',
    'microsoft.network/frontdoors', 'microsoft.cdn/profiles', 'microsoft.network/virtualnetworkgateways',
    'microsoft.network/bastionhosts', 'microsoft.containerservice/managedclusters', 'microsoft.sql/servers/databases',
    'microsoft.web/sites', 'microsoft.apimanagement/service', 'microsoft.recoveryservices/vaults',
    'microsoft.dbforpostgresql/flexibleservers', 'microsoft.dbformysql/flexibleservers',
    'microsoft.documentdb/databaseaccounts', 'microsoft.logic/workflows'
)

# Netsurit logo (PNG, base64) for the report header
$script:LogoBase64 = 'iVBORw0KGgoAAAANSUhEUgAAAMIAAABACAYAAABImajiAAAMGElEQVR42u2dbYwdVRnHf2dm7i5tl20LhRYSELAaU/ggGhH8wG0jVV4CkcBCJBqsSDAaESSiSLSsxlcU0PJBSjQGDYlcoRLEVjC2C4oiiDXSCoSCIorYF7rtdnd778x5/DDPuKfTuXPvdu+Wsvf8k5OZO3Nm5syZ5/+c//PMmV3zWvWcXQih0BqBMcmsMDQjcXzJkY+sXSsDA6Gp1RIOAoSVgWHQjlXPXzwmyUaDGMEIiGl2jDEks4MoHLXxk/OH1lYFjAHBwyOHSMT0RcaYdqzDqg1FgfmeVAdOZcmS0YNtXPUkDoIwmBNgsIDBNG+viISBMZLIbP+oPUqdPIi1CO0UgJE4ifujnsXbZfe3zeCgpVoND3ajLSLaJilvr1grgjHG+kft0YIIkzzAEA3X6/HhUeWqndXzzjJDQ7EMDIS+Kz26igjqkYNERDByp1QH+mqAUKJRPDxmIhGMMcFYYpO5UeWE7bL75ktqteT1kEgeHq8rETKJtLNeT/qjyse3eYnk0a1EQOVQIiIVI6u9RPLoWiJkEqk/qpzoJZJH1xIhJQPRcL0Re4nk0dVEgIksUgUvkTy6mAjGmGA0sUl/ZUIibfASyaPbiADZi7ZGfLhKpGVeInnMdCIISFGxiLEitgJ3yPLlc6gtES+RPGYsESJjTGEJgnDM2qC/Ujlp+97wVsOgZWAg8N3sMeOIICCxlT2JlZGiYoSR4UZ9eE4QXbb1zHOXmVot8RLJ41BHNBkC9JjA1MVua4T2DGywx9jESLD/pwwShjI3OczsMGMxALWan/3pMTOIMGHlkhy94KV/mtrmeruH+I9hPGYeEcBs+9e8XoFGauPlRu5J4DFTiYCtVMSkmSJv6B7dGSx7ePgRwaPMoQQTeQWScrVIlkWzWjKEtPfeJXEkabvHUNK+fPtxzplvY75+0f6ie81fN2jhiEVLWaLFvfcp9YknQmeQN4ay2EmAuMTAJ4up/hWRUM9hW5BXSu6XA7jXds/htrHde590n3giTA2ZgbwLeJNu+y/wSIHxZL+PB07TbX8FntV9AfB+YJbWM00MywC/AV7TbUuBI0uOyR/bANYBdf2dAEuA84BTgPm671Vt3wbgGa1r1HhPBd6s530GeDp3v9n60cCZuj4M/NrZ9x49RxEZGsBO4O/Ac9rGIFc3As4GevV8DwO7dd9ZwNw2+uT/xPFEmLo3jYHrgYud7e8Ensp5sqzue4Ef6raVwJf1YVWAu/UBtsLpwOO6fivw9kmOXouAbWpc3wSu1usXYRy4B1ih92CBTwOX6/7vAtc495fJnkT7oabbXlDDz+pdm+uzIuwFNgFfA+7V82aSaTbwU12iZP6brt8JnOBHhIOPPfrg9+qDWQ28u4msqDvkGMudZyfQp8YWlnh1d+gfze3LS4P8eUacNn0VuM7xtK8CLyspjgUWAIc5I0Eld7/u9Wni2bN6wwXtiHX/Hv2dGXoILFRv/w7gZ8A5OpJFepxof/UWSLs9LeRSvk8angidC5ZDfSixesLrgG85D64ogAyanCcEBtTDBQXy4QVn/aNKnowMPcAa9foAH3LOkwWfO1S2XK1ti4BVwE26D+AolUCXArfkAumg5B6aBcthwb1GWr4E3K7riS6PA76oo4YBPqdEkJxBhwVtu0QlZra9H/i5jrYJcCHwktsnngidRaKGmKhRrQGe58DS1Jucob4MzzaRFBmeanKeJY6x7FajG3aItxV4SIsrq6YD4zqyuKPnMPARjYGOBt6qo+1oG+fbnPvd6zgjAZ7UkW8fVnpMHeJIm7vUS80C7phEwJbH4Uqqw3SZlUqBZ808dKDXda83x/G+gXO860kDvU72O3JKyPRPpQ+dtmX3UlHp+Koj43smMUK7fdLXok8CT4TOYp7q7sfVoJYBVzrD/WSwXeOJcV1mpdEkDWmbpCTz+xJnJIm1zFECn+EEs3HOi04nYm1b3WljQ+XdCbptp45c7ZBysn3ipVGH0aMP8Crgz/pAbwZ+CbzSxvHGMbyvkKZijZMpCVQyrW4SO7SbNQpVGtylMUYdeJ+WTarF7wcebZK6nA4HcpQTI/QCi1VeZt57je6rFDiDKcMToTNwvdQC4AkN/j6lAdrtGqC1KzEE+GCTfX9QIkxFrlg1rqu1fRc5+07Wcp1ea6XGCdOlHmLgRuAzDukrSoAMj5KmUAOm/gKxqZby6PyDNfpwX1Sj+wBwGROpwXbwoga5z2jZpMuNHYpphDTNeDFwAWm+f2uu3unAr4AraJ7S7QR6dVSYq8s52o8N7cflwC6n3R2HHxGmJ3AW1bOfVFlkgdtIU3i7C4Ls/O+ANH26sUCWiJOhmmo7s1HlAS0LNE64gDTtWtE6q4AHncC104HybaQv7bLszseYeGE3W7NgkeNk/IjwBosX1gI/0X4+Sof37W0en2iJnfVWc4KmYowh6dvmBzTAX66BumgmquqQMD91xLRhX1IiK58Dfk86neO3wA1MvBS7kXTqRzKd9uqJMH3IdPi1KjmsjhCXtxnsZam/bOmWTnlFyQXoWUqxRw3yCb2+MPGCDs3gZIZ9rHNsRoqszcc4x+wuiKcyzGLihWSPJhZ+7Oy/gQNPQ3siHAJEMOplr3EMZUWTIDuPMT1Hllp0S6d08jwm0pXWuV72Ge4ixwDdaQt/cgLbs0jz9HVHFmZp0EudbX9xCF7UV+4IaFQuNfTYi0lfqNnpsllPhOlFog/+bo0Vwkno3MWkk9TeosusLM5lVKaS5VoD/EKzRsdpTJDN8/kO8DYm3iU85hz7sMYLonXvI50T1Kfe/UTSaRlnO/f7I8fo2xlJn9W4xOgocf10jgo+WO6c949LAmADfIJ0urL7ut82OY/VLI40IdaFGn+UzdNv1p4s+D6ZdPoCpFOwd6lx71W5cwQTE+2+TzptIfPmu/R+7tXfy0ln1b6s113ExKzQCvAN0mkNoXPPtqQfMtyiGbdYR5eb2HdqhBtDSRtOKXtJKH5EmB70se+0hCIP9w8N/CrqNSPHWDJPO9/R6D1KGrfM1mWrGGN+SXuya40rmbJ5Sf06+pyiJECvcwvp+5DAkTCBjgIXAVscWzoeOMm5r63AZ1Xj57Nfbp/NKjDagPT9wZNapw/4eu4ejtB9vZSndg3pNxsR6VSS0I8InZc+kH4TcL96mhcKPHEWL6wC/qNGbkgnxOHo6hU5cjQLbjcWeNLsejHp2+JMPm3J7bfO9nNJJ9+dqQRY6GSPniZ9w/xciXy5T+ssU2m0SLdvI/2oZ72SwTjXzfeZ0aCcJqPb5aSzeUXjpoxQY8CH1bAF+HdBv2frI1q3V499JV/XE2HqWReA32mhSbrQXb+nyXkSNawDuX6enGvaqJtleDaz/2zNfPYqaUKGkHQ26INa2jlenJjjsZI2ZsRp1r4G6XcK7fRJXeO0pnU9ETqXdAgcQ5Q247JOfLzfzPhMi7ru/KWyD/eTFu0wNE/pJiXHt/sHAMr+MEI791nU7/vV9UToXLDc7ouuuA2p1SnJ1um2UzKaTVefldWbzHXjVp7Mw8MP6YeQ2DZlBTanS2v9/1vwmLlEMCDlpZYYEBv21P1j8+g0DpkYYaR67qI5YSgjwGiS7Of1g0YU2EpsA5sck/jn5jHTiCADA6Gp1ZIGcn5DzKpG0hjvwez/ciSKQQQbEBiM0YS6l0keM0MamVotkZUrg3lDp/1gj238cX6lZ64Y+sLAFJSgLzTBbP/YPGZmjDAIhkFbb8iKURuPVIyJY7E2FpGi4h+bx4wkgmHQrq9Wo4WPrdsyZpMv9EVRhBhrwBQV/9g8ZuaIACwbGorXV6vRgqF1q3bU6xv6K1FkRXxc7NFdRABYunSpFTCxSa7cm9g9PSZA/H/k8eg2IpjBQbuhWg0XDj30/JjEN/RVohDBjwoe3UUElUiJl0geXU8EQFyJNJ4koz2Bl0ge3UeEfSTSuCSf74u8RPLoQiKoRIplYCBcMLRu1WteInl0KxEAWLJEBEzgs0ge3UyETCLN81kkj64eEXISaUe9MTTXSySPaSKCyb7iblGMzm04+FMcVCKFob1iPLG7eoIgpI0283q22eONRQQD4yLsbVWQtB77/n+ugyuR1q/bMib2pllBuBdktFWbDYwnInUQ/zGPRyn+B5+tFh68wBvAAAAAAElFTkSuQmCC'

# ═════════════════════════════════════════════════════════════════════════════
# Setup
# ═════════════════════════════════════════════════════════════════════════════
function Initialize-RequiredModules {
    Write-Host "Checking required modules..." -ForegroundColor Cyan
    $required = @(
        @{ Name = 'Az.Accounts';      Min = [version]'2.13.0' }
        @{ Name = 'Az.ResourceGraph'; Min = [version]'0.13.0' }
    )
    foreach ($m in $required) {
        $installed = Get-Module -ListAvailable -Name $m.Name | Sort-Object Version -Descending | Select-Object -First 1
        if (!$installed -or [version]$installed.Version -lt $m.Min) {
            Write-Host "  Installing $($m.Name)..." -ForegroundColor Yellow
            Install-Module -Name $m.Name -Scope CurrentUser -Force -AllowClobber -ErrorAction SilentlyContinue
        }
        Import-Module $m.Name -ErrorAction SilentlyContinue
    }
}

function Initialize-OutputDirectory {
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $safeName  = ($CustomerName -replace '[^\w\- ]', '').Trim() -replace '\s+', '_'
    $folder    = if ($safeName) { "Azure_Assessment_${safeName}_$timestamp" } else { "Azure_Assessment_$timestamp" }
    $root      = if ($OutputRoot) { $OutputRoot } else { Join-Path ([Environment]::GetFolderPath('UserProfile')) 'Downloads' }
    $path      = Join-Path $root $folder
    New-Item -ItemType Directory -Path $path -Force | Out-Null
    Write-Host "Output folder: $path" -ForegroundColor Green
    return $path
}

function Resolve-SubscriptionTenant {
    <# Finds the tenant that owns a subscription without signing in: ARM answers an
       unauthenticated request with a 401 whose WWW-Authenticate header names the tenant. #>
    param([Parameter(Mandatory)][string]$Id)
    try {
        Invoke-WebRequest -Uri "https://management.azure.com/subscriptions/$Id`?api-version=2022-12-01" -UseBasicParsing -ErrorAction Stop | Out-Null
    }
    catch {
        $resp   = $_.Exception.Response
        $header = if ($resp.Headers.WwwAuthenticate) { "$($resp.Headers.WwwAuthenticate)" } else { "$($resp.Headers['WWW-Authenticate'])" }
        if ($header -match 'login\.windows\.net/([0-9a-fA-F-]{36})') { return $Matches[1] }
    }
    return $null
}

function Connect-AzureServices {
    Write-Host "`nConnecting to Azure..." -ForegroundColor Cyan
    try {
        # Stop Az asking which subscription to use at sign-in; the script reads all of them
        try { Update-AzConfig -LoginExperienceV2 Off -Scope Process -ErrorAction Stop | Out-Null } catch { }

        # Subscription given without a tenant: look the tenant up, so Az signs straight into it
        # instead of trying every tenant the account is a guest in (MFA warnings for each)
        if ($SubscriptionId -and !$TenantId) {
            $found = Resolve-SubscriptionTenant -Id $SubscriptionId[0].Trim()
            if ($found) {
                Write-Host "  Subscription $($SubscriptionId[0].Trim()) belongs to tenant $found" -ForegroundColor Gray
                $script:TenantId = $found
                $TenantId = $found
            }
        }

        $ctx = Get-AzContext -ErrorAction SilentlyContinue
        $reuse = $ctx -and $ctx.Account -and (!$TenantId -or $ctx.Tenant.Id -eq $TenantId) -and
                 (!$SubscriptionId -or $ctx.Subscription.Id -in $SubscriptionId)
        if ($reuse) {
            Write-Host "  Using existing session: $($ctx.Account.Id) -> $($ctx.Tenant.Id)" -ForegroundColor Green
            Write-Host "  (Run Disconnect-AzAccount first to sign in as someone else.)" -ForegroundColor Gray
        }
        else {
            $p = @{ ErrorAction = 'Stop' }
            if ($TenantId)       { $p['TenantId'] = $TenantId }
            if ($SubscriptionId) { $p['Subscription'] = $SubscriptionId[0].Trim() }
            if ($UseDeviceCode) { $p['UseDeviceAuthentication'] = $true }
            Connect-AzAccount @p | Out-Null
            $ctx = Get-AzContext
        }
        if (!$ctx -or !$ctx.Account) { throw "No Azure session was established (sign-in cancelled or failed)." }
        $script:TenantIdUsed = $ctx.Tenant.Id
        $script:Account      = $ctx.Account.Id
        Write-Host "Connected: $($script:Account) -> $($script:TenantIdUsed)" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "Failed to connect: $_" -ForegroundColor Red
        return $false
    }
}

# ═════════════════════════════════════════════════════════════════════════════
# Query helpers
# ═════════════════════════════════════════════════════════════════════════════
function Invoke-ArgQuery {
    <# Runs a Resource Graph query across the assessment scope, following skip tokens. Never throws. #>
    param([Parameter(Mandatory)][string]$Query, [string]$Name = 'Resource Graph query', [switch]$TenantScope)
    $rows = [System.Collections.Generic.List[object]]::new()
    $p = @{ Query = $Query; First = 1000; ErrorAction = 'Stop' }
    if ($script:ScopeSubs.Count -gt 0 -and -not $TenantScope) { $p['Subscription'] = $script:ScopeSubs }
    else { $p['UseTenantScope'] = $true }
    $page = 0
    try {
        do {
            $resp = Search-AzGraph @p
            $hasData = $null -ne $resp -and ($resp.PSObject.Properties.Name -contains 'Data')
            $data = if ($hasData) { $resp.Data } else { $resp }
            foreach ($r in @($data)) { if ($null -ne $r) { $rows.Add($r) } }
            $token = if ($hasData) { $resp.SkipToken } else { $null }
            if ($token) { $p['SkipToken'] = $token }
            $page++
        } while ($token -and $page -lt 1000)
    }
    catch {
        $msg = $_.Exception.Message
        $script:FailedQueries[$Name] = $true
        Write-Host "    Could not read '$Name': $msg" -ForegroundColor Yellow
        Add-NotCollected -Item $Name -Reason $msg -HowToFix 'Grant Reader at the tenant root management group (or on each subscription) and re-run.'
    }
    return , $rows.ToArray()
}

function Test-QueryOk([string]$Name) { return -not $script:FailedQueries.ContainsKey($Name) }

function Invoke-ArmRequest {
    <# ARM REST call through the Az session. Retries on throttling. Throws on HTTP errors. #>
    param([string]$Path, [string]$Uri, [string]$Method = 'GET', [string]$Payload)
    # Read-only guard: GET only, plus POST to the Cost Management query API (a read that takes its query as a POST body)
    $target = "$Path$Uri"
    if ($Method -ne 'GET' -and -not ($Method -eq 'POST' -and $target -match '/providers/Microsoft\.CostManagement/query\b')) {
        throw "Blocked non-read request ($Method $target): this assessment is read-only."
    }
    $attempt = 0
    while ($true) {
        $attempt++
        $p = @{ Method = $Method; ErrorAction = 'Stop' }
        if ($Uri) { $p['Uri'] = $Uri } else { $p['Path'] = $Path }
        if ($Payload) { $p['Payload'] = $Payload }
        $resp = Invoke-AzRestMethod @p
        if ($resp.StatusCode -eq 429 -and $attempt -le 5) {
            $wait = 15
            foreach ($h in 'Retry-After', 'x-ms-ratelimit-microsoft.costmanagement-entity-retry-after', 'x-ms-ratelimit-microsoft.costmanagement-tenant-retry-after') {
                try { $v = @($resp.Headers.GetValues($h))[0]; if ($v) { $wait = [int]$v; break } } catch { }
            }
            Start-Sleep -Seconds ([math]::Min([math]::Max($wait, 5), 60))
            continue
        }
        if ($resp.StatusCode -ge 400) {
            $msg = try { ($resp.Content | ConvertFrom-Json).error.message } catch { "$($resp.Content)" }
            if (!$msg) { $msg = "$($resp.Content)" }
            throw "HTTP $($resp.StatusCode): $msg"
        }
        if ([string]::IsNullOrWhiteSpace($resp.Content)) { return $null }
        return ($resp.Content | ConvertFrom-Json -Depth 50)
    }
}

function Get-ArmList {
    param([string]$Path)
    $items = [System.Collections.Generic.List[object]]::new()
    $r = Invoke-ArmRequest -Path $Path
    while ($r) {
        foreach ($v in @($r.value)) { if ($null -ne $v) { $items.Add($v) } }
        $r = if ($r.nextLink) { Invoke-ArmRequest -Uri $r.nextLink } else { $null }
    }
    return , $items.ToArray()
}

function Resolve-Principals {
    <# Resolves object IDs to names through Microsoft Graph. Missing IDs = deleted principals. #>
    param([string[]]$Ids)
    $map = @{}
    $list = @($Ids | Where-Object { $_ } | Sort-Object -Unique)
    if ($list.Count -eq 0) { return @{ Ok = $true; Map = $map } }
    try {
        $tok = Get-AzAccessToken -ResourceUrl 'https://graph.microsoft.com' -ErrorAction Stop
        $token = $tok.Token
        if ($token -is [securestring]) { $token = ConvertFrom-SecureString -SecureString $token -AsPlainText }
        $headers = @{ Authorization = "Bearer $token" }
        for ($i = 0; $i -lt $list.Count; $i += 1000) {
            $chunk = @($list[$i..([math]::Min($i + 999, $list.Count - 1))])
            $body = @{ ids = $chunk; types = @('user', 'group', 'servicePrincipal') } | ConvertTo-Json -Depth 3
            $r = Invoke-RestMethod -Method POST -Uri 'https://graph.microsoft.com/v1.0/directoryObjects/getByIds' -Headers $headers -Body $body -ContentType 'application/json' -ErrorAction Stop
            foreach ($o in @($r.value)) {
                $type = "$($o.'@odata.type')" -replace '#microsoft\.graph\.', ''
                $upn  = if ($o.userPrincipalName) { $o.userPrincipalName } elseif ($o.appId) { $o.appId } else { '' }
                $map[$o.id] = [pscustomobject]@{ DisplayName = $o.displayName; UPN = $upn; Type = $type }
            }
        }
        return @{ Ok = $true; Map = $map }
    }
    catch {
        return @{ Ok = $false; Map = $map; Error = $_.Exception.Message }
    }
}

# ═════════════════════════════════════════════════════════════════════════════
# Output helpers
# ═════════════════════════════════════════════════════════════════════════════
function Write-SectionHeader {
    param([string]$Title, [int]$Total)
    $script:SectionNum++
    $script:CurrentSection = $Title
    Write-Host "`n[$($script:SectionNum)/$Total] $Title..." -ForegroundColor Cyan
}

function Add-Finding {
    param(
        [Parameter(Mandatory)][string]$Id,
        [Parameter(Mandatory)][ValidateSet('Critical', 'High', 'Medium', 'Low', 'Info')][string]$Severity,
        [Parameter(Mandatory)][string]$Title,
        [int]$Count = 0,
        [string]$Detail = '',
        [string]$Recommendation = '',
        [string]$EvidenceFile = '',
        [ValidateSet('verified', 'reported', 'inferred')][string]$Confidence = 'verified'
    )
    $fw = Get-FrameworkRef $Id
    $script:Findings.Add([pscustomobject]@{
            Id                = $Id
            Section           = $script:CurrentSection
            SuggestedSeverity = $Severity
            Title             = $Title
            Count             = $Count
            Detail            = $Detail
            Recommendation    = $Recommendation
            EvidenceFile      = $EvidenceFile
            Confidence        = $Confidence
            CafPhase          = $fw.CafPhase
            CafDesignArea     = $fw.CafDesignArea
            WafPillar         = $fw.WafPillar
            WafCodes          = $fw.WafCodes
            ReferencedTo      = $fw.ReferencedTo
            References        = $fw.References
        })
}

function Add-NotCollected {
    param([string]$Item, [string]$Reason, [string]$HowToFix)
    $short = if ($Reason -and $Reason.Length -gt 300) { $Reason.Substring(0, 300) + '...' } else { $Reason }
    $script:NotCollected.Add([pscustomobject]@{
            Section  = $script:CurrentSection
            Item     = $Item
            Reason   = $short
            HowToFix = $HowToFix
        })
}

function Export-AssessmentCsv {
    <# Always writes the file. Empty data = header-only file, which means "none exist". #>
    param([object[]]$Data, [Parameter(Mandatory)][string]$Name, [Parameter(Mandatory)][string[]]$Columns)
    $path = Join-Path $script:OutputPath $Name
    $rows = @($Data | Where-Object { $null -ne $_ })
    if ($rows.Count -gt 0) {
        $rows | Select-Object -Property $Columns | Export-Csv -Path $path -NoTypeInformation -Encoding utf8BOM
    }
    else {
        (($Columns | ForEach-Object { '"' + $_ + '"' }) -join ',') | Out-File -FilePath $path -Encoding utf8BOM
    }
}

function ConvertTo-HtmlEncoded([object]$Value) {
    if ($null -eq $Value) { return '' }
    return [System.Net.WebUtility]::HtmlEncode("$Value")
}

function ConvertTo-HtmlTable {
    param([object[]]$Data, [string[]]$Properties = @(), [int]$MaxRows = 25, [string]$CsvName = '')
    $rows = @($Data | Where-Object { $null -ne $_ })
    if ($rows.Count -eq 0) { return "<p class='no-data'>None found</p>" }
    $cols = if ($Properties.Count -gt 0) { $Properties } else { $rows[0].PSObject.Properties.Name }
    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.Append('<table><thead><tr>')
    foreach ($c in $cols) { [void]$sb.Append("<th>$(ConvertTo-HtmlEncoded (($c -creplace '(?<=[a-z])(?=[A-Z])', ' ') -replace 'Ip\b', 'IP'))</th>") }
    [void]$sb.Append('</tr></thead><tbody>')
    foreach ($row in ($rows | Select-Object -First $MaxRows)) {
        [void]$sb.Append('<tr>')
        foreach ($c in $cols) {
            $v = $row.$c
            if ($c -in 'Severity', 'SuggestedSeverity') { [void]$sb.Append("<td><span class='sev sev-$v'>$(ConvertTo-HtmlEncoded $v)</span></td>") }
            else { [void]$sb.Append("<td>$(ConvertTo-HtmlEncoded $v)</td>") }
        }
        [void]$sb.Append('</tr>')
    }
    [void]$sb.Append('</tbody></table>')
    if ($rows.Count -gt $MaxRows) {
        $ref = if ($CsvName) { " Full list in $CsvName." } else { '' }
        [void]$sb.Append("<p class='note'>Showing $MaxRows of $($rows.Count).$ref</p>")
    }
    return $sb.ToString()
}

function New-StatCard {
    param([string]$Label, [object]$Value, [ValidateSet('info', 'good', 'warn', 'bad', 'neutral')][string]$Tone = 'info')
    return "<div class='stat-card tone-$Tone'><div class='stat-value'>$(ConvertTo-HtmlEncoded $Value)</div><div class='stat-label'>$(ConvertTo-HtmlEncoded $Label)</div></div>"
}

function Get-SectionFindingsHtml {
    param([string]$Section)
    $f = @($script:Findings | Where-Object { $_.Section -eq $Section -and $_.SuggestedSeverity -ne 'Info' } | Sort-Object { Get-SeverityRank $_.SuggestedSeverity })
    if ($f.Count -eq 0) { return "<h3>Findings</h3><p class='no-data'>No issues flagged in this section.</p>" }
    $rows = $f | Select-Object @{N = 'Severity'; E = { $_.SuggestedSeverity } }, Title, Count, Recommendation, ReferencedTo
    return "<h3>Findings</h3>" + (ConvertTo-HtmlTable $rows @('Severity', 'Title', 'Count', 'Recommendation', 'ReferencedTo') -MaxRows 50)
}

function Get-SeverityRank([string]$s) {
    switch ($s) { 'Critical' { 0 } 'High' { 1 } 'Medium' { 2 } 'Low' { 3 } default { 4 } }
}

function Add-HtmlSection {
    param([string]$Title, [string]$Content)
    $sev = @($script:Findings | Where-Object { $_.Section -eq $Title } | ForEach-Object { $_.SuggestedSeverity })
    $status = if ($sev -contains 'Critical' -or $sev -contains 'High') { 'bad' } elseif ($sev -contains 'Medium') { 'warn' } else { 'good' }
    $anchor = 'sec-' + ($Title -replace '[^A-Za-z0-9]+', '-').ToLower().Trim('-')
    $script:TocEntries.Add("<a href='#$anchor'>$(ConvertTo-HtmlEncoded $Title)</a>")
    $body = $Content + (Get-SectionFindingsHtml -Section $Title)
    $align = if ($script:SectionAlignment.ContainsKey($Title)) { "<p class='align'>Referenced to $(ConvertTo-HtmlEncoded $script:SectionAlignment[$Title])</p>" } else { '' }
    $script:HtmlSections.Add("<section class='section $status' id='$anchor'><h2>$(ConvertTo-HtmlEncoded $Title)</h2>$align$body</section>")
}

function Add-ErrorSection {
    param([string]$Title, [object]$ErrorRecord)
    Write-Host "  Error: $ErrorRecord" -ForegroundColor Red
    Add-NotCollected -Item "$Title (section failed)" -Reason "$ErrorRecord" -HowToFix 'Re-run the script; if it fails again send the console output to the toolkit owner.'
    Add-HtmlSection $Title "<p class='error'>This section could not be completed: $(ConvertTo-HtmlEncoded $ErrorRecord)</p>"
}

# ═════════════════════════════════════════════════════════════════════════════
# Small utilities
# ═════════════════════════════════════════════════════════════════════════════
function Get-SubName([string]$Id) {
    if ($Id -and $script:SubNames.ContainsKey($Id)) { return $script:SubNames[$Id] }
    return $Id
}

function Get-OfferType([string]$QuotaId) {
    switch -Regex ($QuotaId) {
        '^EnterpriseAgreement' { return 'Enterprise Agreement' }
        '^MSDNDevTest'         { return 'Dev/Test (EA)' }
        '^MSDN'                { return 'Visual Studio subscriber' }
        '^CSP'                 { return 'CSP' }
        '^PayAsYouGo'          { return 'Pay-As-You-Go / MCA' }
        '^FreeTrial'           { return 'Free trial' }
        '^Sponsored'           { return 'Sponsorship' }
        '^AzureForStudents'    { return 'Azure for Students' }
        '^Internal'            { return 'Microsoft internal' }
        default                { return $(if ($QuotaId) { $QuotaId } else { 'Unknown' }) }
    }
}

function Get-ScopeLevel([string]$Scope) {
    if ($Scope -eq '/') { return 'Root' }
    if ($Scope -match '^/providers/Microsoft\.Management/managementGroups/') { return 'Management group' }
    if ($Scope -match '^/subscriptions/[^/]+$') { return 'Subscription' }
    if ($Scope -match '^/subscriptions/[^/]+/resourceGroups/[^/]+$') { return 'Resource group' }
    return 'Resource'
}

function Test-PortInRange([string]$Range, [int]$Port) {
    $r = "$Range".Trim()
    if ($r -eq '*') { return $true }
    if ($r -match '^(\d+)\s*-\s*(\d+)$') { return ($Port -ge [int]$Matches[1] -and $Port -le [int]$Matches[2]) }
    if ($r -match '^\d+$') { return ([int]$r -eq $Port) }
    return $false
}

function ConvertFrom-JsonArray([string]$Json) {
    if ([string]::IsNullOrWhiteSpace($Json) -or $Json -eq 'null') { return @() }
    try { return @($Json | ConvertFrom-Json -ErrorAction Stop) } catch { return @() }
}

function Get-OsLifecycle {
    <# Returns OS label, lifecycle status and suggested severity from what Azure reports about the machine. #>
    param([string]$OsType, [string]$OsName, [string]$OsVersion, [string]$Offer, [string]$Sku)
    $all = "$OsName $OsVersion $Offer $Sku"
    $unknown = [pscustomobject]@{ Os = ($(if ($OsName) { "$OsName $OsVersion".Trim() } else { "$Offer $Sku".Trim() })); Status = 'Unknown - check manually'; Severity = 'Info' }

    if ($OsType -match '(?i)windows' -or $all -match '(?i)windows') {
        if ($all -match '(?i)windows[ -]?1[01]\b|win1[01]\b|windows-1[01]') {
            if ($all -match '(?i)windows[ -]?11|win11|windows-11') {
                return [pscustomobject]@{ Os = 'Windows 11'; Status = 'Supported'; Severity = 'Info' }
            }
            return [pscustomobject]@{ Os = 'Windows 10'; Status = 'End of support (Oct 2025) - confirm ESU coverage'; Severity = 'High' }
        }
        $year = $null
        if ($OsName -match 'Server\s+(20\d\d)') { $year = [int]$Matches[1] }
        elseif ($Offer -match '(?i)ws(20\d\d)') { $year = [int]$Matches[1] }
        elseif ($Sku -match '^(20\d\d)') { $year = [int]$Matches[1] }
        elseif ($Sku -match '(20\d\d)') { $year = [int]$Matches[1] }
        if ($year -and $script:WindowsServerLifecycle.ContainsKey($year)) {
            $l = $script:WindowsServerLifecycle[$year]
            $label = "Windows Server $year" + $(if ($all -match '(?i)R2') { ' R2' } else { '' })
            return [pscustomobject]@{ Os = $label; Status = $l.Status; Severity = $l.Severity }
        }
        return $unknown
    }

    $t = $all.ToLower()
    if ($t -match 'centos') {
        return [pscustomobject]@{ Os = 'CentOS Linux'; Status = 'End of life (CentOS 7: Jun 2024)'; Severity = 'High' }
    }
    if ($t -match 'ubuntu') {
        $v = $null
        if ($t -match '(1[4-9]|2\d)[._]04') { $v = "$($Matches[1]).04" }
        elseif ($t -match 'trusty') { $v = '14.04' } elseif ($t -match 'xenial') { $v = '16.04' }
        elseif ($t -match 'bionic') { $v = '18.04' } elseif ($t -match 'focal') { $v = '20.04' }
        elseif ($t -match 'jammy') { $v = '22.04' } elseif ($t -match 'noble') { $v = '24.04' }
        if (!$v) { return $unknown }
        $major = [int]($v.Split('.')[0])
        if ($major -le 18) { return [pscustomobject]@{ Os = "Ubuntu $v"; Status = 'End of standard support'; Severity = 'High' } }
        if ($major -eq 20) { return [pscustomobject]@{ Os = "Ubuntu $v"; Status = 'Standard support ended 2025 (ESM needs Ubuntu Pro)'; Severity = 'Medium' } }
        return [pscustomobject]@{ Os = "Ubuntu $v"; Status = 'Supported'; Severity = 'Info' }
    }
    if ($t -match 'rhel|red ?hat') {
        $m = $null
        if ($OsVersion -match '^(\d+)') { $m = [int]$Matches[1] } elseif ($Sku -match '^(\d)') { $m = [int]$Matches[1] }
        if (!$m) { return $unknown }
        if ($m -le 7) { return [pscustomobject]@{ Os = "RHEL $m"; Status = 'End of maintenance support (RHEL 7: Jun 2024)'; Severity = 'High' } }
        return [pscustomobject]@{ Os = "RHEL $m"; Status = 'Supported'; Severity = 'Info' }
    }
    if ($t -match 'sles|suse') {
        $m = $null
        if ($OsVersion -match '^(\d+)') { $m = [int]$Matches[1] } elseif ($Sku -match '(?<!\d)(1[1-5])(?!\d)') { $m = [int]$Matches[1] }
        if (!$m) { return $unknown }
        if ($m -le 12) { return [pscustomobject]@{ Os = "SLES $m"; Status = 'General support ended (SLES 12: Oct 2024)'; Severity = 'High' } }
        return [pscustomobject]@{ Os = "SLES $m"; Status = 'Supported'; Severity = 'Info' }
    }
    if ($t -match 'debian') {
        $m = $null
        if ($OsVersion -match '^(\d+)') { $m = [int]$Matches[1] } elseif ($Sku -match '^(\d+)') { $m = [int]$Matches[1] }
        if (!$m) { return $unknown }
        if ($m -le 11) { return [pscustomobject]@{ Os = "Debian $m"; Status = 'End of life (Debian 11 LTS ended Aug 2026)'; Severity = 'High' } }
        return [pscustomobject]@{ Os = "Debian $m"; Status = 'Supported'; Severity = 'Info' }
    }
    return $unknown
}

function Get-PrincipalTypeLabel([string]$Type) {
    switch -Regex ($Type) {
        '^(?i)user$'             { return 'User' }
        '^(?i)group$'            { return 'Group' }
        '^(?i)serviceprincipal$' { return 'ServicePrincipal' }
        '^(?i)foreigngroup$'     { return 'ForeignGroup' }
        '^$'                     { return 'Unknown' }
        default                  { return $Type }
    }
}

function Get-VmGeneration([string]$Size) {
    if (!$Size) { return 'Unknown' }
    if ($Size -match '^Basic_A') { return 'Retired (Basic A-series)' }
    if ($Size -match '^Standard_A\d+$') { return 'Retired (A-series v1)' }
    if ($Size -match '^Standard_(B|M)\d') { return 'Current' }
    if ($Size -match '_v(\d+)') {
        $v = [int]$Matches[1]
        if ($v -le 3) { return "Previous generation (v$v)" }
        return "Current (v$v)"
    }
    return 'Previous generation (v1)'
}

# ═════════════════════════════════════════════════════════════════════════════
# Framework alignment: Cloud Adoption Framework (CAF) and Well-Architected Framework (WAF)
# Every finding ID maps to a CAF phase, a CAF Azure landing zone design area (blank = workload level)
# and one or more WAF design review checklist codes (blank = platform-only, CAF covers it).
# Checked against Microsoft Learn, October 2026. Findings are "referenced to" these, never "compliant".
# ═════════════════════════════════════════════════════════════════════════════
$script:CafPhaseUrl = @{
    'Strategy' = 'strategy'; 'Plan' = 'plan'; 'Ready' = 'ready'; 'Adopt' = 'adopt'
    'Govern' = 'govern'; 'Secure' = 'secure'; 'Manage' = 'manage'
}
$script:CafDesignAreaUrl = [ordered]@{
    'Azure billing and Microsoft Entra tenant' = 'azure-billing-ad-tenant'
    'Identity and access management'           = 'identity-access'
    'Resource organization'                    = 'resource-org'
    'Network topology and connectivity'        = 'network-topology-and-connectivity'
    'Security'                                 = 'security'
    'Management'                               = 'management'
    'Governance'                               = 'governance'
    'Platform automation and DevOps'           = 'platform-automation-devops'
}
$script:WafPillars = [ordered]@{
    'RE' = @{ Name = 'Reliability';            Folder = 'reliability' }
    'SE' = @{ Name = 'Security';               Folder = 'security' }
    'CO' = @{ Name = 'Cost Optimization';      Folder = 'cost-optimization' }
    'OE' = @{ Name = 'Operational Excellence'; Folder = 'operational-excellence' }
    'PE' = @{ Name = 'Performance Efficiency'; Folder = 'performance-efficiency' }
}
# WAF design review checklist codes -> article slug and short title
$script:WafCodes = @{
    'RE:01' = @('simplify', 'Simplicity and efficiency');            'RE:02' = @('identify-flows', 'Identify and rate flows')
    'RE:03' = @('failure-mode-analysis', 'Failure mode analysis');   'RE:04' = @('metrics', 'Reliability and recovery targets')
    'RE:05' = @('redundancy', 'Redundancy');                         'RE:06' = @('scaling', 'Scaling strategy')
    'RE:07' = @('self-preservation', 'Self-preservation and self-healing'); 'RE:08' = @('reliability-test', 'Reliability testing')
    'RE:09' = @('disaster-recovery', 'Disaster recovery');           'RE:10' = @('monitoring', 'Health monitoring')
    'SE:01' = @('establish-baseline', 'Security baseline');          'SE:02' = @('secure-development-lifecycle', 'Secure development lifecycle')
    'SE:03' = @('data-classification', 'Data classification');       'SE:04' = @('segmentation', 'Segmentation')
    'SE:05' = @('identity-access', 'Identity and access management'); 'SE:06' = @('networking', 'Network controls')
    'SE:07' = @('encryption', 'Encryption');                         'SE:08' = @('harden-resources', 'Hardening')
    'SE:09' = @('application-secrets', 'Application secrets');       'SE:10' = @('monitor-threats', 'Threat monitoring')
    'SE:11' = @('test', 'Security testing');                         'SE:12' = @('incident-response', 'Incident response')
    'CO:01' = @('create-culture-financial-responsibility', 'Financial responsibility'); 'CO:02' = @('cost-model', 'Cost model')
    'CO:03' = @('collect-review-cost-data', 'Cost data and review'); 'CO:04' = @('set-spending-guardrails', 'Spending guardrails')
    'CO:05' = @('get-best-rates', 'Best rates');                     'CO:06' = @('align-usage-to-billing-increments', 'Billing increments')
    'CO:07' = @('optimize-component-costs', 'Component costs');      'CO:08' = @('optimize-environment-costs', 'Environment costs')
    'CO:09' = @('optimize-flow-costs', 'Flow costs');                'CO:10' = @('optimize-data-costs', 'Data costs')
    'CO:11' = @('optimize-code-costs', 'Code costs');                'CO:12' = @('optimize-scaling-costs', 'Scaling costs')
    'CO:13' = @('optimize-personnel-time', 'Personnel time');        'CO:14' = @('consolidation', 'Consolidation')
    'OE:01' = @('devops-culture', 'Standard practices');             'OE:02' = @('formalize-operations-tasks', 'Standardised operations')
    'OE:03' = @('formalize-development-practices', 'Development practices'); 'OE:04' = @('tools-processes', 'Tools and processes')
    'OE:05' = @('infrastructure-as-code-design', 'Infrastructure as code'); 'OE:06' = @('workload-supply-chain', 'Supply chain and pipelines')
    'OE:07' = @('observability', 'Monitoring stack');                'OE:08' = @('incident-response', 'Incident management')
    'OE:09' = @('testing', 'Testing');                               'OE:10' = @('enable-automation', 'Automation')
    'OE:11' = @('safe-deployments', 'Safe deployment')
    'PE:01' = @('performance-targets', 'Performance targets');       'PE:02' = @('capacity-planning', 'Capacity planning')
    'PE:03' = @('select-services', 'Right services');                'PE:04' = @('monitoring', 'Performance measurement')
    'PE:05' = @('scale-partition', 'Scaling and partitioning');      'PE:06' = @('performance-test', 'Performance testing')
    'PE:07' = @('optimize-code-infrastructure', 'Code and infrastructure'); 'PE:08' = @('optimize-data-performance', 'Data usage')
    'PE:09' = @('prioritize-critical-flows', 'Critical flows');      'PE:10' = @('optimize-operational-tasks', 'Operational tasks')
    'PE:11' = @('respond-live-performance-issues', 'Live performance issues'); 'PE:12' = @('continuous-performance-optimize', 'Continuous optimisation')
}

# Finding ID -> 'CAF phase | CAF landing zone design area | WAF codes'
$script:FrameworkMap = @{
    'AZ-OVR-001' = 'Govern|Azure billing and Microsoft Entra tenant|CO:07'
    'AZ-GOV-001' = 'Ready|Resource organization|SE:04'
    'AZ-GOV-002' = 'Govern|Governance|SE:01; CO:04'
    'AZ-GOV-003' = 'Govern|Governance|SE:01'
    'AZ-GOV-004' = 'Govern|Governance|SE:01'
    'AZ-GOV-005' = 'Govern|Resource organization|CO:03'
    'AZ-GOV-006' = 'Govern|Governance|'
    'AZ-IAM-001' = 'Secure|Identity and access management|SE:05'
    'AZ-IAM-002' = 'Secure|Identity and access management|SE:05'
    'AZ-IAM-003' = 'Govern|Identity and access management|SE:05'
    'AZ-IAM-004' = 'Ready|Identity and access management|SE:05'
    'AZ-IAM-005' = 'Secure|Identity and access management|SE:05'
    'AZ-IAM-006' = 'Secure|Identity and access management|SE:05'
    'AZ-IAM-007' = 'Secure|Identity and access management|SE:05'
    'AZ-SEC-001' = 'Secure|Security|SE:01'
    'AZ-SEC-002' = 'Secure|Security|SE:10'
    'AZ-SEC-003' = 'Secure|Security|SE:01; SE:10'
    'AZ-SEC-004' = 'Secure|Network topology and connectivity|SE:06'
    'AZ-SEC-005' = 'Secure|Network topology and connectivity|SE:04; SE:06'
    'AZ-SEC-006' = 'Secure|Security|SE:08'
    'AZ-SEC-007' = 'Secure|Security|SE:07'
    'AZ-SEC-008' = 'Secure|Network topology and connectivity|SE:06'
    'AZ-SEC-009' = 'Secure|Security|SE:09'
    'AZ-SEC-010' = 'Secure|Network topology and connectivity|SE:06'
    'AZ-NET-001' = 'Ready|Network topology and connectivity|'
    'AZ-NET-002' = 'Manage|Network topology and connectivity|'
    'AZ-NET-003' = 'Secure|Network topology and connectivity|SE:04; SE:06'
    'AZ-NET-004' = 'Ready|Network topology and connectivity|SE:06'
    'AZ-NET-005' = 'Ready|Network topology and connectivity|'
    'AZ-NET-006' = 'Secure|Network topology and connectivity|SE:06'
    'AZ-NET-007' = 'Manage|Management|OE:07'
    'AZ-RES-001' = 'Manage|Management|RE:09'
    'AZ-RES-002' = 'Manage|Management|RE:09; RE:10'
    'AZ-RES-003' = 'Manage|Management|RE:09'
    'AZ-RES-004' = 'Adopt||RE:05'
    'AZ-RES-005' = 'Adopt||RE:05'
    'AZ-ADV-001' = 'Manage||RE:05'
    'AZ-ADV-002' = 'Manage||PE:12'
    'AZ-ADV-003' = 'Manage|Management|OE:02'
    'AZ-CST-001' = 'Manage||CO:07'
    'AZ-CST-002' = 'Manage||CO:05'
    'AZ-CST-003' = 'Manage||CO:07'
    'AZ-CST-004' = 'Manage||CO:07'
    'AZ-CST-005' = 'Manage||CO:07'
    'AZ-CST-006' = 'Manage||CO:07'
    'AZ-CST-007' = 'Manage||CO:07'
    'AZ-CST-008' = 'Manage||CO:05'
    'AZ-CST-009' = 'Govern|Governance|CO:03; CO:04'
    'AZ-OPS-001' = 'Manage|Management|OE:07'
    'AZ-OPS-002' = 'Manage|Management|OE:07'
    'AZ-OPS-003' = 'Manage|Management|OE:07'
    'AZ-OPS-004' = 'Manage|Management|SE:08'
    'AZ-OPS-005' = 'Manage|Management|SE:08'
    'AZ-OPS-006' = 'Manage|Management|OE:07; SE:10'
    'AZ-OPS-007' = 'Manage|Management|OE:07'
    'AZ-OPS-008' = 'Manage|Management|RE:10; OE:08'
    'AZ-OPS-009' = 'Secure|Security|SE:10'
    'AZ-LCM-001' = 'Adopt||SE:08'
    'AZ-LCM-002' = 'Adopt||SE:08'
    'AZ-LCM-003' = 'Adopt||PE:03; CO:07'
    'AZ-LCM-004' = 'Adopt||SE:08'
    'AZ-LCM-005' = 'Secure|Security|SE:07'
    'AZ-LCM-006' = 'Adopt||SE:08'
}
# Retirement findings (AZ-LCM-R01..) take their WAF codes from the Waf field of $script:Retirements
for ($i = 0; $i -lt $script:Retirements.Count; $i++) {
    $script:FrameworkMap[('AZ-LCM-R{0:D2}' -f ($i + 1))] = "Adopt||$($script:Retirements[$i].Waf)"
}

# What each report section is referenced to (shown under the section heading)
$script:SectionAlignment = @{
    'Estate Overview'           = 'CAF Ready: Azure billing and Microsoft Entra tenant, Resource organization'
    'Governance'                = 'CAF Govern: Governance, Resource organization · WAF Security (SE:01), Cost Optimization (CO:03, CO:04)'
    'Identity & RBAC'           = 'CAF Secure: Identity and access management · WAF Security (SE:05)'
    'Security'                  = 'CAF Secure: Security · WAF Security (SE:01, SE:06-SE:10)'
    'Network Topology'          = 'CAF Ready: Network topology and connectivity · WAF Security (SE:04, SE:06)'
    'Resiliency & Backup'       = 'CAF Manage: Management (protect and recover) · WAF Reliability (RE:05, RE:09, RE:10)'
    'Cost & Waste'              = 'CAF Govern and Manage: cost management · WAF Cost Optimization (CO:03-CO:07); Advisor by WAF pillar'
    'Operations'                = 'CAF Manage: Management (monitor, administer) · WAF Operational Excellence (OE:07, OE:08), Security (SE:08, SE:10)'
    'Lifecycle & Modernisation' = 'CAF Adopt: Modernize · WAF Security (SE:08), Performance Efficiency (PE:03), Reliability (RE:05)'
}

function Get-WafUrl([string]$Code) {
    $c = $Code.Trim()
    if (!$script:WafCodes.ContainsKey($c)) { return '' }
    return "https://learn.microsoft.com/azure/well-architected/$($script:WafPillars[$c.Substring(0, 2)].Folder)/$($script:WafCodes[$c][0])"
}

function Get-FrameworkRef([string]$Id) {
    $raw = if ($script:FrameworkMap.ContainsKey($Id)) { $script:FrameworkMap[$Id] } else { 'Unmapped||' }
    $p = $raw -split '\|', 3
    $phase = $p[0].Trim(); $area = $p[1].Trim()
    $codes = @($p[2] -split ';' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    $pillars = @($codes | ForEach-Object { $script:WafPillars[$_.Substring(0, 2)].Name } | Select-Object -Unique)
    $urls = [System.Collections.Generic.List[string]]::new()
    if ($area -and $script:CafDesignAreaUrl.Contains($area)) { $urls.Add("https://learn.microsoft.com/azure/cloud-adoption-framework/ready/landing-zone/design-area/$($script:CafDesignAreaUrl[$area])") }
    elseif ($script:CafPhaseUrl.ContainsKey($phase)) { $urls.Add("https://learn.microsoft.com/azure/cloud-adoption-framework/$($script:CafPhaseUrl[$phase])/") }
    foreach ($c in $codes) { $u = Get-WafUrl $c; if ($u) { $urls.Add($u) } }
    $short = @()
    if ($phase) { $short += "CAF $phase" + $(if ($area) { ": $area" } else { '' }) }
    if ($codes.Count) { $short += "WAF " + (($codes | ForEach-Object { "$_ $($script:WafCodes[$_][1])" }) -join ', ') }
    return [pscustomobject]@{
        CafPhase      = $phase
        CafDesignArea = $(if ($area) { $area } else { 'Workload' })
        WafPillar     = ($pillars -join '; ')
        WafCodes      = ($codes -join '; ')
        ReferencedTo  = ($short -join ' · ')
        References    = ($urls -join ' ')
    }
}

# What each CAF design area and WAF pillar is assessed by. Anything marked Workshop cannot be seen
# through the Azure APIs and must be covered in the customer workshop (see Azure-Assessment-Toolkit.md).
$script:FrameworkCoverage = @(
    [pscustomobject]@{ Framework = 'CAF design area'; Area = 'Azure billing and Microsoft Entra tenant'; AssessedBy = 'Script + workshop'; Notes = 'Offer types and subscription state here; EA/MCA/CSP billing structure and tenant count in the workshop' }
    [pscustomobject]@{ Framework = 'CAF design area'; Area = 'Identity and access management'; AssessedBy = 'Script + M365 script'; Notes = 'Azure RBAC and PIM here; Entra Conditional Access, MFA and admin roles in the M365 Tenant Assessment' }
    [pscustomobject]@{ Framework = 'CAF design area'; Area = 'Resource organization'; AssessedBy = 'Script'; Notes = 'Management group hierarchy, subscription placement, tagging' }
    [pscustomobject]@{ Framework = 'CAF design area'; Area = 'Network topology and connectivity'; AssessedBy = 'Script + ARI'; Notes = 'Topology, hubs, peering, overlaps, NSGs, firewalls; ARI diagrams. Hybrid connectivity design and DNS in the workshop' }
    [pscustomobject]@{ Framework = 'CAF design area'; Area = 'Security'; AssessedBy = 'Script'; Notes = 'Defender for Cloud plans, secure score, recommendations, data service hardening' }
    [pscustomobject]@{ Framework = 'CAF design area'; Area = 'Management'; AssessedBy = 'Script'; Notes = 'Monitoring, agents, logs, alerting, patching, backup and DR' }
    [pscustomobject]@{ Framework = 'CAF design area'; Area = 'Governance'; AssessedBy = 'Script'; Notes = 'Azure Policy, compliance, locks, budgets' }
    [pscustomobject]@{ Framework = 'CAF design area'; Area = 'Platform automation and DevOps'; AssessedBy = 'Workshop only'; Notes = 'IaC, pipelines and landing zone deployment method are not visible through the Azure APIs' }
    [pscustomobject]@{ Framework = 'WAF pillar'; Area = 'Reliability'; AssessedBy = 'Script + WARA + workshop'; Notes = 'Placement, redundancy, backup, ASR, Advisor here; WARA for production workloads; RE:04 targets, RE:08 testing and RE:09 DR tests in the workshop' }
    [pscustomobject]@{ Framework = 'WAF pillar'; Area = 'Security'; AssessedBy = 'Script + workshop'; Notes = 'Platform configuration here; SE:02 SDL, SE:03 data classification, SE:11 testing and SE:12 incident response in the workshop' }
    [pscustomobject]@{ Framework = 'WAF pillar'; Area = 'Cost Optimization'; AssessedBy = 'Script + workshop'; Notes = 'Spend, waste, Advisor, budgets, rates here; CO:01 culture and CO:02 cost model in the workshop' }
    [pscustomobject]@{ Framework = 'WAF pillar'; Area = 'Operational Excellence'; AssessedBy = 'Script + workshop'; Notes = 'Monitoring stack and alerting here; OE:01-OE:06, OE:09-OE:11 (practices, IaC, pipelines, deployments) in the workshop' }
    [pscustomobject]@{ Framework = 'WAF pillar'; Area = 'Performance Efficiency'; AssessedBy = 'Advisor + workshop'; Notes = 'Advisor performance recommendations and VM size generation here; PE:01 targets, PE:02 capacity and PE:06 testing in the workshop' }
)

# ═════════════════════════════════════════════════════════════════════════════
# Shared VM inventory (used by Security, Resiliency, Cost, Operations, Lifecycle)
# ═════════════════════════════════════════════════════════════════════════════
function Get-VmInventory {
    $rows = Invoke-ArgQuery -Name 'Virtual machines' -Query @'
resources
| where type =~ 'microsoft.compute/virtualmachines'
| extend size = tostring(properties.hardwareProfile.vmSize), osType = tostring(properties.storageProfile.osDisk.osType),
         powerState = tostring(properties.extended.instanceView.powerState.code),
         osName = tostring(properties.extended.instanceView.osName), osVersion = tostring(properties.extended.instanceView.osVersion),
         imgPublisher = tostring(properties.storageProfile.imageReference.publisher), imgOffer = tostring(properties.storageProfile.imageReference.offer),
         imgSku = tostring(properties.storageProfile.imageReference.sku),
         zones = tostring(zones), avset = tostring(properties.availabilitySet.id), vmss = tostring(properties.virtualMachineScaleSet.id),
         managedOs = isnotnull(properties.storageProfile.osDisk.managedDisk), nics = tostring(properties.networkProfile.networkInterfaces),
         winPatch = tostring(properties.osProfile.windowsConfiguration.patchSettings.patchMode),
         linPatch = tostring(properties.osProfile.linuxConfiguration.patchSettings.patchMode),
         licenseType = tostring(properties.licenseType), priority = tostring(properties.priority)
| project id, name, resourceGroup, subscriptionId, location, size, osType, powerState, osName, osVersion, imgPublisher, imgOffer, imgSku,
          zones, avset, vmss, managedOs, nics, winPatch, linPatch, licenseType, priority
'@
    $list = foreach ($r in $rows) {
        $life = Get-OsLifecycle -OsType $r.osType -OsName $r.osName -OsVersion $r.osVersion -Offer $r.imgOffer -Sku $r.imgSku
        $zones = @(ConvertFrom-JsonArray $r.zones)
        $placement = if ($zones.Count -gt 0) { "Zone $($zones -join ',')" } elseif ($r.vmss) { 'Scale set' } elseif ($r.avset) { 'Availability set' } else { 'Single instance' }
        $nicIds = @(ConvertFrom-JsonArray $r.nics | ForEach-Object { "$($_.id)".ToLower() })
        [pscustomobject]@{
            ResourceId      = "$($r.id)".ToLower()
            Name            = $r.name
            Subscription    = Get-SubName $r.subscriptionId
            SubscriptionId  = $r.subscriptionId
            ResourceGroup   = $r.resourceGroup
            Location        = $r.location
            Size            = $r.size
            SizeGeneration  = Get-VmGeneration $r.size
            OsType          = $r.osType
            OsLabel         = $life.Os
            OsLifecycle     = $life.Status
            OsSeverity      = $life.Severity
            PowerState      = ($r.powerState -replace '^PowerState/', '')
            Placement       = $placement
            ManagedDisks    = [bool]$r.managedOs
            PatchMode       = $(if ($r.winPatch) { $r.winPatch } elseif ($r.linPatch) { $r.linPatch } else { '' })
            LicenseType     = $r.licenseType
            Priority        = $r.priority
            NicIds          = $nicIds
            AzureBackup     = $null
            LastBackup      = ''
            AsrReplicated   = $null
            PublicIp        = ''
            AMA             = $null
            LegacyAgent     = $null
            MDE             = $null
            LastPatchAssess = ''
            MissingCritical = $null
            MissingSecurity = $null
        }
    }
    $script:Vms = @($list)
    $script:VmById = @{}
    foreach ($v in $script:Vms) { $script:VmById[$v.ResourceId] = $v }
    Write-Host "  VMs found: $($script:Vms.Count)" -ForegroundColor Gray
}

# ═════════════════════════════════════════════════════════════════════════════
# 1. Estate Overview
# ═════════════════════════════════════════════════════════════════════════════
function Get-EstateOverview {
    param([int]$Total)
    $title = 'Estate Overview'
    Write-SectionHeader $title $Total
    try {
        $subs = Invoke-ArgQuery -Name 'Subscriptions' -Query @'
resourcecontainers
| where type =~ 'microsoft.resources/subscriptions'
| project id, subscriptionId, name, state = tostring(properties.state), quotaId = tostring(properties.subscriptionPolicies.quotaId),
          spendingLimit = tostring(properties.subscriptionPolicies.spendingLimit), mgChain = tostring(properties.managementGroupAncestorsChain)
'@
        foreach ($s in $subs) { $script:SubNames[$s.subscriptionId] = $s.name }

        $subRows = foreach ($s in $subs) {
            $chain = @(ConvertFrom-JsonArray $s.mgChain)
            $names = @($chain | ForEach-Object { if ($_.displayName) { $_.displayName } else { $_.name } })
            [array]::Reverse($names)
            $parent = if ($chain.Count) { "$($chain[0].name)" } else { '' }
            [pscustomobject]@{
                Subscription        = $s.name
                SubscriptionId      = $s.subscriptionId
                State               = $s.state
                OfferType           = Get-OfferType $s.quotaId
                QuotaId             = $s.quotaId
                SpendingLimit       = $s.spendingLimit
                ManagementGroupPath = ($names -join ' > ')
                ParentGroup         = $parent
                MgDepth             = $chain.Count
                UnderTenantRoot     = ($parent -and $parent -eq $script:TenantIdUsed)
            }
        }
        $script:SubRows = @($subRows)
        Export-AssessmentCsv $subRows 'Azure_Subscriptions.csv' @('Subscription', 'SubscriptionId', 'State', 'OfferType', 'QuotaId', 'SpendingLimit', 'ManagementGroupPath', 'ParentGroup', 'MgDepth', 'UnderTenantRoot')

        $mgs = Invoke-ArgQuery -TenantScope -Name 'Management groups' -Query @'
resourcecontainers
| where type =~ 'microsoft.management/managementgroups'
| project id, name, displayName = tostring(properties.displayName), parent = tostring(properties.details.parent.name)
'@
        $mgRows = $mgs | Select-Object @{N = 'Name'; E = { $_.name } }, @{N = 'DisplayName'; E = { $_.displayName } }, @{N = 'Parent'; E = { $_.parent } }, @{N = 'IsTenantRoot'; E = { $_.name -eq $script:TenantIdUsed } }
        Export-AssessmentCsv $mgRows 'Azure_ManagementGroups.csv' @('Name', 'DisplayName', 'Parent', 'IsTenantRoot')

        $byType   = Invoke-ArgQuery -Name 'Resource count by type' -Query "resources | summarize Count = count() by type | order by Count desc"
        $byRegion = Invoke-ArgQuery -Name 'Resource count by region' -Query "resources | summarize Count = count() by location | order by Count desc"
        $bySub    = Invoke-ArgQuery -Name 'Resource count by subscription' -Query "resources | summarize Count = count() by subscriptionId | order by Count desc"
        $rgCount  = Invoke-ArgQuery -Name 'Resource group count' -Query "resourcecontainers | where type =~ 'microsoft.resources/subscriptions/resourcegroups' | summarize Count = count()"

        $byType    = @($byType | Sort-Object { [int]$_.Count } -Descending)
        $byRegion  = @($byRegion | Sort-Object { [int]$_.Count } -Descending)
        $typeRows   = $byType | Select-Object @{N = 'ResourceType'; E = { $_.type } }, Count
        $regionRows = $byRegion | Select-Object @{N = 'Region'; E = { $_.location } }, Count
        $subCountRows = $bySub | Select-Object @{N = 'Subscription'; E = { Get-SubName $_.subscriptionId } }, @{N = 'SubscriptionId'; E = { $_.subscriptionId } }, Count
        Export-AssessmentCsv $typeRows 'Azure_ResourceCounts_ByType.csv' @('ResourceType', 'Count')
        Export-AssessmentCsv $regionRows 'Azure_ResourceCounts_ByRegion.csv' @('Region', 'Count')
        Export-AssessmentCsv $subCountRows 'Azure_ResourceCounts_BySubscription.csv' @('Subscription', 'SubscriptionId', 'Count')

        Get-VmInventory

        $totalRes = [int](($byType | Measure-Object Count -Sum).Sum)
        $enabled  = @($subRows | Where-Object { $_.State -eq 'Enabled' }).Count
        $arc      = [int](($byType | Where-Object { $_.type -eq 'microsoft.hybridcompute/machines' } | Measure-Object Count -Sum).Sum)
        $script:Data['Subscriptions']        = @($subRows).Count
        $script:Data['SubscriptionsEnabled'] = $enabled
        $script:Data['ManagementGroups']     = @($mgRows).Count
        $script:Data['ResourceGroups']       = $(if ($rgCount) { [int]$rgCount[0].Count } else { 0 })
        $script:Data['Resources']            = $totalRes
        $script:Data['ResourceTypes']        = @($typeRows).Count
        $script:Data['Regions']              = @($regionRows).Count
        $script:Data['VirtualMachines']      = $script:Vms.Count
        $script:Data['ArcMachines']          = $arc
        Write-Host "  Subscriptions: $(@($subRows).Count) | Resources: $totalRes | VMs: $($script:Vms.Count) | Regions: $(@($regionRows).Count)" -ForegroundColor Green

        if (@($subRows).Count -eq 0) {
            Add-NotCollected -Item 'Subscriptions' -Reason 'The signed-in account cannot read any subscription.' -HowToFix 'Assign Reader on the subscriptions (or the tenant root management group) and re-run.'
        }
        if (@($mgRows).Count -eq 0) {
            Add-NotCollected -Item 'Management group hierarchy' -Reason 'No management groups were readable.' -HowToFix 'Assign Reader at the tenant root management group (needs a Global Administrator to elevate access once).'
        }
        $disabled = @($subRows | Where-Object { $_.State -and $_.State -ne 'Enabled' })
        if ($disabled.Count -gt 0) {
            Add-Finding -Id 'AZ-OVR-001' -Severity Low -Title 'Subscriptions not in Enabled state' -Count $disabled.Count `
                -Detail (($disabled | ForEach-Object { "$($_.Subscription) ($($_.State))" }) -join '; ') `
                -Recommendation 'Confirm whether disabled/warned subscriptions are still needed; cancel or reactivate them.' -EvidenceFile 'Azure_Subscriptions.csv'
        }

        $html = "<div class='stat-row'>" +
            (New-StatCard 'Subscriptions' @($subRows).Count) +
            (New-StatCard 'Management groups' @($mgRows).Count) +
            (New-StatCard 'Resource groups' $script:Data['ResourceGroups']) +
            (New-StatCard 'Resources' $totalRes) +
            (New-StatCard 'Virtual machines' $script:Vms.Count) +
            (New-StatCard 'Arc-enabled servers' $arc 'neutral') +
            (New-StatCard 'Regions' @($regionRows).Count 'neutral') +
            "</div><h3>Subscriptions</h3>" +
            (ConvertTo-HtmlTable $subRows @('Subscription', 'State', 'OfferType', 'ManagementGroupPath') -MaxRows 30 -CsvName 'Azure_Subscriptions.csv') +
            "<div class='two-col'><div><h3>Top resource types</h3>" +
            (ConvertTo-HtmlTable $typeRows @('ResourceType', 'Count') -MaxRows 15 -CsvName 'Azure_ResourceCounts_ByType.csv') +
            "</div><div><h3>Regions</h3>" +
            (ConvertTo-HtmlTable $regionRows @('Region', 'Count') -MaxRows 15 -CsvName 'Azure_ResourceCounts_ByRegion.csv') +
            "</div></div>"
        Add-HtmlSection $title $html
    }
    catch { Add-ErrorSection $title $_ }
}

# ═════════════════════════════════════════════════════════════════════════════
# 2. Governance
# ═════════════════════════════════════════════════════════════════════════════
function Get-Governance {
    param([int]$Total)
    $title = 'Governance'
    Write-SectionHeader $title $Total
    try {
        # Management group structure
        $subRows   = @($script:SubRows)
        $atRoot    = @($subRows | Where-Object { $_.UnderTenantRoot })
        $maxDepth  = [int](($subRows | Measure-Object MgDepth -Maximum).Maximum)
        $script:Data['SubscriptionsUnderTenantRoot'] = $atRoot.Count
        $script:Data['ManagementGroupMaxDepth']      = $maxDepth

        # Policy assignments
        $pa = Invoke-ArgQuery -TenantScope -Name 'Policy assignments' -Query @'
policyresources
| where type =~ 'microsoft.authorization/policyassignments'
| project id, name, displayName = tostring(properties.displayName), scope = tostring(properties.scope),
          definitionId = tostring(properties.policyDefinitionId), enforcementMode = tostring(properties.enforcementMode)
'@
        if ($script:ScopeSubs.Count -gt 0) {
            $pa = @($pa | Where-Object { $_.scope -notmatch '^/subscriptions/' -or ($script:ScopeSubs -contains ($_.scope -split '/')[2]) })
        }
        $paRows = foreach ($a in $pa) {
            [pscustomobject]@{
                Assignment      = $(if ($a.displayName) { $a.displayName } else { $a.name })
                Name            = $a.name
                ScopeLevel      = Get-ScopeLevel $a.scope
                Scope           = $a.scope
                Kind            = $(if ($a.definitionId -match '/policySetDefinitions/') { 'Initiative' } else { 'Policy' })
                Source          = $(if ($a.definitionId -match '^/providers/Microsoft\.Authorization/') { 'Built-in' } else { 'Custom' })
                EnforcementMode = $(if ($a.enforcementMode) { $a.enforcementMode } else { 'Default' })
                AssignmentId    = "$($a.id)".ToLower()
            }
        }
        $paRows = @($paRows)
        Export-AssessmentCsv $paRows 'Azure_PolicyAssignments.csv' @('Assignment', 'Name', 'ScopeLevel', 'Scope', 'Kind', 'Source', 'EnforcementMode', 'AssignmentId')
        $nonDefault = @($paRows | Where-Object { $_.Name -ne 'SecurityCenterBuiltIn' })
        $doNotEnforce = @($paRows | Where-Object { $_.EnforcementMode -eq 'DoNotEnforce' })

        # Policy compliance
        $comp = Invoke-ArgQuery -Name 'Policy compliance' -Query @'
policyresources
| where type =~ 'microsoft.policyinsights/policystates'
| extend state = tostring(properties.complianceState), rid = tolower(tostring(properties.resourceId))
| summarize nc = max(iff(state =~ 'NonCompliant', 1, 0)) by subscriptionId, rid
| summarize Evaluated = count(), NonCompliant = sum(nc) by subscriptionId
'@
        $compRows = foreach ($c in $comp) {
            [pscustomobject]@{
                Subscription   = Get-SubName $c.subscriptionId
                SubscriptionId = $c.subscriptionId
                Evaluated      = [int]$c.Evaluated
                NonCompliant   = [int]$c.NonCompliant
                CompliancePct  = $(if ([int]$c.Evaluated -gt 0) { [math]::Round((1 - [int]$c.NonCompliant / [int]$c.Evaluated) * 100, 1) } else { $null })
            }
        }
        $compRows = @($compRows)
        Export-AssessmentCsv $compRows 'Azure_PolicyCompliance_BySubscription.csv' @('Subscription', 'SubscriptionId', 'Evaluated', 'NonCompliant', 'CompliancePct')
        $evalTotal = [int](($compRows | Measure-Object Evaluated -Sum).Sum)
        $ncTotal   = [int](($compRows | Measure-Object NonCompliant -Sum).Sum)
        $compPct   = if ($evalTotal -gt 0) { [math]::Round((1 - $ncTotal / $evalTotal) * 100, 1) } else { $null }

        $top = Invoke-ArgQuery -Name 'Top non-compliant policies' -Query @'
policyresources
| where type =~ 'microsoft.policyinsights/policystates'
| where tostring(properties.complianceState) =~ 'NonCompliant'
| extend assignmentId = tolower(tostring(properties.policyAssignmentId)), definitionName = tostring(properties.policyDefinitionName), rid = tolower(tostring(properties.resourceId))
| summarize Resources = dcount(rid) by assignmentId, definitionName
| order by Resources desc
| take 25
'@
        $defNames = @{}
        $names = @($top | ForEach-Object { $_.definitionName } | Where-Object { $_ } | Sort-Object -Unique)
        if ($names.Count -gt 0) {
            $inList = ($names | ForEach-Object { "'" + ($_ -replace "'", "") + "'" }) -join ','
            $defs = Invoke-ArgQuery -TenantScope -Name 'Policy definition names' -Query "policyresources | where type =~ 'microsoft.authorization/policydefinitions' | where name in~ ($inList) | project id, name, displayName = tostring(properties.displayName)"
            foreach ($d in $defs) { $defNames[$d.name.ToLower()] = $d.displayName }
        }
        $paById = @{}; foreach ($a in $paRows) { $paById[$a.AssignmentId] = $a.Assignment }
        $topRows = foreach ($t in $top) {
            [pscustomobject]@{
                Policy                = $(if ($defNames.ContainsKey("$($t.definitionName)".ToLower())) { $defNames["$($t.definitionName)".ToLower()] } else { $t.definitionName })
                Assignment            = $(if ($paById.ContainsKey($t.assignmentId)) { $paById[$t.assignmentId] } else { ($t.assignmentId -split '/')[-1] })
                NonCompliantResources = [int]$t.Resources
            }
        }
        Export-AssessmentCsv $topRows 'Azure_PolicyNonCompliance_Top.csv' @('Policy', 'Assignment', 'NonCompliantResources')

        # Tags
        $tagRes = Invoke-ArgQuery -Name 'Tag coverage (resources)' -Query "resources | summarize Total = count(), Untagged = countif(isnull(tags) or tostring(tags) == '{}' or tostring(tags) == '')"
        $tagRg  = Invoke-ArgQuery -Name 'Tag coverage (resource groups)' -Query "resourcecontainers | where type =~ 'microsoft.resources/subscriptions/resourcegroups' | summarize Total = count(), Untagged = countif(isnull(tags) or tostring(tags) == '{}' or tostring(tags) == '')"
        $tagRows = @(
            foreach ($pair in @(@('Resources', $tagRes), @('Resource groups', $tagRg))) {
                $r = @($pair[1])[0]
                if ($r) {
                    [pscustomobject]@{ Scope = $pair[0]; Total = [int]$r.Total; Untagged = [int]$r.Untagged
                        PctTagged = $(if ([int]$r.Total -gt 0) { [math]::Round((1 - [int]$r.Untagged / [int]$r.Total) * 100, 1) } else { $null }) }
                }
            }
        )
        Export-AssessmentCsv $tagRows 'Azure_TagCoverage.csv' @('Scope', 'Total', 'Untagged', 'PctTagged')
        $keys = Invoke-ArgQuery -Name 'Tag keys' -Query @'
resources
| where isnotempty(tags)
| project tags
| mvexpand tags
| extend tagKey = tostring(bag_keys(tags)[0])
| summarize Resources = count() by tagKey
| order by Resources desc
| take 40
'@
        $keyRows = $keys | Select-Object @{N = 'TagKey'; E = { $_.tagKey } }, Resources
        Export-AssessmentCsv $keyRows 'Azure_TagKeys.csv' @('TagKey', 'Resources')
        $resTagPct = ($tagRows | Where-Object { $_.Scope -eq 'Resources' }).PctTagged

        # Locks (per subscription, REST)
        $lockRows = [System.Collections.Generic.List[object]]::new()
        $lockFailed = 0
        foreach ($s in $subRows) {
            try {
                $locks = Get-ArmList "/subscriptions/$($s.SubscriptionId)/providers/Microsoft.Authorization/locks?api-version=2020-05-01"
                foreach ($l in $locks) {
                    $scope = ($l.id -split '/providers/Microsoft.Authorization/locks/', 2)[0]
                    $lockRows.Add([pscustomobject]@{
                            Subscription = $s.Subscription; Lock = $l.name; Level = $l.properties.level
                            ScopeLevel = Get-ScopeLevel $scope; Scope = $scope; Notes = $l.properties.notes
                        })
                }
            }
            catch { $lockFailed++ }
        }
        if ($lockFailed -gt 0) { Add-NotCollected -Item 'Resource locks' -Reason "Locks could not be read on $lockFailed subscription(s)." -HowToFix 'Reader on the subscription includes lock read access; check the account has Reader on every subscription.' }
        Export-AssessmentCsv $lockRows 'Azure_Locks.csv' @('Subscription', 'Lock', 'Level', 'ScopeLevel', 'Scope', 'Notes')

        $script:Data['PolicyAssignments']        = $paRows.Count
        $script:Data['PolicyAssignmentsCustom']  = @($paRows | Where-Object { $_.Source -eq 'Custom' }).Count
        $script:Data['PolicyCompliancePct']      = $compPct
        $script:Data['PolicyNonCompliantResources'] = $ncTotal
        $script:Data['ResourcesTaggedPct']       = $resTagPct
        $script:Data['ResourceLocks']            = $lockRows.Count
        Write-Host "  Policy assignments: $($paRows.Count) | Compliance: $(if ($null -ne $compPct) { "$compPct%" } else { 'N/A' }) | Tagged: $resTagPct% | Locks: $($lockRows.Count)" -ForegroundColor Green

        # Findings
        if ($atRoot.Count -gt 0) {
            Add-Finding -Id 'AZ-GOV-001' -Severity Medium -Title 'Subscriptions sit directly under the Tenant Root Group' -Count $atRoot.Count `
                -Detail (($atRoot | ForEach-Object { $_.Subscription }) -join '; ') `
                -Recommendation 'Introduce a management group hierarchy (e.g. Platform / Landing Zones / Sandbox) so policy and RBAC are applied once, not per subscription.' -EvidenceFile 'Azure_Subscriptions.csv'
        }
        if ($nonDefault.Count -eq 0 -and (Test-QueryOk 'Policy assignments') -and $subRows.Count -gt 0) {
            Add-Finding -Id 'AZ-GOV-002' -Severity High -Title 'No Azure Policy assignments beyond the Defender for Cloud default' -Count 0 `
                -Recommendation 'Assign baseline guardrails (allowed locations, required tags, deny public IPs on NICs, audit diagnostic settings) at management group scope.' -EvidenceFile 'Azure_PolicyAssignments.csv'
        }
        if ($doNotEnforce.Count -gt 0) {
            Add-Finding -Id 'AZ-GOV-003' -Severity Low -Title 'Policy assignments set to DoNotEnforce' -Count $doNotEnforce.Count `
                -Recommendation 'Review whether these were left in audit-only mode by mistake.' -EvidenceFile 'Azure_PolicyAssignments.csv'
        }
        if ($null -ne $compPct -and $compPct -lt 80) {
            Add-Finding -Id 'AZ-GOV-004' -Severity $(if ($compPct -lt 60) { 'High' } else { 'Medium' }) -Title "Azure Policy resource compliance is $compPct%" -Count $ncTotal `
                -Recommendation 'Work through the top non-compliant policies; remediate or exempt with a documented reason.' -EvidenceFile 'Azure_PolicyNonCompliance_Top.csv'
        }
        if ($null -ne $resTagPct -and $resTagPct -lt 80) {
            Add-Finding -Id 'AZ-GOV-005' -Severity $(if ($resTagPct -lt 50) { 'Medium' } else { 'Low' }) -Title "Only $resTagPct% of resources carry tags" `
                -Count ([int](@($tagRows | Where-Object { $_.Scope -eq 'Resources' })[0].Untagged)) `
                -Recommendation 'Agree a minimum tag set (owner, environment, cost centre, application) and enforce it with Policy (inherit from resource group).' -EvidenceFile 'Azure_TagCoverage.csv'
        }
        if ($lockRows.Count -eq 0 -and $lockFailed -eq 0 -and $subRows.Count -gt 0) {
            Add-Finding -Id 'AZ-GOV-006' -Severity Low -Title 'No resource locks anywhere in the estate' -Count 0 `
                -Recommendation 'Add CanNotDelete locks to shared platform resources (hub networking, Log Analytics, Recovery Services vaults, Key Vaults).' -EvidenceFile 'Azure_Locks.csv'
        }

        $html = "<div class='stat-row'>" +
            (New-StatCard 'Subscriptions under root' $atRoot.Count $(if ($atRoot.Count) { 'warn' } else { 'good' })) +
            (New-StatCard 'Max MG depth' $maxDepth 'neutral') +
            (New-StatCard 'Policy assignments' $paRows.Count $(if ($nonDefault.Count) { 'info' } else { 'bad' })) +
            (New-StatCard 'Policy compliance' $(if ($null -ne $compPct) { "$compPct%" } else { 'N/A' }) $(if ($null -eq $compPct) { 'neutral' } elseif ($compPct -ge 80) { 'good' } elseif ($compPct -ge 60) { 'warn' } else { 'bad' })) +
            (New-StatCard 'Resources tagged' $(if ($null -ne $resTagPct) { "$resTagPct%" } else { 'N/A' }) $(if ($resTagPct -ge 80) { 'good' } elseif ($resTagPct -ge 50) { 'warn' } else { 'bad' })) +
            (New-StatCard 'Resource locks' $lockRows.Count 'neutral') +
            "</div><h3>Policy assignments</h3>" +
            (ConvertTo-HtmlTable $paRows @('Assignment', 'ScopeLevel', 'Kind', 'Source', 'EnforcementMode') -CsvName 'Azure_PolicyAssignments.csv') +
            "<h3>Most non-compliant policies</h3>" +
            (ConvertTo-HtmlTable $topRows @('Policy', 'Assignment', 'NonCompliantResources') -MaxRows 10 -CsvName 'Azure_PolicyNonCompliance_Top.csv') +
            "<h3>Most used tag keys</h3>" +
            (ConvertTo-HtmlTable $keyRows @('TagKey', 'Resources') -MaxRows 15 -CsvName 'Azure_TagKeys.csv')
        Add-HtmlSection $title $html
    }
    catch { Add-ErrorSection $title $_ }
}

# ═════════════════════════════════════════════════════════════════════════════
# 3. Identity & RBAC
# ═════════════════════════════════════════════════════════════════════════════
function Get-IdentityRbac {
    param([int]$Total)
    $title = 'Identity & RBAC'
    Write-SectionHeader $title $Total
    try {
        $defs = Invoke-ArgQuery -TenantScope -Name 'Role definitions' -Query @'
authorizationresources
| where type =~ 'microsoft.authorization/roledefinitions'
| extend roleName = tostring(properties.roleName), roleType = tostring(properties.type),
         actions = tostring(properties.permissions[0].actions), scopes = tostring(properties.assignableScopes)
| project id, roleGuid = tolower(name), roleName, roleType, actions, scopes
'@
        $roleNames = @{} + $script:BuiltInRoleGuids
        $custom = [System.Collections.Generic.List[object]]::new()
        $seenCustom = @{}
        foreach ($d in $defs) {
            if ($d.roleName) { $roleNames[$d.roleGuid] = $d.roleName }
            if ($d.roleType -eq 'CustomRole' -and !$seenCustom.ContainsKey($d.roleGuid)) {
                $seenCustom[$d.roleGuid] = $true
                $acts = @(ConvertFrom-JsonArray $d.actions)
                $custom.Add([pscustomobject]@{
                        Role             = $d.roleName
                        RoleId           = $d.roleGuid
                        Actions          = ($acts -join '; ')
                        WildcardActions  = [bool](@($acts | Where-Object { $_ -eq '*' -or $_ -match '^Microsoft\.Authorization/\*' }).Count)
                        AssignableScopes = ((ConvertFrom-JsonArray $d.scopes) -join '; ')
                    })
            }
        }
        Export-AssessmentCsv $custom 'Azure_CustomRoles.csv' @('Role', 'RoleId', 'WildcardActions', 'Actions', 'AssignableScopes')

        $ra = Invoke-ArgQuery -TenantScope -Name 'Role assignments' -Query @'
authorizationresources
| where type =~ 'microsoft.authorization/roleassignments'
| extend roleDefinitionId = tolower(tostring(properties.roleDefinitionId)), principalId = tostring(properties.principalId),
         principalType = tostring(properties.principalType), scope = tostring(properties.scope), createdOn = tostring(properties.createdOn)
| extend roleGuid = extract('roledefinitions/([0-9a-f-]+)', 1, roleDefinitionId)
| project id, roleGuid, principalId, principalType, scope, createdOn, subscriptionId
'@
        if ($script:ScopeSubs.Count -gt 0) {
            $ra = @($ra | Where-Object { $_.scope -notmatch '^/subscriptions/' -or ($script:ScopeSubs -contains ($_.scope -split '/')[2]) })
        }

        Write-Host "  Resolving $(@($ra | ForEach-Object principalId | Sort-Object -Unique).Count) principals through Microsoft Graph..." -ForegroundColor Gray
        $res = Resolve-Principals -Ids @($ra | ForEach-Object { $_.principalId })
        if (!$res.Ok) {
            Add-NotCollected -Item 'Principal names (Microsoft Graph)' -Reason $res.Error -HowToFix 'Assign Directory Readers in Entra ID to the assessment account so names, guests and deleted principals can be identified.'
        }

        $raRows = foreach ($a in $ra) {
            $p = $res.Map[$a.principalId]
            $role = if ($roleNames.ContainsKey($a.roleGuid)) { $roleNames[$a.roleGuid] } else { $a.roleGuid }
            $scopeLevel = Get-ScopeLevel $a.scope
            $subId = if ($a.scope -match '^/subscriptions/([^/]+)') { $Matches[1] } else { '' }
            [pscustomobject]@{
                Role           = $role
                Privileged     = ($script:PrivilegedRoles -contains $role)
                PrincipalType  = (Get-PrincipalTypeLabel $(if ($p) { $p.Type } else { $a.principalType }))
                Principal      = $(if ($p) { $p.DisplayName } elseif ($res.Ok) { "Deleted identity ($($a.principalId))" } else { '' })
                PrincipalUPN   = $(if ($p) { $p.UPN } else { '' })
                PrincipalId    = $a.principalId
                Guest          = [bool]($p -and $p.UPN -match '#EXT#')
                Orphaned       = $(if ($res.Ok) { -not [bool]$p } else { $null })
                ScopeLevel     = $scopeLevel
                Subscription   = $(if ($subId) { Get-SubName $subId } else { '' })
                Scope          = $a.scope
                CreatedOn      = $a.createdOn
            }
        }
        $raRows = @($raRows)
        Export-AssessmentCsv $raRows 'Azure_RoleAssignments.csv' @('Role', 'Privileged', 'PrincipalType', 'Principal', 'PrincipalUPN', 'PrincipalId', 'Guest', 'Orphaned', 'ScopeLevel', 'Subscription', 'Scope', 'CreatedOn')

        # PIM eligible assignments (per subscription)
        $pimRows = [System.Collections.Generic.List[object]]::new()
        $pimFailed = 0
        foreach ($s in @($script:SubRows)) {
            try {
                $el = Get-ArmList "/subscriptions/$($s.SubscriptionId)/providers/Microsoft.Authorization/roleEligibilityScheduleInstances?api-version=2020-10-01"
                foreach ($e in $el) {
                    $g = ("$($e.properties.roleDefinitionId)" -split '/')[-1].ToLower()
                    $pimRows.Add([pscustomobject]@{
                            Subscription = $s.Subscription
                            Role         = $(if ($roleNames.ContainsKey($g)) { $roleNames[$g] } else { $g })
                            PrincipalId  = $e.properties.principalId
                            PrincipalType = $e.properties.principalType
                            Scope        = $e.properties.scope
                            EndDateTime  = $e.properties.endDateTime
                        })
                }
            }
            catch { $pimFailed++ }
        }
        if ($pimFailed -gt 0) { Add-NotCollected -Item 'PIM eligible assignments' -Reason "Could not read eligibility schedules on $pimFailed subscription(s) (PIM may not be licensed or readable)." -HowToFix 'Confirm Entra ID P2 / Governance licensing; Reader on the subscription is enough to read eligibility.' }
        Export-AssessmentCsv $pimRows 'Azure_PIM_Eligible.csv' @('Subscription', 'Role', 'PrincipalId', 'PrincipalType', 'Scope', 'EndDateTime')

        $priv          = @($raRows | Where-Object { $_.Privileged })
        $privHigh      = @($priv | Where-Object { $_.ScopeLevel -in 'Root', 'Management group', 'Subscription' })
        $guestPriv     = @($priv | Where-Object { $_.Guest })
        $orphaned      = @($raRows | Where-Object { $_.Orphaned -eq $true })
        $userDirect    = @($raRows | Where-Object { $_.PrincipalType -eq 'User' })
        $wildcard      = @($custom | Where-Object { $_.WildcardActions })
        $ownersPerSub  = $raRows | Where-Object { $_.Role -eq 'Owner' -and $_.ScopeLevel -eq 'Subscription' } | Group-Object Subscription
        $tooManyOwners = @($ownersPerSub | Where-Object { $_.Count -gt 3 })
        $spOwners      = @($priv | Where-Object { $_.PrincipalType -eq 'ServicePrincipal' -and $_.ScopeLevel -in 'Root', 'Management group', 'Subscription' })

        $script:Data['RoleAssignments']             = $raRows.Count
        $script:Data['PrivilegedAssignments']       = $priv.Count
        $script:Data['PrivilegedAtSubscriptionOrAbove'] = $privHigh.Count
        $script:Data['GuestPrivilegedAssignments']  = $(if ($res.Ok) { $guestPriv.Count } else { $null })
        $script:Data['OrphanedAssignments']         = $(if ($res.Ok) { $orphaned.Count } else { $null })
        $script:Data['CustomRoles']                 = $custom.Count
        $script:Data['PimEligibleAssignments']      = $pimRows.Count
        Write-Host "  Assignments: $($raRows.Count) | Privileged: $($priv.Count) | Guests privileged: $($guestPriv.Count) | Orphaned: $($orphaned.Count) | PIM eligible: $($pimRows.Count)" -ForegroundColor Green

        if ($tooManyOwners.Count -gt 0) {
            Add-Finding -Id 'AZ-IAM-001' -Severity Medium -Title 'Subscriptions with more than 3 direct Owners' -Count $tooManyOwners.Count `
                -Detail (($tooManyOwners | ForEach-Object { "$($_.Name): $($_.Count)" }) -join '; ') `
                -Recommendation 'Keep Owner to 2-3 named break-glass/platform accounts per subscription; move day-to-day work to Contributor or narrower roles.' -EvidenceFile 'Azure_RoleAssignments.csv'
        }
        if ($guestPriv.Count -gt 0) {
            Add-Finding -Id 'AZ-IAM-002' -Severity High -Title 'Guest accounts hold privileged roles' -Count $guestPriv.Count `
                -Detail (($guestPriv | Select-Object -First 10 | ForEach-Object { "$($_.Principal) - $($_.Role) @ $($_.ScopeLevel)" }) -join '; ') `
                -Recommendation 'Remove standing Owner/User Access Administrator from guests; use Lighthouse or PIM-eligible access for partners.' -EvidenceFile 'Azure_RoleAssignments.csv'
        }
        if ($orphaned.Count -gt 0) {
            Add-Finding -Id 'AZ-IAM-003' -Severity Low -Title 'Role assignments for deleted identities' -Count $orphaned.Count `
                -Recommendation 'Remove assignments whose principal no longer exists (shown as "Identity not found" in the portal).' -EvidenceFile 'Azure_RoleAssignments.csv'
        }
        if ($raRows.Count -gt 0 -and ($userDirect.Count / $raRows.Count) -gt 0.5) {
            Add-Finding -Id 'AZ-IAM-004' -Severity Low -Title 'Most role assignments are made directly to users, not groups' -Count $userDirect.Count `
                -Recommendation 'Assign roles to Entra groups so access follows joiner/mover/leaver processes.' -EvidenceFile 'Azure_RoleAssignments.csv'
        }
        if ($wildcard.Count -gt 0) {
            Add-Finding -Id 'AZ-IAM-005' -Severity Medium -Title 'Custom roles with wildcard (*) actions' -Count $wildcard.Count `
                -Detail (($wildcard | ForEach-Object { $_.Role }) -join '; ') `
                -Recommendation 'Replace wildcard custom roles with built-in roles or tightly scoped action lists.' -EvidenceFile 'Azure_CustomRoles.csv'
        }
        if ($privHigh.Count -gt 0 -and $pimRows.Count -eq 0 -and $pimFailed -eq 0) {
            Add-Finding -Id 'AZ-IAM-006' -Severity Medium -Title 'All privileged Azure access is permanent (no PIM eligible assignments)' -Count $privHigh.Count `
                -Recommendation 'Make Owner and User Access Administrator eligible through PIM with approval and MFA on activation.' -EvidenceFile 'Azure_PIM_Eligible.csv'
        }
        if ($spOwners.Count -gt 0) {
            Add-Finding -Id 'AZ-IAM-007' -Severity Medium -Title 'Service principals / managed identities with privileged roles at subscription scope or above' -Count $spOwners.Count `
                -Recommendation 'Confirm each automation identity needs Owner/UAA; scope down to the resource groups it manages.' -EvidenceFile 'Azure_RoleAssignments.csv'
        }

        $privTable = $priv | Sort-Object { switch ($_.ScopeLevel) { 'Root' { 0 } 'Management group' { 1 } 'Subscription' { 2 } 'Resource group' { 3 } default { 4 } } } |
            Select-Object Role, @{N = 'Principal'; E = { if ($_.Principal) { $_.Principal } else { $_.PrincipalId } } }, PrincipalType, Guest, ScopeLevel, Subscription
        $html = "<div class='stat-row'>" +
            (New-StatCard 'Role assignments' $raRows.Count) +
            (New-StatCard 'Privileged assignments' $priv.Count $(if ($privHigh.Count -gt 10) { 'warn' } else { 'info' })) +
            (New-StatCard 'Guests with privileged roles' $(if ($res.Ok) { $guestPriv.Count } else { 'N/A' }) $(if ($guestPriv.Count) { 'bad' } elseif ($res.Ok) { 'good' } else { 'neutral' })) +
            (New-StatCard 'Orphaned assignments' $(if ($res.Ok) { $orphaned.Count } else { 'N/A' }) $(if ($orphaned.Count) { 'warn' } else { 'neutral' })) +
            (New-StatCard 'PIM eligible' $pimRows.Count $(if ($pimRows.Count) { 'good' } else { 'warn' })) +
            (New-StatCard 'Custom roles' $custom.Count 'neutral') +
            "</div><h3>Privileged role assignments (Owner, User Access Administrator, RBAC Administrator)</h3>" +
            (ConvertTo-HtmlTable $privTable @('Role', 'Principal', 'PrincipalType', 'Guest', 'ScopeLevel', 'Subscription') -MaxRows 30 -CsvName 'Azure_RoleAssignments.csv') +
            "<p class='note'>Group assignments are not expanded; members of privileged groups also hold the role.</p>"
        Add-HtmlSection $title $html
    }
    catch { Add-ErrorSection $title $_ }
}

# ═════════════════════════════════════════════════════════════════════════════
# 4. Security
# ═════════════════════════════════════════════════════════════════════════════
function Get-SecurityPosture {
    param([int]$Total)
    $title = 'Security'
    Write-SectionHeader $title $Total
    try {
        # Defender for Cloud plans
        $pr = Invoke-ArgQuery -Name 'Defender for Cloud plans' -Query @'
securityresources
| where type =~ 'microsoft.security/pricings'
| project id, subscriptionId, plan = name, tier = tostring(properties.pricingTier), subPlan = tostring(properties.subPlan), deprecated = tostring(properties.deprecated)
'@
        $planRows = $pr | Where-Object { $_.deprecated -ne 'true' } | ForEach-Object {
            [pscustomobject]@{
                Subscription   = Get-SubName $_.subscriptionId
                SubscriptionId = $_.subscriptionId
                Plan           = $(if ($script:DefenderPlanNames.Contains($_.plan)) { $script:DefenderPlanNames[$_.plan] } else { $_.plan })
                PlanId         = $_.plan
                Tier           = $_.tier
                SubPlan        = $_.subPlan
            }
        }
        $planRows = @($planRows)
        Export-AssessmentCsv $planRows 'Azure_Defender_Plans.csv' @('Subscription', 'SubscriptionId', 'Plan', 'PlanId', 'Tier', 'SubPlan')

        $matrix = foreach ($g in ($planRows | Group-Object SubscriptionId)) {
            $o = [ordered]@{ Subscription = Get-SubName $g.Name }
            foreach ($k in 'VirtualMachines', 'CloudPosture', 'StorageAccounts', 'SqlServers', 'KeyVaults', 'AppServices', 'Containers', 'Arm') {
                $p = $g.Group | Where-Object { $_.PlanId -eq $k } | Select-Object -First 1
                $o[$script:DefenderPlanNames[$k]] = if (!$p) { '-' } elseif ($p.Tier -eq 'Standard') { if ($p.SubPlan) { "On ($($p.SubPlan))" } else { 'On' } } else { 'Off' }
            }
            [pscustomobject]$o
        }
        $matrix = @($matrix)

        # Secure score
        $ss = Invoke-ArgQuery -Name 'Secure score' -Query @'
securityresources
| where type =~ 'microsoft.security/securescores'
| project id, subscriptionId, current = todouble(properties.score.current), max = todouble(properties.score.max), pct = todouble(properties.score.percentage)
'@
        $ssRows = $ss | ForEach-Object {
            [pscustomobject]@{ Subscription = Get-SubName $_.subscriptionId; SubscriptionId = $_.subscriptionId
                Current = [math]::Round([double]$_.current, 2); Max = [double]$_.max; Percent = [math]::Round([double]$_.pct * 100, 1) }
        }
        $ssRows = @($ssRows)
        Export-AssessmentCsv $ssRows 'Azure_SecureScore.csv' @('Subscription', 'SubscriptionId', 'Current', 'Max', 'Percent')
        $sumMax = ($ssRows | Measure-Object Max -Sum).Sum
        $scorePct = if ($sumMax -gt 0) { [math]::Round(($ssRows | Measure-Object Current -Sum).Sum / $sumMax * 100, 1) } else { $null }

        # Unhealthy recommendations
        $as = Invoke-ArgQuery -Name 'Defender for Cloud recommendations' -Query @'
securityresources
| where type =~ 'microsoft.security/assessments'
| extend status = tostring(properties.status.code), severity = tostring(properties.metadata.severity), displayName = tostring(properties.displayName)
| where status =~ 'Unhealthy'
| summarize UnhealthyResources = count() by displayName, severity
| order by UnhealthyResources desc
'@
        $recRows = @($as | Select-Object @{N = 'Recommendation'; E = { $_.displayName } }, @{N = 'Severity'; E = { if ($_.severity) { $_.severity } else { 'Unknown' } } }, UnhealthyResources |
            Sort-Object @{E = { Get-SeverityRank $_.Severity } }, @{E = 'UnhealthyResources'; Descending = $true })
        Export-AssessmentCsv $recRows 'Azure_Defender_Recommendations.csv' @('Recommendation', 'Severity', 'UnhealthyResources')
        $recHigh = @($recRows | Where-Object { $_.Severity -eq 'High' })

        # NSG rules open to the internet
        $nsg = Invoke-ArgQuery -Name 'NSG inbound rules' -Query @'
resources
| where type =~ 'microsoft.network/networksecuritygroups'
| extend attachedNics = array_length(properties.networkInterfaces), attachedSubnets = array_length(properties.subnets)
| mv-expand rule = properties.securityRules
| extend direction = tostring(rule.properties.direction), access = tostring(rule.properties.access)
| where direction =~ 'Inbound' and access =~ 'Allow'
| project id, nsgName = name, resourceGroup, subscriptionId, ruleName = tostring(rule.name), priority = toint(rule.properties.priority),
          protocol = tostring(rule.properties.protocol), src = tostring(rule.properties.sourceAddressPrefix), srcs = tostring(rule.properties.sourceAddressPrefixes),
          port = tostring(rule.properties.destinationPortRange), ports = tostring(rule.properties.destinationPortRanges),
          attachedNics, attachedSubnets
'@
        $openSources = @('*', '0.0.0.0/0', 'internet', 'any', '::/0')
        $exposed = foreach ($r in $nsg) {
            $sources = @($r.src) + @(ConvertFrom-JsonArray $r.srcs) | Where-Object { $_ }
            if (-not ($sources | Where-Object { $openSources -contains "$_".ToLower() })) { continue }
            $ranges = @(@($r.port) + @(ConvertFrom-JsonArray $r.ports) | Where-Object { $_ })
            $allPorts = [bool]($ranges | Where-Object { $_ -eq '*' -or $_ -eq '0-65535' })
            $hits = @(foreach ($sp in $script:SensitivePorts.GetEnumerator()) { if ($ranges | Where-Object { Test-PortInRange $_ ([int]$sp.Key) }) { "$($sp.Value) ($($sp.Key))" } })
            $attached = ([int]$r.attachedNics + [int]$r.attachedSubnets) -gt 0
            [pscustomobject]@{
                NSG = $r.nsgName; Rule = $r.ruleName; Priority = $r.priority; Protocol = $r.protocol
                Ports = ($ranges -join ','); SensitivePorts = $(if ($allPorts) { 'ALL PORTS' } else { $hits -join '; ' })
                Sensitive = ($allPorts -or $hits.Count -gt 0); Attached = $attached
                Subscription = Get-SubName $r.subscriptionId; ResourceGroup = $r.resourceGroup
            }
        }
        $exposed = @($exposed)
        Export-AssessmentCsv $exposed 'Azure_NSG_InternetExposed.csv' @('NSG', 'Rule', 'Priority', 'Protocol', 'Ports', 'SensitivePorts', 'Sensitive', 'Attached', 'Subscription', 'ResourceGroup')
        $exposedSensitive = @($exposed | Where-Object { $_.Sensitive -and $_.Attached })

        # Public IPs directly on NICs (VMs reachable from the internet)
        $pip = Invoke-ArgQuery -Name 'Public IPs on NICs' -Query @'
resources
| where type =~ 'microsoft.network/publicipaddresses'
| extend ipcfg = tolower(tostring(properties.ipConfiguration.id))
| where ipcfg contains '/networkinterfaces/'
| extend nicId = extract('(.*/networkinterfaces/[^/]+)/', 1, ipcfg)
| project id, name, ip = tostring(properties.ipAddress), sku = tostring(sku.name), nicId, resourceGroup, subscriptionId
'@
        $nicToVm = @{}
        foreach ($v in $script:Vms) { foreach ($n in $v.NicIds) { $nicToVm[$n] = $v } }
        $pipRows = foreach ($p in $pip) {
            $vm = $nicToVm[$p.nicId]
            if ($vm) { $vm.PublicIp = $p.ip }
            [pscustomobject]@{ PublicIp = $p.name; Address = $p.ip; Sku = $p.sku; VM = $(if ($vm) { $vm.Name } else { '' })
                NIC = ($p.nicId -split '/')[-1]; Subscription = Get-SubName $p.subscriptionId; ResourceGroup = $p.resourceGroup }
        }
        $pipRows = @($pipRows)
        Export-AssessmentCsv $pipRows 'Azure_PublicIPs_OnNICs.csv' @('PublicIp', 'Address', 'Sku', 'VM', 'NIC', 'Subscription', 'ResourceGroup')

        # Storage accounts
        $st = Invoke-ArgQuery -Name 'Storage accounts' -Query @'
resources
| where type =~ 'microsoft.storage/storageaccounts'
| project id, name, resourceGroup, subscriptionId, location, kind, sku = tostring(sku.name),
          allowBlobPublicAccess = tostring(properties.allowBlobPublicAccess), minTls = tostring(properties.minimumTlsVersion),
          httpsOnly = tostring(properties.supportsHttpsTrafficOnly), sharedKey = tostring(properties.allowSharedKeyAccess),
          publicNetworkAccess = tostring(properties.publicNetworkAccess), defaultAction = tostring(properties.networkAcls.defaultAction),
          privateEndpoints = array_length(properties.privateEndpointConnections)
'@
        $stRows = foreach ($s in $st) {
            [pscustomobject]@{
                StorageAccount = $s.name; Kind = $s.kind; Sku = $s.sku; Location = $s.location
                BlobPublicAccess = $(if ($s.allowBlobPublicAccess -eq 'false') { 'Disabled' } else { 'Permitted' })
                MinTls = $(if ($s.minTls) { $s.minTls } else { 'Default' }); HttpsOnly = $s.httpsOnly
                SharedKeyAccess = $(if ($s.sharedKey -eq 'false') { 'Disabled' } else { 'Enabled' })
                PublicNetworkAccess = $(if ($s.publicNetworkAccess) { $s.publicNetworkAccess } else { 'Enabled' })
                FirewallDefault = $(if ($s.defaultAction) { $s.defaultAction } else { 'Allow' })
                PrivateEndpoints = [int]$s.privateEndpoints
                Subscription = Get-SubName $s.subscriptionId; ResourceGroup = $s.resourceGroup
            }
        }
        $script:StorageRows = @($stRows)
        Export-AssessmentCsv $script:StorageRows 'Azure_StorageAccounts.csv' @('StorageAccount', 'Kind', 'Sku', 'Location', 'BlobPublicAccess', 'MinTls', 'HttpsOnly', 'SharedKeyAccess', 'PublicNetworkAccess', 'FirewallDefault', 'PrivateEndpoints', 'Subscription', 'ResourceGroup')
        $stPublicBlob = @($script:StorageRows | Where-Object { $_.BlobPublicAccess -eq 'Permitted' })
        $stOldTls     = @($script:StorageRows | Where-Object { $_.MinTls -in 'TLS1_0', 'TLS1_1' })
        $stOpenNet    = @($script:StorageRows | Where-Object { $_.PublicNetworkAccess -ne 'Disabled' -and $_.FirewallDefault -eq 'Allow' -and $_.PrivateEndpoints -eq 0 })

        # Key Vaults
        $kv = Invoke-ArgQuery -Name 'Key Vaults' -Query @'
resources
| where type =~ 'microsoft.keyvault/vaults'
| project id, name, resourceGroup, subscriptionId, location,
          purgeProtection = tostring(properties.enablePurgeProtection), softDelete = tostring(properties.enableSoftDelete),
          rbac = tostring(properties.enableRbacAuthorization), publicNetworkAccess = tostring(properties.publicNetworkAccess),
          defaultAction = tostring(properties.networkAcls.defaultAction), privateEndpoints = array_length(properties.privateEndpointConnections)
'@
        $kvRows = foreach ($k in $kv) {
            [pscustomobject]@{
                KeyVault = $k.name; Location = $k.location
                PurgeProtection = ($k.purgeProtection -eq 'true'); SoftDelete = ($k.softDelete -ne 'false')
                PermissionModel = $(if ($k.rbac -eq 'true') { 'Azure RBAC' } else { 'Access policies' })
                PublicNetworkAccess = $(if ($k.publicNetworkAccess) { $k.publicNetworkAccess } else { 'Enabled' })
                FirewallDefault = $(if ($k.defaultAction) { $k.defaultAction } else { 'Allow' })
                PrivateEndpoints = [int]$k.privateEndpoints
                Subscription = Get-SubName $k.subscriptionId; ResourceGroup = $k.resourceGroup
            }
        }
        $kvRows = @($kvRows)
        Export-AssessmentCsv $kvRows 'Azure_KeyVaults.csv' @('KeyVault', 'Location', 'PurgeProtection', 'SoftDelete', 'PermissionModel', 'PublicNetworkAccess', 'FirewallDefault', 'PrivateEndpoints', 'Subscription', 'ResourceGroup')
        $kvNoPurge = @($kvRows | Where-Object { -not $_.PurgeProtection })

        # PaaS network exposure
        $paas = Invoke-ArgQuery -Name 'PaaS network access' -Query @'
resources
| where type in~ ('microsoft.sql/servers','microsoft.dbforpostgresql/flexibleservers','microsoft.dbformysql/flexibleservers',
                  'microsoft.documentdb/databaseaccounts','microsoft.cache/redis','microsoft.containerregistry/registries',
                  'microsoft.cognitiveservices/accounts','microsoft.servicebus/namespaces','microsoft.eventhub/namespaces',
                  'microsoft.search/searchservices','microsoft.web/sites','microsoft.machinelearningservices/workspaces')
| extend pna = coalesce(tostring(properties.publicNetworkAccess), tostring(properties.network.publicNetworkAccess)),
         pe = array_length(properties.privateEndpointConnections)
| project id, name, type, resourceGroup, subscriptionId, location, pna, pe
'@
        $paasRows = foreach ($p in $paas) {
            [pscustomobject]@{
                Resource = $p.name; Type = $p.type
                PublicNetworkAccess = $(if ($p.pna) { $p.pna } else { 'Enabled (default)' })
                PrivateEndpoints = [int]$p.pe
                Subscription = Get-SubName $p.subscriptionId; ResourceGroup = $p.resourceGroup; Location = $p.location
            }
        }
        $paasRows = @($paasRows)
        Export-AssessmentCsv $paasRows 'Azure_PaaS_NetworkAccess.csv' @('Resource', 'Type', 'PublicNetworkAccess', 'PrivateEndpoints', 'Subscription', 'ResourceGroup', 'Location')
        $paasPublic = @($paasRows | Where-Object { $_.PublicNetworkAccess -notmatch '^Disabled' -and $_.PrivateEndpoints -eq 0 -and $_.Type -ne 'microsoft.web/sites' })

        $script:Data['SecureScorePct']               = $scorePct
        $script:Data['DefenderHighRecommendations']  = $recHigh.Count
        $script:Data['InternetExposedSensitiveRules'] = $exposedSensitive.Count
        $script:Data['PublicIpsOnNics']              = $pipRows.Count
        $script:Data['StorageBlobPublicAccessPermitted'] = $stPublicBlob.Count
        $script:Data['KeyVaultsWithoutPurgeProtection']  = $kvNoPurge.Count
        Write-Host "  Secure score: $(if ($null -ne $scorePct) { "$scorePct%" } else { 'N/A' }) | Exposed mgmt rules: $($exposedSensitive.Count) | Public IPs on NICs: $($pipRows.Count) | KV without purge protection: $($kvNoPurge.Count)" -ForegroundColor Green

        if ($planRows.Count -eq 0) {
            Add-NotCollected -Item 'Defender for Cloud plans and secure score' -Reason 'No Defender for Cloud data was readable.' -HowToFix 'Assign Security Reader on the subscriptions (or root management group) and re-run.'
        }
        # Findings
        if ($null -ne $scorePct -and $scorePct -lt 70) {
            Add-Finding -Id 'AZ-SEC-001' -Severity $(if ($scorePct -lt 50) { 'High' } else { 'Medium' }) -Title "Defender for Cloud secure score is $scorePct%" -Count $recHigh.Count `
                -Recommendation 'Work the high-severity recommendations first; most are configuration changes with no licence cost.' -EvidenceFile 'Azure_Defender_Recommendations.csv'
        }
        $subsWithVms = @($script:Vms | Group-Object SubscriptionId | ForEach-Object { $_.Name })
        $serversOff = @($planRows | Where-Object { $_.PlanId -eq 'VirtualMachines' -and $_.Tier -ne 'Standard' -and $subsWithVms -contains $_.SubscriptionId })
        if ($serversOff.Count -gt 0) {
            Add-Finding -Id 'AZ-SEC-002' -Severity High -Title 'Subscriptions with VMs but Defender for Servers off' -Count $serversOff.Count `
                -Detail (($serversOff | ForEach-Object { $_.Subscription }) -join '; ') `
                -Recommendation 'Enable Defender for Servers (Plan 1 for MDE + inventory, Plan 2 for vulnerability assessment, FIM and JIT).' -EvidenceFile 'Azure_Defender_Plans.csv'
        }
        $cspmOn = @($planRows | Where-Object { $_.PlanId -eq 'CloudPosture' -and $_.Tier -eq 'Standard' })
        if ($planRows.Count -gt 0 -and $cspmOn.Count -eq 0) {
            Add-Finding -Id 'AZ-SEC-003' -Severity Low -Title 'Defender CSPM (paid posture management) not enabled' -Count 0 `
                -Recommendation 'Consider Defender CSPM for attack path analysis and agentless scanning on production subscriptions.' -EvidenceFile 'Azure_Defender_Plans.csv'
        }
        if ($exposedSensitive.Count -gt 0) {
            Add-Finding -Id 'AZ-SEC-004' -Severity High -Title 'NSG rules allow management/database ports from the internet' -Count $exposedSensitive.Count `
                -Detail (($exposedSensitive | Select-Object -First 10 | ForEach-Object { "$($_.NSG)/$($_.Rule): $($_.SensitivePorts)" }) -join '; ') `
                -Recommendation 'Remove internet-sourced allow rules for RDP/SSH/SMB/database ports; use Azure Bastion, JIT access or VPN.' -EvidenceFile 'Azure_NSG_InternetExposed.csv'
        }
        if ($pipRows.Count -gt 0) {
            Add-Finding -Id 'AZ-SEC-005' -Severity Medium -Title 'Public IPs attached directly to VM network interfaces' -Count $pipRows.Count `
                -Recommendation 'Publish through a load balancer, Application Gateway or Azure Firewall instead of per-VM public IPs; deny with Policy.' -EvidenceFile 'Azure_PublicIPs_OnNICs.csv'
        }
        if ($stPublicBlob.Count -gt 0) {
            Add-Finding -Id 'AZ-SEC-006' -Severity Medium -Title 'Storage accounts that permit anonymous blob access' -Count $stPublicBlob.Count `
                -Recommendation 'Set "Allow Blob anonymous access" to Disabled unless a container is deliberately public.' -EvidenceFile 'Azure_StorageAccounts.csv'
        }
        if ($stOldTls.Count -gt 0) {
            Add-Finding -Id 'AZ-SEC-007' -Severity Medium -Title 'Storage accounts allowing TLS 1.0/1.1' -Count $stOldTls.Count `
                -Recommendation 'Set minimum TLS version to 1.2.' -EvidenceFile 'Azure_StorageAccounts.csv'
        }
        if ($stOpenNet.Count -gt 0) {
            Add-Finding -Id 'AZ-SEC-008' -Severity Low -Title 'Storage accounts reachable from all networks with no private endpoint' -Count $stOpenNet.Count `
                -Recommendation 'Restrict to selected networks or private endpoints where the data is not meant to be public.' -EvidenceFile 'Azure_StorageAccounts.csv'
        }
        if ($kvNoPurge.Count -gt 0) {
            Add-Finding -Id 'AZ-SEC-009' -Severity Medium -Title 'Key Vaults without purge protection' -Count $kvNoPurge.Count `
                -Recommendation 'Enable purge protection so deleted keys and secrets cannot be permanently destroyed during the retention period.' -EvidenceFile 'Azure_KeyVaults.csv'
        }
        if ($paasPublic.Count -gt 0) {
            Add-Finding -Id 'AZ-SEC-010' -Severity Low -Title 'PaaS data services with public network access and no private endpoint' -Count $paasPublic.Count `
                -Recommendation 'Review whether these need public endpoints; use private endpoints or service firewall rules.' -EvidenceFile 'Azure_PaaS_NetworkAccess.csv' -Confidence 'reported'
        }

        $sensTable = $exposed | Where-Object { $_.Sensitive } | Sort-Object { -not $_.Attached }, NSG
        $html = "<div class='stat-row'>" +
            (New-StatCard 'Secure score' $(if ($null -ne $scorePct) { "$scorePct%" } else { 'N/A' }) $(if ($null -eq $scorePct) { 'neutral' } elseif ($scorePct -ge 80) { 'good' } elseif ($scorePct -ge 50) { 'warn' } else { 'bad' })) +
            (New-StatCard 'High-severity recommendations' $recHigh.Count $(if ($recHigh.Count) { 'bad' } else { 'good' })) +
            (New-StatCard 'Internet-exposed mgmt/DB rules' $exposedSensitive.Count $(if ($exposedSensitive.Count) { 'bad' } else { 'good' })) +
            (New-StatCard 'Public IPs on VM NICs' $pipRows.Count $(if ($pipRows.Count) { 'warn' } else { 'good' })) +
            (New-StatCard 'Storage: anonymous blob permitted' $stPublicBlob.Count $(if ($stPublicBlob.Count) { 'warn' } else { 'good' })) +
            (New-StatCard 'Key Vaults w/o purge protection' $kvNoPurge.Count $(if ($kvNoPurge.Count) { 'warn' } else { 'good' })) +
            "</div><h3>Defender for Cloud plans by subscription</h3>" +
            (ConvertTo-HtmlTable $matrix -MaxRows 30 -CsvName 'Azure_Defender_Plans.csv') +
            "<h3>Top Defender for Cloud recommendations</h3>" +
            (ConvertTo-HtmlTable $recRows @('Severity', 'Recommendation', 'UnhealthyResources') -MaxRows 15 -CsvName 'Azure_Defender_Recommendations.csv') +
            "<h3>Inbound rules open to the internet on sensitive ports</h3>" +
            (ConvertTo-HtmlTable $sensTable @('NSG', 'Rule', 'SensitivePorts', 'Attached', 'Subscription') -MaxRows 20 -CsvName 'Azure_NSG_InternetExposed.csv')
        Add-HtmlSection $title $html
    }
    catch { Add-ErrorSection $title $_ }
}

# ═════════════════════════════════════════════════════════════════════════════
# 5. Network Topology (CAF Ready: network topology and connectivity design area)
# ═════════════════════════════════════════════════════════════════════════════
function ConvertTo-Ipv4Range([string]$Cidr) {
    if ($Cidr -notmatch '^(\d{1,3}(?:\.\d{1,3}){3})/(\d{1,2})$') { return $null }
    try {
        $bytes = ([System.Net.IPAddress]::Parse($Matches[1])).GetAddressBytes()
        [array]::Reverse($bytes)
        $ip = [uint64][BitConverter]::ToUInt32($bytes, 0)
        $bits = [int]$Matches[2]
        $size = [uint64][math]::Pow(2, 32 - $bits)
        $start = $ip - ($ip % $size)
        return [pscustomobject]@{ Start = $start; End = $start + $size - 1 }
    }
    catch { return $null }
}

function Get-NetworkTopology {
    param([int]$Total)
    $title = 'Network Topology'
    Write-SectionHeader $title $Total
    try {
        $vnets = Invoke-ArgQuery -Name 'Virtual networks' -Query @'
resources
| where type =~ 'microsoft.network/virtualnetworks'
| project id, name, resourceGroup, subscriptionId, location, addressSpace = tostring(properties.addressSpace.addressPrefixes),
          ddos = tostring(properties.enableDdosProtection), peerings = array_length(properties.virtualNetworkPeerings),
          subnetCount = array_length(properties.subnets)
'@
        $peer = Invoke-ArgQuery -Name 'VNet peerings' -Query @'
resources
| where type =~ 'microsoft.network/virtualnetworks'
| mv-expand p = properties.virtualNetworkPeerings
| where isnotempty(p)
| project id, vnet = name, subscriptionId, peering = tostring(p.name), state = tostring(p.properties.peeringState),
          remote = tolower(tostring(p.properties.remoteVirtualNetwork.id)), gatewayTransit = tostring(p.properties.allowGatewayTransit),
          useRemoteGateways = tostring(p.properties.useRemoteGateways)
'@
        $subnets = Invoke-ArgQuery -Name 'Subnets' -Query @'
resources
| where type =~ 'microsoft.network/virtualnetworks'
| mv-expand s = properties.subnets
| where isnotempty(s)
| project id, vnet = name, resourceGroup, subscriptionId, subnet = tostring(s.name),
          prefix = coalesce(tostring(s.properties.addressPrefix), tostring(s.properties.addressPrefixes[0])),
          nsg = tostring(s.properties.networkSecurityGroup.id), routeTable = tostring(s.properties.routeTable.id),
          natGateway = tostring(s.properties.natGateway.id), ipConfigs = array_length(s.properties.ipConfigurations),
          privateEndpoints = array_length(s.properties.privateEndpoints)
'@
        $comp = Invoke-ArgQuery -Name 'Network platform components' -Query @'
resources
| where type in~ ('microsoft.network/azurefirewalls','microsoft.network/virtualnetworkgateways','microsoft.network/bastionhosts',
                  'microsoft.network/virtualhubs','microsoft.network/virtualwans','microsoft.network/ddosprotectionplans',
                  'microsoft.network/privatednszones','microsoft.network/privateendpoints','microsoft.network/networkwatchers',
                  'microsoft.network/natgateways','microsoft.network/expressroutecircuits','microsoft.network/firewallpolicies',
                  'microsoft.network/publicipaddresses')
| project id, name, type, resourceGroup, subscriptionId, location,
          subnetId = tolower(tostring(properties.ipConfigurations[0].properties.subnet.id)),
          gatewayType = tostring(properties.gatewayType), sku = coalesce(tostring(properties.sku.tier), tostring(sku.name))
'@
        $nva = Invoke-ArgQuery -Name 'Network virtual appliances (IP forwarding NICs)' -Query @'
resources
| where type =~ 'microsoft.network/networkinterfaces'
| where tostring(properties.enableIPForwarding) =~ 'true'
| project id, name, vm = tostring(properties.virtualMachine.id), subnetId = tolower(tostring(properties.ipConfigurations[0].properties.subnet.id)), subscriptionId, resourceGroup
'@

        $vnetOf = { param($subnetId) if ($subnetId -match '^(.*/virtualnetworks/[^/]+)/subnets/') { $Matches[1] } else { '' } }
        $hubRoles = @{}
        foreach ($c in @($comp | Where-Object { $_.type -in 'microsoft.network/azurefirewalls', 'microsoft.network/virtualnetworkgateways', 'microsoft.network/bastionhosts' })) {
            $v = & $vnetOf $c.subnetId
            if ($v) {
                if (!$hubRoles.ContainsKey($v)) { $hubRoles[$v] = [System.Collections.Generic.List[string]]::new() }
                $label = switch ($c.type) { 'microsoft.network/azurefirewalls' { 'Azure Firewall' } 'microsoft.network/virtualnetworkgateways' { "$($c.gatewayType) gateway" } default { 'Bastion' } }
                $hubRoles[$v].Add($label)
            }
        }
        foreach ($n in $nva) {
            $v = & $vnetOf $n.subnetId
            if ($v) { if (!$hubRoles.ContainsKey($v)) { $hubRoles[$v] = [System.Collections.Generic.List[string]]::new() }; $hubRoles[$v].Add('NVA') }
        }

        $vnetRows = foreach ($v in $vnets) {
            $id = "$($v.id)".ToLower()
            $roles = if ($hubRoles.ContainsKey($id)) { @($hubRoles[$id] | Sort-Object -Unique) } else { @() }
            [pscustomobject]@{
                VNet = $v.name; AddressSpace = ((ConvertFrom-JsonArray $v.addressSpace) -join ', '); Location = $v.location
                Peerings = [int]$v.peerings; Subnets = [int]$v.subnetCount; DdosProtection = ($v.ddos -eq 'true')
                HubComponents = ($roles -join ', '); Role = $(if ($roles | Where-Object { $_ -match 'Firewall|gateway|NVA' }) { 'Hub' } elseif ([int]$v.peerings -gt 0) { 'Spoke / peered' } else { 'Isolated' })
                Subscription = Get-SubName $v.subscriptionId; ResourceGroup = $v.resourceGroup; ResourceId = $id
            }
        }
        $vnetRows = @($vnetRows)
        Export-AssessmentCsv $vnetRows 'Azure_VNets.csv' @('VNet', 'AddressSpace', 'Location', 'Role', 'HubComponents', 'Peerings', 'Subnets', 'DdosProtection', 'Subscription', 'ResourceGroup', 'ResourceId')

        $peerRows = @($peer | ForEach-Object {
                [pscustomobject]@{ VNet = $_.vnet; Peering = $_.peering; State = $_.state; RemoteVNet = ($_.remote -split '/')[-1]
                    AllowGatewayTransit = $_.gatewayTransit; UseRemoteGateways = $_.useRemoteGateways; Subscription = Get-SubName $_.subscriptionId }
            })
        Export-AssessmentCsv $peerRows 'Azure_VNetPeerings.csv' @('VNet', 'Peering', 'State', 'RemoteVNet', 'AllowGatewayTransit', 'UseRemoteGateways', 'Subscription')

        $platformSubnets = @('GatewaySubnet', 'AzureFirewallSubnet', 'AzureFirewallManagementSubnet', 'AzureBastionSubnet', 'RouteServerSubnet')
        $subnetRows = @($subnets | ForEach-Object {
                [pscustomobject]@{ VNet = $_.vnet; Subnet = $_.subnet; Prefix = $_.prefix; NSG = (($_.nsg -split '/')[-1]); RouteTable = (($_.routeTable -split '/')[-1])
                    NatGateway = (($_.natGateway -split '/')[-1]); InUse = (([int]$_.ipConfigs + [int]$_.privateEndpoints) -gt 0)
                    PlatformSubnet = ($platformSubnets -contains $_.subnet); Subscription = Get-SubName $_.subscriptionId; ResourceGroup = $_.resourceGroup }
            })
        Export-AssessmentCsv $subnetRows 'Azure_Subnets.csv' @('VNet', 'Subnet', 'Prefix', 'NSG', 'RouteTable', 'NatGateway', 'InUse', 'PlatformSubnet', 'Subscription', 'ResourceGroup')

        # Overlapping address spaces (between VNets)
        $ranges = foreach ($v in $vnetRows) { foreach ($cidr in ($v.AddressSpace -split ',\s*' | Where-Object { $_ })) { $r = ConvertTo-Ipv4Range $cidr; if ($r) { [pscustomobject]@{ VNet = $v.VNet; Id = $v.ResourceId; Cidr = $cidr; Start = $r.Start; End = $r.End } } } }
        $ranges = @($ranges)
        $overlaps = [System.Collections.Generic.List[object]]::new()
        for ($i = 0; $i -lt $ranges.Count; $i++) {
            for ($j = $i + 1; $j -lt $ranges.Count; $j++) {
                $a = $ranges[$i]; $b = $ranges[$j]
                if ($a.Id -ne $b.Id -and $a.Start -le $b.End -and $b.Start -le $a.End) {
                    $overlaps.Add([pscustomobject]@{ VNetA = $a.VNet; RangeA = $a.Cidr; VNetB = $b.VNet; RangeB = $b.Cidr })
                }
            }
        }
        Export-AssessmentCsv $overlaps 'Azure_VNet_Overlaps.csv' @('VNetA', 'RangeA', 'VNetB', 'RangeB')

        $count = { param($t) @($comp | Where-Object { $_.type -eq $t }).Count }
        $vhubs = & $count 'microsoft.network/virtualhubs'
        $firewalls = & $count 'microsoft.network/azurefirewalls'
        $ddosPlans = & $count 'microsoft.network/ddosprotectionplans'
        $pdns = & $count 'microsoft.network/privatednszones'
        $pe = & $count 'microsoft.network/privateendpoints'
        $pips = & $count 'microsoft.network/publicipaddresses'
        $er = & $count 'microsoft.network/expressroutecircuits'
        $gateways = @($comp | Where-Object { $_.type -eq 'microsoft.network/virtualnetworkgateways' })
        $hubs = @($vnetRows | Where-Object { $_.Role -eq 'Hub' })
        $topology = if ($vnetRows.Count -eq 0) { 'No virtual networks' }
                    elseif ($vhubs -gt 0) { 'Virtual WAN' }
                    elseif ($hubs.Count -gt 0 -and @($vnetRows | Where-Object { $_.Role -eq 'Spoke / peered' }).Count -gt 0) { 'Hub and spoke' }
                    elseif ($hubs.Count -gt 0) { 'Single hub VNet, no spokes' }
                    elseif (@($vnetRows | Where-Object { $_.Peerings -gt 0 }).Count -gt 0) { 'Peered VNets without a hub' }
                    else { 'Isolated VNets (no peering)' }
        $egress = ($firewalls -gt 0) -or ($nva.Count -gt 0) -or ($vhubs -gt 0)
        $notConnected = @($peerRows | Where-Object { $_.State -and $_.State -ne 'Connected' })
        $noNsg = @($subnetRows | Where-Object { -not $_.NSG -and -not $_.PlatformSubnet -and $_.InUse })
        $ddosOn = @($vnetRows | Where-Object { $_.DdosProtection }).Count

        # Network Watcher in every region that has VNets
        $nwKeys = @{}; foreach ($w in @($comp | Where-Object { $_.type -eq 'microsoft.network/networkwatchers' })) { $nwKeys["$($w.subscriptionId)|$($w.location)"] = $true }
        $nwMissing = @($vnets | ForEach-Object { "$($_.subscriptionId)|$($_.location)" } | Sort-Object -Unique | Where-Object { -not $nwKeys.ContainsKey($_) })

        $script:Data['NetworkTopology']           = $topology
        $script:Data['VirtualNetworks']           = $vnetRows.Count
        $script:Data['HubVNets']                  = $hubs.Count
        $script:Data['VirtualHubs']               = $vhubs
        $script:Data['AzureFirewalls']            = $firewalls
        $script:Data['NetworkVirtualAppliances']  = @($nva).Count
        $script:Data['VNetAddressOverlaps']       = $overlaps.Count
        $script:Data['PeeringsNotConnected']      = $notConnected.Count
        $script:Data['InUseSubnetsWithoutNsg']    = $noNsg.Count
        $script:Data['PrivateEndpoints']          = $pe
        $script:Data['PrivateDnsZones']           = $pdns
        Write-Host "  Topology: $topology | VNets: $($vnetRows.Count) | Hubs: $($hubs.Count) | Overlaps: $($overlaps.Count) | Subnets without NSG: $($noNsg.Count)" -ForegroundColor Green

        if ($overlaps.Count -gt 0) {
            Add-Finding -Id 'AZ-NET-001' -Severity Medium -Title 'Virtual networks with overlapping address spaces' -Count $overlaps.Count `
                -Detail (($overlaps | Select-Object -First 8 | ForEach-Object { "$($_.VNetA) $($_.RangeA) / $($_.VNetB) $($_.RangeB)" }) -join '; ') `
                -Recommendation 'Plan a non-overlapping IP address scheme across Azure and on-premises; overlapping VNets cannot be peered or routed to each other.' -EvidenceFile 'Azure_VNet_Overlaps.csv'
        }
        if ($notConnected.Count -gt 0) {
            Add-Finding -Id 'AZ-NET-002' -Severity Medium -Title 'VNet peerings not in Connected state' -Count $notConnected.Count `
                -Recommendation 'Re-sync or recreate the peering on both sides; a disconnected peering means no traffic flows between the VNets.' -EvidenceFile 'Azure_VNetPeerings.csv'
        }
        if ($noNsg.Count -gt 0) {
            Add-Finding -Id 'AZ-NET-003' -Severity Medium -Title 'In-use subnets without a network security group' -Count $noNsg.Count `
                -Recommendation 'Associate an NSG with every workload subnet (deny by default) and enforce it with Policy.' -EvidenceFile 'Azure_Subnets.csv'
        }
        if ($vnetRows.Count -gt 0 -and -not $egress) {
            Add-Finding -Id 'AZ-NET-004' -Severity Medium -Title 'No central traffic inspection (no Azure Firewall, NVA or Virtual WAN hub)' -Count 0 `
                -Recommendation 'Route internet egress and spoke-to-spoke traffic through a central firewall in a hub (Azure Firewall or a firewall NVA such as FortiGate).' -EvidenceFile 'Azure_VNets.csv'
        }
        if ($topology -in 'Isolated VNets (no peering)', 'Peered VNets without a hub' -and $vnetRows.Count -gt 1) {
            Add-Finding -Id 'AZ-NET-005' -Severity Low -Title "Network topology is '$topology'" -Count $vnetRows.Count `
                -Recommendation 'Consolidate onto a hub-and-spoke or Virtual WAN topology so connectivity, inspection and DNS are shared.' -EvidenceFile 'Azure_VNets.csv'
        }
        if ($pips -gt 0 -and $ddosOn -eq 0 -and $ddosPlans -eq 0) {
            Add-Finding -Id 'AZ-NET-006' -Severity Low -Title 'Public IPs in use but no DDoS Network Protection plan' -Count $pips `
                -Recommendation 'Assess whether internet-facing workloads need DDoS Network or IP Protection beyond the default infrastructure protection.' -EvidenceFile 'Azure_VNets.csv' -Confidence 'inferred'
        }
        if ($nwMissing.Count -gt 0) {
            Add-Finding -Id 'AZ-NET-007' -Severity Low -Title 'Regions with VNets but no Network Watcher' -Count $nwMissing.Count `
                -Recommendation 'Enable Network Watcher in every region with VNets (needed for flow logs, connection troubleshooting and topology).' -EvidenceFile 'Azure_VNets.csv'
        }

        $compTable = @(
            [pscustomobject]@{ Component = 'Virtual WAN hubs'; Count = $vhubs }
            [pscustomobject]@{ Component = 'Azure Firewalls'; Count = $firewalls }
            [pscustomobject]@{ Component = 'Network virtual appliances (IP forwarding)'; Count = @($nva).Count }
            [pscustomobject]@{ Component = 'VPN / ExpressRoute gateways'; Count = $gateways.Count }
            [pscustomobject]@{ Component = 'ExpressRoute circuits'; Count = $er }
            [pscustomobject]@{ Component = 'Private endpoints'; Count = $pe }
            [pscustomobject]@{ Component = 'Private DNS zones'; Count = $pdns }
            [pscustomobject]@{ Component = 'DDoS protection plans'; Count = $ddosPlans }
        )
        $html = "<div class='stat-row'>" +
            (New-StatCard 'Topology' $topology 'info') +
            (New-StatCard 'Virtual networks' $vnetRows.Count) +
            (New-StatCard 'Hub VNets' $hubs.Count 'neutral') +
            (New-StatCard 'Address overlaps' $overlaps.Count $(if ($overlaps.Count) { 'warn' } else { 'good' })) +
            (New-StatCard 'In-use subnets without NSG' $noNsg.Count $(if ($noNsg.Count) { 'warn' } else { 'good' })) +
            (New-StatCard 'Peerings not connected' $notConnected.Count $(if ($notConnected.Count) { 'warn' } else { 'good' })) +
            "</div><div class='two-col'><div><h3>Platform components</h3>" + (ConvertTo-HtmlTable $compTable @('Component', 'Count')) +
            "</div><div><h3>Virtual networks</h3>" +
            (ConvertTo-HtmlTable $vnetRows @('VNet', 'AddressSpace', 'Role', 'Peerings') -MaxRows 15 -CsvName 'Azure_VNets.csv') + "</div></div>" +
            $(if ($noNsg.Count) { "<h3>In-use subnets without an NSG</h3>" + (ConvertTo-HtmlTable $noNsg @('VNet', 'Subnet', 'Prefix', 'Subscription') -MaxRows 15 -CsvName 'Azure_Subnets.csv') } else { '' }) +
            "<p class='note'>Topology is inferred from where firewalls, gateways and IP-forwarding NICs sit and how VNets are peered. Open the ARI network diagram to confirm.</p>"
        Add-HtmlSection $title $html
    }
    catch { Add-ErrorSection $title $_ }
}

# ═════════════════════════════════════════════════════════════════════════════
# 5. Resiliency & Backup
# ═════════════════════════════════════════════════════════════════════════════
function Get-Resiliency {
    param([int]$Total)
    $title = 'Resiliency & Backup'
    Write-SectionHeader $title $Total
    try {
        $items = Invoke-ArgQuery -Name 'Azure Backup protected items' -Query @'
recoveryservicesresources
| where type =~ 'microsoft.recoveryservices/vaults/backupfabrics/protectioncontainers/protecteditems'
| extend src = tolower(tostring(properties.sourceResourceId)), workload = tostring(properties.workloadType),
         friendlyName = tostring(properties.friendlyName), policy = tostring(properties.policyName),
         protectionState = tostring(properties.protectionState), lastBackup = tostring(properties.lastBackupTime),
         lastStatus = tostring(properties.lastBackupStatus), health = tostring(properties.healthStatus),
         vault = tostring(split(id, '/')[8])
| project id, src, workload, friendlyName, vault, policy, protectionState, lastBackup, lastStatus, health, subscriptionId
'@
        $now = (Get-Date).ToUniversalTime()
        $itemRows = foreach ($i in $items) {
            $lb = $null
            if ($i.lastBackup) { try { $lb = ([datetime]::Parse($i.lastBackup, [cultureinfo]::InvariantCulture)).ToUniversalTime() } catch { } }
            $ageH = if ($lb) { [math]::Round(($now - $lb).TotalHours, 1) } else { $null }
            [pscustomobject]@{
                Item = $i.friendlyName; Workload = $i.workload; Vault = $i.vault; Policy = $i.policy
                ProtectionState = $i.protectionState; LastBackupUtc = $(if ($lb) { $lb.ToString('yyyy-MM-dd HH:mm') } else { '' })
                HoursSinceBackup = $ageH; LastStatus = $i.lastStatus; Health = $i.health
                Stale = ($i.protectionState -match 'Protected' -and ($null -eq $ageH -or $ageH -gt 48))
                Subscription = Get-SubName $i.subscriptionId; SourceResourceId = $i.src
            }
        }
        $itemRows = @($itemRows)
        Export-AssessmentCsv $itemRows 'Azure_BackupItems.csv' @('Item', 'Workload', 'Vault', 'Policy', 'ProtectionState', 'LastBackupUtc', 'HoursSinceBackup', 'LastStatus', 'Health', 'Stale', 'Subscription', 'SourceResourceId')
        $bySource = @{}
        foreach ($r in $itemRows) { if ($r.SourceResourceId) { $bySource[$r.SourceResourceId] = $r } }

        $bi = Invoke-ArgQuery -Name 'Backup vault (Data Protection) instances' -Query @'
recoveryservicesresources
| where type =~ 'microsoft.dataprotection/backupvaults/backupinstances'
| summarize Count = count() by datasourceType = tostring(properties.dataSourceInfo.datasourceType)
'@

        $vaults = Invoke-ArgQuery -Name 'Recovery Services vaults' -Query @'
recoveryservicesresources
| where type =~ 'microsoft.recoveryservices/vaults'
| project id, name, resourceGroup, subscriptionId, location, sku = tostring(sku.name),
          redundancy = tostring(properties.redundancySettings.standardTierStorageRedundancy),
          crossRegionRestore = tostring(properties.redundancySettings.crossRegionRestore),
          softDelete = tostring(properties.securitySettings.softDeleteSettings.softDeleteState),
          immutability = tostring(properties.securitySettings.immutabilitySettings.state),
          publicNetworkAccess = tostring(properties.publicNetworkAccess)
'@
        $vaultRows = foreach ($v in $vaults) {
            [pscustomobject]@{
                Vault = $v.name; Location = $v.location; StorageRedundancy = $v.redundancy; CrossRegionRestore = $v.crossRegionRestore
                SoftDelete = $(if ($v.softDelete) { $v.softDelete } else { 'Enabled (default)' })
                Immutability = $(if ($v.immutability) { $v.immutability } else { 'Disabled' })
                PublicNetworkAccess = $v.publicNetworkAccess
                Subscription = Get-SubName $v.subscriptionId; ResourceGroup = $v.resourceGroup
            }
        }
        $vaultRows = @($vaultRows)
        Export-AssessmentCsv $vaultRows 'Azure_RecoveryVaults.csv' @('Vault', 'Location', 'StorageRedundancy', 'CrossRegionRestore', 'SoftDelete', 'Immutability', 'PublicNetworkAccess', 'Subscription', 'ResourceGroup')

        $asr = Invoke-ArgQuery -Name 'Site Recovery replicated items' -Query @'
recoveryservicesresources
| where type =~ 'microsoft.recoveryservices/vaults/replicationfabrics/replicationprotectioncontainers/replicationprotecteditems'
| project id, name, friendlyName = tostring(properties.friendlyName), health = tostring(properties.replicationHealth),
          state = tostring(properties.protectionState), primary = tostring(properties.primaryFabricFriendlyName),
          recovery = tostring(properties.recoveryFabricFriendlyName), subscriptionId
'@
        $asrRows = @($asr | Select-Object @{N = 'Item'; E = { $_.friendlyName } }, @{N = 'ReplicationHealth'; E = { $_.health } }, @{N = 'ProtectionState'; E = { $_.state } },
            @{N = 'PrimaryFabric'; E = { $_.primary } }, @{N = 'RecoveryFabric'; E = { $_.recovery } }, @{N = 'Subscription'; E = { Get-SubName $_.subscriptionId } })
        Export-AssessmentCsv $asrRows 'Azure_ASR_ReplicatedItems.csv' @('Item', 'ReplicationHealth', 'ProtectionState', 'PrimaryFabric', 'RecoveryFabric', 'Subscription')
        $asrNames = @{}; foreach ($a in $asrRows) { if ($a.Item) { $asrNames["$($a.Item)".ToLower()] = $true } }

        # Enrich VM inventory
        foreach ($vm in $script:Vms) {
            $b = $bySource[$vm.ResourceId]
            $vm.AzureBackup = [bool]$b
            $vm.LastBackup  = $(if ($b) { $b.LastBackupUtc } else { '' })
            $vm.AsrReplicated = $asrNames.ContainsKey($vm.Name.ToLower())
        }
        $running      = @($script:Vms | Where-Object { $_.PowerState -eq 'running' })
        $noBackup     = @($running | Where-Object { -not $_.AzureBackup })
        $noProtection = @($noBackup | Where-Object { -not $_.AsrReplicated })
        $single       = @($script:Vms | Where-Object { $_.Placement -eq 'Single instance' })
        $stale        = @($itemRows | Where-Object { $_.Stale })
        $failed       = @($itemRows | Where-Object { $_.LastStatus -and $_.LastStatus -notmatch '^(Completed|Healthy|IRPending)$' })
        $weakVaults   = @($vaultRows | Where-Object { $_.SoftDelete -match '^Disabled' -or $_.Immutability -eq 'Disabled' })

        $sql = Invoke-ArgQuery -Name 'Azure SQL databases' -Query @'
resources
| where type =~ 'microsoft.sql/servers/databases' and name != 'master'
| project id, name, server = tostring(split(id, '/')[8]), resourceGroup, subscriptionId, location, sku = tostring(sku.name), tier = tostring(sku.tier),
          zoneRedundant = tostring(properties.zoneRedundant), backupRedundancy = tostring(properties.currentBackupStorageRedundancy)
'@
        $sqlRows = @($sql | Select-Object @{N = 'Database'; E = { $_.name } }, @{N = 'Server'; E = { $_.server } }, @{N = 'Sku'; E = { $_.sku } }, @{N = 'Tier'; E = { $_.tier } },
            @{N = 'ZoneRedundant'; E = { $_.zoneRedundant } }, @{N = 'BackupRedundancy'; E = { $_.backupRedundancy } }, @{N = 'Location'; E = { $_.location } },
            @{N = 'Subscription'; E = { Get-SubName $_.subscriptionId } }, @{N = 'ResourceGroup'; E = { $_.resourceGroup } })
        Export-AssessmentCsv $sqlRows 'Azure_SQLDatabases.csv' @('Database', 'Server', 'Sku', 'Tier', 'ZoneRedundant', 'BackupRedundancy', 'Location', 'Subscription', 'ResourceGroup')

        $redundancy = @($script:StorageRows | Group-Object Sku | Sort-Object Count -Descending | Select-Object @{N = 'Redundancy'; E = { $_.Name } }, Count)
        $lrs = [int](($script:StorageRows | Where-Object { $_.Sku -match '_LRS$' } | Measure-Object).Count)
        $placement = @($script:Vms | Group-Object Placement | Sort-Object Count -Descending | Select-Object @{N = 'Placement'; E = { $_.Name } }, Count)

        $script:Data['RunningVMs']                = $running.Count
        $script:Data['RunningVMsWithoutBackup']   = $noBackup.Count
        $script:Data['BackupProtectedItems']      = $itemRows.Count
        $script:Data['BackupStaleOrFailed']       = @($stale + $failed | Sort-Object SourceResourceId -Unique).Count
        $script:Data['AsrReplicatedItems']        = $asrRows.Count
        $script:Data['SingleInstanceVMs']         = $single.Count
        $script:Data['StorageAccountsLRS']        = $lrs
        Write-Host "  Running VMs: $($running.Count) | Without Azure Backup: $($noBackup.Count) | ASR items: $($asrRows.Count) | Single-instance VMs: $($single.Count)" -ForegroundColor Green

        if ($noProtection.Count -gt 0) {
            Add-Finding -Id 'AZ-RES-001' -Severity High -Title 'Running VMs with no Azure Backup and no Site Recovery' -Count $noProtection.Count `
                -Detail (($noProtection | Select-Object -First 15 | ForEach-Object { $_.Name.ToUpper() }) -join ', ') `
                -Recommendation 'Confirm whether a third-party backup covers these VMs; if not, add them to an Azure Backup policy.' -EvidenceFile 'Azure_VMs.csv'
        }
        if ($stale.Count + $failed.Count -gt 0) {
            Add-Finding -Id 'AZ-RES-002' -Severity Medium -Title 'Backup items with failed or stale (>48h) backups' -Count (@($stale + $failed | Sort-Object SourceResourceId -Unique).Count) `
                -Recommendation 'Investigate the failing jobs and confirm backup alerts reach someone who acts on them.' -EvidenceFile 'Azure_BackupItems.csv'
        }
        if ($weakVaults.Count -gt 0) {
            Add-Finding -Id 'AZ-RES-003' -Severity Medium -Title 'Recovery Services vaults without immutability or with soft delete off' -Count $weakVaults.Count `
                -Recommendation 'Enable immutability (and lock it once validated) and keep soft delete on to protect backups from ransomware and admin error.' -EvidenceFile 'Azure_RecoveryVaults.csv'
        }
        if ($single.Count -gt 0) {
            Add-Finding -Id 'AZ-RES-004' -Severity Low -Title 'VMs deployed as single instances (no zone, availability set or scale set)' -Count $single.Count `
                -Recommendation 'For production workloads, pair VMs across availability zones or put them behind a scale set; single-instance VMs get the lowest SLA.' -EvidenceFile 'Azure_VMs.csv'
        }
        if ($lrs -gt 0) {
            Add-Finding -Id 'AZ-RES-005' -Severity Info -Title 'Storage accounts using locally redundant storage (LRS)' -Count $lrs `
                -Recommendation 'Check that LRS is a deliberate choice; production data usually warrants ZRS or GRS.' -EvidenceFile 'Azure_StorageAccounts.csv'
        }

        $gapTable = $noBackup | Select-Object @{N = 'VM'; E = { $_.Name.ToUpper() } }, Subscription, ResourceGroup, Placement, AsrReplicated
        $html = "<div class='stat-row'>" +
            (New-StatCard 'Running VMs' $running.Count) +
            (New-StatCard 'Without Azure Backup' $noBackup.Count $(if ($noBackup.Count) { 'bad' } else { 'good' })) +
            (New-StatCard 'Backup items' $itemRows.Count 'neutral') +
            (New-StatCard 'Failed / stale backups' $script:Data['BackupStaleOrFailed'] $(if ($script:Data['BackupStaleOrFailed']) { 'warn' } else { 'good' })) +
            (New-StatCard 'ASR replicated items' $asrRows.Count 'neutral') +
            (New-StatCard 'Single-instance VMs' $single.Count $(if ($single.Count) { 'warn' } else { 'good' })) +
            "</div><h3>Running VMs without Azure Backup</h3>" +
            (ConvertTo-HtmlTable $gapTable @('VM', 'Subscription', 'ResourceGroup', 'Placement', 'AsrReplicated') -CsvName 'Azure_VMs.csv') +
            "<div class='two-col'><div><h3>VM placement</h3>" + (ConvertTo-HtmlTable $placement @('Placement', 'Count')) +
            "</div><div><h3>Storage redundancy</h3>" + (ConvertTo-HtmlTable $redundancy @('Redundancy', 'Count')) + "</div></div>" +
            "<h3>Recovery Services vaults</h3>" +
            (ConvertTo-HtmlTable $vaultRows @('Vault', 'StorageRedundancy', 'CrossRegionRestore', 'SoftDelete', 'Immutability', 'Subscription') -CsvName 'Azure_RecoveryVaults.csv') +
            $(if (@($bi).Count) { "<p class='note'>Backup vault (Data Protection) instances: " + (ConvertTo-HtmlEncoded (($bi | ForEach-Object { "$($_.datasourceType): $($_.Count)" }) -join '; ')) + "</p>" } else { '' }) +
            "<p class='note'>Only Azure Backup and Azure Site Recovery are visible to this script. Third-party backup (Veeam, Commvault, etc.) must be confirmed with the customer.</p>"
        Add-HtmlSection $title $html
    }
    catch { Add-ErrorSection $title $_ }
}

# ═════════════════════════════════════════════════════════════════════════════
# 6. Cost & Waste
# ═════════════════════════════════════════════════════════════════════════════
function Get-CostAndWaste {
    param([int]$Total)
    $title = 'Cost & Waste'
    Write-SectionHeader $title $Total
    try {
        # Actual cost per subscription (Cost Management)
        $svcRows = [System.Collections.Generic.List[object]]::new()
        $costFailed = [System.Collections.Generic.List[string]]::new()
        $to   = (Get-Date).ToUniversalTime().Date.AddDays(-1)
        $from = $to.AddDays(-($CostDays - 1))
        if (-not $SkipCost) {
            $body = @{
                type       = 'ActualCost'
                timeframe  = 'Custom'
                timePeriod = @{ from = $from.ToString('yyyy-MM-ddT00:00:00Z'); to = $to.ToString('yyyy-MM-ddT23:59:59Z') }
                dataset    = @{
                    granularity = 'Monthly'
                    aggregation = @{ totalCost = @{ name = 'Cost'; function = 'Sum' } }
                    grouping    = @(@{ type = 'Dimension'; name = 'ServiceName' })
                }
            } | ConvertTo-Json -Depth 10
            foreach ($s in @($script:SubRows | Where-Object { $_.State -eq 'Enabled' })) {
                Write-Host "    Cost: $($s.Subscription)" -ForegroundColor Gray
                try {
                    $r = Invoke-ArmRequest -Method POST -Path "/subscriptions/$($s.SubscriptionId)/providers/Microsoft.CostManagement/query?api-version=2023-11-01" -Payload $body
                    while ($r) {
                        $cols = @($r.properties.columns | ForEach-Object { $_.name })
                        $iCost = [array]::IndexOf($cols, 'Cost'); if ($iCost -lt 0) { $iCost = [array]::IndexOf($cols, 'PreTaxCost') }
                        $iSvc  = [array]::IndexOf($cols, 'ServiceName')
                        $iCur  = [array]::IndexOf($cols, 'Currency')
                        $iMon  = [array]::IndexOf($cols, 'BillingMonth'); if ($iMon -lt 0) { $iMon = [array]::IndexOf($cols, 'UsageDate') }
                        foreach ($row in @($r.properties.rows)) {
                            $rawMonth = if ($iMon -ge 0) { $row[$iMon] } else { $null }
                            $month = if ($rawMonth -is [datetime]) { $rawMonth.ToString('yyyy-MM') } else { "$rawMonth" }
                            if ($month -match '^(\d{4})-?(\d{2})') { $month = "$($Matches[1])-$($Matches[2])" }
                            $svcRows.Add([pscustomobject]@{
                                    Subscription = $s.Subscription; SubscriptionId = $s.SubscriptionId; Month = $month
                                    Service = $(if ($iSvc -ge 0 -and "$($row[$iSvc])") { "$($row[$iSvc])" } else { 'Unassigned' })
                                    Cost = [math]::Round([double]$row[$iCost], 2); Currency = $(if ($iCur -ge 0) { "$($row[$iCur])" } else { '' })
                                })
                        }
                        $r = if ($r.properties.nextLink) { Invoke-ArmRequest -Method POST -Uri $r.properties.nextLink -Payload $body } else { $null }
                    }
                }
                catch { $costFailed.Add("$($s.Subscription): $($_.Exception.Message)") }
            }
            if ($costFailed.Count -gt 0) {
                Add-NotCollected -Item 'Actual cost (Cost Management)' -Reason (($costFailed | Select-Object -First 5) -join ' | ') `
                    -HowToFix 'Assign Cost Management Reader (or Billing Reader) on the subscriptions. CSP subscriptions need cost visibility enabled by the partner, or provide an invoice/cost export instead.'
            }
        }
        else {
            Add-NotCollected -Item 'Actual cost (Cost Management)' -Reason 'Skipped with -SkipCost.' -HowToFix 'Re-run without -SkipCost, or provide a cost export.'
        }
        $byService = @($svcRows | Group-Object Service, Currency | ForEach-Object {
                [pscustomobject]@{ Service = $_.Group[0].Service; Currency = $_.Group[0].Currency; Cost = [math]::Round(($_.Group | Measure-Object Cost -Sum).Sum, 2) }
            } | Sort-Object Cost -Descending)
        $byMonth = @($svcRows | Group-Object Month, Currency | ForEach-Object {
                [pscustomobject]@{ Month = $_.Group[0].Month; Currency = $_.Group[0].Currency; Cost = [math]::Round(($_.Group | Measure-Object Cost -Sum).Sum, 2) }
            } | Sort-Object Month)
        $bySub = @($svcRows | Group-Object Subscription, Currency | ForEach-Object {
                [pscustomobject]@{ Subscription = $_.Group[0].Subscription; Currency = $_.Group[0].Currency; Cost = [math]::Round(($_.Group | Measure-Object Cost -Sum).Sum, 2) }
            } | Sort-Object Cost -Descending)
        Export-AssessmentCsv $svcRows 'Azure_Cost_Detail.csv' @('Subscription', 'SubscriptionId', 'Month', 'Service', 'Cost', 'Currency')
        Export-AssessmentCsv $byService 'Azure_Cost_ByService.csv' @('Service', 'Currency', 'Cost')
        Export-AssessmentCsv $byMonth 'Azure_Cost_ByMonth.csv' @('Month', 'Currency', 'Cost')
        Export-AssessmentCsv $bySub 'Azure_Cost_BySubscription.csv' @('Subscription', 'Currency', 'Cost')
        $currencies = @($svcRows | ForEach-Object Currency | Where-Object { $_ } | Sort-Object -Unique)
        $costTotal = if ($currencies.Count -eq 1) { [math]::Round(($svcRows | Measure-Object Cost -Sum).Sum, 2) } else { $null }

        # Advisor
        $adv = Invoke-ArgQuery -Name 'Azure Advisor recommendations' -Query @'
advisorresources
| where type =~ 'microsoft.advisor/recommendations'
| extend category = tostring(properties.category), impact = tostring(properties.impact),
         problem = tostring(properties.shortDescription.problem), solution = tostring(properties.shortDescription.solution),
         impactedType = tostring(properties.impactedField), impactedName = tostring(properties.impactedValue),
         annualSavings = todouble(properties.extendedProperties.annualSavingsAmount), monthlySavings = todouble(properties.extendedProperties.savingsAmount),
         currency = tostring(properties.extendedProperties.savingsCurrency), resourceId = tostring(properties.resourceMetadata.resourceId)
| project id, subscriptionId, category, impact, problem, solution, impactedType, impactedName, annualSavings, monthlySavings, currency, resourceId
'@
        $advRows = foreach ($a in $adv) {
            $annual = if ($a.annualSavings) { [double]$a.annualSavings } elseif ($a.monthlySavings) { [double]$a.monthlySavings * 12 } else { $null }
            [pscustomobject]@{
                Category = $a.category; Impact = $a.impact; Problem = $a.problem; Solution = $a.solution
                ImpactedType = $a.impactedType; ImpactedResource = $a.impactedName
                AnnualSavings = $(if ($null -ne $annual) { [math]::Round($annual, 2) } else { $null }); Currency = $a.currency
                Subscription = Get-SubName $a.subscriptionId; ResourceId = $a.resourceId
            }
        }
        $advRows = @($advRows)
        $script:AdvisorRows = $advRows
        Export-AssessmentCsv $advRows 'Azure_Advisor.csv' @('Category', 'Impact', 'Problem', 'Solution', 'ImpactedType', 'ImpactedResource', 'AnnualSavings', 'Currency', 'Subscription', 'ResourceId')
        $advCost = @($advRows | Where-Object { $_.Category -eq 'Cost' })
        $advSavings = @($advCost | Where-Object { $null -ne $_.AnnualSavings } | Group-Object Currency | ForEach-Object {
                [pscustomobject]@{ Currency = $_.Name; AnnualSavings = [math]::Round(($_.Group | Measure-Object AnnualSavings -Sum).Sum, 2) } })
        $commitRecs = @($advCost | Where-Object { "$($_.Problem) $($_.Solution)" -match '(?i)reserv|savings plan' })
        $advSummary = @($advCost | Group-Object Problem | Sort-Object Count -Descending | ForEach-Object {
                [pscustomobject]@{ Recommendation = $_.Name; Resources = $_.Count
                    AnnualSavings = [math]::Round(($_.Group | Where-Object { $null -ne $_.AnnualSavings } | Measure-Object AnnualSavings -Sum).Sum, 2)
                    Currency = ($_.Group | Where-Object Currency | Select-Object -First 1).Currency }
            })

        # Advisor by WAF pillar (Advisor categories are the WAF pillars)
        $pillarOf = @{ 'HighAvailability' = 'Reliability'; 'Security' = 'Security'; 'Cost' = 'Cost Optimization'; 'OperationalExcellence' = 'Operational Excellence'; 'Performance' = 'Performance Efficiency' }
        $advByPillar = @(foreach ($cat in 'HighAvailability', 'Security', 'Cost', 'OperationalExcellence', 'Performance') {
                $g = @($advRows | Where-Object { $_.Category -eq $cat })
                [pscustomobject]@{ WafPillar = $pillarOf[$cat]; AdvisorCategory = $cat; Recommendations = $g.Count
                    HighImpact = @($g | Where-Object { $_.Impact -eq 'High' }).Count; DistinctProblems = @($g | ForEach-Object Problem | Sort-Object -Unique).Count }
            })
        Export-AssessmentCsv $advByPillar 'Azure_Advisor_ByPillar.csv' @('WafPillar', 'AdvisorCategory', 'Recommendations', 'HighImpact', 'DistinctProblems')
        $advFind = @(
            @{ Id = 'AZ-ADV-001'; Cat = 'HighAvailability'; Title = 'Azure Advisor reliability recommendations'; Rec = 'Review the Advisor reliability (high availability) recommendations with the workload owners; run WARA for production workloads.' }
            @{ Id = 'AZ-ADV-002'; Cat = 'Performance'; Title = 'Azure Advisor performance recommendations'; Rec = 'Review the Advisor performance recommendations against the workload performance targets.' }
            @{ Id = 'AZ-ADV-003'; Cat = 'OperationalExcellence'; Title = 'Azure Advisor operational excellence recommendations'; Rec = 'Review the Advisor operational excellence recommendations; many are quick configuration fixes.' }
        )
        foreach ($af in $advFind) {
            $row = $advByPillar | Where-Object { $_.AdvisorCategory -eq $af.Cat }
            if ($row.Recommendations -gt 0) {
                Add-Finding -Id $af.Id -Severity $(if ($row.HighImpact -gt 0) { 'Medium' } else { 'Low' }) -Title "$($af.Title)$(if ($row.HighImpact) { " ($($row.HighImpact) high impact)" })" -Count $row.Recommendations `
                    -Recommendation $af.Rec -EvidenceFile 'Azure_Advisor.csv' -Confidence 'reported'
            }
        }

        # Budgets (cost guardrails)
        $budgetRows = [System.Collections.Generic.List[object]]::new()
        $budgetFailed = 0
        $subsNoBudget = [System.Collections.Generic.List[string]]::new()
        foreach ($s in @($script:SubRows | Where-Object { $_.State -eq 'Enabled' })) {
            try {
                $b = Get-ArmList "/subscriptions/$($s.SubscriptionId)/providers/Microsoft.Consumption/budgets?api-version=2023-05-01"
                if ($b.Count -eq 0) { $subsNoBudget.Add($s.Subscription) }
                foreach ($x in $b) {
                    $budgetRows.Add([pscustomobject]@{ Subscription = $s.Subscription; Budget = $x.name; Amount = $x.properties.amount; TimeGrain = $x.properties.timeGrain
                            Alerts = @($x.properties.notifications.PSObject.Properties).Count })
                }
            }
            catch { $budgetFailed++ }
        }
        if ($budgetFailed -gt 0) { Add-NotCollected -Item 'Budgets' -Reason "Budgets could not be read on $budgetFailed subscription(s)." -HowToFix 'Assign Cost Management Reader on the subscriptions.' }
        Export-AssessmentCsv $budgetRows 'Azure_Budgets.csv' @('Subscription', 'Budget', 'Amount', 'TimeGrain', 'Alerts')
        if ($subsNoBudget.Count -gt 0) {
            Add-Finding -Id 'AZ-CST-009' -Severity Medium -Title 'Subscriptions with no budget or cost alert' -Count $subsNoBudget.Count `
                -Detail ($subsNoBudget -join '; ') `
                -Recommendation 'Create a budget with alert thresholds (e.g. 80%, 100%, forecast 110%) on every subscription, owned by someone who acts on it.' -EvidenceFile 'Azure_Budgets.csv'
        }

        # Waste
        $waste = [System.Collections.Generic.List[object]]::new()
        $addWaste = {
            param($Category, $Rows, $DetailScript)
            foreach ($r in $Rows) {
                $waste.Add([pscustomobject]@{
                        Category = $Category; Name = $r.name; Subscription = Get-SubName $r.subscriptionId; ResourceGroup = $r.resourceGroup
                        Location = $r.location; Detail = (& $DetailScript $r); ResourceId = $r.id
                    })
            }
        }
        $disks = Invoke-ArgQuery -Name 'Unattached managed disks' -Query @'
resources
| where type =~ 'microsoft.compute/disks'
| where tostring(properties.diskState) =~ 'Unattached'
| project id, name, resourceGroup, subscriptionId, location, sizeGB = toint(properties.diskSizeGB), sku = tostring(sku.name), timeCreated = tostring(properties.timeCreated), tags = tostring(tags)
'@
        & $addWaste 'Unattached disk' $disks { param($r) "$($r.sizeGB) GB $($r.sku)" + $(if ("$($r.tags) $($r.name)" -match '(?i)asr-?replica|asrseed') { ' (looks like an ASR replica/seed disk - check before deleting)' } else { '' }) }

        $pips = Invoke-ArgQuery -Name 'Unassociated public IPs' -Query @'
resources
| where type =~ 'microsoft.network/publicipaddresses'
| where isempty(tostring(properties.ipConfiguration.id)) and isempty(tostring(properties.natGateway.id))
| project id, name, resourceGroup, subscriptionId, location, sku = tostring(sku.name), ip = tostring(properties.ipAddress), allocation = tostring(properties.publicIPAllocationMethod)
'@
        & $addWaste 'Unassociated public IP' $pips { param($r) "$($r.sku) $($r.allocation) $($r.ip)".Trim() }

        $nics = Invoke-ArgQuery -Name 'Unattached NICs' -Query @'
resources
| where type =~ 'microsoft.network/networkinterfaces'
| where isempty(tostring(properties.virtualMachine.id)) and isempty(tostring(properties.privateEndpoint.id)) and isempty(tostring(properties.privateLinkService.id))
| project id, name, resourceGroup, subscriptionId, location
'@
        & $addWaste 'Unattached NIC' $nics { param($r) 'No cost on its own; housekeeping' }

        $snaps = Invoke-ArgQuery -Name 'Snapshots older than 90 days' -Query @'
resources
| where type =~ 'microsoft.compute/snapshots'
| extend created = todatetime(properties.timeCreated)
| where created < ago(90d)
| project id, name, resourceGroup, subscriptionId, location, sizeGB = toint(properties.diskSizeGB), created = tostring(properties.timeCreated)
'@
        & $addWaste 'Snapshot older than 90 days' $snaps { param($r) "$($r.sizeGB) GB, created $("$($r.created)".Substring(0, [math]::Min(10, "$($r.created)".Length)))" }

        $asp = Invoke-ArgQuery -Name 'Empty App Service plans' -Query @'
resources
| where type =~ 'microsoft.web/serverfarms'
| where toint(properties.numberOfSites) == 0 and tostring(sku.tier) !in~ ('Free','Shared','Dynamic')
| project id, name, resourceGroup, subscriptionId, location, sku = tostring(sku.name), tier = tostring(sku.tier)
'@
        & $addWaste 'Empty App Service plan' $asp { param($r) "$($r.tier) $($r.sku)" }

        $rgs = Invoke-ArgQuery -Name 'Empty resource groups' -Query @'
resourcecontainers
| where type =~ 'microsoft.resources/subscriptions/resourcegroups'
| extend rgKey = tolower(id)
| join kind=leftouter (
    resources
    | extend rgKey = tolower(strcat('/subscriptions/', subscriptionId, '/resourcegroups/', resourceGroup))
    | summarize Resources = count() by rgKey
  ) on rgKey
| where isnull(Resources)
| project id, name, resourceGroup = name, subscriptionId, location
'@
        & $addWaste 'Empty resource group' $rgs { param($r) 'No resources' }

        foreach ($vm in @($script:Vms | Where-Object { $_.PowerState -eq 'stopped' })) {
            $waste.Add([pscustomobject]@{ Category = 'VM stopped but not deallocated'; Name = $vm.Name; Subscription = $vm.Subscription; ResourceGroup = $vm.ResourceGroup
                    Location = $vm.Location; Detail = "$($vm.Size) - compute is still billed"; ResourceId = $vm.ResourceId })
        }
        foreach ($vm in @($script:Vms | Where-Object { $_.PowerState -eq 'deallocated' })) {
            $waste.Add([pscustomobject]@{ Category = 'VM deallocated (disks still billed)'; Name = $vm.Name; Subscription = $vm.Subscription; ResourceGroup = $vm.ResourceGroup
                    Location = $vm.Location; Detail = "$($vm.Size) - confirm whether still needed"; ResourceId = $vm.ResourceId })
        }
        Export-AssessmentCsv $waste 'Azure_Waste.csv' @('Category', 'Name', 'Subscription', 'ResourceGroup', 'Location', 'Detail', 'ResourceId')
        $wasteSummary = @($waste | Group-Object Category | Sort-Object Count -Descending | Select-Object @{N = 'Category'; E = { $_.Name } }, Count)

        # Azure Hybrid Benefit (Windows Server VMs without a Windows_Server licence type)
        $ahb = @($script:Vms | Where-Object { $_.OsType -eq 'Windows' -and !$_.LicenseType -and $_.OsLabel -like 'Windows Server*' -and $_.Priority -ne 'Spot' })

        $script:Data['CostPeriodDays']        = $(if ($SkipCost) { $null } else { $CostDays })
        $script:Data['CostTotal']             = $costTotal
        $script:Data['CostCurrency']          = $(if ($currencies.Count -eq 1) { $currencies[0] } elseif ($currencies.Count -gt 1) { 'Mixed' } else { $null })
        $script:Data['AdvisorCostRecommendations'] = $advCost.Count
        $script:Data['AdvisorAnnualSavings']  = $(if ($advSavings.Count -eq 1) { $advSavings[0].AnnualSavings } else { $null })
        $script:Data['AdvisorSavingsCurrency'] = $(if ($advSavings.Count -eq 1) { $advSavings[0].Currency } else { $null })
        $script:Data['WasteItems']            = $waste.Count
        $script:Data['AhbOpportunityVMs']     = $ahb.Count
        $script:Data['SubscriptionsWithoutBudget'] = $(if ($budgetFailed) { $null } else { $subsNoBudget.Count })
        foreach ($r in $advByPillar) { $script:Data["Advisor$($r.AdvisorCategory)"] = $r.Recommendations }
        Write-Host "  Spend ($CostDays d): $(if ($null -ne $costTotal) { "$costTotal $($currencies[0])" } else { 'N/A' }) | Advisor cost recs: $($advCost.Count) | Waste items: $($waste.Count)" -ForegroundColor Green

        if ($advCost.Count -gt 0) {
            $sv = ($advSavings | ForEach-Object { "{0:N0} {1}/year" -f $_.AnnualSavings, $_.Currency }) -join ' + '
            Add-Finding -Id 'AZ-CST-001' -Severity Medium -Title "Azure Advisor cost recommendations ($sv)" -Count $advCost.Count `
                -Recommendation 'Right-size or shut down the flagged resources; validate utilisation with the workload owner first.' -EvidenceFile 'Azure_Advisor.csv' -Confidence 'reported'
        }
        if ($commitRecs.Count -gt 0) {
            Add-Finding -Id 'AZ-CST-002' -Severity Low -Title 'Reservation / savings plan purchase opportunities' -Count $commitRecs.Count `
                -Recommendation 'Review 1-year/3-year reservations or a compute savings plan for steady-state workloads.' -EvidenceFile 'Azure_Advisor.csv' -Confidence 'reported'
        }
        $wasteCats = @{ 'Unattached disk' = 'AZ-CST-003'; 'Unassociated public IP' = 'AZ-CST-004'; 'Snapshot older than 90 days' = 'AZ-CST-005'; 'Empty App Service plan' = 'AZ-CST-006'; 'VM stopped but not deallocated' = 'AZ-CST-007' }
        foreach ($w in $wasteSummary) {
            if ($wasteCats.ContainsKey($w.Category)) {
                Add-Finding -Id $wasteCats[$w.Category] -Severity $(if ($w.Category -eq 'VM stopped but not deallocated') { 'Medium' } else { 'Low' }) -Title "$($w.Category) ($($w.Count))" -Count $w.Count `
                    -Recommendation 'Confirm with the owner, then delete or deallocate. Snapshot/disk deletion is irreversible.' -EvidenceFile 'Azure_Waste.csv'
            }
        }
        if ($ahb.Count -gt 0) {
            Add-Finding -Id 'AZ-CST-008' -Severity Low -Title 'Windows Server VMs not using Azure Hybrid Benefit' -Count $ahb.Count `
                -Recommendation 'If the customer has Windows Server licences with Software Assurance, apply Azure Hybrid Benefit to cut compute cost.' -EvidenceFile 'Azure_VMs.csv' -Confidence 'inferred'
        }

        $html = "<div class='stat-row'>" +
            (New-StatCard "Actual cost (last $CostDays days)" $(if ($null -ne $costTotal) { "{0:N0} {1}" -f $costTotal, $currencies[0] } elseif ($currencies.Count -gt 1) { 'Mixed currencies' } else { 'N/A' }) 'info') +
            (New-StatCard 'Advisor cost recommendations' $advCost.Count $(if ($advCost.Count) { 'warn' } else { 'good' })) +
            (New-StatCard 'Advisor savings / year' $(if ($advSavings.Count -eq 1) { "{0:N0} {1}" -f $advSavings[0].AnnualSavings, $advSavings[0].Currency } elseif ($advSavings.Count) { 'Mixed currencies' } else { 'N/A' }) 'warn') +
            (New-StatCard 'Waste items' $waste.Count $(if ($waste.Count) { 'warn' } else { 'good' })) +
            (New-StatCard 'AHB opportunities' $ahb.Count 'neutral') +
            "</div><div class='two-col'><div><h3>Spend by month</h3>" +
            (ConvertTo-HtmlTable ($byMonth | Select-Object Month, @{N = 'Cost'; E = { '{0:N2}' -f $_.Cost } }, Currency) @('Month', 'Cost', 'Currency') -CsvName 'Azure_Cost_ByMonth.csv') +
            "</div><div><h3>Top services</h3>" +
            (ConvertTo-HtmlTable ($byService | Select-Object Service, @{N = 'Cost'; E = { '{0:N2}' -f $_.Cost } }, Currency) @('Service', 'Cost', 'Currency') -MaxRows 10 -CsvName 'Azure_Cost_ByService.csv') +
            "</div></div><h3>Spend by subscription</h3>" +
            (ConvertTo-HtmlTable ($bySub | Select-Object Subscription, @{N = 'Cost'; E = { '{0:N2}' -f $_.Cost } }, Currency) @('Subscription', 'Cost', 'Currency') -MaxRows 15 -CsvName 'Azure_Cost_BySubscription.csv') +
            "<h3>Azure Advisor recommendations by WAF pillar</h3>" +
            (ConvertTo-HtmlTable $advByPillar @('WafPillar', 'Recommendations', 'HighImpact', 'DistinctProblems') -CsvName 'Azure_Advisor_ByPillar.csv') +
            "<h3>Advisor cost recommendations</h3>" +
            (ConvertTo-HtmlTable ($advSummary | Select-Object Recommendation, Resources, @{N = 'AnnualSavings'; E = { if ($_.AnnualSavings) { '{0:N0} {1}' -f $_.AnnualSavings, $_.Currency } else { '' } } }) @('Recommendation', 'Resources', 'AnnualSavings') -MaxRows 15 -CsvName 'Azure_Advisor.csv') +
            "<h3>Waste and housekeeping</h3>" +
            (ConvertTo-HtmlTable $wasteSummary @('Category', 'Count') -CsvName 'Azure_Waste.csv') +
            "<p class='note'>Costs are actual (billed) cost in the billing currency from Cost Management. Reservation and savings plan utilisation needs billing-account access and is not collected.</p>"
        Add-HtmlSection $title $html
    }
    catch { Add-ErrorSection $title $_ }
}

# ═════════════════════════════════════════════════════════════════════════════
# 7. Operations
# ═════════════════════════════════════════════════════════════════════════════
function Get-Operations {
    param([int]$Total)
    $title = 'Operations'
    Write-SectionHeader $title $Total
    try {
        $law = Invoke-ArgQuery -Name 'Log Analytics workspaces' -Query @'
resources
| where type =~ 'microsoft.operationalinsights/workspaces'
| project id, name, resourceGroup, subscriptionId, location, sku = tostring(properties.sku.name), retentionDays = toint(properties.retentionInDays),
          dailyCapGb = todouble(properties.workspaceCapping.dailyQuotaGb)
'@
        $sentinel = Invoke-ArgQuery -Name 'Microsoft Sentinel' -Query "resources | where type =~ 'microsoft.operationsmanagement/solutions' | where name startswith 'SecurityInsights' | project id, name, workspace = tolower(tostring(properties.workspaceResourceId))"
        $sentinelWs = @{}; foreach ($s in $sentinel) { if ($s.workspace) { $sentinelWs[$s.workspace] = $true } }
        $lawRows = @($law | ForEach-Object {
                [pscustomobject]@{ Workspace = $_.name; Location = $_.location; Sku = $_.sku; RetentionDays = $_.retentionDays
                    DailyCapGb = $(if ($_.dailyCapGb -lt 0) { 'None' } else { $_.dailyCapGb }); Sentinel = $sentinelWs.ContainsKey("$($_.id)".ToLower())
                    Subscription = Get-SubName $_.subscriptionId; ResourceGroup = $_.resourceGroup }
            })
        Export-AssessmentCsv $lawRows 'Azure_LogAnalytics.csv' @('Workspace', 'Location', 'Sku', 'RetentionDays', 'DailyCapGb', 'Sentinel', 'Subscription', 'ResourceGroup')

        # Agents on VMs and Arc machines
        $ext = Invoke-ArgQuery -Name 'Monitoring and security agents' -Query @'
resources
| where type in~ ('microsoft.compute/virtualmachines/extensions', 'microsoft.hybridcompute/machines/extensions')
| extend machineId = tostring(split(tolower(id), '/extensions/')[0]), extType = tostring(properties.type)
| where extType in~ ('AzureMonitorWindowsAgent','AzureMonitorLinuxAgent','MicrosoftMonitoringAgent','OmsAgentForLinux','MDE.Windows','MDE.Linux')
| project id, machineId, extType, state = tostring(properties.provisioningState)
'@
        $extByMachine = @{}
        foreach ($e in $ext) {
            if (!$extByMachine.ContainsKey($e.machineId)) { $extByMachine[$e.machineId] = [System.Collections.Generic.List[string]]::new() }
            $extByMachine[$e.machineId].Add($e.extType)
        }
        foreach ($vm in $script:Vms) {
            $types = if ($extByMachine.ContainsKey($vm.ResourceId)) { @($extByMachine[$vm.ResourceId]) } else { @() }
            $vm.AMA         = [bool]($types | Where-Object { $_ -match '^AzureMonitor(Windows|Linux)Agent$' })
            $vm.LegacyAgent = [bool]($types | Where-Object { $_ -match '^(MicrosoftMonitoringAgent|OmsAgentForLinux)$' })
            $vm.MDE         = [bool]($types | Where-Object { $_ -match '^MDE\.' })
        }

        # Arc-enabled servers
        $arc = Invoke-ArgQuery -Name 'Arc-enabled servers' -Query @'
resources
| where type =~ 'microsoft.hybridcompute/machines'
| project id, name, resourceGroup, subscriptionId, location, status = tostring(properties.status), osType = tostring(properties.osType),
          osName = tostring(properties.osName), osSku = tostring(properties.osSku), osVersion = tostring(properties.osVersion),
          agentVersion = tostring(properties.agentVersion), lastStatusChange = tostring(properties.lastStatusChange)
'@
        $script:ArcRows = @($arc | ForEach-Object {
                $id = "$($_.id)".ToLower()
                $types = if ($extByMachine.ContainsKey($id)) { @($extByMachine[$id]) } else { @() }
                $life = Get-OsLifecycle -OsType $_.osType -OsName $_.osSku -OsVersion $_.osVersion -Offer '' -Sku ''
                [pscustomobject]@{ Machine = $_.name; Status = $_.status; Os = $(if ($_.osSku) { $_.osSku } else { $_.osName }); OsVersion = $_.osVersion
                    OsLabel = $life.Os; OsLifecycle = $life.Status; OsSeverity = $life.Severity; AgentVersion = $_.agentVersion
                    AMA = [bool]($types | Where-Object { $_ -match '^AzureMonitor' }); MDE = [bool]($types | Where-Object { $_ -match '^MDE\.' })
                    LastStatusChange = $_.lastStatusChange; Subscription = Get-SubName $_.subscriptionId; ResourceGroup = $_.resourceGroup; ResourceId = $id }
            })
        Export-AssessmentCsv $script:ArcRows 'Azure_ArcMachines.csv' @('Machine', 'Status', 'Os', 'OsVersion', 'OsLabel', 'OsLifecycle', 'OsSeverity', 'AgentVersion', 'AMA', 'MDE', 'LastStatusChange', 'Subscription', 'ResourceGroup', 'ResourceId')

        # Patch assessment (Azure Update Manager)
        $pa = Invoke-ArgQuery -Name 'Update Manager patch assessments' -Query @'
patchassessmentresources
| where type in~ ('microsoft.compute/virtualmachines/patchassessmentresults', 'microsoft.hybridcompute/machines/patchassessmentresults')
| extend machineId = tostring(split(tolower(id), '/patchassessmentresults/')[0])
| project id, machineId, lastAssess = tostring(properties.lastModifiedDateTime),
          critical = toint(properties.availablePatchCountByClassification.critical), security = toint(properties.availablePatchCountByClassification.security)
'@
        foreach ($p in $pa) {
            $vm = $script:VmById[$p.machineId]
            if ($vm) { $vm.LastPatchAssess = "$($p.lastAssess)"; $vm.MissingCritical = $p.critical; $vm.MissingSecurity = $p.security }
        }

        # Activity log export (per subscription)
        $alRows = [System.Collections.Generic.List[object]]::new()
        foreach ($s in @($script:SubRows | Where-Object { $_.State -eq 'Enabled' })) {
            try {
                $ds = Get-ArmList "/subscriptions/$($s.SubscriptionId)/providers/Microsoft.Insights/diagnosticSettings?api-version=2021-05-01-preview"
                $alRows.Add([pscustomobject]@{ Subscription = $s.Subscription; SubscriptionId = $s.SubscriptionId; Exported = ($ds.Count -gt 0)
                        Destinations = (($ds | ForEach-Object { @($_.properties.workspaceId, $_.properties.storageAccountId, $_.properties.eventHubAuthorizationRuleId) | Where-Object { $_ } | ForEach-Object { ($_ -split '/')[-1] } }) -join '; ') })
            }
            catch { $alRows.Add([pscustomobject]@{ Subscription = $s.Subscription; SubscriptionId = $s.SubscriptionId; Exported = $null; Destinations = "Not readable: $($_.Exception.Message)" }) }
        }
        Export-AssessmentCsv $alRows 'Azure_ActivityLogExport.csv' @('Subscription', 'SubscriptionId', 'Exported', 'Destinations')

        # Diagnostic settings on key resources
        $diagRows = [System.Collections.Generic.List[object]]::new()
        if ($MaxDiagnosticChecks -gt 0) {
            $typeList = ($script:DiagnosticTypes | ForEach-Object { "'$_'" }) -join ','
            $cand = Invoke-ArgQuery -Name 'Resources for diagnostic settings check' -Query "resources | where type in~ ($typeList) | where not(type =~ 'microsoft.sql/servers/databases' and name =~ 'master') | project id, name, type, resourceGroup, subscriptionId"
            $ordered = @($cand | Sort-Object { [array]::IndexOf($script:DiagnosticTypes, "$($_.type)".ToLower()) }, name)
            $check = @($ordered | Select-Object -First $MaxDiagnosticChecks)
            if ($ordered.Count -gt $check.Count) {
                Add-NotCollected -Item 'Diagnostic settings (partial)' -Reason "Checked $($check.Count) of $($ordered.Count) resources (limit -MaxDiagnosticChecks $MaxDiagnosticChecks)." -HowToFix 'Re-run with a higher -MaxDiagnosticChecks for full coverage.'
            }
            $i = 0
            foreach ($c in $check) {
                $i++
                if ($i % 25 -eq 0) { Write-Host "    Diagnostic settings: $i / $($check.Count)" -ForegroundColor Gray }
                try {
                    $ds = Get-ArmList "$($c.id)/providers/Microsoft.Insights/diagnosticSettings?api-version=2021-05-01-preview"
                    $diagRows.Add([pscustomobject]@{ Resource = $c.name; Type = $c.type; Configured = ($ds.Count -gt 0)
                            Destinations = (($ds | ForEach-Object { @($_.properties.workspaceId, $_.properties.storageAccountId, $_.properties.eventHubAuthorizationRuleId) | Where-Object { $_ } | ForEach-Object { ($_ -split '/')[-1] } }) -join '; ')
                            Subscription = Get-SubName $c.subscriptionId; ResourceGroup = $c.resourceGroup })
                }
                catch { $diagRows.Add([pscustomobject]@{ Resource = $c.name; Type = $c.type; Configured = $null; Destinations = 'Not readable'; Subscription = Get-SubName $c.subscriptionId; ResourceGroup = $c.resourceGroup }) }
            }
        }
        Export-AssessmentCsv $diagRows 'Azure_DiagnosticSettings.csv' @('Resource', 'Type', 'Configured', 'Destinations', 'Subscription', 'ResourceGroup')

        # Alerting
        $al = Invoke-ArgQuery -Name 'Alert rules and action groups' -Query @'
resources
| where type in~ ('microsoft.insights/actiongroups','microsoft.insights/metricalerts','microsoft.insights/activitylogalerts','microsoft.insights/scheduledqueryrules')
| extend isServiceHealth = iff(type =~ 'microsoft.insights/activitylogalerts' and tostring(properties.condition) has 'ServiceHealth', 1, 0)
| summarize Count = count(), ServiceHealth = sum(isServiceHealth) by type, subscriptionId
'@
        $alertRows = @($al | ForEach-Object { [pscustomobject]@{ Subscription = Get-SubName $_.subscriptionId; Type = $_.type; Count = [int]$_.Count; ServiceHealthAlerts = [int]$_.ServiceHealth } })
        Export-AssessmentCsv $alertRows 'Azure_Alerting.csv' @('Subscription', 'Type', 'Count', 'ServiceHealthAlerts')
        $svcHealth = [int](($alertRows | Measure-Object ServiceHealthAlerts -Sum).Sum)
        $actionGroups = [int](($alertRows | Where-Object { $_.Type -eq 'microsoft.insights/actiongroups' } | Measure-Object Count -Sum).Sum)

        $running   = @($script:Vms | Where-Object { $_.PowerState -eq 'running' })
        $noAma     = @($running | Where-Object { -not $_.AMA })
        $legacy    = @($script:Vms | Where-Object { $_.LegacyAgent }) + @()
        $legacyArc = @($ext | Where-Object { $_.machineId -match '/microsoft\.hybridcompute/' -and $_.extType -match '^(MicrosoftMonitoringAgent|OmsAgentForLinux)$' })
        $noMde     = @($running | Where-Object { -not $_.MDE })
        $notAssessed = @($running | Where-Object { -not $_.LastPatchAssess })
        $missingPatches = @($script:Vms | Where-Object { ([int]$_.MissingCritical + [int]$_.MissingSecurity) -gt 0 })
        $notExported = @($alRows | Where-Object { $_.Exported -eq $false })
        $diagChecked = @($diagRows | Where-Object { $null -ne $_.Configured })
        $diagPct = if ($diagChecked.Count) { [math]::Round(@($diagChecked | Where-Object { $_.Configured }).Count / $diagChecked.Count * 100, 1) } else { $null }

        $script:Data['LogAnalyticsWorkspaces']  = $lawRows.Count
        $script:Data['SentinelWorkspaces']      = @($lawRows | Where-Object { $_.Sentinel }).Count
        $script:Data['RunningVMsWithoutAMA']    = $noAma.Count
        $script:Data['MachinesWithLegacyAgent'] = $legacy.Count + $legacyArc.Count
        $script:Data['VMsMissingCriticalOrSecurityPatches'] = $missingPatches.Count
        $script:Data['SubscriptionsActivityLogNotExported'] = $notExported.Count
        $script:Data['DiagnosticSettingsCoveragePct'] = $diagPct
        $script:Data['ServiceHealthAlerts']     = $svcHealth
        Write-Host "  Workspaces: $($lawRows.Count) | Running VMs without AMA: $($noAma.Count) | Legacy agent: $($legacy.Count + $legacyArc.Count) | Diag coverage: $(if ($null -ne $diagPct) { "$diagPct%" } else { 'N/A' })" -ForegroundColor Green

        if ($lawRows.Count -eq 0 -and (Test-QueryOk 'Log Analytics workspaces') -and @($script:SubRows).Count -gt 0) {
            Add-Finding -Id 'AZ-OPS-001' -Severity Medium -Title 'No Log Analytics workspace in the estate' -Count 0 `
                -Recommendation 'Create a central workspace for platform and security logs (activity log, diagnostics, VM insights, Sentinel).' -EvidenceFile 'Azure_LogAnalytics.csv'
        }
        if ($noAma.Count -gt 0) {
            Add-Finding -Id 'AZ-OPS-002' -Severity Medium -Title 'Running VMs without the Azure Monitor Agent' -Count $noAma.Count `
                -Recommendation 'Deploy AMA with data collection rules (via Policy) so VM logs and performance data are collected.' -EvidenceFile 'Azure_VMs.csv'
        }
        if (($legacy.Count + $legacyArc.Count) -gt 0) {
            Add-Finding -Id 'AZ-OPS-003' -Severity Medium -Title 'Legacy Log Analytics agent (MMA/OMS) still installed - retired Aug 2024' -Count ($legacy.Count + $legacyArc.Count) `
                -Recommendation 'Migrate to AMA and remove the legacy agent; it no longer receives support or security fixes.' -EvidenceFile 'Azure_VMs.csv'
        }
        if ($missingPatches.Count -gt 0) {
            Add-Finding -Id 'AZ-OPS-004' -Severity High -Title 'VMs missing critical or security updates (Update Manager assessment)' -Count $missingPatches.Count `
                -Recommendation 'Schedule patching through Azure Update Manager maintenance configurations.' -EvidenceFile 'Azure_VMs.csv'
        }
        if ($running.Count -gt 0 -and $notAssessed.Count -gt 0) {
            Add-Finding -Id 'AZ-OPS-005' -Severity Low -Title 'Running VMs with no Update Manager patch assessment' -Count $notAssessed.Count `
                -Recommendation 'Enable periodic assessment (Policy: "Configure periodic checking for missing system updates").' -EvidenceFile 'Azure_VMs.csv'
        }
        if ($notExported.Count -gt 0) {
            Add-Finding -Id 'AZ-OPS-006' -Severity Medium -Title 'Subscriptions whose activity log is not exported' -Count $notExported.Count `
                -Recommendation 'Send the activity log to Log Analytics; the portal only keeps 90 days and it is the main audit trail for control-plane changes.' -EvidenceFile 'Azure_ActivityLogExport.csv'
        }
        if ($null -ne $diagPct -and $diagPct -lt 70) {
            Add-Finding -Id 'AZ-OPS-007' -Severity Medium -Title "Diagnostic settings on only $diagPct% of key resources checked" -Count @($diagChecked | Where-Object { -not $_.Configured }).Count `
                -Recommendation 'Use the built-in "Enable allLogs category group" policy initiative to send resource logs to Log Analytics.' -EvidenceFile 'Azure_DiagnosticSettings.csv'
        }
        if ($svcHealth -eq 0 -and (Test-QueryOk 'Alert rules and action groups') -and @($script:SubRows).Count -gt 0) {
            Add-Finding -Id 'AZ-OPS-008' -Severity Medium -Title 'No Service Health alerts configured' -Count 0 `
                -Recommendation 'Create Service Health alerts (outages, planned maintenance, health advisories, security advisories) to an action group.' -EvidenceFile 'Azure_Alerting.csv'
        }
        if ($noMde.Count -gt 0) {
            Add-Finding -Id 'AZ-OPS-009' -Severity Info -Title 'Running VMs without the Defender for Endpoint extension' -Count $noMde.Count `
                -Recommendation 'Check whether MDE is deployed another way (Intune, GPO) or enable it through Defender for Servers.' -EvidenceFile 'Azure_VMs.csv' -Confidence 'inferred'
        }

        $agentTable = @(
            [pscustomobject]@{ Check = 'Running VMs'; Count = $running.Count }
            [pscustomobject]@{ Check = 'With Azure Monitor Agent'; Count = $running.Count - $noAma.Count }
            [pscustomobject]@{ Check = 'With legacy MMA/OMS agent'; Count = $legacy.Count }
            [pscustomobject]@{ Check = 'With Defender for Endpoint extension'; Count = $running.Count - $noMde.Count }
            [pscustomobject]@{ Check = 'With an Update Manager assessment'; Count = $running.Count - $notAssessed.Count }
            [pscustomobject]@{ Check = 'Missing critical/security updates'; Count = $missingPatches.Count }
            [pscustomobject]@{ Check = 'Arc-enabled servers'; Count = $script:ArcRows.Count }
        )
        $html = "<div class='stat-row'>" +
            (New-StatCard 'Log Analytics workspaces' $lawRows.Count $(if ($lawRows.Count) { 'info' } else { 'bad' })) +
            (New-StatCard 'Sentinel workspaces' $script:Data['SentinelWorkspaces'] 'neutral') +
            (New-StatCard 'Diagnostic settings coverage' $(if ($null -ne $diagPct) { "$diagPct%" } else { 'N/A' }) $(if ($null -eq $diagPct) { 'neutral' } elseif ($diagPct -ge 70) { 'good' } else { 'warn' })) +
            (New-StatCard 'Activity log not exported' $notExported.Count $(if ($notExported.Count) { 'warn' } else { 'good' })) +
            (New-StatCard 'Service Health alerts' $svcHealth $(if ($svcHealth) { 'good' } else { 'warn' })) +
            (New-StatCard 'Action groups' $actionGroups 'neutral') +
            "</div><div class='two-col'><div><h3>Agents and patching</h3>" + (ConvertTo-HtmlTable $agentTable @('Check', 'Count')) +
            "</div><div><h3>Log Analytics workspaces</h3>" +
            (ConvertTo-HtmlTable $lawRows @('Workspace', 'RetentionDays', 'Sentinel', 'Subscription') -MaxRows 10 -CsvName 'Azure_LogAnalytics.csv') + "</div></div>" +
            "<h3>Activity log export</h3>" + (ConvertTo-HtmlTable $alRows @('Subscription', 'Exported', 'Destinations') -CsvName 'Azure_ActivityLogExport.csv')
        Add-HtmlSection $title $html
    }
    catch { Add-ErrorSection $title $_ }
}

# ═════════════════════════════════════════════════════════════════════════════
# 8. Lifecycle & Modernisation
# ═════════════════════════════════════════════════════════════════════════════
function Get-Lifecycle {
    param([int]$Total)
    $title = 'Lifecycle & Modernisation'
    Write-SectionHeader $title $Total
    try {
        # Operating systems (VMs + Arc)
        $osAll = @($script:Vms | Select-Object @{N = 'Machine'; E = { $_.Name } }, @{N = 'Kind'; E = { 'Azure VM' } }, OsLabel, OsLifecycle, OsSeverity, Subscription) +
                 @($script:ArcRows | Select-Object @{N = 'Machine'; E = { $_.Machine } }, @{N = 'Kind'; E = { 'Arc server' } }, OsLabel, OsLifecycle, OsSeverity, Subscription)
        $osSummary = @($osAll | Group-Object OsLabel, OsLifecycle | Sort-Object { Get-SeverityRank $_.Group[0].OsSeverity }, @{E = 'Count'; Descending = $true } | ForEach-Object {
                [pscustomobject]@{ OperatingSystem = $_.Group[0].OsLabel; Lifecycle = $_.Group[0].OsLifecycle; Severity = $_.Group[0].OsSeverity; Machines = $_.Count } })
        $eol = @($osAll | Where-Object { $_.OsSeverity -eq 'High' })
        $nearEol = @($osAll | Where-Object { $_.OsSeverity -eq 'Medium' })
        $unknownOs = @($osAll | Where-Object { $_.OsLifecycle -like 'Unknown*' })

        # VM size generations
        $gen = @($script:Vms | Group-Object SizeGeneration | Sort-Object Count -Descending | Select-Object @{N = 'Generation'; E = { $_.Name } }, Count)
        $oldGen = @($script:Vms | Where-Object { $_.SizeGeneration -match '^(Previous|Retired)' })

        # Retired / retiring features
        $retRows = [System.Collections.Generic.List[object]]::new()
        foreach ($r in $script:Retirements) {
            $hits = Invoke-ArgQuery -Name $r.Label -Query $r.Query
            foreach ($h in $hits) {
                $retRows.Add([pscustomobject]@{ Item = $r.Label; Status = $r.Status; When = $r.When; Resource = $h.name; Type = $h.type
                        Subscription = Get-SubName $h.subscriptionId; ResourceGroup = $h.resourceGroup; Location = $h.location; ResourceId = $h.id })
            }
        }
        Export-AssessmentCsv $retRows 'Azure_Lifecycle_Retirements.csv' @('Item', 'Status', 'When', 'Resource', 'Type', 'Subscription', 'ResourceGroup', 'Location', 'ResourceId')
        $retSummary = @($retRows | Group-Object Item | ForEach-Object { [pscustomobject]@{ Item = $_.Name; Status = $_.Group[0].Status; When = $_.Group[0].When; Resources = $_.Count } } |
            Sort-Object @{E = { if ($_.Status -eq 'Retired') { 0 } else { 1 } } }, @{E = 'Resources'; Descending = $true })

        # App Service runtimes
        $sites = Invoke-ArgQuery -Name 'App Service apps' -Query @'
resources
| where type =~ 'microsoft.web/sites'
| project id, name, kind, resourceGroup, subscriptionId, location, state = tostring(properties.state), httpsOnly = tostring(properties.httpsOnly),
          plan = tostring(split(tostring(properties.serverFarmId), '/')[8])
'@
        $rt = Invoke-ArgQuery -Name 'App Service runtimes' -Query @'
resources
| where type =~ 'microsoft.web/sites'
| mv-expand sp = properties.siteProperties.properties
| where tostring(sp.name) in~ ('LinuxFxVersion','WindowsFxVersion') and isnotempty(tostring(sp.value))
| project id, runtime = tostring(sp.value)
'@
        $rtById = @{}; foreach ($x in $rt) { $rtById["$($x.id)".ToLower()] = $x.runtime }
        $siteRows = @($sites | ForEach-Object {
                $runtime = $rtById["$($_.id)".ToLower()]
                $outdated = [bool]($runtime -and ($script:OutdatedRuntimePatterns | Where-Object { $runtime -match $_ }))
                [pscustomobject]@{ App = $_.name; Kind = $_.kind; State = $_.state; Runtime = $runtime; OutdatedRuntime = $outdated
                    HttpsOnly = $_.httpsOnly; Plan = $_.plan; Subscription = Get-SubName $_.subscriptionId; ResourceGroup = $_.resourceGroup; Location = $_.location }
            })
        Export-AssessmentCsv $siteRows 'Azure_AppServices.csv' @('App', 'Kind', 'State', 'Runtime', 'OutdatedRuntime', 'HttpsOnly', 'Plan', 'Subscription', 'ResourceGroup', 'Location')
        $oldRuntime = @($siteRows | Where-Object { $_.OutdatedRuntime })
        $noHttps = @($siteRows | Where-Object { $_.HttpsOnly -eq 'false' })

        # AKS versions
        $aks = Invoke-ArgQuery -Name 'AKS clusters' -Query @'
resources
| where type =~ 'microsoft.containerservice/managedclusters'
| project id, name, resourceGroup, subscriptionId, location, version = tostring(properties.kubernetesVersion), tier = tostring(sku.tier),
          nodePools = array_length(properties.agentPoolProfiles), powerState = tostring(properties.powerState.code)
'@
        $supported = @{}
        $aksRows = @($aks | ForEach-Object {
                $key = "$($_.subscriptionId)|$($_.location)"
                if (!$supported.ContainsKey($key)) {
                    try {
                        $kv = Invoke-ArmRequest -Path "/subscriptions/$($_.subscriptionId)/providers/Microsoft.ContainerService/locations/$($_.location)/kubernetesVersions?api-version=2024-02-01"
                        $supported[$key] = @($kv.values | Where-Object { -not $_.isPreview } | ForEach-Object { "$($_.version)" })
                    }
                    catch { $supported[$key] = $null }
                }
                $minor = ("$($_.version)" -split '\.')[0..1] -join '.'
                $status = if ($null -eq $supported[$key]) { 'Unknown' } elseif ($supported[$key] -contains $minor) { 'Supported' } else { 'Out of support' }
                [pscustomobject]@{ Cluster = $_.name; Version = $_.version; Status = $status; Tier = $_.tier; NodePools = $_.nodePools; PowerState = $_.powerState
                    Subscription = Get-SubName $_.subscriptionId; ResourceGroup = $_.resourceGroup; Location = $_.location }
            })
        Export-AssessmentCsv $aksRows 'Azure_AKS.csv' @('Cluster', 'Version', 'Status', 'Tier', 'NodePools', 'PowerState', 'Subscription', 'ResourceGroup', 'Location')
        $aksOld = @($aksRows | Where-Object { $_.Status -eq 'Out of support' })

        $script:Data['MachinesOnEndOfSupportOs'] = $eol.Count
        $script:Data['MachinesOnNearEndOfSupportOs'] = $nearEol.Count
        $script:Data['VMsOnPreviousGenerationSizes'] = $oldGen.Count
        $script:Data['RetiredFeatureResources'] = @($retRows | Where-Object { $_.Status -eq 'Retired' }).Count
        $script:Data['AppsOnOutdatedRuntime'] = $oldRuntime.Count
        $script:Data['AksClustersOutOfSupport'] = $aksOld.Count
        Write-Host "  End-of-support OS: $($eol.Count) | Retired-feature resources: $($script:Data['RetiredFeatureResources']) | Outdated runtimes: $($oldRuntime.Count) | AKS out of support: $($aksOld.Count)" -ForegroundColor Green

        if ($eol.Count -gt 0) {
            Add-Finding -Id 'AZ-LCM-001' -Severity High -Title 'Machines on end-of-support operating systems' -Count $eol.Count `
                -Detail (($eol | Group-Object OsLabel | ForEach-Object { "$($_.Name): $($_.Count)" }) -join '; ') `
                -Recommendation 'Upgrade or rebuild; where that is not possible, confirm Extended Security Updates coverage (free for some versions on Azure and via Azure Arc).' -EvidenceFile 'Azure_VMs.csv'
        }
        if ($nearEol.Count -gt 0) {
            Add-Finding -Id 'AZ-LCM-002' -Severity Medium -Title 'Machines on operating systems close to end of support' -Count $nearEol.Count `
                -Detail (($nearEol | Group-Object OsLabel | ForEach-Object { "$($_.Name): $($_.Count)" }) -join '; ') `
                -Recommendation 'Add these to the upgrade roadmap now (Windows Server 2016 extended support ends Jan 2027).' -EvidenceFile 'Azure_VMs.csv'
        }
        if ($unknownOs.Count -gt 0) {
            Add-NotCollected -Item 'Operating system version' -Reason "$($unknownOs.Count) machine(s) did not report a recognisable OS (VM agent not running, custom image or marketplace image without version)." -HowToFix 'Confirm OS versions with the customer or from ARI / the VM itself.'
        }
        foreach ($rs in $retSummary) {
            $rIdx = [array]::IndexOf(@($script:Retirements | ForEach-Object Label), $rs.Item) + 1
            Add-Finding -Id ('AZ-LCM-R{0:D2}' -f $rIdx) `
                -Severity $(if ($rs.Status -eq 'Retired') { 'High' } else { 'Medium' }) -Title "$($rs.Item) ($($rs.Status): $($rs.When))" -Count $rs.Resources `
                -Recommendation 'Migrate to the supported replacement; retired resources lose support and may stop working.' -EvidenceFile 'Azure_Lifecycle_Retirements.csv'
        }
        if ($oldGen.Count -gt 0) {
            Add-Finding -Id 'AZ-LCM-003' -Severity Low -Title 'VMs on previous-generation or retired sizes' -Count $oldGen.Count `
                -Recommendation 'Move to current v5/v6 sizes during the next maintenance window; usually better price-performance.' -EvidenceFile 'Azure_VMs.csv'
        }
        if ($oldRuntime.Count -gt 0) {
            Add-Finding -Id 'AZ-LCM-004' -Severity Medium -Title 'App Service apps on out-of-support language runtimes' -Count $oldRuntime.Count `
                -Recommendation 'Upgrade the runtime stack; end-of-life runtimes stop receiving security patches.' -EvidenceFile 'Azure_AppServices.csv'
        }
        if ($noHttps.Count -gt 0) {
            Add-Finding -Id 'AZ-LCM-005' -Severity Low -Title 'App Service apps not enforcing HTTPS only' -Count $noHttps.Count `
                -Recommendation 'Turn on HTTPS Only.' -EvidenceFile 'Azure_AppServices.csv'
        }
        if ($aksOld.Count -gt 0) {
            Add-Finding -Id 'AZ-LCM-006' -Severity High -Title 'AKS clusters on unsupported Kubernetes versions' -Count $aksOld.Count `
                -Recommendation 'Upgrade to a supported minor version and enable an auto-upgrade channel.' -EvidenceFile 'Azure_AKS.csv'
        }

        $html = "<div class='stat-row'>" +
            (New-StatCard 'End-of-support OS' $eol.Count $(if ($eol.Count) { 'bad' } else { 'good' })) +
            (New-StatCard 'OS near end of support' $nearEol.Count $(if ($nearEol.Count) { 'warn' } else { 'good' })) +
            (New-StatCard 'Retired-feature resources' $script:Data['RetiredFeatureResources'] $(if ($script:Data['RetiredFeatureResources']) { 'bad' } else { 'good' })) +
            (New-StatCard 'Previous-gen VM sizes' $oldGen.Count $(if ($oldGen.Count) { 'warn' } else { 'good' })) +
            (New-StatCard 'Outdated app runtimes' $oldRuntime.Count $(if ($oldRuntime.Count) { 'warn' } else { 'good' })) +
            (New-StatCard 'AKS out of support' $aksOld.Count $(if ($aksOld.Count) { 'bad' } else { 'good' })) +
            "</div><h3>Operating systems</h3>" +
            (ConvertTo-HtmlTable $osSummary @('Severity', 'OperatingSystem', 'Lifecycle', 'Machines') -MaxRows 30 -CsvName 'Azure_VMs.csv') +
            "<h3>Retired and retiring platform features</h3>" +
            (ConvertTo-HtmlTable $retSummary @('Item', 'Status', 'When', 'Resources') -CsvName 'Azure_Lifecycle_Retirements.csv') +
            "<div class='two-col'><div><h3>VM size generations</h3>" + (ConvertTo-HtmlTable $gen @('Generation', 'Count')) +
            "</div><div><h3>AKS clusters</h3>" + (ConvertTo-HtmlTable $aksRows @('Cluster', 'Version', 'Status') -MaxRows 10 -CsvName 'Azure_AKS.csv') + "</div></div>" +
            $(if ($oldRuntime.Count) { "<h3>Apps on outdated runtimes</h3>" + (ConvertTo-HtmlTable $oldRuntime @('App', 'Runtime', 'Subscription') -CsvName 'Azure_AppServices.csv') } else { '' }) +
            "<p class='note'>OS versions come from the VM agent where it reports, otherwise from the marketplace image. Lifecycle dates are maintained in the configuration block at the top of the script.</p>"
        Add-HtmlSection $title $html
    }
    catch { Add-ErrorSection $title $_ }
}

# ═════════════════════════════════════════════════════════════════════════════
# Final exports
# ═════════════════════════════════════════════════════════════════════════════
function Export-FinalData {
    Write-Host "`nWriting evidence files..." -ForegroundColor Cyan
    Export-AssessmentCsv $script:Vms 'Azure_VMs.csv' @('Name', 'Subscription', 'SubscriptionId', 'ResourceGroup', 'Location', 'Size', 'SizeGeneration',
        'OsType', 'OsLabel', 'OsLifecycle', 'OsSeverity', 'PowerState', 'Placement', 'ManagedDisks', 'PatchMode', 'LicenseType', 'Priority', 'PublicIp',
        'AzureBackup', 'LastBackup', 'AsrReplicated', 'AMA', 'LegacyAgent', 'MDE', 'LastPatchAssess', 'MissingCritical', 'MissingSecurity', 'ResourceId')

    $sorted = @($script:Findings | Sort-Object @{E = { Get-SeverityRank $_.SuggestedSeverity } }, Section, Id)
    Export-AssessmentCsv $sorted 'Findings.csv' @('Id', 'Section', 'SuggestedSeverity', 'Title', 'Count', 'Detail', 'Recommendation', 'EvidenceFile', 'Confidence',
        'CafPhase', 'CafDesignArea', 'WafPillar', 'WafCodes', 'ReferencedTo', 'References')
    Export-AssessmentCsv $script:FrameworkCoverage 'Framework_Coverage.csv' @('Framework', 'Area', 'AssessedBy', 'Notes')
    $byPillar = [ordered]@{}
    foreach ($k in $script:WafPillars.Keys) { $n = $script:WafPillars[$k].Name; $byPillar[$n] = @($script:Findings | Where-Object { $_.WafPillar -split '; ' -contains $n }).Count }
    $byArea = [ordered]@{}
    foreach ($a in @($script:CafDesignAreaUrl.Keys) + 'Workload') { $byArea[$a] = @($script:Findings | Where-Object { $_.CafDesignArea -eq $a }).Count }
    $byPhase = [ordered]@{}
    foreach ($ph in 'Ready', 'Adopt', 'Govern', 'Secure', 'Manage') { $byPhase[$ph] = @($script:Findings | Where-Object { $_.CafPhase -eq $ph }).Count }
    Export-AssessmentCsv $script:NotCollected 'NotCollected.csv' @('Section', 'Item', 'Reason', 'HowToFix')

    $bySev = [ordered]@{}
    foreach ($s in 'Critical', 'High', 'Medium', 'Low', 'Info') { $bySev[$s] = @($script:Findings | Where-Object { $_.SuggestedSeverity -eq $s }).Count }
    $summary = [ordered]@{
        tool          = 'Azure-Tenant-Assessment'
        version       = $script:ToolVersion
        generatedUtc  = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
        customer      = $CustomerName
        tenantId      = $script:TenantIdUsed
        account       = $script:Account
        mode          = $(if ($QuickMode) { 'Quick' } else { 'Full' })
        scope         = $(if ($script:ScopeSubs.Count) { @($script:ScopeSubs) } else { 'All readable subscriptions' })
        costDays      = $(if ($SkipCost) { $null } else { $CostDays })
        metrics       = $script:Data
        findings      = $bySev
        frameworks    = [ordered]@{ wafPillar = $byPillar; cafDesignArea = $byArea; cafPhase = $byPhase; unmapped = @($script:Findings | Where-Object { $_.CafPhase -eq 'Unmapped' } | ForEach-Object Id) }
        notCollected  = $script:NotCollected.Count
        files         = @(Get-ChildItem -Path $script:OutputPath -File | Where-Object { $_.Name -ne 'Assessment_Summary.json' } | ForEach-Object Name) + 'Assessment_Summary.json'
    }
    $summary | ConvertTo-Json -Depth 6 | Out-File -FilePath (Join-Path $script:OutputPath 'Assessment_Summary.json') -Encoding utf8
}

# ═════════════════════════════════════════════════════════════════════════════
# HTML report
# ═════════════════════════════════════════════════════════════════════════════
function New-HtmlReport {
    Write-Host "Generating HTML report..." -ForegroundColor Cyan
    $d = $script:Data
    $customer   = if ($CustomerName) { $CustomerName } else { 'Azure estate' }
    $generated  = Get-Date -Format 'yyyy-MM-dd HH:mm'
    $scopeText  = if ($script:ScopeSubs.Count) { "$($script:ScopeSubs.Count) selected subscription(s)" } else { "$($d['Subscriptions']) subscription(s)" }
    $modeText   = if ($QuickMode) { 'Quick assessment' } else { 'Full assessment' }
    $fmt = { param($v, $suffix = '') if ($null -eq $v -or "$v" -eq '') { 'N/A' } else { "$v$suffix" } }

    $tone = {
        param($v, $goodMax, $warnMax)
        if ($null -eq $v -or "$v" -eq '') { return 'neutral' }
        if ([double]$v -le $goodMax) { return 'good' } elseif ([double]$v -le $warnMax) { return 'warn' } else { return 'bad' }
    }
    $pctTone = {
        param($v)
        if ($null -eq $v -or "$v" -eq '') { return 'neutral' }
        if ([double]$v -ge 80) { return 'good' } elseif ([double]$v -ge 50) { return 'warn' } else { return 'bad' }
    }
    $spend = if ($null -ne $d['CostTotal']) { '{0:N0} {1}' -f $d['CostTotal'], $d['CostCurrency'] } elseif ($d['CostCurrency'] -eq 'Mixed') { 'Mixed currencies' } else { 'N/A' }
    $savings = if ($null -ne $d['AdvisorAnnualSavings']) { '{0:N0} {1}' -f $d['AdvisorAnnualSavings'], $d['AdvisorSavingsCurrency'] } else { 'N/A' }

    $cards = "<div class='stat-row'>" +
        (New-StatCard 'Subscriptions' (& $fmt $d['Subscriptions'])) +
        (New-StatCard 'Resources' (& $fmt $d['Resources'])) +
        (New-StatCard 'Virtual machines' (& $fmt $d['VirtualMachines'])) +
        (New-StatCard 'Secure score' (& $fmt $d['SecureScorePct'] '%') (& $pctTone $d['SecureScorePct'])) +
        (New-StatCard 'Policy compliance' (& $fmt $d['PolicyCompliancePct'] '%') (& $pctTone $d['PolicyCompliancePct'])) +
        "</div><div class='stat-row'>" +
        (New-StatCard 'Internet-exposed mgmt/DB rules' (& $fmt $d['InternetExposedSensitiveRules']) (& $tone $d['InternetExposedSensitiveRules'] 0 0)) +
        (New-StatCard 'Running VMs without Azure Backup' (& $fmt $d['RunningVMsWithoutBackup']) (& $tone $d['RunningVMsWithoutBackup'] 0 2)) +
        (New-StatCard 'End-of-support OS' (& $fmt $d['MachinesOnEndOfSupportOs']) (& $tone $d['MachinesOnEndOfSupportOs'] 0 0)) +
        (New-StatCard "Spend (last $(& $fmt $d['CostPeriodDays']) days)" $spend) +
        (New-StatCard 'Advisor savings / year' $savings 'warn') +
        "</div>"

    $keyFindings = @($script:Findings | Where-Object { $_.SuggestedSeverity -ne 'Info' } | Sort-Object @{E = { Get-SeverityRank $_.SuggestedSeverity } }, Section) |
        Select-Object @{N = 'Severity'; E = { $_.SuggestedSeverity } }, Section, Title, Count, @{N = 'CAF'; E = { if ($_.CafDesignArea -eq 'Workload') { "$($_.CafPhase) (workload)" } else { "$($_.CafPhase): $($_.CafDesignArea)" } } }, @{N = 'WAF'; E = { $_.WafCodes } }

    # Alignment overview: findings by WAF pillar and by CAF design area, split by severity
    $sevCols = 'High', 'Medium', 'Low'
    $pillarRows = foreach ($k in $script:WafPillars.Keys) {
        $n = $script:WafPillars[$k].Name
        $f = @($script:Findings | Where-Object { $_.WafPillar -split '; ' -contains $n })
        $o = [ordered]@{ 'WAF pillar' = $n }
        foreach ($s in $sevCols) { $o[$s] = @($f | Where-Object { $_.SuggestedSeverity -eq $s -or ($s -eq 'High' -and $_.SuggestedSeverity -eq 'Critical') }).Count }
        $o['Assessed by'] = ($script:FrameworkCoverage | Where-Object { $_.Framework -eq 'WAF pillar' -and $_.Area -eq $n }).AssessedBy
        [pscustomobject]$o
    }
    $areaRows = foreach ($a in @($script:CafDesignAreaUrl.Keys) + 'Workload') {
        $f = @($script:Findings | Where-Object { $_.CafDesignArea -eq $a })
        $o = [ordered]@{ 'CAF design area' = $(if ($a -eq 'Workload') { 'Workload level (WAF only)' } else { $a }) }
        foreach ($s in $sevCols) { $o[$s] = @($f | Where-Object { $_.SuggestedSeverity -eq $s -or ($s -eq 'High' -and $_.SuggestedSeverity -eq 'Critical') }).Count }
        $o['Assessed by'] = $(if ($a -eq 'Workload') { 'Script' } else { ($script:FrameworkCoverage | Where-Object { $_.Framework -eq 'CAF design area' -and $_.Area -eq $a }).AssessedBy })
        [pscustomobject]$o
    }
    $alignHtml = "<h3>Referenced to the Cloud Adoption Framework and Well-Architected Framework</h3>" +
        "<div class='two-col'><div>" + (ConvertTo-HtmlTable $areaRows -MaxRows 20) + "</div><div>" + (ConvertTo-HtmlTable $pillarRows -MaxRows 10) + "</div></div>" +
        "<p class='note'>Each finding is referenced to a CAF phase, a CAF Azure landing zone design area and one or more WAF design review checklist recommendations. Areas marked workshop are covered in the customer workshop; they cannot be read from the Azure APIs. See the CAF and WAF alignment appendix.</p>"

    $sevCounts = foreach ($s in 'Critical', 'High', 'Medium', 'Low') {
        $n = @($script:Findings | Where-Object { $_.SuggestedSeverity -eq $s }).Count
        if ($n) { "<span class='sev sev-$s'>$n $s</span>" }
    }
    $summaryHtml = "<section class='section summary' id='sec-summary'><h2>Executive summary</h2>" +
        "<p class='lead'>Read-only, point-in-time assessment of $(ConvertTo-HtmlEncoded $scopeText) in tenant <code>$(ConvertTo-HtmlEncoded $script:TenantIdUsed)</code>. " +
        "Findings carry a suggested severity; confirm context with the customer before they go into a report.</p>" +
        $(if (-not $d['Subscriptions']) { "<p class='callout'><strong>No subscriptions were readable.</strong> The account needs Reader on the subscriptions or the tenant root management group. See 'What could not be collected'.</p>" } else { '' }) +
        $cards +
        "<h3>Key findings $($sevCounts -join ' ')</h3>" +
        (ConvertTo-HtmlTable $keyFindings @('Severity', 'Section', 'Title', 'Count', 'CAF', 'WAF') -MaxRows 60 -CsvName 'Findings.csv') +
        $alignHtml +
        "</section>"

    $ncHtml = if ($script:NotCollected.Count) {
        "<section class='section warn' id='sec-not-collected'><h2>What could not be collected</h2>" +
        "<p>These gaps should become items on the customer information request.</p>" +
        (ConvertTo-HtmlTable $script:NotCollected @('Section', 'Item', 'Reason', 'HowToFix') -MaxRows 100) + "</section>"
    }
    else { "<section class='section good' id='sec-not-collected'><h2>What could not be collected</h2><p>Everything in scope was readable.</p></section>" }

    # Appendix: every finding with its CAF and WAF references, linked to Microsoft Learn
    $apx = [System.Text.StringBuilder]::new()
    [void]$apx.Append("<section class='section summary' id='sec-alignment'><h2>Appendix: CAF and WAF alignment</h2>")
    [void]$apx.Append("<p>Findings are referenced to the Microsoft Cloud Adoption Framework (CAF) and Azure Well-Architected Framework (WAF). A reference shows which guidance the recommendation follows; it is not a statement of compliance.</p>")
    [void]$apx.Append("<table><thead><tr><th>ID</th><th>Severity</th><th>Finding</th><th>CAF phase</th><th>CAF design area</th><th>WAF recommendations</th></tr></thead><tbody>")
    foreach ($f in @($script:Findings | Sort-Object @{E = { Get-SeverityRank $_.SuggestedSeverity } }, Id)) {
        $areaCell = if ($f.CafDesignArea -ne 'Workload' -and $script:CafDesignAreaUrl.Contains($f.CafDesignArea)) {
            "<a href='https://learn.microsoft.com/azure/cloud-adoption-framework/ready/landing-zone/design-area/$($script:CafDesignAreaUrl[$f.CafDesignArea])'>$(ConvertTo-HtmlEncoded $f.CafDesignArea)</a>" } else { 'Workload level' }
        $phaseCell = if ($script:CafPhaseUrl.ContainsKey($f.CafPhase)) { "<a href='https://learn.microsoft.com/azure/cloud-adoption-framework/$($script:CafPhaseUrl[$f.CafPhase])/'>$(ConvertTo-HtmlEncoded $f.CafPhase)</a>" } else { ConvertTo-HtmlEncoded $f.CafPhase }
        $wafCell = (@($f.WafCodes -split '; ' | Where-Object { $_ }) | ForEach-Object { "<a href='$(Get-WafUrl $_)'>$_</a> $(ConvertTo-HtmlEncoded $script:WafCodes[$_][1])" }) -join '<br>'
        if (!$wafCell) { $wafCell = '<span class=''no-data''>Platform only (CAF)</span>' }
        [void]$apx.Append("<tr><td class='nowrap'>$(ConvertTo-HtmlEncoded $f.Id)</td><td><span class='sev sev-$($f.SuggestedSeverity)'>$($f.SuggestedSeverity)</span></td><td>$(ConvertTo-HtmlEncoded $f.Title)</td><td>$phaseCell</td><td>$areaCell</td><td>$wafCell</td></tr>")
    }
    [void]$apx.Append("</tbody></table><h3>Coverage</h3>")
    [void]$apx.Append((ConvertTo-HtmlTable $script:FrameworkCoverage @('Framework', 'Area', 'AssessedBy', 'Notes') -MaxRows 20 -CsvName 'Framework_Coverage.csv'))
    [void]$apx.Append("</section>")
    $alignAppendix = $apx.ToString()

    $toc = "<nav class='toc no-print'><a href='#sec-summary'>Summary</a>" + ($script:TocEntries -join '') + "<a href='#sec-not-collected'>Not collected</a><a href='#sec-alignment'>CAF / WAF</a></nav>"
    $logo = if ($script:LogoBase64 -and $script:LogoBase64 -notmatch '^__') { "<img src='data:image/png;base64,$($script:LogoBase64)' alt='Netsurit'>" } else { "<strong>NETSURIT</strong>" }

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>$(ConvertTo-HtmlEncoded $customer) - Azure Assessment</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Oxygen:wght@400;700&display=swap" rel="stylesheet">
<style>
:root{--red:#F4364C;--black:#2C2A29;--grey:#6E6E6E;--line:#E1E1E1;--soft:#F7F7F7;--orange:#FF8200;--yellow:#FFB81C;--teal:#00C7B1;--blue:#0072CE}
*{box-sizing:border-box}
body{margin:0;font-family:'Oxygen',Arial,sans-serif;font-size:13px;line-height:1.45;color:var(--black);background:#EFEFEF}
.page{max-width:1180px;margin:0 auto;background:#fff;padding:28px 36px 36px}
header.brand{display:flex;align-items:flex-end;justify-content:space-between;gap:24px;border-bottom:3px solid var(--red);padding-bottom:14px}
header.brand img{height:40px;display:block}
header.brand .meta{text-align:right;color:var(--grey);font-size:12px}
h1{font-size:26px;margin:18px 0 2px}
.subtitle{color:var(--grey);margin:0 0 6px}
nav.toc{position:sticky;top:0;z-index:5;background:#fff;border-bottom:1px solid var(--line);padding:8px 0;margin:10px 0 0;display:flex;flex-wrap:wrap;gap:6px 16px;font-size:12px}
nav.toc a{color:var(--black);text-decoration:none;border-bottom:2px solid transparent}
nav.toc a:hover{border-bottom-color:var(--red)}
.section{border-left:4px solid var(--blue);padding:2px 0 6px 18px;margin:28px 0}
.section.good{border-left-color:var(--teal)}.section.warn{border-left-color:var(--orange)}.section.bad{border-left-color:var(--red)}
.section.summary{border-left-color:var(--black)}
h2{font-size:18px;margin:0 0 10px}
h3{font-size:11.5px;text-transform:uppercase;letter-spacing:.5px;color:var(--grey);margin:18px 0 6px}
.lead{margin:0 0 10px}
code{font-size:12px;background:var(--soft);padding:1px 4px;border-radius:3px}
.stat-row{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:10px;margin:10px 0}
.stat-card{border:1px solid var(--line);border-top:4px solid var(--blue);border-radius:4px;padding:9px 12px;background:var(--soft)}
.stat-card.tone-good{border-top-color:var(--teal)}.stat-card.tone-warn{border-top-color:var(--orange)}.stat-card.tone-bad{border-top-color:var(--red)}.stat-card.tone-neutral{border-top-color:#BDBDBD}
.stat-value{font-size:19px;font-weight:700}
.stat-label{font-size:11px;color:var(--grey);margin-top:2px}
.two-col{display:grid;grid-template-columns:1fr 1fr;gap:20px}
table{width:100%;border-collapse:collapse;font-size:12px;margin:4px 0}
th{background:var(--black);color:#fff;text-align:left;padding:6px 8px;font-weight:700}
td{padding:5px 8px;border-bottom:1px solid var(--line);vertical-align:top;word-break:break-word}
tbody tr:nth-child(even) td{background:#FBFBFB}
.sev{display:inline-block;padding:1px 8px;border-radius:10px;font-size:11px;font-weight:700;color:#fff;white-space:nowrap;margin-right:4px}
.sev-Critical{background:#8E0F1F}.sev-High{background:var(--red)}.sev-Medium{background:var(--orange)}.sev-Low{background:var(--yellow);color:var(--black)}.sev-Info{background:var(--blue)}.sev-Unknown{background:#9E9E9E}
.no-data{color:var(--grey);font-style:italic;margin:4px 0}
.note{color:var(--grey);font-size:11px;font-style:italic;margin:6px 0}
.error{color:var(--red)}
.nowrap{white-space:nowrap}
.align{color:var(--grey);font-size:11.5px;margin:-6px 0 8px}
td a{color:var(--blue);text-decoration:none}
.callout{background:#FFF4E5;border:1px solid var(--orange);border-radius:4px;padding:8px 12px;margin:10px 0}
footer{margin-top:32px;padding-top:12px;border-top:1px solid var(--line);color:var(--grey);font-size:11px;display:flex;justify-content:space-between;gap:16px}
footer .tag{color:var(--red);font-weight:700}
@media (max-width:760px){.page{padding:16px}.two-col{grid-template-columns:1fr}header.brand{flex-direction:column;align-items:flex-start}header.brand .meta{text-align:left}}
@page{size:A4;margin:15mm 12mm}
@media print{
  body{background:#fff;color:#000;font-size:11px}
  .page{max-width:none;padding:0}
  .no-print,nav.toc{display:none!important}
  header.brand{border-bottom:2px solid #000}
  .section{border-left-width:3px;margin:16px 0;padding-left:12px}
  h2{break-after:avoid;page-break-after:avoid}
  h3{break-after:avoid;page-break-after:avoid}
  .stat-row,.stat-card,.two-col>div{break-inside:avoid;page-break-inside:avoid}
  .stat-row{grid-template-columns:repeat(auto-fit,minmax(105px,1fr));gap:6px}
  .stat-card{background:#fff;border-color:#999;padding:6px 8px}
  .stat-value{font-size:15px}
  .stat-label,.note,.subtitle,header.brand .meta,footer{color:#000}
  table{page-break-inside:auto}
  thead{display:table-header-group}
  tr{break-inside:avoid;page-break-inside:avoid}
  th{background:#fff!important;color:#000!important;border-bottom:2px solid #000}
  tbody tr:nth-child(even) td{background:#fff}
  td{border-bottom:1px solid #BBB}
  .sev{background:#fff!important;color:#000!important;border:1px solid #000}
  code{background:none}
  a,td a{color:#000!important;text-decoration:none}
  .align{color:#000}
}
</style>
</head>
<body>
<div class="page">
<header class="brand">
  <div>$logo</div>
  <div class="meta">Netsurit Professional Services<br>$modeText &middot; $generated</div>
</header>
<h1>Azure Assessment &ndash; $(ConvertTo-HtmlEncoded $customer)</h1>
<p class="subtitle">Tenant $(ConvertTo-HtmlEncoded $script:TenantIdUsed) &middot; $(ConvertTo-HtmlEncoded $scopeText) &middot; run as $(ConvertTo-HtmlEncoded $script:Account)</p>
$toc
$summaryHtml
$($script:HtmlSections -join "`n")
$ncHtml
$alignAppendix
<footer>
  <div>Netsurit Professional Services &middot; Azure Tenant Assessment v$($script:ToolVersion) &middot; Read-only collection; evidence files are in the same folder as this report.</div>
  <div class="tag">Supporting the Dreams of the Doers.</div>
</footer>
</div>
</body>
</html>
"@
    $reportPath = Join-Path $script:OutputPath 'Azure_Assessment_Report.html'
    $html | Out-File -FilePath $reportPath -Encoding utf8
    Write-Host "  Report: $reportPath" -ForegroundColor Green
    return $reportPath
}

# ═════════════════════════════════════════════════════════════════════════════
# Main
# ═════════════════════════════════════════════════════════════════════════════
function Start-AzureTenantAssessment {
    $modeLabel = if ($QuickMode) { 'QUICK (Overview, Security, Cost & Waste)' } else { 'FULL (9 sections)' }
    $custLabel = if ($CustomerName) { $CustomerName } else { '(no -CustomerName specified)' }
    Write-Host @"
=================================================================
       AZURE TENANT ASSESSMENT TOOL  v$($script:ToolVersion)
       Netsurit Professional Services
=================================================================
Customer : $custLabel
Mode     : $modeLabel
Required : Reader (root MG or each subscription), Security Reader,
           Cost Management Reader, Directory Readers (Entra)
=================================================================
"@ -ForegroundColor Cyan

    try {
        Initialize-RequiredModules
        if (!(Connect-AzureServices)) { Write-Host "Cannot connect to Azure. Exiting." -ForegroundColor Red; return }
        $script:OutputPath = Initialize-OutputDirectory
        $script:ScopeSubs  = @($SubscriptionId | Where-Object { $_ } | ForEach-Object { $_.Trim() })

        $total = if ($QuickMode) { 3 } else { 9 }
        Write-Host "`n=== Starting data collection ($total sections) ===" -ForegroundColor Cyan
        Get-EstateOverview -Total $total
        if (!$QuickMode) {
            Get-Governance   -Total $total
            Get-IdentityRbac -Total $total
        }
        Get-SecurityPosture -Total $total
        if (!$QuickMode) { Get-NetworkTopology -Total $total; Get-Resiliency -Total $total }
        Get-CostAndWaste -Total $total
        if (!$QuickMode) {
            Get-Operations -Total $total
            Get-Lifecycle  -Total $total
        }

        Export-FinalData
        $reportPath = New-HtmlReport

        $high = @($script:Findings | Where-Object { $_.SuggestedSeverity -in 'Critical', 'High' }).Count
        Write-Host @"

=================================================================
  ASSESSMENT COMPLETE
  Findings      : $($script:Findings.Count) ($high critical/high)
  Not collected : $($script:NotCollected.Count) item(s) - see NotCollected.csv
  Output        : $($script:OutputPath)
  Report        : $reportPath
=================================================================
"@ -ForegroundColor Green

        if (!$NoOpen) {
            if ($IsWindows) { Start-Process explorer.exe $script:OutputPath; Start-Process $reportPath }
            elseif ($IsMacOS) { & open $reportPath }
        }
    }
    finally {
        [System.Threading.Thread]::CurrentThread.CurrentCulture = $script:OriginalCulture
    }
}

Start-AzureTenantAssessment
