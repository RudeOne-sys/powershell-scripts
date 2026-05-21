#Requires -Version 5.1
<#
.SYNOPSIS
    Microsoft 365 Tenant Assessment Tool v3.0
.DESCRIPTION
    Comprehensive M365 tenant assessment for existing and potential customers.
    Outputs CSV data files and a self-contained HTML executive summary report.
.PARAMETER CustomerName
    Customer name used for output folder naming and report header.
.PARAMETER TenantId
    Target tenant ID for partner/delegated access.
.PARAMETER QuickMode
    Runs only Tenant Overview, Licensing, Secure Score, and User Health.
.EXAMPLE
    .\M365TenantAssessment.ps1 -CustomerName "Contoso"
.EXAMPLE
    .\M365TenantAssessment.ps1 -CustomerName "Fabrikam" -TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -QuickMode
#>
param(
    [string]$CustomerName = "",
    [string]$TenantId = "",
    [switch]$QuickMode
)

$ErrorActionPreference = 'Continue'

# ── Script-level state ────────────────────────────────────────────────────────
$script:ExchangeConnected = $false
$script:HtmlSections      = [System.Collections.Generic.List[string]]::new()
$script:Data              = @{}
$script:SectionNum        = 0

$script:SkuMap = @{
    'O365_BUSINESS_ESSENTIALS'   = 'Microsoft 365 Business Basic'
    'O365_BUSINESS_PREMIUM'      = 'Microsoft 365 Business Standard'
    'SPB'                        = 'Microsoft 365 Business Premium'
    'O365_BUSINESS'              = 'Microsoft 365 Apps for Business'
    'OFFICESUBSCRIPTION'         = 'Microsoft 365 Apps for Enterprise'
    'STANDARDPACK'               = 'Office 365 E1'
    'ENTERPRISEPACK'             = 'Office 365 E3'
    'ENTERPRISEPREMIUM'          = 'Office 365 E5'
    'DESKLESSPACK'               = 'Office 365 F3'
    'SPE_E3'                     = 'Microsoft 365 E3'
    'SPE_E5'                     = 'Microsoft 365 E5'
    'SPE_F1'                     = 'Microsoft 365 F1'
    'SPE_F3'                     = 'Microsoft 365 F3'
    'POWER_BI_STANDARD'          = 'Power BI (Free)'
    'POWER_BI_PRO'               = 'Power BI Pro'
    'EXCHANGESTANDARD'           = 'Exchange Online Plan 1'
    'EXCHANGEENTERPRISE'         = 'Exchange Online Plan 2'
    'PROJECTPROFESSIONAL'        = 'Project Plan 3'
    'VISIOCLIENT'                = 'Visio Plan 2'
    'ATP_ENTERPRISE'             = 'Defender for Office 365 Plan 1'
    'THREAT_INTELLIGENCE'        = 'Defender for Office 365 Plan 2'
    'EMS'                        = 'EMS E3'
    'EMSPREMIUM'                 = 'EMS E5'
    'AAD_PREMIUM'                = 'Entra ID P1'
    'AAD_PREMIUM_P2'             = 'Entra ID P2'
    'RIGHTSMANAGEMENT'           = 'Azure Information Protection P1'
    'SHAREPOINTSTANDARD'         = 'SharePoint Online Plan 1'
    'SHAREPOINTENTERPRISE'       = 'SharePoint Online Plan 2'
    'MCOSTANDARD'                = 'Skype for Business Online Plan 2'
    'FLOW_FREE'                  = 'Power Automate Free'
    'TEAMS_EXPLORATORY'          = 'Teams Exploratory'
    'TEAMS_COMMERCIAL_TRIAL'     = 'Microsoft Teams Trial'
    'DEVELOPERPACK_E5'           = 'Microsoft 365 E5 Developer'
}

# ── Module initialisation ─────────────────────────────────────────────────────
function Initialize-RequiredModules {
    Write-Host "Checking required modules..." -ForegroundColor Cyan

    $graphModules = @(
        'Microsoft.Graph.Authentication',
        'Microsoft.Graph.Users',
        'Microsoft.Graph.Groups',
        'Microsoft.Graph.Sites',
        'Microsoft.Graph.Teams',
        'Microsoft.Graph.Applications',
        'Microsoft.Graph.Security',
        'Microsoft.Graph.Reports',
        'Microsoft.Graph.Identity.DirectoryManagement',
        'Microsoft.Graph.Identity.SignIns'
    )

    foreach ($module in $graphModules) {
        if (!(Get-Module -ListAvailable -Name $module)) {
            Write-Host "  Installing $module..." -ForegroundColor Yellow
            Install-Module -Name $module -Force -AllowClobber -Scope CurrentUser -ErrorAction SilentlyContinue
        }
        Import-Module $module -ErrorAction SilentlyContinue
    }

    if (Get-Module -ListAvailable -Name 'ExchangeOnlineManagement') {
        Import-Module ExchangeOnlineManagement -ErrorAction SilentlyContinue
        Write-Host "  ExchangeOnlineManagement module available" -ForegroundColor Green
    } else {
        Write-Host "  ExchangeOnlineManagement not found — shared mailbox data will use Graph heuristic" -ForegroundColor Yellow
    }
}

# ── Output directory ──────────────────────────────────────────────────────────
function Initialize-OutputDirectory {
    $timestamp     = Get-Date -Format "yyyyMMdd_HHmmss"
    $folderName    = if ($CustomerName) { "M365_Assessment_${CustomerName}_$timestamp" } else { "M365_Assessment_$timestamp" }
    $downloadsPath = [Environment]::GetFolderPath("UserProfile") + "\Downloads"
    $outputPath    = Join-Path $downloadsPath $folderName

    New-Item -ItemType Directory -Path $outputPath -Force | Out-Null
    Write-Host "Output folder: $outputPath" -ForegroundColor Green
    return $outputPath
}

# ── Graph connection ──────────────────────────────────────────────────────────
function Connect-MicrosoftGraphServices {
    Write-Host "`nConnecting to Microsoft Graph..." -ForegroundColor Cyan

    $scopes = @(
        "User.Read.All",
        "Group.Read.All",
        "Sites.Read.All",
        "TeamSettings.Read.All",
        "TeamMember.Read.All",
        "Channel.ReadBasic.All",
        "Application.Read.All",
        "SecurityEvents.Read.All",
        "Reports.Read.All",
        "Directory.Read.All",
        "MailboxSettings.Read",
        "Organization.Read.All",
        "Policy.Read.All",
        "RoleManagement.Read.Directory",
        "AuditLog.Read.All",
        "UserAuthenticationMethod.Read.All",
        "InformationProtectionPolicy.Read"
    )

    try {
        $connectParams = @{ Scopes = $scopes; NoWelcome = $true }
        if ($TenantId) { $connectParams['TenantId'] = $TenantId }
        Connect-MgGraph @connectParams

        $ctx = Get-MgContext
        Write-Host "Connected: $($ctx.Account) -> $($ctx.TenantId)" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "Failed to connect: $_" -ForegroundColor Red
        return $false
    }
}

# ── EXO connection ────────────────────────────────────────────────────────────
function Connect-ExchangeOnlineIfAvailable {
    if (!(Get-Module -Name ExchangeOnlineManagement -ErrorAction SilentlyContinue)) { return }
    Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
    try {
        $connectParams = @{ ShowBanner = $false }
        if ($TenantId) { $connectParams['DelegatedOrganization'] = $TenantId }
        Connect-ExchangeOnline @connectParams -ErrorAction Stop
        $script:ExchangeConnected = $true
        Write-Host "  Exchange Online connected" -ForegroundColor Green
    }
    catch {
        Write-Host "  Could not connect to Exchange Online: $_" -ForegroundColor Yellow
    }
}

# ── Section / HTML helpers ────────────────────────────────────────────────────
function Write-SectionHeader {
    param([string]$Title, [int]$Total)
    $script:SectionNum++
    Write-Host "`n[$($script:SectionNum)/$Total] $Title..." -ForegroundColor Cyan
}

function Add-HtmlSection {
    param([string]$Title, [string]$Content, [string]$Status = "info")
    $cls = switch ($Status) {
        'good' { 'section-good' }
        'warn' { 'section-warn' }
        'bad'  { 'section-bad'  }
        default { '' }
    }
    $script:HtmlSections.Add("<div class='section $cls'><h2>$Title</h2>$Content</div>")
}

function ConvertTo-HtmlTable {
    param([object[]]$Data, [string[]]$Properties = @())
    if (!$Data -or $Data.Count -eq 0) { return "<p class='no-data'>No data available</p>" }
    $cols = if ($Properties.Count -gt 0) { $Properties } else { $Data[0].PSObject.Properties.Name }
    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.Append('<table><thead><tr>')
    foreach ($col in $cols) { [void]$sb.Append("<th>$col</th>") }
    [void]$sb.Append('</tr></thead><tbody>')
    foreach ($row in $Data) {
        [void]$sb.Append('<tr>')
        foreach ($col in $cols) { [void]$sb.Append("<td>$($row.$col)</td>") }
        [void]$sb.Append('</tr>')
    }
    [void]$sb.Append('</tbody></table>')
    return $sb.ToString()
}

function New-StatCard {
    param([string]$Label, [string]$Value, [string]$Color = "#0078d4")
    return "<div class='stat-card' style='border-top-color:$Color'><div class='stat-value'>$Value</div><div class='stat-label'>$Label</div></div>"
}

function Get-DataValue {
    param([string]$Key)
    if ($script:Data.ContainsKey($Key) -and $null -ne $script:Data[$Key]) {
        return "$($script:Data[$Key])"
    }
    return 'N/A'
}

# ═════════════════════════════════════════════════════════════════════════════
# 1. Tenant Overview
# ═════════════════════════════════════════════════════════════════════════════
function Get-TenantOverview {
    param([string]$OutputPath, [int]$Total)
    Write-SectionHeader "Tenant Overview" $Total

    try {
        $org    = Get-MgOrganization
        $ctx    = Get-MgContext

        $primaryDomain = ($org.VerifiedDomains | Where-Object { $_.IsDefault }).Name
        $allDomains    = (($org.VerifiedDomains | Where-Object { $_.IsVerified }).Name) -join ', '
        $created       = if ($org.CreatedDateTime) { $org.CreatedDateTime.ToString('yyyy-MM-dd') } else { 'N/A' }

        [PSCustomObject]@{
            'Organization Name'  = $org.DisplayName
            'Tenant ID'          = $ctx.TenantId
            'Primary Domain'     = $primaryDomain
            'All Domains'        = $allDomains
            'Country'            = $org.CountryLetterCode
            'Technical Contact'  = ($org.TechnicalNotificationMails -join ', ')
            'Tenant Created'     = $created
        } | Export-Csv -Path "$OutputPath\Tenant_Overview.csv" -NoTypeInformation

        $script:Data['OrgName']       = $org.DisplayName
        $script:Data['TenantId']      = $ctx.TenantId
        $script:Data['PrimaryDomain'] = $primaryDomain

        Write-Host "  $($org.DisplayName) | $primaryDomain | $($ctx.TenantId)" -ForegroundColor Green

        $domainRows = $org.VerifiedDomains | Where-Object { $_.IsVerified } |
            Select-Object Name, Type, @{N='Default';E={$_.IsDefault}}

        $html = "<div class='stat-row'>" +
                (New-StatCard "Organization" $org.DisplayName) +
                (New-StatCard "Primary Domain" $primaryDomain) +
                (New-StatCard "Country" $org.CountryLetterCode) +
                (New-StatCard "Tenant Created" $created) +
                "</div><h3>Verified Domains</h3>" +
                (ConvertTo-HtmlTable $domainRows @('Name','Type','Default'))

        Add-HtmlSection "Tenant Overview" $html
    }
    catch {
        Write-Host "  Error: $_" -ForegroundColor Red
        Add-HtmlSection "Tenant Overview" "<p class='error'>Error: $_</p>" "bad"
    }
}

# ═════════════════════════════════════════════════════════════════════════════
# 2. Licensing
# ═════════════════════════════════════════════════════════════════════════════
function Get-M365LicensingInfo {
    param([string]$OutputPath, [int]$Total)
    Write-SectionHeader "Licensing" $Total

    try {
        $licenses = Get-MgSubscribedSku | Select-Object @{
            N='SKU Part Number'; E={$_.SkuPartNumber}
        }, @{
            N='Friendly Name'; E={
                if ($script:SkuMap.ContainsKey($_.SkuPartNumber)) { $script:SkuMap[$_.SkuPartNumber] } else { $_.SkuPartNumber }
            }
        }, @{
            N='Total';     E={$_.PrepaidUnits.Enabled}
        }, @{
            N='Consumed';  E={$_.ConsumedUnits}
        }, @{
            N='Available'; E={$_.PrepaidUnits.Enabled - $_.ConsumedUnits}
        }, @{
            N='Warning';   E={$_.PrepaidUnits.Warning}
        }, @{
            N='Suspended'; E={$_.PrepaidUnits.Suspended}
        }

        $licenses | Export-Csv -Path "$OutputPath\Licenses.csv" -NoTypeInformation

        $totalPurchased = ($licenses | Measure-Object Total -Sum).Sum
        $totalConsumed  = ($licenses | Measure-Object Consumed -Sum).Sum
        $totalAvailable = $totalPurchased - $totalConsumed
        $unusedPct      = if ($totalPurchased -gt 0) { [math]::Round(($totalAvailable / $totalPurchased) * 100, 1) } else { 0 }

        $script:Data['TotalLicenses']  = $totalPurchased
        $script:Data['UsedLicenses']   = $totalConsumed
        $script:Data['UnusedLicenses'] = $totalAvailable

        Write-Host "  $($licenses.Count) SKUs | $totalConsumed/$totalPurchased consumed | $totalAvailable available ($unusedPct% unused)" -ForegroundColor Green

        $unusedColor  = if ($unusedPct -gt 20) { "#ff8c00" } else { "#107c10" }
        $unusedStatus = if ($unusedPct -gt 20) { "warn" } else { "good" }

        $html = "<div class='stat-row'>" +
                (New-StatCard "Purchased" $totalPurchased) +
                (New-StatCard "Consumed" $totalConsumed "#107c10") +
                (New-StatCard "Available ($unusedPct%)" $totalAvailable $unusedColor) +
                (New-StatCard "SKUs" $licenses.Count) +
                "</div><h3>License Details</h3>" +
                (ConvertTo-HtmlTable $licenses @('Friendly Name','SKU Part Number','Total','Consumed','Available','Warning','Suspended'))

        Add-HtmlSection "Licensing" $html $unusedStatus
    }
    catch {
        Write-Host "  Error: $_" -ForegroundColor Red
        Add-HtmlSection "Licensing" "<p class='error'>Error: $_</p>" "bad"
    }
}

# ═════════════════════════════════════════════════════════════════════════════
# 3. Secure Score
# ═════════════════════════════════════════════════════════════════════════════
function Get-SecureScoreInfo {
    param([string]$OutputPath, [int]$Total)
    Write-SectionHeader "Secure Score" $Total

    try {
        $secureScore = Get-MgSecuritySecureScore | Select-Object -First 1
        $pct = 0

        if ($secureScore) {
            $pct = if ($secureScore.MaxScore -gt 0) {
                [math]::Round(($secureScore.CurrentScore / $secureScore.MaxScore) * 100, 1)
            } else { 0 }

            [PSCustomObject]@{
                'Assessment Date'  = Get-Date -Format 'yyyy-MM-dd'
                'Current Score'    = $secureScore.CurrentScore
                'Max Score'        = $secureScore.MaxScore
                'Score Percentage' = "$pct%"
                'Licensed Users'   = $secureScore.LicensedUserCount
                'Active Users'     = $secureScore.ActiveUserCount
            } | Export-Csv -Path "$OutputPath\SecureScore_Current.csv" -NoTypeInformation

            $script:Data['SecureScore']    = $secureScore.CurrentScore
            $script:Data['SecureScoreMax'] = $secureScore.MaxScore
            $script:Data['SecureScorePct'] = $pct
        }

        $controlProfiles = Get-MgSecuritySecureScoreControlProfile
        $maxPossible     = [math]::Max(($controlProfiles | Measure-Object -Property MaxScore -Sum).Sum, 1)
        $rankCounter     = 0

        $recommendations = $controlProfiles | Sort-Object MaxScore -Descending | ForEach-Object {
            $rankCounter++
            $title    = $_.Title
            $category = $_.ControlCategory
            $id       = $_.Id
            $combined = "$title|$id"

            $product = if ($combined -match 'Entra|Azure AD|MFA|password|sign.?in|identity|consent|AAD') {
                'Entra ID'
            } elseif ($combined -match 'Defender Antivirus|Windows Defender|BitLocker|firewall|device|Intune|endpoint|MDM') {
                'Defender for Endpoint'
            } elseif ($combined -match 'phishing|spam|safe.*attach|safe.*link|spoof|DKIM|DMARC|SPF|AntiPhish') {
                'Defender for Office 365'
            } elseif ($combined -match 'DLP|data loss|sensitivity|IRM|information protect') {
                'Information Protection'
            } elseif ($combined -match 'Exchange|mailbox|email|mail flow|forward') {
                'Exchange Online'
            } elseif ($combined -match 'SharePoint|OneDrive|external sharing|SPO') {
                'SharePoint Online'
            } elseif ($combined -match 'Teams|meeting|chat') {
                'Microsoft Teams'
            } elseif ($combined -match 'Cloud App|MCAS|shadow IT') {
                'Defender for Cloud Apps'
            } elseif ($category -eq 'Identity') {
                'Entra ID'
            } elseif ($category -eq 'Device') {
                'Defender for Endpoint'
            } else {
                'Microsoft 365'
            }

            [PSCustomObject]@{
                'Rank'               = $rankCounter
                'Recommended Action' = $title
                'Points'             = $_.MaxScore
                'Score Impact'       = "$([math]::Round(($_.MaxScore / $maxPossible) * 100, 2))%"
                'Status'             = if ($_.ImplementationStatus) { $_.ImplementationStatus } else { 'To address' }
                'Category'           = if ($category) { $category } else { 'Security' }
                'Product'            = $product
                'Last Synced'        = if ($_.LastModifiedDateTime) { $_.LastModifiedDateTime.ToString('yyyy-MM-dd') } else { 'N/A' }
            }
        }

        $recommendations | Export-Csv -Path "$OutputPath\SecureScore_Recommendations.csv" -NoTypeInformation

        $notDone    = $recommendations | Where-Object { $_.Status -notin @('Completed','Resolved','Ignored') }
        $highImpact = $notDone | Where-Object { $_.Points -ge 5 } | Select-Object -First 10

        $script:Data['SecureScoreGaps'] = $notDone.Count

        Write-Host "  Score: $($secureScore.CurrentScore)/$($secureScore.MaxScore) ($pct%) | $($notDone.Count) gaps | $($highImpact.Count) high-impact" -ForegroundColor Green

        $scoreColor  = if ($pct -ge 80) { "#107c10" } elseif ($pct -ge 50) { "#ff8c00" } else { "#d13438" }
        $scoreStatus = if ($pct -ge 80) { "good" } elseif ($pct -ge 50) { "warn" } else { "bad" }

        $html = "<div class='stat-row'>" +
                (New-StatCard "Secure Score" "$($secureScore.CurrentScore)/$($secureScore.MaxScore)" $scoreColor) +
                (New-StatCard "Score %" "$pct%" $scoreColor) +
                (New-StatCard "Total Controls" $recommendations.Count) +
                (New-StatCard "Gaps to Address" $notDone.Count $(if ($notDone.Count -gt 20) { "#d13438" } else { "#ff8c00" })) +
                "</div><h3>Top High-Impact Recommendations</h3>" +
                (ConvertTo-HtmlTable $highImpact @('Rank','Recommended Action','Points','Score Impact','Status','Product','Category')) +
                "<p><em>Full list exported to SecureScore_Recommendations.csv</em></p>"

        Add-HtmlSection "Secure Score" $html $scoreStatus
    }
    catch {
        Write-Host "  Error: $_" -ForegroundColor Red
        Add-HtmlSection "Secure Score" "<p class='error'>Error: $_</p>" "bad"
    }
}

# ═════════════════════════════════════════════════════════════════════════════
# 4. User Health
# ═════════════════════════════════════════════════════════════════════════════
function Get-UserHealth {
    param([string]$OutputPath, [int]$Total)
    Write-SectionHeader "User Health" $Total

    try {
        Write-Host "  Gathering users..." -ForegroundColor Gray
        $users = Get-MgUser -All -Property Id,DisplayName,UserPrincipalName,UserType,AccountEnabled,AssignedLicenses,CreatedDateTime,SignInActivity,Mail

        $enabledMembers = $users | Where-Object { $_.AccountEnabled -and $_.UserType -eq 'Member' }
        $guestUsers     = $users | Where-Object { $_.UserType -eq 'Guest' }
        $licensedUsers  = ($users | Where-Object { $_.AssignedLicenses.Count -gt 0 }).Count
        $unlicensed     = ($enabledMembers | Where-Object { $_.AssignedLicenses.Count -eq 0 }).Count

        $thirtyDaysAgo = (Get-Date).AddDays(-30)
        $ninetyDaysAgo = (Get-Date).AddDays(-90)
        $newUsers      = ($users | Where-Object { $_.CreatedDateTime -gt $thirtyDaysAgo }).Count

        $inactive = ($enabledMembers | Where-Object {
            $_.SignInActivity.LastSignInDateTime -and
            $_.SignInActivity.LastSignInDateTime -lt $ninetyDaysAgo
        }).Count

        Write-Host "  Gathering admin roles..." -ForegroundColor Gray
        $adminRoles = [System.Collections.Generic.List[PSCustomObject]]::new()
        $globalAdmins = 0

        try {
            $directoryRoles = Get-MgDirectoryRole -All
            foreach ($role in $directoryRoles) {
                try {
                    $members = Get-MgDirectoryRoleMember -DirectoryRoleId $role.Id -All -ErrorAction SilentlyContinue
                    foreach ($m in $members) {
                        if ($m.AdditionalProperties.'@odata.type' -eq '#microsoft.graph.user') {
                            $adminRoles.Add([PSCustomObject]@{
                                'Role'         = $role.DisplayName
                                'Display Name' = $m.AdditionalProperties.displayName
                                'UPN'          = $m.AdditionalProperties.userPrincipalName
                            })
                        }
                    }
                } catch { }
            }
            $globalAdmins = ($adminRoles | Where-Object { $_.Role -eq 'Global Administrator' }).Count
        } catch {
            Write-Host "  Could not enumerate admin roles" -ForegroundColor Yellow
        }

        $users | Select-Object @{N='Display Name';E={$_.DisplayName}},
            @{N='UPN';E={$_.UserPrincipalName}},
            @{N='Type';E={$_.UserType}},
            @{N='Enabled';E={$_.AccountEnabled}},
            @{N='Licensed';E={$_.AssignedLicenses.Count -gt 0}},
            @{N='Created';E={if ($_.CreatedDateTime) { $_.CreatedDateTime.ToString('yyyy-MM-dd') } else { 'N/A' }}},
            @{N='Last Sign-In';E={if ($_.SignInActivity.LastSignInDateTime) { $_.SignInActivity.LastSignInDateTime.ToString('yyyy-MM-dd') } else { 'N/A' }}} |
            Export-Csv -Path "$OutputPath\Users.csv" -NoTypeInformation

        if ($adminRoles.Count -gt 0) {
            $adminRoles | Export-Csv -Path "$OutputPath\Users_AdminRoles.csv" -NoTypeInformation
        }

        $script:Data['TotalUsers']    = $users.Count
        $script:Data['GuestUsers']    = $guestUsers.Count
        $script:Data['GlobalAdmins']  = $globalAdmins
        $script:Data['InactiveUsers'] = $inactive

        Write-Host "  Total: $($users.Count) | Enabled members: $($enabledMembers.Count) | Guests: $($guestUsers.Count) | Global Admins: $globalAdmins" -ForegroundColor Green

        $adminColor  = if ($globalAdmins -gt 10) { "#d13438" } elseif ($globalAdmins -gt 5) { "#ff8c00" } else { "#107c10" }
        $adminStatus = if ($globalAdmins -gt 10) { "bad" } elseif ($globalAdmins -gt 5) { "warn" } else { "good" }

        $gaRows = $adminRoles | Where-Object { $_.Role -eq 'Global Administrator' } | Select-Object 'Display Name','UPN' | Select-Object -First 15

        $adminWarn = if ($globalAdmins -gt 5) {
            "<p class='warn-text'>Warning: $globalAdmins Global Administrator accounts detected. Microsoft recommends no more than 4-5 for most organizations. Excess global admins increase the attack surface significantly.</p>"
        } else { "" }

        $html = "<div class='stat-row'>" +
                (New-StatCard "Total Users" $users.Count) +
                (New-StatCard "Enabled Members" $enabledMembers.Count "#107c10") +
                (New-StatCard "Guest Users" $guestUsers.Count "#ff8c00") +
                (New-StatCard "Unlicensed Active" $unlicensed $(if ($unlicensed -gt 0) { "#ff8c00" } else { "#107c10" })) +
                "</div><div class='stat-row'>" +
                (New-StatCard "Licensed" $licensedUsers) +
                (New-StatCard "Disabled Accounts" ($users.Count - $enabledMembers.Count - $guestUsers.Count) "#767676") +
                (New-StatCard "New (30 days)" $newUsers) +
                (New-StatCard "Inactive (90 days)" $inactive $(if ($inactive -gt 10) { "#ff8c00" } else { "#107c10" })) +
                "</div>" +
                "<h3>Global Administrators ($globalAdmins)</h3>" +
                (ConvertTo-HtmlTable $gaRows @('Display Name','UPN')) +
                $adminWarn

        Add-HtmlSection "User Health" $html $adminStatus
    }
    catch {
        Write-Host "  Error: $_" -ForegroundColor Red
        Add-HtmlSection "User Health" "<p class='error'>Error: $_</p>" "bad"
    }
}

# ═════════════════════════════════════════════════════════════════════════════
# 5. Entra ID Security
# ═════════════════════════════════════════════════════════════════════════════
function Get-EntraIDSecurity {
    param([string]$OutputPath, [int]$Total)
    Write-SectionHeader "Entra ID Security" $Total

    try {
        # Conditional Access
        Write-Host "  Checking Conditional Access policies..." -ForegroundColor Gray
        $caPolicies   = @()
        $caEnabled    = 0
        $caDisabled   = 0
        $caReportOnly = 0
        try {
            $caPolicies   = Get-MgIdentityConditionalAccessPolicy -All -ErrorAction Stop
            $caEnabled    = ($caPolicies | Where-Object { $_.State -eq 'enabled' }).Count
            $caDisabled   = ($caPolicies | Where-Object { $_.State -eq 'disabled' }).Count
            $caReportOnly = ($caPolicies | Where-Object { $_.State -eq 'enabledForReportingButNotEnforced' }).Count

            $caPolicies | Select-Object @{N='Policy Name';E={$_.DisplayName}},
                @{N='State';E={$_.State}},
                @{N='Created';E={if ($_.CreatedDateTime) { $_.CreatedDateTime.ToString('yyyy-MM-dd') } else { 'N/A' }}},
                @{N='Modified';E={if ($_.ModifiedDateTime) { $_.ModifiedDateTime.ToString('yyyy-MM-dd') } else { 'N/A' }}} |
                Export-Csv -Path "$OutputPath\EntraID_ConditionalAccess.csv" -NoTypeInformation
        } catch {
            Write-Host "  CA policies not accessible (requires Policy.Read.All)" -ForegroundColor Yellow
        }

        # MFA registration
        Write-Host "  Checking MFA registration stats..." -ForegroundColor Gray
        $mfaRegistered = 'N/A'
        $mfaTotal      = 'N/A'
        $ssprRegistered = 'N/A'
        try {
            $mfaReport = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/reports/authenticationMethods/usersRegisteredByFeature" -Method GET -ErrorAction Stop
            $counts = $mfaReport.userRegistrationFeatureCounts
            $mfaEntry  = $counts | Where-Object { $_.feature -eq 'mfa' }
            $ssprEntry = $counts | Where-Object { $_.feature -eq 'selfServicePasswordReset' }
            if ($mfaEntry)  { $mfaRegistered = $mfaEntry.registered; $mfaTotal = $mfaEntry.total }
            if ($ssprEntry) { $ssprRegistered = $ssprEntry.registered }
        } catch {
            Write-Host "  MFA stats not available (may require Entra ID P1/P2)" -ForegroundColor Yellow
        }

        # Risky users
        $riskyUsers = 0
        try {
            $risky = Get-MgRiskyUser -All -ErrorAction Stop
            $riskyUsers = ($risky | Where-Object { $_.RiskLevel -ne 'none' }).Count
        } catch { }

        $script:Data['CAPolicies']    = $caEnabled
        $script:Data['MFARegistered'] = $mfaRegistered
        $script:Data['RiskyUsers']    = $riskyUsers

        Write-Host "  CA (enabled): $caEnabled | MFA registered: $mfaRegistered/$mfaTotal | Risky users: $riskyUsers" -ForegroundColor Green

        $caStatus = if ($caEnabled -eq 0) { "bad" } elseif ($caEnabled -lt 3) { "warn" } else { "good" }

        $mfaPct = if ($mfaTotal -ne 'N/A' -and [int]$mfaTotal -gt 0) {
            "$([math]::Round(([int]$mfaRegistered / [int]$mfaTotal) * 100, 1))%"
        } else { 'N/A' }

        $mfaColor = if ($mfaPct -eq 'N/A') { "#767676" } elseif ([double]($mfaPct -replace '%','') -lt 80) { "#d13438" } else { "#107c10" }

        $caPolicyRows = $caPolicies | Select-Object @{N='Policy Name';E={$_.DisplayName}}, @{N='State';E={$_.State}}

        $caWarn = if ($caEnabled -eq 0) {
            "<p class='warn-text'>Critical: No enabled Conditional Access policies found. All sign-ins are unrestricted. This is the single highest-impact security gap to address.</p>"
        } elseif ($caReportOnly -gt 0) {
            "<p class='warn-text'>Note: $caReportOnly policies are in Report-Only mode. Review and enforce them to improve security posture.</p>"
        } else { "" }

        $html = "<div class='stat-row'>" +
                (New-StatCard "CA Policies (Enabled)" $caEnabled $(if ($caEnabled -eq 0) { "#d13438" } elseif ($caEnabled -lt 3) { "#ff8c00" } else { "#107c10" })) +
                (New-StatCard "CA Policies (Report-Only)" $caReportOnly "#ff8c00") +
                (New-StatCard "CA Policies (Disabled)" $caDisabled "#767676") +
                (New-StatCard "Risky Users" $riskyUsers $(if ($riskyUsers -gt 0) { "#d13438" } else { "#107c10" })) +
                "</div><div class='stat-row'>" +
                (New-StatCard "MFA Registered" $mfaRegistered $mfaColor) +
                (New-StatCard "MFA Registration %" $mfaPct $mfaColor) +
                (New-StatCard "SSPR Registered" $ssprRegistered) +
                (New-StatCard "Total (MFA Report)" $mfaTotal "#767676") +
                "</div><h3>Conditional Access Policies</h3>" +
                (ConvertTo-HtmlTable $caPolicyRows @('Policy Name','State')) +
                $caWarn

        Add-HtmlSection "Entra ID Security" $html $caStatus
    }
    catch {
        Write-Host "  Error: $_" -ForegroundColor Red
        Add-HtmlSection "Entra ID Security" "<p class='error'>Error: $_</p>" "bad"
    }
}

# ═════════════════════════════════════════════════════════════════════════════
# 6. Exchange Online
# ═════════════════════════════════════════════════════════════════════════════
function Get-ExchangeOnlineInfo {
    param([string]$OutputPath, [int]$Total)
    Write-SectionHeader "Exchange Online" $Total

    try {
        Write-Host "  Gathering users with mail addresses..." -ForegroundColor Gray
        $users = Get-MgUser -All -Property DisplayName,Mail,UserPrincipalName,UserType,AccountEnabled,AssignedLicenses |
                 Where-Object { $_.Mail -ne $null }

        # User mailboxes: enabled Member accounts with licenses
        $userMailboxes = $users | Where-Object {
            $_.AccountEnabled -and $_.UserType -eq 'Member' -and $_.AssignedLicenses.Count -gt 0
        } | Select-Object @{N='Display Name';E={$_.DisplayName}},
            @{N='Email';E={$_.Mail}},
            @{N='UPN';E={$_.UserPrincipalName}},
            @{N='Mailbox Type';E={'User Mailbox'}},
            @{N='Enabled';E={$_.AccountEnabled}}

        # Shared mailboxes
        $sharedMailboxes = @()
        $sharedSource    = 'Graph (heuristic)'

        if ($script:ExchangeConnected) {
            Write-Host "  Getting shared mailboxes via Exchange Online..." -ForegroundColor Gray
            try {
                $exoShared = Get-EXOMailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited `
                    -Properties DisplayName,PrimarySmtpAddress,WhenCreated -ErrorAction Stop
                $sharedMailboxes = $exoShared | Select-Object @{N='Display Name';E={$_.DisplayName}},
                    @{N='Email';E={$_.PrimarySmtpAddress}},
                    @{N='Mailbox Type';E={'Shared Mailbox'}},
                    @{N='Created';E={if ($_.WhenCreated) { $_.WhenCreated.ToString('yyyy-MM-dd') } else { 'N/A' }}}
                $sharedSource = 'Exchange Online (accurate)'
            } catch {
                Write-Host "  EXO shared mailbox query failed, falling back to Graph heuristic" -ForegroundColor Yellow
            }
        }

        if ($sharedMailboxes.Count -eq 0) {
            # Heuristic: disabled Member accounts with mail and no licenses
            $sharedMailboxes = $users | Where-Object {
                !$_.AccountEnabled -and $_.AssignedLicenses.Count -eq 0 -and $_.UserType -eq 'Member'
            } | Select-Object @{N='Display Name';E={$_.DisplayName}},
                @{N='Email';E={$_.Mail}},
                @{N='Mailbox Type';E={'Shared Mailbox (estimated)'}},
                @{N='Created';E={'N/A'}}
        }

        # Groups
        Write-Host "  Gathering groups..." -ForegroundColor Gray
        $allGroups    = Get-MgGroup -All -Property DisplayName,Mail,MailEnabled,SecurityEnabled,GroupTypes,Description
        $m365Groups   = $allGroups | Where-Object { $_.GroupTypes -contains "Unified" }
        $distGroups   = $allGroups | Where-Object { $_.MailEnabled -and !$_.SecurityEnabled -and $_.GroupTypes.Count -eq 0 }
        $mailSecGroups = $allGroups | Where-Object { $_.SecurityEnabled -and $_.MailEnabled -and $_.GroupTypes.Count -eq 0 }
        $secGroups    = $allGroups | Where-Object { $_.SecurityEnabled -and !$_.MailEnabled }

        ($m365Groups + $distGroups + $mailSecGroups) | Select-Object @{N='Display Name';E={$_.DisplayName}},
            @{N='Email';E={$_.Mail}},
            @{N='Type';E={
                if ($_.GroupTypes -contains "Unified") { "M365 Group" }
                elseif ($_.MailEnabled -and !$_.SecurityEnabled) { "Distribution" }
                else { "Mail-Enabled Security" }
            }},
            @{N='Description';E={$_.Description}} |
            Export-Csv -Path "$OutputPath\Exchange_Groups.csv" -NoTypeInformation

        $userMailboxes  | Export-Csv -Path "$OutputPath\Exchange_Mailboxes.csv" -NoTypeInformation
        $sharedMailboxes | Export-Csv -Path "$OutputPath\Exchange_SharedMailboxes.csv" -NoTypeInformation
        ($secGroups | Select-Object DisplayName,Description) | Export-Csv -Path "$OutputPath\Security_Groups.csv" -NoTypeInformation

        $script:Data['UserMailboxes']   = $userMailboxes.Count
        $script:Data['SharedMailboxes'] = $sharedMailboxes.Count

        Write-Host "  User mailboxes: $($userMailboxes.Count) | Shared: $($sharedMailboxes.Count) | M365 Groups: $($m365Groups.Count) | Dist: $($distGroups.Count)" -ForegroundColor Green

        $sharedWarn = if (!$script:ExchangeConnected) {
            "<p class='warn-text'>Note: Shared mailbox count is an estimate (Graph heuristic). Install ExchangeOnlineManagement module for accurate data. Source: $sharedSource</p>"
        } else { "<p>Data source: $sharedSource</p>" }

        $html = "<div class='stat-row'>" +
                (New-StatCard "User Mailboxes" $userMailboxes.Count) +
                (New-StatCard "Shared Mailboxes" $sharedMailboxes.Count) +
                (New-StatCard "M365 Groups" $m365Groups.Count) +
                (New-StatCard "Distribution Groups" $distGroups.Count) +
                "</div><div class='stat-row'>" +
                (New-StatCard "Mail-Enabled Security" $mailSecGroups.Count) +
                (New-StatCard "Security Groups" $secGroups.Count "#767676") +
                (New-StatCard "Total Groups" $allGroups.Count "#767676") +
                "</div>" + $sharedWarn

        Add-HtmlSection "Exchange Online" $html
    }
    catch {
        Write-Host "  Error: $_" -ForegroundColor Red
        Add-HtmlSection "Exchange Online" "<p class='error'>Error: $_</p>" "bad"
    }
}

# ═════════════════════════════════════════════════════════════════════════════
# 7. Email Security (DNS)
# ═════════════════════════════════════════════════════════════════════════════
function Get-EmailSecurity {
    param([string]$OutputPath, [int]$Total)
    Write-SectionHeader "Email Security (DNS)" $Total

    try {
        $domain = $script:Data['PrimaryDomain']
        if (!$domain) {
            Write-Host "  Primary domain not available" -ForegroundColor Yellow
            Add-HtmlSection "Email Security (DNS)" "<p class='warn-text'>Primary domain not available — skipped.</p>" "warn"
            return
        }

        Write-Host "  Checking DNS records for $domain..." -ForegroundColor Gray

        # SPF
        $spfStatus = 'Not Found'; $spfRecord = ''
        try {
            $spf = Resolve-DnsName -Name $domain -Type TXT -ErrorAction Stop | Where-Object { $_.Strings -match 'v=spf1' }
            if ($spf) {
                $spfRecord = ($spf.Strings -join ' ')
                $spfStatus = if ($spfRecord -match '-all') { 'Hard Fail (-all)' }
                              elseif ($spfRecord -match '~all') { 'Soft Fail (~all)' }
                              else { 'Permissive (+all)' }
            }
        } catch { }

        # DMARC
        $dmarcStatus = 'Not Found'; $dmarcRecord = ''
        try {
            $dmarc = Resolve-DnsName -Name "_dmarc.$domain" -Type TXT -ErrorAction Stop | Where-Object { $_.Strings -match 'v=DMARC1' }
            if ($dmarc) {
                $dmarcRecord = ($dmarc.Strings -join ' ')
                $dmarcStatus = if ($dmarcRecord -match 'p=reject') { 'Reject (enforced)' }
                               elseif ($dmarcRecord -match 'p=quarantine') { 'Quarantine (enforced)' }
                               else { 'None (not enforced)' }
            }
        } catch { }

        # DKIM — check common selectors
        $dkimStatus = 'Not detected'
        foreach ($sel in @('selector1','selector2','google','k1','mail','default')) {
            try {
                $dkim = Resolve-DnsName -Name "$sel._domainkey.$domain" -Type TXT -ErrorAction Stop |
                        Where-Object { $_.Strings -match 'v=DKIM1' }
                if ($dkim) { $dkimStatus = "Found (selector: $sel)"; break }
            } catch { }
        }

        # MTA-STS
        $mtaStsStatus = 'Not Found'
        try {
            $mta = Resolve-DnsName -Name "_mta-sts.$domain" -Type TXT -ErrorAction Stop
            if ($mta) { $mtaStsStatus = 'Found' }
        } catch { }

        $results = @(
            [PSCustomObject]@{ Check='SPF';     Status=$spfStatus;   Details=$spfRecord }
            [PSCustomObject]@{ Check='DMARC';   Status=$dmarcStatus; Details=$dmarcRecord }
            [PSCustomObject]@{ Check='DKIM';    Status=$dkimStatus;  Details='' }
            [PSCustomObject]@{ Check='MTA-STS'; Status=$mtaStsStatus; Details='' }
        )

        $results | Export-Csv -Path "$OutputPath\EmailSecurity_DNS.csv" -NoTypeInformation

        $hasSPF   = $spfStatus -ne 'Not Found'
        $hasDMARC = $dmarcStatus -notin @('Not Found','None (not enforced)')
        $hasDKIM  = $dkimStatus -ne 'Not detected'

        $status = if ($hasSPF -and $hasDMARC -and $hasDKIM) { "good" } elseif ($hasSPF) { "warn" } else { "bad" }

        Write-Host "  SPF: $spfStatus | DMARC: $dmarcStatus | DKIM: $dkimStatus" -ForegroundColor Green

        $spfColor   = if ($hasSPF) { "#107c10" } else { "#d13438" }
        $dmarcColor = if ($hasDMARC) { "#107c10" } elseif ($dmarcStatus -match 'None') { "#ff8c00" } else { "#d13438" }
        $dkimColor  = if ($hasDKIM) { "#107c10" } else { "#ff8c00" }

        $warnings = ""
        if (!$hasDMARC) { $warnings += "<p class='warn-text'>DMARC is not enforced. Without p=reject or p=quarantine, spoofed emails from your domain cannot be blocked at recipient servers.</p>" }
        if (!$hasSPF)   { $warnings += "<p class='warn-text'>SPF record not found. Your domain is vulnerable to email spoofing.</p>" }
        if (!$hasDKIM)  { $warnings += "<p class='warn-text'>DKIM not detected on common selectors. It may be configured on a custom selector — verify in your DNS provider.</p>" }

        $html = "<div class='stat-row'>" +
                (New-StatCard "SPF" $spfStatus $spfColor) +
                (New-StatCard "DMARC" $dmarcStatus $dmarcColor) +
                (New-StatCard "DKIM" $dkimStatus $dkimColor) +
                (New-StatCard "MTA-STS" $mtaStsStatus $(if ($mtaStsStatus -ne 'Not Found') { "#107c10" } else { "#767676" })) +
                "</div><h3>DNS Record Details</h3>" +
                (ConvertTo-HtmlTable $results @('Check','Status','Details')) +
                $warnings

        Add-HtmlSection "Email Security (DNS)" $html $status
    }
    catch {
        Write-Host "  Error: $_" -ForegroundColor Red
        Add-HtmlSection "Email Security (DNS)" "<p class='error'>Error: $_</p>" "bad"
    }
}

# ═════════════════════════════════════════════════════════════════════════════
# 8. SharePoint Online
# ═════════════════════════════════════════════════════════════════════════════
function Get-SharePointSites {
    param([string]$OutputPath, [int]$Total)
    Write-SectionHeader "SharePoint Online" $Total

    try {
        $siteCount    = 0
        $activeSites  = 0
        $totalStorage = 0

        try {
            $tempFile = Join-Path $env:TEMP "SPUsage_$(Get-Date -Format 'yyyyMMddHHmmss').csv"
            Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/reports/getSharePointSiteUsageDetail(period='D30')" -Method GET -OutputFilePath $tempFile

            if (Test-Path $tempFile) {
                $raw = Import-Csv $tempFile
                $siteUsage = $raw | Select-Object @{N='Site URL';E={$_.'Site URL'}},
                    @{N='Owner';E={if ($_.'Owner Display Name') { $_.'Owner Display Name' } else { $_.'Owner Principal Name' }}},
                    @{N='Storage Used (GB)';E={[math]::Round([double]$_.'Storage Used (Byte)' / 1GB, 2)}},
                    @{N='File Count';E={$_.'File Count'}},
                    @{N='Active Files';E={$_.'Active File Count'}},
                    @{N='Last Activity';E={$_.'Last Activity Date'}},
                    @{N='Site Type';E={if ($_.'Root Web Template') { $_.'Root Web Template' } else { 'Standard' }}}

                $siteUsage | Export-Csv -Path "$OutputPath\SharePoint_Usage.csv" -NoTypeInformation
                $siteCount    = $siteUsage.Count
                $totalStorage = [math]::Round(($siteUsage | Measure-Object 'Storage Used (GB)' -Sum).Sum, 2)
                $activeSites  = ($siteUsage | Where-Object { $_.'Last Activity' -ne '' }).Count
                Remove-Item $tempFile -Force -ErrorAction SilentlyContinue

                Write-Host "  Sites: $siteCount | Active (30d): $activeSites | Storage: $totalStorage GB" -ForegroundColor Green

                $top10 = $siteUsage | Sort-Object 'Storage Used (GB)' -Descending | Select-Object -First 10
                $html = "<div class='stat-row'>" +
                        (New-StatCard "Total Sites" $siteCount) +
                        (New-StatCard "Active (30 days)" $activeSites "#107c10") +
                        (New-StatCard "Total Storage" "$totalStorage GB") +
                        (New-StatCard "Inactive Sites" ($siteCount - $activeSites) "#ff8c00") +
                        "</div><h3>Top 10 Sites by Storage</h3>" +
                        (ConvertTo-HtmlTable $top10 @('Site URL','Owner','Storage Used (GB)','File Count','Active Files','Last Activity'))

                Add-HtmlSection "SharePoint Online" $html
                return
            }
        } catch {
            Write-Host "  Could not retrieve SharePoint usage report: $_" -ForegroundColor Yellow
        }

        Add-HtmlSection "SharePoint Online" "<p class='warn-text'>Usage data not available. Requires Reports.Read.All permission.</p>" "warn"
    }
    catch {
        Write-Host "  Error: $_" -ForegroundColor Red
        Add-HtmlSection "SharePoint Online" "<p class='error'>Error: $_</p>" "bad"
    }
}

# ═════════════════════════════════════════════════════════════════════════════
# 9. OneDrive for Business
# ═════════════════════════════════════════════════════════════════════════════
function Get-OneDriveUsage {
    param([string]$OutputPath, [int]$Total)
    Write-SectionHeader "OneDrive for Business" $Total

    try {
        $tempFile = Join-Path $env:TEMP "ODUsage_$(Get-Date -Format 'yyyyMMddHHmmss').csv"
        Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/reports/getOneDriveUsageAccountDetail(period='D30')" -Method GET -OutputFilePath $tempFile

        if (Test-Path $tempFile) {
            $raw = Import-Csv $tempFile
            $usage = $raw | Select-Object @{N='Display Name';E={
                    if ($_.'Owner Display Name') { $_.'Owner Display Name' }
                    else { ($_.'Owner Principal Name' -split '@')[0] }
                }},
                @{N='UPN';E={$_.'Owner Principal Name'}},
                @{N='Last Activity';E={$_.'Last Activity Date'}},
                @{N='Files';E={$_.'File Count'}},
                @{N='Active Files';E={$_.'Active File Count'}},
                @{N='Storage Used (GB)';E={[math]::Round([double]$_.'Storage Used (Byte)' / 1GB, 2)}},
                @{N='Allocated (GB)';E={[math]::Round([double]$_.'Storage Allocated (Byte)' / 1GB, 2)}}

            $usage | Export-Csv -Path "$OutputPath\OneDrive_Usage.csv" -NoTypeInformation

            $activeUsers  = ($usage | Where-Object { $_.'Last Activity' -ne '' }).Count
            $totalStorage = [math]::Round(($usage | Measure-Object 'Storage Used (GB)' -Sum).Sum, 2)
            $avgStorage   = [math]::Round($totalStorage / [Math]::Max($usage.Count, 1), 2)
            $top10        = $usage | Sort-Object 'Storage Used (GB)' -Descending | Select-Object -First 10

            Write-Host "  Accounts: $($usage.Count) | Active: $activeUsers | Storage: $totalStorage GB" -ForegroundColor Green

            $html = "<div class='stat-row'>" +
                    (New-StatCard "Accounts" $usage.Count) +
                    (New-StatCard "Active (30 days)" $activeUsers "#107c10") +
                    (New-StatCard "Total Storage" "$totalStorage GB") +
                    (New-StatCard "Avg per User" "$avgStorage GB" "#767676") +
                    "</div><h3>Top 10 Storage Users</h3>" +
                    (ConvertTo-HtmlTable $top10 @('Display Name','UPN','Storage Used (GB)','Files','Last Activity'))

            Add-HtmlSection "OneDrive for Business" $html
            Remove-Item $tempFile -Force -ErrorAction SilentlyContinue
        }
    }
    catch {
        Write-Host "  Error: $_" -ForegroundColor Red
        Add-HtmlSection "OneDrive for Business" "<p class='error'>Error: $_</p>" "bad"
    }
}

# ═════════════════════════════════════════════════════════════════════════════
# 10. Microsoft Teams
# ═════════════════════════════════════════════════════════════════════════════
function Get-TeamsInfo {
    param([string]$OutputPath, [int]$Total)
    Write-SectionHeader "Microsoft Teams" $Total

    try {
        $teams = Get-MgTeam -All | Select-Object @{N='Team Name';E={$_.DisplayName}},
            @{N='Team ID';E={$_.Id}},
            @{N='Visibility';E={$_.Visibility}},
            @{N='Archived';E={$_.IsArchived}},
            @{N='Description';E={$_.Description}}

        $activeTeams   = 0
        $totalMessages = 0

        try {
            $tempFile = Join-Path $env:TEMP "TeamsUsage_$(Get-Date -Format 'yyyyMMddHHmmss').csv"
            Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/reports/getTeamsTeamActivityDetail(period='D30')" -Method GET -OutputFilePath $tempFile
            if (Test-Path $tempFile) {
                $raw = Import-Csv $tempFile
                $teamsUsage = $raw | Select-Object @{N='Team Name';E={$_.'Team Name'}},
                    @{N='Last Activity';E={$_.'Last Activity Date'}},
                    @{N='Active Users';E={$_.'Active Users'}},
                    @{N='Active Channels';E={$_.'Active Channels'}},
                    @{N='Guests';E={$_.'Guests'}},
                    @{N='Channel Messages';E={$_.'Channel Messages'}},
                    @{N='Meetings Organized';E={$_.'Meetings Organized'}}
                $teamsUsage | Export-Csv -Path "$OutputPath\Teams_Usage.csv" -NoTypeInformation
                $activeTeams   = ($teamsUsage | Where-Object { $_.'Last Activity' -ne '' }).Count
                $totalMessages = ($teamsUsage | Measure-Object 'Channel Messages' -Sum).Sum
                Remove-Item $tempFile -Force -ErrorAction SilentlyContinue
            }
        } catch { }

        $teams | Export-Csv -Path "$OutputPath\Teams_List.csv" -NoTypeInformation

        $activeCount   = ($teams | Where-Object { !$_.Archived }).Count
        $archivedCount = ($teams | Where-Object { $_.Archived }).Count
        $privateCount  = ($teams | Where-Object { $_.Visibility -eq 'Private' }).Count
        $publicCount   = ($teams | Where-Object { $_.Visibility -eq 'Public' }).Count

        Write-Host "  Total: $($teams.Count) | Active: $activeCount | Archived: $archivedCount | Active (30d usage): $activeTeams" -ForegroundColor Green

        $usageRow = if ($activeTeams -gt 0) {
            "<div class='stat-row'>" +
            (New-StatCard "Active (30 days)" $activeTeams "#107c10") +
            (New-StatCard "Channel Messages (30d)" $totalMessages) +
            "</div>"
        } else { "" }

        $previewRows = $teams | Select-Object -First 25

        $html = "<div class='stat-row'>" +
                (New-StatCard "Total Teams" $teams.Count) +
                (New-StatCard "Active" $activeCount "#107c10") +
                (New-StatCard "Archived" $archivedCount "#767676") +
                (New-StatCard "Private / Public" "$privateCount / $publicCount") +
                "</div>" + $usageRow +
                "<h3>Teams List</h3>" +
                (ConvertTo-HtmlTable $previewRows @('Team Name','Visibility','Archived','Description')) +
                $(if ($teams.Count -gt 25) { "<p><em>Showing 25 of $($teams.Count). See Teams_List.csv for full list.</em></p>" } else { "" })

        Add-HtmlSection "Microsoft Teams" $html
    }
    catch {
        Write-Host "  Error: $_" -ForegroundColor Red
        Add-HtmlSection "Microsoft Teams" "<p class='error'>Error: $_</p>" "bad"
    }
}

# ═════════════════════════════════════════════════════════════════════════════
# 11. Entra ID Applications
# ═════════════════════════════════════════════════════════════════════════════
function Get-EntraIDApps {
    param([string]$OutputPath, [int]$Total)
    Write-SectionHeader "Entra ID Applications" $Total

    try {
        $apps = Get-MgApplication -All | Select-Object @{N='App Name';E={$_.DisplayName}},
            @{N='App ID';E={$_.AppId}},
            @{N='Created';E={if ($_.CreatedDateTime) { $_.CreatedDateTime.ToString('yyyy-MM-dd') } else { 'N/A' }}},
            @{N='Sign-In Audience';E={$_.SignInAudience}},
            @{N='Publisher Domain';E={$_.PublisherDomain}},
            @{N='API Permission Count';E={($_.RequiredResourceAccess.ResourceAccess | Measure-Object).Count}}

        $sps = Get-MgServicePrincipal -All | Select-Object @{N='App Name';E={$_.DisplayName}},
            @{N='App ID';E={$_.AppId}},
            @{N='Type';E={$_.ServicePrincipalType}},
            @{N='Enabled';E={$_.AccountEnabled}},
            @{N='Publisher';E={$_.PublisherName}},
            @{N='Verified Publisher';E={$_.VerifiedPublisher.DisplayName}}

        $apps | Export-Csv -Path "$OutputPath\EntraID_AppRegistrations.csv" -NoTypeInformation
        $sps  | Export-Csv -Path "$OutputPath\EntraID_ServicePrincipals.csv" -NoTypeInformation

        $thirdParty    = $sps | Where-Object { $_.Type -eq 'Application' -and $_.Publisher -ne 'Microsoft Corporation' -and $_.Publisher }
        $microsoftApps = ($sps | Where-Object { $_.Publisher -eq 'Microsoft Corporation' }).Count

        Write-Host "  App Registrations: $($apps.Count) | Service Principals: $($sps.Count) | Third-Party: $($thirdParty.Count)" -ForegroundColor Green

        $thirdPartyWarn = if ($thirdParty.Count -gt 30) {
            "<p class='warn-text'>$($thirdParty.Count) third-party applications detected. Review these for unnecessary permissions and unused integrations.</p>"
        } else { "" }

        $html = "<div class='stat-row'>" +
                (New-StatCard "App Registrations" $apps.Count) +
                (New-StatCard "Service Principals" $sps.Count) +
                (New-StatCard "Third-Party Apps" $thirdParty.Count $(if ($thirdParty.Count -gt 30) { "#ff8c00" } else { "#767676" })) +
                (New-StatCard "Microsoft Apps" $microsoftApps "#767676") +
                "</div><h3>Third-Party Applications</h3>" +
                (ConvertTo-HtmlTable ($thirdParty | Sort-Object 'App Name' | Select-Object -First 20) @('App Name','Publisher','Type','Enabled','Verified Publisher')) +
                $(if ($thirdParty.Count -gt 20) { "<p><em>Showing 20 of $($thirdParty.Count). See EntraID_ServicePrincipals.csv for full list.</em></p>" } else { "" }) +
                $thirdPartyWarn

        Add-HtmlSection "Entra ID Applications" $html
    }
    catch {
        Write-Host "  Error: $_" -ForegroundColor Red
        Add-HtmlSection "Entra ID Applications" "<p class='error'>Error: $_</p>" "bad"
    }
}

# ═════════════════════════════════════════════════════════════════════════════
# 12. Compliance Posture
# ═════════════════════════════════════════════════════════════════════════════
function Get-CompliancePosture {
    param([string]$OutputPath, [int]$Total)
    Write-SectionHeader "Compliance Posture" $Total

    try {
        # Sensitivity labels
        $labelCount = 0
        try {
            $resp = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/security/informationProtection/sensitivityLabels" -Method GET -ErrorAction Stop
            $labelCount = $resp.value.Count
            if ($labelCount -gt 0) {
                ($resp.value | Select-Object @{N='Label Name';E={$_.name}},
                    @{N='Order';E={$_.sensitivityOrder}},
                    @{N='Active';E={$_.isActive}},
                    @{N='Description';E={$_.description}}) |
                    Export-Csv -Path "$OutputPath\Compliance_SensitivityLabels.csv" -NoTypeInformation
            }
        } catch {
            Write-Host "  Sensitivity labels not accessible" -ForegroundColor Yellow
        }

        # Audit log accessibility
        $auditStatus = 'Unknown'
        try {
            Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/security/auditLog/queries" -Method GET -ErrorAction Stop | Out-Null
            $auditStatus = 'Accessible'
        } catch { }

        # Check if Purview DLP policies are accessible (requires Compliance admin)
        $dlpNote = 'Requires Compliance Administrator role — not assessed with Global Reader'

        Write-Host "  Sensitivity Labels: $labelCount | Audit Log: $auditStatus" -ForegroundColor Green

        $checks = @(
            [PSCustomObject]@{ 'Control'='Sensitivity Labels';  'Result'=$labelCount; 'Status'=if($labelCount -gt 0){'Configured'}else{'None found'}; 'Risk'=if($labelCount -eq 0){'High'}else{'Low'} }
            [PSCustomObject]@{ 'Control'='Audit Log';           'Result'=$auditStatus; 'Status'=$auditStatus; 'Risk'=if($auditStatus -eq 'Accessible'){'Low'}else{'Review'} }
            [PSCustomObject]@{ 'Control'='DLP Policies';        'Result'='N/A'; 'Status'=$dlpNote; 'Risk'='Review in Purview' }
            [PSCustomObject]@{ 'Control'='Retention Policies';  'Result'='N/A'; 'Status'='Requires Compliance Administrator role'; 'Risk'='Review in Purview' }
        )

        $status = if ($labelCount -gt 0) { "good" } else { "warn" }

        $html = "<div class='stat-row'>" +
                (New-StatCard "Sensitivity Labels" $labelCount $(if ($labelCount -gt 0) { "#107c10" } else { "#ff8c00" })) +
                (New-StatCard "Audit Log" $auditStatus $(if ($auditStatus -eq 'Accessible') { "#107c10" } else { "#ff8c00" })) +
                "</div><h3>Compliance Checks</h3>" +
                (ConvertTo-HtmlTable $checks @('Control','Result','Status','Risk')) +
                "<p class='warn-text'>Note: DLP policies, retention policies, and communication compliance require Compliance Administrator role and cannot be read with Global Reader alone. Review these in the Microsoft Purview portal.</p>"

        Add-HtmlSection "Compliance Posture" $html $status
    }
    catch {
        Write-Host "  Error: $_" -ForegroundColor Red
        Add-HtmlSection "Compliance Posture" "<p class='error'>Error: $_</p>" "bad"
    }
}

# ═════════════════════════════════════════════════════════════════════════════
# HTML Report Generator
# ═════════════════════════════════════════════════════════════════════════════
function Generate-HtmlReport {
    param([string]$OutputPath)
    Write-Host "`nGenerating HTML report..." -ForegroundColor Cyan

    $orgName    = Get-DataValue 'OrgName'
    $tenantId   = Get-DataValue 'TenantId'
    $assessDate = Get-Date -Format "yyyy-MM-dd HH:mm"
    $title      = if ($CustomerName) { "$CustomerName - M365 Assessment" } else { "$orgName - M365 Assessment" }

    $scorePct = Get-DataValue 'SecureScorePct'
    $scoreNum = Get-DataValue 'SecureScore'
    $scoreMax = Get-DataValue 'SecureScoreMax'
    $scoreColor = if ($scorePct -ne 'N/A' -and [double]$scorePct -ge 80) { "#107c10" }
                  elseif ($scorePct -ne 'N/A' -and [double]$scorePct -ge 50) { "#ff8c00" }
                  else { "#d13438" }

    $sectionsHtml = $script:HtmlSections -join "`n"

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>$title</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Segoe UI',Arial,sans-serif;font-size:14px;background:#f3f2f1;color:#323130}
.header{background:#0078d4;color:#fff;padding:24px 32px}
.header h1{font-size:22px;font-weight:600}
.header .meta{margin-top:8px;font-size:13px;opacity:.9}
.header .meta span{margin-right:20px}
.container{max-width:1200px;margin:0 auto;padding:20px 16px}
.overview-bar{display:flex;gap:12px;flex-wrap:wrap;margin-bottom:20px}
.ov-card{background:#fff;border-radius:4px;padding:14px 18px;flex:1;min-width:120px;box-shadow:0 1px 3px rgba(0,0,0,.1);text-align:center}
.ov-value{font-size:26px;font-weight:700;color:#0078d4}
.ov-label{font-size:11px;color:#605e5c;margin-top:3px}
.section{background:#fff;border-radius:4px;margin-bottom:14px;padding:18px 22px;box-shadow:0 1px 3px rgba(0,0,0,.1);border-left:4px solid #0078d4}
.section-good{border-left-color:#107c10}
.section-warn{border-left-color:#ff8c00}
.section-bad{border-left-color:#d13438}
.section h2{font-size:15px;font-weight:600;margin-bottom:14px}
.section h3{font-size:12px;font-weight:600;text-transform:uppercase;letter-spacing:.5px;color:#605e5c;margin:14px 0 7px}
.stat-row{display:flex;gap:10px;flex-wrap:wrap;margin-bottom:10px}
.stat-card{background:#faf9f8;border-radius:4px;padding:10px 14px;flex:1;min-width:120px;border-top:4px solid #0078d4}
.stat-value{font-size:18px;font-weight:700}
.stat-label{font-size:11px;color:#605e5c;margin-top:2px}
table{width:100%;border-collapse:collapse;font-size:13px;margin-top:4px}
th{background:#f3f2f1;padding:7px 10px;text-align:left;font-weight:600;border-bottom:2px solid #edebe9}
td{padding:6px 10px;border-bottom:1px solid #edebe9}
tr:hover td{background:#faf9f8}
.no-data{color:#605e5c;font-style:italic;padding:6px 0}
.error{color:#d13438}
.warn-text{color:#7a4100;background:#fff4ce;border:1px solid #ffd335;border-radius:4px;padding:8px 12px;margin-top:8px;font-size:13px}
.footer{text-align:center;padding:20px;color:#605e5c;font-size:11px}
</style>
</head>
<body>
<div class="header">
  <h1>Microsoft 365 Tenant Assessment</h1>
  <div class="meta">
    <span><strong>Organization:</strong> $orgName</span>
    <span><strong>Tenant ID:</strong> $tenantId</span>
    <span><strong>Date:</strong> $assessDate</span>
    $(if ($CustomerName) { "<span><strong>Customer:</strong> $CustomerName</span>" })
    $(if ($QuickMode) { "<span><strong>Mode:</strong> Quick Assessment</span>" })
  </div>
</div>
<div class="container">
  <div class="overview-bar">
    <div class="ov-card"><div class="ov-value" style="color:$scoreColor">$(if($scoreNum -ne 'N/A'){"$scoreNum/$scoreMax"}else{'N/A'})</div><div class="ov-label">Secure Score</div></div>
    <div class="ov-card"><div class="ov-value" style="color:$scoreColor">$(if($scorePct -ne 'N/A'){"$scorePct%"}else{'N/A'})</div><div class="ov-label">Score %</div></div>
    <div class="ov-card"><div class="ov-value">$(Get-DataValue 'TotalUsers')</div><div class="ov-label">Total Users</div></div>
    <div class="ov-card"><div class="ov-value">$(Get-DataValue 'TotalLicenses')</div><div class="ov-label">Licenses Purchased</div></div>
    <div class="ov-card"><div class="ov-value">$(Get-DataValue 'UnusedLicenses')</div><div class="ov-label">Unused Licenses</div></div>
    <div class="ov-card"><div class="ov-value">$(Get-DataValue 'CAPolicies')</div><div class="ov-label">CA Policies (Active)</div></div>
    <div class="ov-card"><div class="ov-value">$(Get-DataValue 'GlobalAdmins')</div><div class="ov-label">Global Admins</div></div>
    <div class="ov-card"><div class="ov-value">$(Get-DataValue 'SecureScoreGaps')</div><div class="ov-label">Security Gaps</div></div>
  </div>
  $sectionsHtml
</div>
<div class="footer">M365 Tenant Assessment Tool v3.0 &bull; Generated $assessDate</div>
</body>
</html>
"@

    $reportPath = "$OutputPath\M365_Assessment_Report.html"
    $html | Out-File -FilePath $reportPath -Encoding UTF8
    Write-Host "  Report: $reportPath" -ForegroundColor Green
    return $reportPath
}

# ═════════════════════════════════════════════════════════════════════════════
# Main
# ═════════════════════════════════════════════════════════════════════════════
function Start-M365TenantAssessment {
    $modeLabel = if ($QuickMode) { "QUICK (Tenant Overview, Licensing, Secure Score)" } else { "FULL (12 sections)" }
    $custLabel = if ($CustomerName) { $CustomerName } else { "(no -CustomerName specified)" }

    Write-Host @"
=================================================================
       MICROSOFT 365 TENANT ASSESSMENT TOOL  v3.0
=================================================================
Customer : $custLabel
Mode     : $modeLabel
Required : Global Reader (minimum)
=================================================================
"@ -ForegroundColor Cyan

    Initialize-RequiredModules

    $outputPath = Initialize-OutputDirectory

    $graphOk = Connect-MicrosoftGraphServices
    if (!$graphOk) {
        Write-Host "Cannot connect to Microsoft Graph. Exiting." -ForegroundColor Red
        return
    }

    if (!$QuickMode) {
        Connect-ExchangeOnlineIfAvailable
    }

    $total = if ($QuickMode) { 3 } else { 12 }
    Write-Host "`n=== Starting Data Collection ($total sections) ===" -ForegroundColor Cyan

    # Always run
    Get-TenantOverview    -OutputPath $outputPath -Total $total
    Get-M365LicensingInfo -OutputPath $outputPath -Total $total
    Get-SecureScoreInfo   -OutputPath $outputPath -Total $total

    if (!$QuickMode) {
        Get-UserHealth         -OutputPath $outputPath -Total $total
        Get-EntraIDSecurity    -OutputPath $outputPath -Total $total
        Get-ExchangeOnlineInfo -OutputPath $outputPath -Total $total
        Get-EmailSecurity      -OutputPath $outputPath -Total $total
        Get-SharePointSites    -OutputPath $outputPath -Total $total
        Get-OneDriveUsage      -OutputPath $outputPath -Total $total
        Get-TeamsInfo          -OutputPath $outputPath -Total $total
        Get-EntraIDApps        -OutputPath $outputPath -Total $total
        Get-CompliancePosture  -OutputPath $outputPath -Total $total
    }

    $reportPath = Generate-HtmlReport -OutputPath $outputPath

    Write-Host "`nDisconnecting..." -ForegroundColor Cyan
    Disconnect-MgGraph -ErrorAction SilentlyContinue
    if ($script:ExchangeConnected) {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
    }

    Write-Host @"

=================================================================
  ASSESSMENT COMPLETE
  Output : $outputPath
  Report : $reportPath
=================================================================
"@ -ForegroundColor Green

    Start-Process explorer.exe $outputPath
    Start-Process $reportPath
}

Start-M365TenantAssessment
