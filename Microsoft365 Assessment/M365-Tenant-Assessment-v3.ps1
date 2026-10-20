#Requires -Version 5.1
<#
.SYNOPSIS
    Microsoft 365 Tenant Assessment Tool v3.1
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
    .\M365-Tenant-Assessment-v3.ps1 -CustomerName "Contoso"
.EXAMPLE
    .\M365-Tenant-Assessment-v3.ps1 -CustomerName "Fabrikam" -TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -QuickMode
#>
param(
    [string]$CustomerName = "",
    [string]$TenantId = "",
    [switch]$QuickMode,
    [string]$ClientId = ""   # Optional: App Registration client ID for pre-consented delegated access
)

$ErrorActionPreference = 'Continue'

# ── Script-level state ────────────────────────────────────────────────────────
$script:ExchangeConnected = $false
$script:HtmlSections      = [System.Collections.Generic.List[object]]::new()
$script:Findings          = [System.Collections.Generic.List[object]]::new()
$script:NotCollected      = [System.Collections.Generic.List[object]]::new()
$script:Data              = @{}
$script:SectionNum        = 0
$script:CurrentSection    = ''
$script:ToolVersion       = '3.1'

# Netsurit logo (PNG, base64) for the report header
$script:LogoBase64 = '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'

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
        Write-Host "  ExchangeOnlineManagement not found - shared mailbox data will use Graph heuristic" -ForegroundColor Yellow
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
        if ($TenantId)  { $connectParams['TenantId']  = $TenantId }
        if ($ClientId)  { $connectParams['ClientId']  = $ClientId }
        Connect-MgGraph @connectParams -ErrorAction Stop

        $ctx = Get-MgContext
        if (-not $ctx -or -not $ctx.Account) { throw "No Graph session was established (sign-in cancelled or failed)." }
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
        # DisableWAM: the Windows broker (WAM) throws NullReferenceException when no
        # console window handle is available (e.g. VS Code terminal) — use browser sign-in
        $connectParams = @{ ShowBanner = $false; DisableWAM = $true }
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
    $script:CurrentSection = $Title
    Write-Host "`n[$($script:SectionNum)/$Total] $Title..." -ForegroundColor Cyan
}

function ConvertTo-HtmlEncoded([object]$Value) {
    if ($null -eq $Value) { return '' }
    return [System.Net.WebUtility]::HtmlEncode("$Value")
}

function Get-SeverityRank([string]$s) {
    switch ($s) { 'Critical' { 0 } 'High' { 1 } 'Medium' { 2 } 'Low' { 3 } default { 4 } }
}

function Add-Finding {
    param(
        [Parameter(Mandatory)][string]$Id,
        [Parameter(Mandatory)][ValidateSet('Critical', 'High', 'Medium', 'Low', 'Info')][string]$Severity,
        [Parameter(Mandatory)][string]$Section,
        [Parameter(Mandatory)][string]$Title,
        [object]$Count = '',
        [string]$Recommendation = '',
        [string]$EvidenceFile = ''
    )
    $script:Findings.Add([PSCustomObject]@{
        Id = $Id; Severity = $Severity; Section = $Section; Title = $Title
        Count = $Count; Recommendation = $Recommendation; EvidenceFile = $EvidenceFile
    })
}

function Add-NotCollected {
    param([string]$Item, [string]$Reason, [string]$HowToFix)
    $short = if ($Reason -and $Reason.Length -gt 300) { $Reason.Substring(0, 300) + '...' } else { $Reason }
    $script:NotCollected.Add([PSCustomObject]@{
        Section = $script:CurrentSection; Item = $Item; Reason = $short; HowToFix = $HowToFix
    })
}

function Add-HtmlSection {
    # Sections are stored and rendered at report time, so each one can show the findings raised against it
    param([string]$Title, [string]$Content, [string]$Status = "info")
    $script:HtmlSections.Add([PSCustomObject]@{ Title = $Title; Content = $Content; Status = $Status })
}

function ConvertTo-HtmlTable {
    param([object[]]$Data, [string[]]$Properties = @(), [int]$MaxRows = 200, [string]$CsvName = '')
    $rows = @($Data | Where-Object { $null -ne $_ })
    if ($rows.Count -eq 0) { return "<p class='no-data'>None found</p>" }
    $cols = if ($Properties.Count -gt 0) { $Properties } else { $rows[0].PSObject.Properties.Name }
    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.Append('<table><thead><tr>')
    foreach ($col in $cols) { [void]$sb.Append("<th>$(ConvertTo-HtmlEncoded $col)</th>") }
    [void]$sb.Append('</tr></thead><tbody>')
    foreach ($row in ($rows | Select-Object -First $MaxRows)) {
        [void]$sb.Append('<tr>')
        foreach ($col in $cols) {
            $v = $row.$col
            if ($col -eq 'Severity') { [void]$sb.Append("<td><span class='sev sev-$v'>$(ConvertTo-HtmlEncoded $v)</span></td>") }
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
    # $Tone: info/good/warn/bad/neutral. Older hex colours are still accepted and mapped to a tone.
    param([string]$Label, [object]$Value, [string]$Tone = 'info')
    $t = switch ($Tone.ToLower()) {
        '#107c10' { 'good' } '#ff8c00' { 'warn' } '#d13438' { 'bad' } '#767676' { 'neutral' } '#0078d4' { 'info' }
        { $_ -in 'info', 'good', 'warn', 'bad', 'neutral' } { $_ }
        default { 'info' }
    }
    if ($null -eq $Value -or "$Value" -eq '') { $Value = 'N/A' }
    return "<div class='stat-card tone-$t'><div class='stat-value'>$(ConvertTo-HtmlEncoded $Value)</div><div class='stat-label'>$(ConvertTo-HtmlEncoded $Label)</div></div>"
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
        $script:Data['UnusedPct']      = $unusedPct

        Write-Host ("  {0} SKUs | {1}/{2} consumed | {3} available ({4}% unused)" -f $licenses.Count, $totalConsumed, $totalPurchased, $totalAvailable, $unusedPct) -ForegroundColor Green

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
        $secureScore = Get-MgSecuritySecureScore -ErrorAction Stop | Select-Object -First 1
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

        $controlProfiles = Get-MgSecuritySecureScoreControlProfile -ErrorAction Stop
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
        if ("$_" -match 'AADSTS500011') {
            # The Defender/Security Center service principal has never been provisioned in this tenant
            $msg = "Secure Score not available: the Microsoft Defender portal has not been initialised in this tenant. Sign in once at https://security.microsoft.com as an admin, wait a few hours, then re-run."
            Write-Host "  $msg" -ForegroundColor Yellow
            Add-NotCollected -Item 'Secure Score' -Reason 'Microsoft Defender portal not initialised in this tenant (AADSTS500011).' -HowToFix 'Sign in once at https://security.microsoft.com as an admin, wait a few hours, then re-run.'
            Add-HtmlSection "Secure Score" "<p class='warn-text'>$msg</p>" "warn"
        }
        else {
            Write-Host "  Error: $_" -ForegroundColor Red
            Add-NotCollected -Item 'Secure Score' -Reason "$_" -HowToFix 'Confirm SecurityEvents.Read.All was consented and the account has Security Reader or Global Reader.'
            Add-HtmlSection "Secure Score" "<p class='error'>Error: $(ConvertTo-HtmlEncoded $_)</p>" "bad"
        }
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
        # SignInActivity needs Entra ID P1/P2 — without it Graph rejects the whole request (403),
        # so fall back to fetching users without sign-in data
        $userProps     = @('Id','DisplayName','UserPrincipalName','UserType','AccountEnabled','AssignedLicenses','CreatedDateTime','Mail')
        $hasSignInData = $true
        try {
            $users = Get-MgUser -All -Property ($userProps + 'SignInActivity') -ErrorAction Stop
        }
        catch {
            Write-Host "  Sign-in activity not available (requires Entra ID P1/P2) — continuing without it" -ForegroundColor Yellow
            Add-NotCollected -Item 'User sign-in activity (inactive users)' -Reason 'Sign-in activity requires an Entra ID P1/P2 licence in the tenant.' -HowToFix 'Confirm inactive accounts with the customer, or re-run once P1/P2 is licensed.'
            $hasSignInData = $false
            $users = Get-MgUser -All -Property $userProps -ErrorAction Stop
        }

        $enabledMembers = $users | Where-Object { $_.AccountEnabled -and $_.UserType -eq 'Member' }
        $guestUsers     = $users | Where-Object { $_.UserType -eq 'Guest' }
        $licensedUsers  = ($users | Where-Object { $_.AssignedLicenses.Count -gt 0 }).Count
        $unlicensed     = ($enabledMembers | Where-Object { $_.AssignedLicenses.Count -eq 0 }).Count

        $thirtyDaysAgo = (Get-Date).AddDays(-30)
        $ninetyDaysAgo = (Get-Date).AddDays(-90)
        $newUsers      = ($users | Where-Object { $_.CreatedDateTime -gt $thirtyDaysAgo }).Count

        $inactive = if ($hasSignInData) {
            ($enabledMembers | Where-Object {
                $_.SignInActivity.LastSignInDateTime -and
                $_.SignInActivity.LastSignInDateTime -lt $ninetyDaysAgo
            }).Count
        } else { 'N/A' }

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
        $script:Data['UnlicensedActive'] = $unlicensed

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
                (New-StatCard "Inactive (90 days)" $inactive $(if (-not $hasSignInData) { "#767676" } elseif ($inactive -gt 10) { "#ff8c00" } else { "#107c10" })) +
                (New-StatCard "Global Admins" $globalAdmins $adminColor) +
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
        $caReadable   = $false
        try {
            $caPolicies   = Get-MgIdentityConditionalAccessPolicy -All -ErrorAction Stop
            $caReadable   = $true
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
            Add-NotCollected -Item 'Conditional Access policies' -Reason "$_" -HowToFix 'Consent Policy.Read.All and run as Global Reader or Security Reader.'
        }

        # Security defaults (the baseline MFA enforcement when Conditional Access is not used)
        $securityDefaults = $null
        try {
            $sd = Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/policies/identitySecurityDefaultsEnforcementPolicy" -Method GET -ErrorAction Stop
            $securityDefaults = [bool]$sd.isEnabled
        } catch {
            Add-NotCollected -Item 'Security defaults' -Reason "$_" -HowToFix 'Consent Policy.Read.All and re-run, or check Entra admin centre > Properties > Manage security defaults.'
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
            Add-NotCollected -Item 'MFA registration statistics' -Reason 'The authentication methods report requires Entra ID P1/P2.' -HowToFix 'Check MFA registration per user in the Entra admin centre, or re-run once P1/P2 is licensed.'
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
        $script:Data['CAReadable']    = $caReadable
        $script:Data['CAReportOnly']  = $caReportOnly
        $script:Data['SecurityDefaults'] = $securityDefaults
        $script:Data['MFAPct'] = if ($mfaTotal -ne 'N/A' -and [int]$mfaTotal -gt 0) { [math]::Round(([int]$mfaRegistered / [int]$mfaTotal) * 100, 1) } else { $null }

        Write-Host "  CA (enabled): $caEnabled | MFA registered: $mfaRegistered/$mfaTotal | Risky users: $riskyUsers" -ForegroundColor Green

        $caStatus = if ($caEnabled -eq 0 -and $securityDefaults) { "warn" } elseif ($caEnabled -eq 0) { "bad" } elseif ($caEnabled -lt 3) { "warn" } else { "good" }

        $mfaPct = if ($mfaTotal -ne 'N/A' -and [int]$mfaTotal -gt 0) {
            "$([math]::Round(([int]$mfaRegistered / [int]$mfaTotal) * 100, 1))%"
        } else { 'N/A' }

        $mfaColor = if ($mfaPct -eq 'N/A') { "#767676" } elseif ([double]($mfaPct -replace '%','') -lt 80) { "#d13438" } else { "#107c10" }

        $caPolicyRows = $caPolicies | Select-Object @{N='Policy Name';E={$_.DisplayName}}, @{N='State';E={$_.State}}

        $caWarn = if ($caEnabled -eq 0 -and $securityDefaults) {
            "<p class='warn-text'>No Conditional Access policies, but security defaults are enabled, so all users must register for MFA and legacy authentication is blocked. Conditional Access (Entra ID P1) gives finer control.</p>"
        } elseif ($caEnabled -eq 0) {
            "<p class='warn-text'>Critical: No enabled Conditional Access policies found. All sign-ins are unrestricted. This is the single highest-impact security gap to address.</p>"
        } elseif ($caReportOnly -gt 0) {
            "<p class='warn-text'>Note: $caReportOnly policies are in Report-Only mode. Review and enforce them to improve security posture.</p>"
        } else { "" }

        $html = "<div class='stat-row'>" +
                (New-StatCard "CA Policies (Enabled)" $caEnabled $(if ($caEnabled -eq 0) { "#d13438" } elseif ($caEnabled -lt 3) { "#ff8c00" } else { "#107c10" })) +
                (New-StatCard "CA Policies (Report-Only)" $caReportOnly "#ff8c00") +
                (New-StatCard "CA Policies (Disabled)" $caDisabled "#767676") +
                (New-StatCard "Risky Users" $riskyUsers $(if ($riskyUsers -gt 0) { "#d13438" } else { "#107c10" })) +
                (New-StatCard "Security Defaults" $(if ($null -eq $securityDefaults) { 'N/A' } elseif ($securityDefaults) { 'Enabled' } else { 'Disabled' }) $(if ($null -eq $securityDefaults) { 'neutral' } elseif ($securityDefaults -or $caEnabled -gt 0) { 'good' } else { 'bad' })) +
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
                 Where-Object { $null -ne $_.Mail }

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
        if (!$script:ExchangeConnected) {
            Add-NotCollected -Item 'Exchange Online (accurate shared mailbox count)' -Reason 'Not connected to Exchange Online.' -HowToFix 'Install the ExchangeOnlineManagement module and re-run; the shared mailbox count is an estimate.'
        }

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
            Add-HtmlSection "Email Security (DNS)" "<p class='warn-text'>Primary domain not available - skipped.</p>" "warn"
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

        # DKIM - check common selectors
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

        $script:Data['SPFStatus']   = $spfStatus
        $script:Data['DMARCStatus'] = $dmarcStatus
        $script:Data['DKIMStatus']  = $dkimStatus

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
        if (!$hasDKIM)  { $warnings += "<p class='warn-text'>DKIM not detected on common selectors. It may be configured on a custom selector - verify in your DNS provider.</p>" }

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
            Add-NotCollected -Item 'SharePoint site usage' -Reason "$_" -HowToFix 'Consent Reports.Read.All and run as Global Reader or Reports Reader.'
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
        $script:Data['ThirdPartyApps'] = @($thirdParty).Count

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
            $script:Data['SensitivityLabels'] = $labelCount
        } catch {
            Write-Host "  Sensitivity labels not accessible" -ForegroundColor Yellow
            Add-NotCollected -Item 'Sensitivity labels' -Reason 'Not readable with the granted permissions or licence (needs InformationProtectionPolicy.Read and Purview Information Protection).' -HowToFix 'Review labels in the Microsoft Purview portal.'
        }

        # Audit log accessibility
        $auditStatus = 'Unknown'
        try {
            Invoke-MgGraphRequest -Uri "https://graph.microsoft.com/v1.0/security/auditLog/queries" -Method GET -ErrorAction Stop | Out-Null
            $auditStatus = 'Accessible'
        } catch {
            Add-NotCollected -Item 'Audit log status' -Reason 'The audit log query API was not readable (needs AuditLogsQuery.Read.All).' -HowToFix 'Confirm unified audit logging is on in Microsoft Purview > Audit.'
        }

        # Check if Purview DLP policies are accessible (requires Compliance admin)
        $dlpNote = 'Requires Compliance Administrator role - not assessed with Global Reader'

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
# Findings (suggested severities, derived from the collected data)
# ═════════════════════════════════════════════════════════════════════════════
function Get-AssessmentFindings {
    $d = $script:Data
    $has = { param($k) $d.ContainsKey($k) -and $null -ne $d[$k] -and "$($d[$k])" -ne 'N/A' }

    # Identity
    if ((& $has 'CAReadable') -and $d['CAReadable'] -and [int]$d['CAPolicies'] -eq 0) {
        if ($d['SecurityDefaults'] -eq $false) {
            Add-Finding -Id 'M365-ID-001' -Severity Critical -Section 'Entra ID Security' -Title 'No MFA enforcement: no Conditional Access policies and security defaults are off' `
                -Recommendation 'Turn on security defaults now, or deploy Conditional Access policies requiring MFA for all users (Entra ID P1).' -EvidenceFile 'EntraID_ConditionalAccess.csv'
        }
        elseif ($d['SecurityDefaults'] -eq $true) {
            Add-Finding -Id 'M365-ID-002' -Severity Low -Section 'Entra ID Security' -Title 'Relying on security defaults instead of Conditional Access' `
                -Recommendation 'Security defaults cover the basics. Consider Entra ID P1 and Conditional Access for location, device and app-based controls.'
        }
        else {
            Add-Finding -Id 'M365-ID-001' -Severity High -Section 'Entra ID Security' -Title 'No enabled Conditional Access policies (security defaults status unknown)' `
                -Recommendation 'Confirm security defaults are on, or deploy Conditional Access policies requiring MFA.' -EvidenceFile 'EntraID_ConditionalAccess.csv'
        }
    }
    if ((& $has 'CAReportOnly') -and [int]$d['CAReportOnly'] -gt 0) {
        Add-Finding -Id 'M365-ID-003' -Severity Low -Section 'Entra ID Security' -Title 'Conditional Access policies left in report-only mode' -Count $d['CAReportOnly'] `
            -Recommendation 'Review the sign-in impact of each report-only policy and switch it on, or remove it.' -EvidenceFile 'EntraID_ConditionalAccess.csv'
    }
    if ((& $has 'MFAPct') -and [double]$d['MFAPct'] -lt 80) {
        $sev = if ([double]$d['MFAPct'] -lt 50) { 'High' } else { 'Medium' }
        Add-Finding -Id 'M365-ID-004' -Severity $sev -Section 'Entra ID Security' -Title "Only $($d['MFAPct'])% of users are registered for MFA" `
            -Recommendation 'Run an MFA registration campaign and require registration through Conditional Access or security defaults.'
    }
    if ((& $has 'RiskyUsers') -and [int]$d['RiskyUsers'] -gt 0) {
        Add-Finding -Id 'M365-ID-005' -Severity High -Section 'Entra ID Security' -Title 'Users flagged as risky by Entra ID Protection' -Count $d['RiskyUsers'] `
            -Recommendation 'Investigate each risky user, reset credentials where compromise is confirmed, and dismiss false positives.'
    }
    if (& $has 'GlobalAdmins') {
        $ga = [int]$d['GlobalAdmins']
        if ($ga -gt 5) {
            Add-Finding -Id 'M365-ID-006' -Severity Medium -Section 'User Health' -Title 'Too many Global Administrators' -Count $ga `
                -Recommendation 'Reduce to 2-4 Global Administrators and move other admins to least-privilege roles.' -EvidenceFile 'Users_AdminRoles.csv'
        }
        elseif ($ga -eq 1) {
            Add-Finding -Id 'M365-ID-007' -Severity Low -Section 'User Health' -Title 'Only one Global Administrator (no break-glass account)' -Count $ga `
                -Recommendation 'Create a second, cloud-only emergency access account excluded from MFA-blocking policies and monitored for sign-ins.' -EvidenceFile 'Users_AdminRoles.csv'
        }
    }
    if ((& $has 'InactiveUsers') -and [int]$d['InactiveUsers'] -gt 0) {
        Add-Finding -Id 'M365-ID-008' -Severity Low -Section 'User Health' -Title 'Enabled accounts with no sign-in for 90+ days' -Count $d['InactiveUsers'] `
            -Recommendation 'Confirm with the customer and disable or remove stale accounts; reclaim their licences.' -EvidenceFile 'Users.csv'
    }

    # Secure Score
    if ((& $has 'SecureScorePct') -and [double]$d['SecureScorePct'] -lt 70) {
        $sev = if ([double]$d['SecureScorePct'] -lt 40) { 'High' } else { 'Medium' }
        Add-Finding -Id 'M365-SS-001' -Severity $sev -Section 'Secure Score' -Title "Microsoft Secure Score is $($d['SecureScorePct'])%" -Count $d['SecureScoreGaps'] `
            -Recommendation 'Work through the high-impact recommendations first; most are configuration changes with no licence cost.' -EvidenceFile 'SecureScore_Recommendations.csv'
    }

    # Licensing
    if ((& $has 'UnusedPct') -and [double]$d['UnusedPct'] -gt 20) {
        Add-Finding -Id 'M365-LIC-001' -Severity Low -Section 'Licensing' -Title "$($d['UnusedPct'])% of purchased licences are unassigned" -Count $d['UnusedLicenses'] `
            -Recommendation 'Reduce licence counts at the next renewal or reassign them.' -EvidenceFile 'Licenses.csv'
    }
    if ((& $has 'UnlicensedActive') -and [int]$d['UnlicensedActive'] -gt 0) {
        Add-Finding -Id 'M365-LIC-002' -Severity Low -Section 'User Health' -Title 'Enabled member accounts without a licence' -Count $d['UnlicensedActive'] `
            -Recommendation 'Check whether these are service, shared or leaver accounts; disable sign-in where they are not used interactively.' -EvidenceFile 'Users.csv'
    }

    # Email authentication
    if (& $has 'SPFStatus') {
        switch -Wildcard ($d['SPFStatus']) {
            'Not Found'  { Add-Finding -Id 'M365-MAIL-001' -Severity High   -Section 'Email Security (DNS)' -Title 'No SPF record on the primary domain' -Recommendation 'Publish an SPF record ending in -all that lists Exchange Online (include:spf.protection.outlook.com) and any other senders.' -EvidenceFile 'EmailSecurity_DNS.csv' }
            'Permissive*'{ Add-Finding -Id 'M365-MAIL-001' -Severity High   -Section 'Email Security (DNS)' -Title 'SPF record allows any sender (+all or no all mechanism)' -Recommendation 'Change the SPF record to end in -all.' -EvidenceFile 'EmailSecurity_DNS.csv' }
            'Soft Fail*' { Add-Finding -Id 'M365-MAIL-002' -Severity Low    -Section 'Email Security (DNS)' -Title 'SPF uses soft fail (~all)' -Recommendation 'Once DMARC is enforced, consider moving SPF to -all.' -EvidenceFile 'EmailSecurity_DNS.csv' }
        }
    }
    if (& $has 'DMARCStatus') {
        if ($d['DMARCStatus'] -eq 'Not Found') {
            Add-Finding -Id 'M365-MAIL-003' -Severity High -Section 'Email Security (DNS)' -Title 'No DMARC record on the primary domain' `
                -Recommendation 'Publish DMARC with p=none and reporting, review the reports, then move to p=quarantine and p=reject.' -EvidenceFile 'EmailSecurity_DNS.csv'
        }
        elseif ($d['DMARCStatus'] -like 'None*') {
            Add-Finding -Id 'M365-MAIL-004' -Severity Medium -Section 'Email Security (DNS)' -Title 'DMARC is published but not enforced (p=none)' `
                -Recommendation 'Review DMARC reports, then move the policy to p=quarantine and p=reject so spoofed mail is blocked.' -EvidenceFile 'EmailSecurity_DNS.csv'
        }
    }
    if ((& $has 'DKIMStatus') -and $d['DKIMStatus'] -eq 'Not detected') {
        Add-Finding -Id 'M365-MAIL-005' -Severity Medium -Section 'Email Security (DNS)' -Title 'DKIM not detected on common selectors' `
            -Recommendation 'Enable DKIM signing for the domain in Microsoft Defender > Email authentication settings (it may use a custom selector).' -EvidenceFile 'EmailSecurity_DNS.csv'
    }

    # Apps and compliance
    if ((& $has 'ThirdPartyApps') -and [int]$d['ThirdPartyApps'] -gt 30) {
        Add-Finding -Id 'M365-APP-001' -Severity Low -Section 'Entra ID Applications' -Title 'Large number of third-party applications' -Count $d['ThirdPartyApps'] `
            -Recommendation 'Review third-party apps and their permissions; remove unused ones and restrict user consent.' -EvidenceFile 'EntraID_ServicePrincipals.csv'
    }
    if ((& $has 'SensitivityLabels') -and [int]$d['SensitivityLabels'] -eq 0) {
        Add-Finding -Id 'M365-CMP-001' -Severity Low -Section 'Compliance Posture' -Title 'No sensitivity labels published' `
            -Recommendation 'Define a basic labelling scheme (e.g. Public, Internal, Confidential) in Microsoft Purview if the licences include it.'
    }
}

# ═════════════════════════════════════════════════════════════════════════════
# HTML Report Generator
# ═════════════════════════════════════════════════════════════════════════════
function New-HtmlReport {
    param([string]$OutputPath)
    Write-Host "`nGenerating HTML report..." -ForegroundColor Cyan

    Get-AssessmentFindings

    $d          = $script:Data
    $orgName    = Get-DataValue 'OrgName'
    $tenantId   = Get-DataValue 'TenantId'
    $customer   = if ($CustomerName) { $CustomerName } else { $orgName }
    $generated  = Get-Date -Format 'yyyy-MM-dd HH:mm'
    $modeText   = if ($QuickMode) { 'Quick assessment' } else { 'Full assessment' }
    $account    = (Get-MgContext).Account

    # Evidence files
    $findingCols = 'Id', 'Severity', 'Section', 'Title', 'Count', 'Recommendation', 'EvidenceFile'
    if ($script:Findings.Count) { $script:Findings | Select-Object $findingCols | Export-Csv -Path "$OutputPath\Findings.csv" -NoTypeInformation }
    if ($script:NotCollected.Count) { $script:NotCollected | Export-Csv -Path "$OutputPath\NotCollected.csv" -NoTypeInformation }

    # Headline cards
    $pctTone = { param($v) if ($null -eq $v -or "$v" -eq 'N/A') { 'neutral' } elseif ([double]$v -ge 80) { 'good' } elseif ([double]$v -ge 50) { 'warn' } else { 'bad' } }
    $scorePct = if ($d.ContainsKey('SecureScorePct')) { $d['SecureScorePct'] } else { $null }
    $mfaPct   = if ($d.ContainsKey('MFAPct')) { $d['MFAPct'] } else { $null }
    $caText   = if ($d.ContainsKey('CAPolicies')) {
                    if ([int]$d['CAPolicies'] -gt 0) { "$($d['CAPolicies']) enabled" }
                    elseif ($d['SecurityDefaults'] -eq $true) { 'Security defaults' }
                    else { 'None' }
                } else { 'N/A' }
    $caTone   = switch ($caText) { 'N/A' { 'neutral' } 'None' { 'bad' } 'Security defaults' { 'warn' } default { 'good' } }
    $dmarc    = Get-DataValue 'DMARCStatus'
    $dmarcTone = if ($dmarc -match 'enforced\)$' -and $dmarc -notmatch 'not enforced') { 'good' } elseif ($dmarc -eq 'N/A') { 'neutral' } elseif ($dmarc -match 'None') { 'warn' } else { 'bad' }
    $ga = Get-DataValue 'GlobalAdmins'
    $gaTone = if ($ga -eq 'N/A') { 'neutral' } elseif ([int]$ga -gt 5 -or [int]$ga -lt 2) { 'warn' } else { 'good' }

    $cards = "<div class='stat-row'>" +
        (New-StatCard 'Users' (Get-DataValue 'TotalUsers')) +
        (New-StatCard 'Licences used / purchased' "$(Get-DataValue 'UsedLicenses') / $(Get-DataValue 'TotalLicenses')") +
        (New-StatCard 'Secure Score' $(if ($null -ne $scorePct) { "$scorePct%" } else { 'N/A' }) (& $pctTone $scorePct)) +
        (New-StatCard 'MFA registered' $(if ($null -ne $mfaPct) { "$mfaPct%" } else { 'N/A' }) (& $pctTone $mfaPct)) +
        "</div><div class='stat-row'>" +
        (New-StatCard 'Conditional Access' $caText $caTone) +
        (New-StatCard 'Global Admins' $ga $gaTone) +
        (New-StatCard 'DMARC' $dmarc $dmarcTone) +
        (New-StatCard 'Guest users' (Get-DataValue 'GuestUsers')) +
        "</div>"

    $keyFindings = @($script:Findings | Where-Object { $_.Severity -ne 'Info' } | Sort-Object @{ E = { Get-SeverityRank $_.Severity } }, Section)
    $sevCounts = foreach ($s in 'Critical', 'High', 'Medium', 'Low') {
        $n = @($script:Findings | Where-Object { $_.Severity -eq $s }).Count
        if ($n) { "<span class='sev sev-$s'>$n $s</span>" }
    }

    $summaryHtml = "<section class='section summary' id='sec-summary'><h2>Executive summary</h2>" +
        "<p class='lead'>Read-only, point-in-time assessment of the Microsoft 365 tenant <strong>$(ConvertTo-HtmlEncoded $orgName)</strong> (<code>$(ConvertTo-HtmlEncoded $tenantId)</code>). " +
        "Findings carry a suggested severity; confirm context with the customer before they go into a report.</p>" +
        $cards +
        "<h3>Key findings $($sevCounts -join ' ')</h3>" +
        (ConvertTo-HtmlTable $keyFindings @('Severity', 'Section', 'Title', 'Count', 'Recommendation') -MaxRows 60 -CsvName 'Findings.csv') +
        "</section>"

    # Sections, each with the findings raised against it
    $toc = [System.Collections.Generic.List[string]]::new()
    $sectionsHtml = foreach ($s in $script:HtmlSections) {
        $anchor = 'sec-' + ($s.Title -replace '[^A-Za-z0-9]+', '-').ToLower().Trim('-')
        $toc.Add("<a href='#$anchor'>$(ConvertTo-HtmlEncoded $s.Title)</a>")
        $f = @($script:Findings | Where-Object { $_.Section -eq $s.Title -and $_.Severity -ne 'Info' } | Sort-Object { Get-SeverityRank $_.Severity })
        $sev = @($f | ForEach-Object { $_.Severity })
        $status = if ($sev -contains 'Critical' -or $sev -contains 'High') { 'bad' }
                  elseif ($sev -contains 'Medium') { 'warn' }
                  elseif ($s.Status -in 'good', 'warn', 'bad') { $s.Status }
                  else { 'good' }
        $findingsHtml = if ($f.Count) { "<h3>Findings</h3>" + (ConvertTo-HtmlTable $f @('Severity', 'Title', 'Count', 'Recommendation')) }
                        else { "<h3>Findings</h3><p class='no-data'>No issues flagged in this section.</p>" }
        "<section class='section $status' id='$anchor'><h2>$(ConvertTo-HtmlEncoded $s.Title)</h2>$($s.Content)$findingsHtml</section>"
    }

    $ncHtml = if ($script:NotCollected.Count) {
        "<section class='section warn' id='sec-not-collected'><h2>What could not be collected</h2>" +
        "<p>These gaps should become items on the customer information request.</p>" +
        (ConvertTo-HtmlTable $script:NotCollected @('Section', 'Item', 'Reason', 'HowToFix') -MaxRows 100 -CsvName 'NotCollected.csv') + "</section>"
    }
    else { "<section class='section good' id='sec-not-collected'><h2>What could not be collected</h2><p>Everything in scope was readable.</p></section>" }

    $tocHtml = "<nav class='toc no-print'><a href='#sec-summary'>Summary</a>" + ($toc -join '') + "<a href='#sec-not-collected'>Not collected</a></nav>"
    $logo = if ($script:LogoBase64) { "<img src='data:image/png;base64,$($script:LogoBase64)' alt='Netsurit'>" } else { "<strong>NETSURIT</strong>" }

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>$(ConvertTo-HtmlEncoded $customer) - Microsoft 365 Assessment</title>
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
.stat-value{font-size:19px;font-weight:700;word-break:break-word}
.stat-label{font-size:11px;color:var(--grey);margin-top:2px}
table{width:100%;border-collapse:collapse;font-size:12px;margin:4px 0}
th{background:var(--black);color:#fff;text-align:left;padding:6px 8px;font-weight:700}
td{padding:5px 8px;border-bottom:1px solid var(--line);vertical-align:top;word-break:break-word}
tbody tr:nth-child(even) td{background:#FBFBFB}
.sev{display:inline-block;padding:1px 8px;border-radius:10px;font-size:11px;font-weight:700;color:#fff;white-space:nowrap;margin-right:4px}
.sev-Critical{background:#8E0F1F}.sev-High{background:var(--red)}.sev-Medium{background:var(--orange)}.sev-Low{background:var(--yellow);color:var(--black)}.sev-Info{background:var(--blue)}
.no-data{color:var(--grey);font-style:italic;margin:4px 0}
.note{color:var(--grey);font-size:11px;font-style:italic;margin:6px 0}
.error{color:var(--red)}
.warn-text{background:#FFF4E5;border:1px solid var(--orange);border-radius:4px;padding:8px 12px;margin:10px 0}
footer{margin-top:32px;padding-top:12px;border-top:1px solid var(--line);color:var(--grey);font-size:11px;display:flex;justify-content:space-between;gap:16px}
footer .tag{color:var(--red);font-weight:700}
@media (max-width:760px){.page{padding:16px}header.brand{flex-direction:column;align-items:flex-start}header.brand .meta{text-align:left}}
@page{size:A4;margin:15mm 12mm}
@media print{
  body{background:#fff;color:#000;font-size:11px}
  .page{max-width:none;padding:0}
  .no-print,nav.toc{display:none!important}
  header.brand{border-bottom:2px solid #000}
  .section{border-left-width:3px;margin:16px 0;padding-left:12px}
  h2,h3{break-after:avoid;page-break-after:avoid}
  .stat-row,.stat-card{break-inside:avoid;page-break-inside:avoid}
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
  .warn-text{background:#fff;border-color:#000}
  code{background:none}
}
</style>
</head>
<body>
<div class="page">
<header class="brand">
  <div>$logo</div>
  <div class="meta">Netsurit Professional Services<br>$modeText &middot; $generated</div>
</header>
<h1>Microsoft 365 Assessment &ndash; $(ConvertTo-HtmlEncoded $customer)</h1>
<p class="subtitle">$(ConvertTo-HtmlEncoded $orgName) &middot; $(ConvertTo-HtmlEncoded (Get-DataValue 'PrimaryDomain')) &middot; tenant $(ConvertTo-HtmlEncoded $tenantId)$(if ($account) { " &middot; run as $(ConvertTo-HtmlEncoded $account)" })</p>
$tocHtml
$summaryHtml
$($sectionsHtml -join "`n")
$ncHtml
<footer>
  <div>Netsurit Professional Services &middot; Microsoft 365 Tenant Assessment v$($script:ToolVersion) &middot; Read-only collection; evidence files are in the same folder as this report.</div>
  <div class="tag">Supporting the Dreams of the Doers.</div>
</footer>
</div>
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
       MICROSOFT 365 TENANT ASSESSMENT TOOL  v3.1
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

    $reportPath = New-HtmlReport -OutputPath $outputPath

    Write-Host "`nDisconnecting..." -ForegroundColor Cyan
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
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
