# Microsoft 365 Tenant Assessment Script - Graph API Only Version
# Version: 2.0
# Purpose: Gather comprehensive tenant information for migration and security assessments
# Required: Global Reader permissions in the target tenant
# Issues: 
# Secure score information may be incomplete or unavailable, trying to find a way to export the raw report from the tenant
# Double check the shared mailbox info, does not look accurate

#Requires -Version 5.1

# Function to install and import required modules
function Initialize-RequiredModules {
    Write-Host "Checking and installing required modules..." -ForegroundColor Cyan
    
    $requiredModules = @(
        'Microsoft.Graph.Authentication',
        'Microsoft.Graph.Users',
        'Microsoft.Graph.Groups',
        'Microsoft.Graph.Sites',
        'Microsoft.Graph.Teams',
        'Microsoft.Graph.Applications',
        'Microsoft.Graph.Security',
        'Microsoft.Graph.Reports',
        'Microsoft.Graph.Identity.DirectoryManagement'
    )
    
    foreach ($module in $requiredModules) {
        if (!(Get-Module -ListAvailable -Name $module)) {
            Write-Host "Installing $module..." -ForegroundColor Yellow
            Install-Module -Name $module -Force -AllowClobber -Scope CurrentUser
        }
        Import-Module $module -ErrorAction SilentlyContinue
    }
}

# Function to create output directory
function Initialize-OutputDirectory {
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $downloadsPath = [Environment]::GetFolderPath("UserProfile") + "\Downloads"
    $outputPath = Join-Path $downloadsPath "M365_Assessment_$timestamp"
    
    if (!(Test-Path $outputPath)) {
        New-Item -ItemType Directory -Path $outputPath | Out-Null
    }
    
    Write-Host "Output directory created: $outputPath" -ForegroundColor Green
    return $outputPath
}

# Function to connect to Microsoft Graph
function Connect-MicrosoftGraphServices {
    Write-Host "`nConnecting to Microsoft Graph..." -ForegroundColor Cyan
    
    try {
        # Define required scopes for comprehensive access
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
            "Mail.Read",
            "MailboxSettings.Read",
            "Organization.Read.All"
        )
        
        Connect-MgGraph -Scopes $scopes -NoWelcome
        
        $context = Get-MgContext
        Write-Host "Connected to tenant: $($context.TenantId)" -ForegroundColor Green
        Write-Host "Account: $($context.Account)" -ForegroundColor Green
        
        return $true
    }
    catch {
        Write-Host "Failed to connect to Microsoft Graph: $_" -ForegroundColor Red
        return $false
    }
}

# 1. Get Microsoft 365 Licensing Information
function Get-M365LicensingInfo {
    param($OutputPath)
    
    Write-Host "`n[1/7] Gathering Microsoft 365 Licensing Information..." -ForegroundColor Cyan
    
    try {
        # SKU friendly name mapping
        $skuMapping = @{
            'O365_BUSINESS_ESSENTIALS' = 'Microsoft 365 Business Basic'
            'O365_BUSINESS_PREMIUM' = 'Microsoft 365 Business Standard'
            'SPB' = 'Microsoft 365 Business Premium'
            'O365_BUSINESS' = 'Microsoft 365 Apps for Business'
            'OFFICESUBSCRIPTION' = 'Microsoft 365 Apps for Enterprise'
            'STANDARDPACK' = 'Office 365 E1'
            'STANDARDWOFFPACK' = 'Office 365 E2'
            'ENTERPRISEPACK' = 'Office 365 E3'
            'ENTERPRISEPREMIUM' = 'Office 365 E5'
            'DESKLESSPACK' = 'Office 365 F3'
            'SPE_E3' = 'Microsoft 365 E3'
            'SPE_E5' = 'Microsoft 365 E5'
            'SPE_F1' = 'Microsoft 365 F1'
            'SPE_F3' = 'Microsoft 365 F3'
            'FLOW_FREE' = 'Power Automate Free'
            'POWER_BI_STANDARD' = 'Power BI (free)'
            'POWER_BI_PRO' = 'Power BI Pro'
            'TEAMS_EXPLORATORY' = 'Teams Exploratory'
            'EXCHANGESTANDARD' = 'Exchange Online Plan 1'
            'EXCHANGEENTERPRISE' = 'Exchange Online Plan 2'
            'PROJECTPROFESSIONAL' = 'Project Plan 3'
            'VISIOCLIENT' = 'Visio Plan 2'
            'STREAM' = 'Microsoft Stream'
            'ATP_ENTERPRISE' = 'Microsoft Defender for Office 365 (Plan 1)'
            'THREAT_INTELLIGENCE' = 'Microsoft Defender for Office 365 (Plan 2)'
            'EMS' = 'Enterprise Mobility + Security E3'
            'EMSPREMIUM' = 'Enterprise Mobility + Security E5'
            'AAD_PREMIUM' = 'Azure Active Directory Premium P1'
            'AAD_PREMIUM_P2' = 'Azure Active Directory Premium P2'
            'RIGHTSMANAGEMENT' = 'Azure Information Protection Plan 1'
            'ENTERPRISEPREMIUM_NOPSTNCONF' = 'Office 365 E5 without Audio Conferencing'
            'SHAREPOINTSTANDARD' = 'SharePoint Online Plan 1'
            'SHAREPOINTENTERPRISE' = 'SharePoint Online Plan 2'
            'MCOIMP' = 'Skype for Business Online (Plan 1)'
            'MCOSTANDARD' = 'Skype for Business Online (Plan 2)'
            'TEAMS_COMMERCIAL_TRIAL' = 'Microsoft Teams Commercial Cloud Trial'
            'WINDOWS_STORE' = 'Windows Store for Business'
            'DEVELOPERPACK' = 'Office 365 E3 Developer'
            'DEVELOPERPACK_E5' = 'Microsoft 365 E5 Developer'
        }
        
        $licenses = Get-MgSubscribedSku | Select-Object @{
            Name = 'SKU Part Number'; Expression = {$_.SkuPartNumber}
        }, @{
            Name = 'Friendly Name'; Expression = {
                $sku = $_.SkuPartNumber
                if ($skuMapping.ContainsKey($sku)) {
                    $skuMapping[$sku]
                } else {
                    $sku
                }
            }
        }, @{
            Name = 'SKU ID'; Expression = {$_.SkuId}
        }, @{
            Name = 'Total Licenses'; Expression = {$_.PrepaidUnits.Enabled}
        }, @{
            Name = 'Consumed Licenses'; Expression = {$_.ConsumedUnits}
        }, @{
            Name = 'Available Licenses'; Expression = {$_.PrepaidUnits.Enabled - $_.ConsumedUnits}
        }, @{
            Name = 'Warning Units'; Expression = {$_.PrepaidUnits.Warning}
        }, @{
            Name = 'Suspended Units'; Expression = {$_.PrepaidUnits.Suspended}
        }
        
        $licenses | Export-Csv -Path "$OutputPath\M365_Licenses.csv" -NoTypeInformation
        Write-Host "  ✓ Exported $($licenses.Count) license SKUs" -ForegroundColor Green
        
        # Create a summary report
        $summary = @"
Microsoft 365 Licensing Summary
Generated: $(Get-Date)
================================

Total License SKUs: $($licenses.Count)
Total Licenses Purchased: $(($licenses | Measure-Object -Property 'Total Licenses' -Sum).Sum)
Total Licenses Consumed: $(($licenses | Measure-Object -Property 'Consumed Licenses' -Sum).Sum)
Total Licenses Available: $(($licenses | Measure-Object -Property 'Available Licenses' -Sum).Sum)

Top 5 Most Consumed Licenses:
$($licenses | Sort-Object 'Consumed Licenses' -Descending | Select-Object -First 5 | ForEach-Object {"  - $($_.'Friendly Name'): $($_.'Consumed Licenses')/$($_.'Total Licenses')"} | Out-String)
"@
        
        $summary | Out-File -FilePath "$OutputPath\M365_Licenses_Summary.txt"
    }
    catch {
        Write-Host "  ✗ Error gathering licensing info: $_" -ForegroundColor Red
    }
}

# 2. Get Secure Score and Recommendations
function Get-SecureScoreInfo {
    param($OutputPath)
    
    Write-Host "`n[2/7] Gathering Secure Score Information..." -ForegroundColor Cyan
    
    try {
        # Get current secure score
        $secureScore = Get-MgSecuritySecureScore | Select-Object -First 1
        
        if ($secureScore) {
            $scoreData = @{
                'Assessment Date' = Get-Date
                'Current Score' = $secureScore.CurrentScore
                'Max Score' = $secureScore.MaxScore
                'Score Percentage' = [math]::Round(($secureScore.CurrentScore / $secureScore.MaxScore) * 100, 2)
                'Licensed User Count' = $secureScore.LicensedUserCount
                'Active User Count' = $secureScore.ActiveUserCount
            }
            
            $scoreData | Export-Csv -Path "$OutputPath\SecureScore_Current.csv" -NoTypeInformation
            Write-Host "  ✓ Current Secure Score: $($scoreData['Current Score'])/$($scoreData['Max Score']) ($($scoreData['Score Percentage'])%)" -ForegroundColor Green
        }
        
        # Get secure score control profiles (recommendations) with enhanced columns
        $controlProfiles = Get-MgSecuritySecureScoreControlProfile
        
        # Calculate the maximum possible score for percentage calculations
        $maxPossibleScore = ($controlProfiles | Measure-Object -Property MaxScore -Sum).Sum
        if ($maxPossibleScore -eq 0) { $maxPossibleScore = 1 } # Prevent division by zero
        
        # Sort by MaxScore descending and add rank
        $sortedProfiles = $controlProfiles | Sort-Object -Property MaxScore -Descending
        $rankCounter = 0
        
        $recommendations = $sortedProfiles | ForEach-Object {
            $rankCounter++
            
            # Determine the product based on control title, ID, and category
            $product = 'Microsoft 365'
            $controlId = $_.Id
            $title = $_.Title
            $category = $_.ControlCategory
            
            # More accurate product mapping based on control patterns from your example
            if ($title -match 'Defender Antivirus|Windows Defender|Antivirus|malware' -or 
                $controlId -match 'DefenderAntivirus|WindowsDefender') {
                $product = 'Defender for Endpoint'
            }
            elseif ($title -match 'Entra ID|Azure AD|MFA|authentication|password|sign.?in|identity|consent' -or 
                    $category -eq 'Identity' -or
                    $controlId -match 'AAD|Entra|Identity|MFA|Authentication') {
                $product = 'Microsoft Entra ID'
            }
            elseif ($title -match 'BitLocker|encryption|device|endpoint|Intune|AutoRun|UAC|firewall|network protection' -or 
                    $category -eq 'Device' -or
                    $controlId -match 'Endpoint|Device|Intune|MDM|BitLocker') {
                $product = 'Defender for Endpoint'
            }
            elseif ($title -match 'phishing|spam|safe.*attachment|safe.*link|spoof|DKIM|DMARC|SPF' -or
                    $controlId -match 'AntiPhishing|SafeAttachment|SafeLink|EmailAuth') {
                $product = 'Defender for Office'
            }
            elseif ($title -match 'DLP|data loss|information protection|sensitivity|IRM' -or
                    $controlId -match 'InformationProtection|DLP|IRM') {
                $product = 'Microsoft Information Protection'
            }
            elseif ($title -match 'Exchange|mailbox|email|mail flow|journaling' -or
                    $controlId -match 'Exchange|Mailbox|EmailForwarding') {
                $product = 'Exchange Online'
            }
            elseif ($title -match 'SharePoint|OneDrive|external sharing|file' -or
                    $controlId -match 'SharePoint|OneDrive|SPO') {
                $product = 'SharePoint Online'
            }
            elseif ($title -match 'Teams|meeting|chat|communication' -or
                    $controlId -match 'Teams|Meeting') {
                $product = 'Microsoft Teams'
            }
            elseif ($title -match 'Cloud App|MCAS|shadow IT' -or
                    $controlId -match 'CloudApp|MCAS') {
                $product = 'Defender for Cloud Apps'
            }
            elseif ($title -match 'Power Platform|Power Apps|Power Automate|Power BI' -or
                    $controlId -match 'PowerPlatform|PowerApps') {
                $product = 'Power Platform'
            }
            elseif ($title -match 'Azure|subscription|resource' -or
                    $controlId -match 'Azure|Subscription') {
                $product = 'Azure Security'
            }
            elseif ($category -eq 'Apps') {
                # Default Apps category items
                if ($title -match 'Office|Word|Excel|PowerPoint|Outlook') {
                    $product = 'Microsoft 365 Apps'
                } else {
                    $product = 'Microsoft 365 Apps'
                }
            }
            elseif ($category -eq 'Data') {
                # Default Data category items
                if ($title -match 'retention|archive|audit') {
                    $product = 'Microsoft 365 Compliance'
                } else {
                    $product = 'Microsoft Information Protection'
                }
            }
            
            # Calculate score impact as percentage
            $scoreImpactPercent = [math]::Round(($_.MaxScore / $maxPossibleScore) * 100, 2)
            
            [PSCustomObject]@{
                'Rank' = $rankCounter
                'Recommended action' = $_.Title
                'Score impact' = "$scoreImpactPercent%"
                'Status' = if ($_.ImplementationStatus) { $_.ImplementationStatus } else { 'To address' }
                'Regressed' = if ($_.ComplianceInformation) {
                    $compliance = $_.ComplianceInformation | Where-Object {$_.ComplianceStatus -eq 'Regressed'}
                    if ($compliance) { 'Yes' } else { 'No' }
                } else { 'No' }
                'Have license?' = if ($_.ControlStateUpdates -and $_.ControlStateUpdates.Count -gt 0) {
                    'Yes'
                } elseif ($_.IsApplicable -eq $false) {
                    'No'
                } else {
                    'Yes'
                }
                'Category' = if ($category) { $category } else { 'Security' }
                'Product' = $product
                'Last synced' = if ($_.LastModifiedDateTime) {
                    $_.LastModifiedDateTime.ToString('dd MM yyyy')
                } else {
                    (Get-Date).ToString('dd MM yyyy')
                }
                'User Impact' = if ($_.UserImpact) { $_.UserImpact } else { 'Low' }
                'Max Score Points' = $_.MaxScore
            }
        }
        
        # Export without the helper column
        $recommendations | Select-Object -Property Rank, 'Recommended action', 'Score impact', Status, Regressed, 'Have license?', Category, Product, 'Last synced', 'User Impact' |
            Export-Csv -Path "$OutputPath\SecureScore_Recommendations.csv" -NoTypeInformation
        
        Write-Host "  ✓ Exported $($recommendations.Count) security recommendations" -ForegroundColor Green
        
        # Create recommendations summary
        $notImplemented = $recommendations | Where-Object {$_.Status -ne 'Completed' -and $_.Status -ne 'Resolved'}
        $highImpact = $notImplemented | Where-Object {
            $_.'Max Score Points' -ge 10
        }
        
        $summary = @"
Secure Score Recommendations Summary
=====================================

Total Recommendations: $($recommendations.Count)
Implemented: $(($recommendations | Where-Object {$_.Status -eq 'Completed' -or $_.Status -eq 'Resolved'}).Count)
Not Implemented: $($notImplemented.Count)
High Impact (10+ points) Not Implemented: $($highImpact.Count)

Top 10 High-Value Improvements:
$($highImpact | Select-Object -First 10 | ForEach-Object {"  - $($_.'Recommended action'): $($_.'Score impact') impact ($($_.'Max Score Points') points)"} | Out-String)

Product Breakdown:
$($recommendations | Group-Object Product | Sort-Object Count -Descending | ForEach-Object {"  - $($_.Name): $($_.Count) recommendations"} | Out-String)

Category Breakdown:
$($recommendations | Group-Object Category | Sort-Object Count -Descending | ForEach-Object {"  - $($_.Name): $($_.Count) recommendations"} | Out-String)
"@
        
        $summary | Out-File -FilePath "$OutputPath\SecureScore_Summary.txt"
    }
    catch {
        Write-Host "  ✗ Error gathering secure score: $_" -ForegroundColor Red
    }
}

# 3. Get Exchange Online Information
function Get-ExchangeOnlineInfo {
    param($OutputPath)
    
    Write-Host "`n[3/7] Gathering Exchange Online Information..." -ForegroundColor Cyan
    
    try {
        # Get all users with mailbox information
        Write-Host "  - Gathering user mailboxes..." -ForegroundColor Gray
        $users = Get-MgUser -All -Property DisplayName,Mail,UserPrincipalName,UserType,AccountEnabled,LicenseDetails,AssignedLicenses,ProxyAddresses | 
                 Where-Object { $_.Mail -ne $null }
        
        $mailboxes = $users | Select-Object @{
            Name = 'Display Name'; Expression = {$_.DisplayName}
        }, @{
            Name = 'Email Address'; Expression = {$_.Mail}
        }, @{
            Name = 'User Principal Name'; Expression = {$_.UserPrincipalName}
        }, @{
            Name = 'Mailbox Type'; Expression = {'User Mailbox'}
        }, @{
            Name = 'Account Enabled'; Expression = {$_.AccountEnabled}
        }, @{
            Name = 'User Type'; Expression = {$_.UserType}
        }
        
        # Get all groups
        Write-Host "  - Gathering groups..." -ForegroundColor Gray
        $allGroups = Get-MgGroup -All -Property DisplayName,Mail,MailEnabled,SecurityEnabled,GroupTypes,MembershipRule,ProxyAddresses,Description
        
        # Separate different group types
        $m365Groups = $allGroups | Where-Object { $_.GroupTypes -contains "Unified" } | Select-Object @{
            Name = 'Display Name'; Expression = {$_.DisplayName}
        }, @{
            Name = 'Email Address'; Expression = {$_.Mail}
        }, @{
            Name = 'Group Type'; Expression = {'Microsoft 365 Group'}
        }, @{
            Name = 'Description'; Expression = {$_.Description}
        }, @{
            Name = 'Mail Enabled'; Expression = {$_.MailEnabled}
        }, @{
            Name = 'Security Enabled'; Expression = {$_.SecurityEnabled}
        }
        
        $distGroups = $allGroups | Where-Object { $_.MailEnabled -eq $true -and $_.SecurityEnabled -eq $false -and $_.GroupTypes.Count -eq 0 } | Select-Object @{
            Name = 'Display Name'; Expression = {$_.DisplayName}
        }, @{
            Name = 'Email Address'; Expression = {$_.Mail}
        }, @{
            Name = 'Group Type'; Expression = {'Distribution Group'}
        }, @{
            Name = 'Description'; Expression = {$_.Description}
        }
        
        $securityGroups = $allGroups | Where-Object { $_.SecurityEnabled -eq $true -and $_.MailEnabled -eq $false } | Select-Object @{
            Name = 'Display Name'; Expression = {$_.DisplayName}
        }, @{
            Name = 'Group Type'; Expression = {'Security Group'}
        }, @{
            Name = 'Description'; Expression = {$_.Description}
        }
        
        $mailSecurityGroups = $allGroups | Where-Object { $_.SecurityEnabled -eq $true -and $_.MailEnabled -eq $true -and $_.GroupTypes.Count -eq 0 } | Select-Object @{
            Name = 'Display Name'; Expression = {$_.DisplayName}
        }, @{
            Name = 'Email Address'; Expression = {$_.Mail}
        }, @{
            Name = 'Group Type'; Expression = {'Mail-Enabled Security Group'}
        }, @{
            Name = 'Description'; Expression = {$_.Description}
        }
        
        # Get shared mailboxes (these are typically disabled user accounts with specific licenses)
        Write-Host "  - Identifying shared mailboxes..." -ForegroundColor Gray
        $sharedMailboxes = $users | Where-Object { $_.AccountEnabled -eq $false } | Select-Object @{
            Name = 'Display Name'; Expression = {$_.DisplayName}
        }, @{
            Name = 'Email Address'; Expression = {$_.Mail}
        }, @{
            Name = 'User Principal Name'; Expression = {$_.UserPrincipalName}
        }, @{
            Name = 'Mailbox Type'; Expression = {'Shared Mailbox (Potential)'}
        }, @{
            Name = 'Account Enabled'; Expression = {$_.AccountEnabled}
        }
        
        # Export all Exchange-related data
        $mailboxes | Export-Csv -Path "$OutputPath\Exchange_Mailboxes.csv" -NoTypeInformation
        
        if ($sharedMailboxes.Count -gt 0) {
            $sharedMailboxes | Export-Csv -Path "$OutputPath\Exchange_SharedMailboxes.csv" -NoTypeInformation
        }
        
        $allGroupsExport = $m365Groups + $distGroups + $mailSecurityGroups
        if ($allGroupsExport.Count -gt 0) {
            $allGroupsExport | Export-Csv -Path "$OutputPath\Exchange_Groups.csv" -NoTypeInformation
        }
        
        if ($securityGroups.Count -gt 0) {
            $securityGroups | Export-Csv -Path "$OutputPath\Security_Groups.csv" -NoTypeInformation
        }
        
        Write-Host "  ✓ User Mailboxes: $($mailboxes.Count)" -ForegroundColor Green
        Write-Host "  ✓ Potential Shared Mailboxes: $($sharedMailboxes.Count)" -ForegroundColor Green
        Write-Host "  ✓ Microsoft 365 Groups: $($m365Groups.Count)" -ForegroundColor Green
        Write-Host "  ✓ Distribution Groups: $($distGroups.Count)" -ForegroundColor Green
        Write-Host "  ✓ Mail-Enabled Security Groups: $($mailSecurityGroups.Count)" -ForegroundColor Green
        Write-Host "  ✓ Security Groups: $($securityGroups.Count)" -ForegroundColor Green
        
        # Create Exchange summary
        $summary = @"
Exchange Online Summary
========================================

Mailboxes:
  - User Mailboxes: $($mailboxes.Count)
  - Potential Shared Mailboxes: $($sharedMailboxes.Count)
  - Total Mail-Enabled Users: $($users.Count)

Groups:
  - Microsoft 365 Groups: $($m365Groups.Count)
  - Distribution Groups: $($distGroups.Count)
  - Mail-Enabled Security Groups: $($mailSecurityGroups.Count)
  - Security Groups (non-mail): $($securityGroups.Count)
  - Total Groups: $($allGroups.Count)

Note: Shared mailbox identification is based on account status.
For precise shared mailbox details, Exchange Online PowerShell may be required.
"@
        
        $summary | Out-File -FilePath "$OutputPath\Exchange_Summary.txt"
    }
    catch {
        Write-Host "  ✗ Error gathering Exchange info: $_" -ForegroundColor Red
    }
}

# 4. Get SharePoint Online Sites
function Get-SharePointSites {
    param($OutputPath)
    
    Write-Host "`n[4/7] Gathering SharePoint Online Sites..." -ForegroundColor Cyan
    
    try {
        # First try to get all sites
        $sites = @()
        
        # Get the root site and all sites if possible
        try {
            $allSites = Get-MgSite -All -ErrorAction Stop
            $sites = $allSites | Select-Object @{
                Name = 'Site Name'; Expression = {$_.DisplayName}
            }, @{
                Name = 'Site URL'; Expression = {$_.WebUrl}
            }, @{
                Name = 'Site ID'; Expression = {$_.Id}
            }, @{
                Name = 'Created Date'; Expression = {$_.CreatedDateTime}
            }, @{
                Name = 'Last Modified'; Expression = {$_.LastModifiedDateTime}
            }, @{
                Name = 'Root Site'; Expression = {$_.IsRoot}
            }
        }
        catch {
            Write-Host "  - Note: Unable to enumerate all sites. Getting root site only." -ForegroundColor Yellow
            
            # Try to get at least the root site
            try {
                $rootSite = Get-MgSite -SiteId "root" -ErrorAction Stop
                $sites = @($rootSite | Select-Object @{
                    Name = 'Site Name'; Expression = {$_.DisplayName}
                }, @{
                    Name = 'Site URL'; Expression = {$_.WebUrl}
                }, @{
                    Name = 'Site ID'; Expression = {$_.Id}
                }, @{
                    Name = 'Created Date'; Expression = {$_.CreatedDateTime}
                }, @{
                    Name = 'Last Modified'; Expression = {$_.LastModifiedDateTime}
                }, @{
                    Name = 'Root Site'; Expression = {$true}
                })
            }
            catch {
                Write-Host "  - Warning: Could not access SharePoint sites" -ForegroundColor Yellow
            }
        }
        
        # Try to get SharePoint usage report
        Write-Host "  - Gathering SharePoint usage data..." -ForegroundColor Gray
        try {
            $tempFile = Join-Path $env:TEMP "SharePointUsage_$(Get-Date -Format 'yyyyMMddHHmmss').csv"
            $uri = "https://graph.microsoft.com/v1.0/reports/getSharePointSiteUsageDetail(period='D30')"
            Invoke-MgGraphRequest -Uri $uri -Method GET -OutputFilePath $tempFile
            
            if (Test-Path $tempFile) {
                $usageData = Import-Csv -Path $tempFile
                
                $siteUsage = $usageData | Select-Object @{
                    Name = 'Site URL'; Expression = {$_.'Site URL' -replace 'https://', '' -replace 'http://', ''}
                }, @{
                    Name = 'Site Owner'; Expression = {
                        $owner = $_.'Owner Display Name'
                        if ([string]::IsNullOrWhiteSpace($owner)) {
                            $_.'Owner Principal Name'
                        } else {
                            $owner
                        }
                    }
                }, @{
                    Name = 'Storage Used (GB)'; Expression = {
                        if ($_.'Storage Used (Byte)') {
                            [math]::Round([double]$_.'Storage Used (Byte)' / 1GB, 2)
                        } else { 0 }
                    }
                }, @{
                    Name = 'Storage Allocated (GB)'; Expression = {
                        if ($_.'Storage Allocated (Byte)') {
                            [math]::Round([double]$_.'Storage Allocated (Byte)' / 1GB, 2)
                        } else { 0 }
                    }
                }, @{
                    Name = 'Files Count'; Expression = {$_.'File Count'}
                }, @{
                    Name = 'Active Files Count'; Expression = {$_.'Active File Count'}
                }, @{
                    Name = 'Page View Count'; Expression = {$_.'Page View Count'}
                }, @{
                    Name = 'Visited Page Count'; Expression = {$_.'Visited Page Count'}
                }, @{
                    Name = 'Last Activity Date'; Expression = {$_.'Last Activity Date'}
                }, @{
                    Name = 'Site Type'; Expression = {
                        if ($_.'Root Web Template') {
                            $_.'Root Web Template'
                        } else {
                            'Standard'
                        }
                    }
                }
                
                $siteUsage | Export-Csv -Path "$OutputPath\SharePoint_Sites_Usage.csv" -NoTypeInformation
                
                $totalStorage = ($siteUsage | Measure-Object -Property 'Storage Used (GB)' -Sum).Sum
                $activeSites = ($siteUsage | Where-Object {$_.'Last Activity Date' -ne ''}).Count
                
                Write-Host "  ✓ Sites with usage data: $($siteUsage.Count)" -ForegroundColor Green
                Write-Host "  ✓ Active sites (30 days): $activeSites" -ForegroundColor Green
                Write-Host "  ✓ Total storage used: $([math]::Round($totalStorage, 2)) GB" -ForegroundColor Green
                
                Remove-Item -Path $tempFile -Force -ErrorAction SilentlyContinue
            }
        }
        catch {
            Write-Host "  - Could not retrieve SharePoint usage data" -ForegroundColor Yellow
        }
        
        if ($sites.Count -gt 0) {
            $sites | Export-Csv -Path "$OutputPath\SharePoint_Sites.csv" -NoTypeInformation
            Write-Host "  ✓ Total Sites Found: $($sites.Count)" -ForegroundColor Green
        }
        
        # Create SharePoint summary
        $summary = @"
SharePoint Online Summary
=========================

$(if ($sites.Count -gt 0) { "Sites Found: $($sites.Count)" } else { "Sites: Unable to enumerate (requires additional permissions)" })
$(if ($siteUsage) { "Sites with Usage Data: $($siteUsage.Count)" })
$(if ($totalStorage) { "Total Storage Used: $([math]::Round($totalStorage, 2)) GB" })
$(if ($activeSites) { "Active Sites (Last 30 days): $activeSites" })

Note: Full site enumeration may require additional permissions.
Usage report provides comprehensive data for all sites.
"@
        
        $summary | Out-File -FilePath "$OutputPath\SharePoint_Summary.txt"
    }
    catch {
        Write-Host "  ✗ Error gathering SharePoint sites: $_" -ForegroundColor Red
    }
}

# 5. Get OneDrive for Business Usage
function Get-OneDriveUsage {
    param($OutputPath)
    
    Write-Host "`n[5/7] Gathering OneDrive for Business Usage..." -ForegroundColor Cyan
    
    try {
        # Get OneDrive usage report (last 30 days) with output to file
        $tempFile = Join-Path $env:TEMP "OneDriveUsage_$(Get-Date -Format 'yyyyMMddHHmmss').csv"
        
        # Use the -OutputFilePath parameter to handle the CSV response
        $uri = "https://graph.microsoft.com/v1.0/reports/getOneDriveUsageAccountDetail(period='D30')"
        Invoke-MgGraphRequest -Uri $uri -Method GET -OutputFilePath $tempFile
        
        # Read and process the CSV file
        if (Test-Path $tempFile) {
            $oneDriveData = Import-Csv -Path $tempFile
            
            $oneDriveUsage = $oneDriveData | Select-Object @{
                Name = 'Display Name'; Expression = {
                    $name = $_.'Owner Display Name'
                    if ([string]::IsNullOrWhiteSpace($name)) {
                        # Try to extract from UPN
                        $upn = $_.'Owner Principal Name'
                        if ($upn) {
                            $upn.Split('@')[0]
                        } else {
                            'Unknown'
                        }
                    } else {
                        $name
                    }
                }
            }, @{
                Name = 'User Principal Name'; Expression = {$_.'Owner Principal Name'}
            }, @{
                Name = 'Last Activity Date'; Expression = {$_.'Last Activity Date'}
            }, @{
                Name = 'Files Count'; Expression = {$_.'File Count'}
            }, @{
                Name = 'Active Files Count'; Expression = {$_.'Active File Count'}
            }, @{
                Name = 'Storage Used (GB)'; Expression = {
                    if ($_.'Storage Used (Byte)') {
                        [math]::Round([double]$_.'Storage Used (Byte)' / 1GB, 2)
                    } else { 0 }
                }
            }, @{
                Name = 'Storage Allocated (GB)'; Expression = {
                    if ($_.'Storage Allocated (Byte)') {
                        [math]::Round([double]$_.'Storage Allocated (Byte)' / 1GB, 2)
                    } else { 0 }
                }
            }, @{
                Name = 'Storage Used Percentage'; Expression = {
                    if ($_.'Storage Used (Byte)' -and $_.'Storage Allocated (Byte)') {
                        $used = [double]$_.'Storage Used (Byte)'
                        $allocated = [double]$_.'Storage Allocated (Byte)'
                        if ($allocated -gt 0) {
                            [math]::Round(($used / $allocated) * 100, 2)
                        } else { 0 }
                    } else { 0 }
                }
            }
            
            $oneDriveUsage | Export-Csv -Path "$OutputPath\OneDrive_Usage.csv" -NoTypeInformation
            
            $activeUsers = ($oneDriveUsage | Where-Object {$_.'Last Activity Date' -ne ''}).Count
            $totalStorage = ($oneDriveUsage | Measure-Object -Property 'Storage Used (GB)' -Sum).Sum
            $totalFiles = ($oneDriveUsage | Measure-Object -Property 'Files Count' -Sum).Sum
            
            Write-Host "  ✓ Total OneDrive Users: $($oneDriveUsage.Count)" -ForegroundColor Green
            Write-Host "  ✓ Active Users (30 days): $activeUsers" -ForegroundColor Green
            Write-Host "  ✓ Total Storage Used: $([math]::Round($totalStorage, 2)) GB" -ForegroundColor Green
            
            # Create OneDrive summary
            $summary = @"
OneDrive for Business Summary
==============================

Total OneDrive Accounts: $($oneDriveUsage.Count)
Active Users (Last 30 days): $activeUsers
Total Files: $totalFiles
Total Storage Used: $([math]::Round($totalStorage, 2)) GB
Average Storage per User: $([math]::Round($totalStorage / [Math]::Max($oneDriveUsage.Count, 1), 2)) GB

Top 10 Storage Consumers:
$(($oneDriveUsage | Sort-Object 'Storage Used (GB)' -Descending | Select-Object -First 10 | ForEach-Object {"  - $($_.'Display Name'): $($_.'Storage Used (GB)') GB"}) -join "`n")
"@
            
            $summary | Out-File -FilePath "$OutputPath\OneDrive_Summary.txt"
            
            # Clean up temp file
            Remove-Item -Path $tempFile -Force -ErrorAction SilentlyContinue
        }
    }
    catch {
        Write-Host "  ✗ Error gathering OneDrive usage: $_" -ForegroundColor Red
    }
}

# 6. Get Microsoft Teams Information
function Get-TeamsInfo {
    param($OutputPath)
    
    Write-Host "`n[6/7] Gathering Microsoft Teams Information..." -ForegroundColor Cyan
    
    try {
        # Get all teams
        $teams = Get-MgTeam -All | Select-Object @{
            Name = 'Team Name'; Expression = {$_.DisplayName}
        }, @{
            Name = 'Team ID'; Expression = {$_.Id}
        }, @{
            Name = 'Description'; Expression = {$_.Description}
        }, @{
            Name = 'Visibility'; Expression = {$_.Visibility}
        }, @{
            Name = 'Created Date'; Expression = {$_.CreatedDateTime}
        }, @{
            Name = 'Archived'; Expression = {$_.IsArchived}
        }
        
        # Get Teams usage report for additional insights
        Write-Host "  - Gathering Teams usage data..." -ForegroundColor Gray
        $teamsUsage = $null
        try {
            $tempFile = Join-Path $env:TEMP "TeamsUsage_$(Get-Date -Format 'yyyyMMddHHmmss').csv"
            $uri = "https://graph.microsoft.com/v1.0/reports/getTeamsTeamActivityDetail(period='D30')"
            Invoke-MgGraphRequest -Uri $uri -Method GET -OutputFilePath $tempFile
            
            if (Test-Path $tempFile) {
                $teamsUsageData = Import-Csv -Path $tempFile
                
                $teamsUsage = $teamsUsageData | Select-Object @{
                    Name = 'Team Name'; Expression = {$_.'Team Name'}
                }, @{
                    Name = 'Team ID'; Expression = {$_.'Team Id'}
                }, @{
                    Name = 'Last Activity Date'; Expression = {$_.'Last Activity Date'}
                }, @{
                    Name = 'Active Users'; Expression = {$_.'Active Users'}
                }, @{
                    Name = 'Active Channels'; Expression = {$_.'Active Channels'}
                }, @{
                    Name = 'Guests'; Expression = {$_.'Guests'}
                }, @{
                    Name = 'Active Guests'; Expression = {$_.'Active Guests'}
                }, @{
                    Name = 'Channel Messages'; Expression = {$_.'Channel Messages'}
                }, @{
                    Name = 'Post Messages'; Expression = {$_.'Post Messages'}
                }, @{
                    Name = 'Reply Messages'; Expression = {$_.'Reply Messages'}
                }, @{
                    Name = 'Meetings Organized'; Expression = {$_.'Meetings Organized'}
                }
                
                $teamsUsage | Export-Csv -Path "$OutputPath\Teams_Usage.csv" -NoTypeInformation
                
                $activeTeams = ($teamsUsage | Where-Object {$_.'Last Activity Date' -ne ''}).Count
                $totalMessages = ($teamsUsage | Measure-Object -Property 'Channel Messages' -Sum).Sum
                
                Write-Host "  ✓ Teams with usage data: $($teamsUsage.Count)" -ForegroundColor Green
                Write-Host "  ✓ Active teams (30 days): $activeTeams" -ForegroundColor Green
                
                Remove-Item -Path $tempFile -Force -ErrorAction SilentlyContinue
            }
        }
        catch {
            Write-Host "  - Could not retrieve Teams usage data" -ForegroundColor Yellow
        }
        
        # Get detailed team information with error handling
        $teamDetails = @()
        $teamsWithErrors = 0
        
        foreach ($team in $teams) {
            try {
                # Try to get team members and channels
                $members = $null
                $channels = $null
                
                try {
                    $members = Get-MgTeamMember -TeamId $team.'Team ID' -All -ErrorAction Stop
                }
                catch {
                    # If we can't get members due to permissions, continue without member data
                    $teamsWithErrors++
                }
                
                try {
                    $channels = Get-MgTeamChannel -TeamId $team.'Team ID' -All -ErrorAction SilentlyContinue
                }
                catch {
                    # If we can't get channels, continue without channel data
                }
                
                # Try to match with usage data
                $usageData = $teamsUsage | Where-Object { $_.'Team ID' -eq $team.'Team ID' }
                
                $teamDetails += [PSCustomObject]@{
                    'Team Name' = $team.'Team Name'
                    'Team ID' = $team.'Team ID'
                    'Member Count' = if ($members) { $members.Count } elseif ($usageData) { $usageData.'Active Users' } else { 'N/A' }
                    'Channel Count' = if ($channels) { $channels.Count } elseif ($usageData) { $usageData.'Active Channels' } else { 'N/A' }
                    'Owners' = if ($members) { ($members | Where-Object {$_.Roles -contains 'owner'}).Count } else { 'N/A' }
                    'Guests' = if ($members) { ($members | Where-Object {$_.Roles -contains 'guest'}).Count } elseif ($usageData) { $usageData.'Guests' } else { 'N/A' }
                    'Visibility' = $team.Visibility
                    'Archived' = $team.Archived
                    'Last Activity' = if ($usageData) { $usageData.'Last Activity Date' } else { 'N/A' }
                }
            }
            catch {
                # If any error occurs, add team with limited data
                $teamDetails += [PSCustomObject]@{
                    'Team Name' = $team.'Team Name'
                    'Team ID' = $team.'Team ID'
                    'Member Count' = 'N/A'
                    'Channel Count' = 'N/A'
                    'Owners' = 'N/A'
                    'Guests' = 'N/A'
                    'Visibility' = $team.Visibility
                    'Archived' = $team.Archived
                    'Last Activity' = 'N/A'
                }
            }
        }
        
        $teams | Export-Csv -Path "$OutputPath\Teams_List.csv" -NoTypeInformation
        if ($teamDetails.Count -gt 0) {
            $teamDetails | Export-Csv -Path "$OutputPath\Teams_Details.csv" -NoTypeInformation
        }
        
        Write-Host "  ✓ Total Teams: $($teams.Count)" -ForegroundColor Green
        Write-Host "  ✓ Active Teams: $(($teams | Where-Object {$_.Archived -eq $false}).Count)" -ForegroundColor Green
        Write-Host "  ✓ Archived Teams: $(($teams | Where-Object {$_.Archived -eq $true}).Count)" -ForegroundColor Green
        
        if ($teamsWithErrors -gt 0) {
            Write-Host "  ⚠ Note: Could not retrieve member details for $teamsWithErrors team(s) due to permissions" -ForegroundColor Yellow
        }
        
        # Create Teams summary
        $summary = @"
Microsoft Teams Summary
========================

Total Teams: $($teams.Count)
Active Teams: $(($teams | Where-Object {$_.Archived -eq $false}).Count)
Archived Teams: $(($teams | Where-Object {$_.Archived -eq $true}).Count)
Private Teams: $(($teams | Where-Object {$_.Visibility -eq 'Private'}).Count)
Public Teams: $(($teams | Where-Object {$_.Visibility -eq 'Public'}).Count)

$(if ($activeTeams) {
"Teams Activity (Last 30 days):
  Active Teams: $activeTeams
  Total Channel Messages: $totalMessages"
})

$(if ($teamsWithErrors -gt 0) {
"Note: Full member details could not be retrieved for $teamsWithErrors team(s).
This typically requires TeamMember.Read.All permission or Team owner rights."
})
"@
        
        $summary | Out-File -FilePath "$OutputPath\Teams_Summary.txt"
    }
    catch {
        Write-Host "  ✗ Error gathering Teams info: $_" -ForegroundColor Red
    }
}

# 7. Get Entra ID App Registrations
function Get-EntraIDApps {
    param($OutputPath)
    
    Write-Host "`n[7/7] Gathering Entra ID App Registrations..." -ForegroundColor Cyan
    
    try {
        $apps = Get-MgApplication -All | Select-Object @{
            Name = 'App Name'; Expression = {$_.DisplayName}
        }, @{
            Name = 'App ID'; Expression = {$_.AppId}
        }, @{
            Name = 'Object ID'; Expression = {$_.Id}
        }, @{
            Name = 'Created Date'; Expression = {$_.CreatedDateTime}
        }, @{
            Name = 'Sign-In Audience'; Expression = {$_.SignInAudience}
        }, @{
            Name = 'Publisher Domain'; Expression = {$_.PublisherDomain}
        }, @{
            Name = 'API Permissions'; Expression = {
                ($_.RequiredResourceAccess | ForEach-Object {
                    $_.ResourceAccess.Count
                } | Measure-Object -Sum).Sum
            }
        }, @{
            Name = 'Redirect URIs'; Expression = {
                ($_.Web.RedirectUris + $_.PublicClient.RedirectUris + $_.Spa.RedirectUris) -join '; '
            }
        }
        
        # Get service principals (Enterprise Applications)
        $servicePrincipals = Get-MgServicePrincipal -All | Select-Object @{
            Name = 'App Name'; Expression = {$_.DisplayName}
        }, @{
            Name = 'App ID'; Expression = {$_.AppId}
        }, @{
            Name = 'Object ID'; Expression = {$_.Id}
        }, @{
            Name = 'Type'; Expression = {$_.ServicePrincipalType}
        }, @{
            Name = 'Enabled'; Expression = {$_.AccountEnabled}
        }, @{
            Name = 'Publisher'; Expression = {$_.PublisherName}
        }, @{
            Name = 'Verified Publisher'; Expression = {$_.VerifiedPublisher.DisplayName}
        }
        
        $apps | Export-Csv -Path "$OutputPath\EntraID_AppRegistrations.csv" -NoTypeInformation
        $servicePrincipals | Export-Csv -Path "$OutputPath\EntraID_ServicePrincipals.csv" -NoTypeInformation
        
        Write-Host "  ✓ App Registrations: $($apps.Count)" -ForegroundColor Green
        Write-Host "  ✓ Service Principals: $($servicePrincipals.Count)" -ForegroundColor Green
        
        # Identify third-party apps
        $thirdPartyApps = $servicePrincipals | Where-Object {
            $_.Type -eq 'Application' -and 
            $_.'Publisher' -ne 'Microsoft Corporation' -and
            $_.'Publisher' -ne $null
        }
        
        # Create Entra ID summary
        $summary = @"
Entra ID Applications Summary
==============================

App Registrations: $($apps.Count)
Service Principals (Enterprise Apps): $($servicePrincipals.Count)
Third-Party Applications: $($thirdPartyApps.Count)
Microsoft Applications: $(($servicePrincipals | Where-Object {$_.'Publisher' -eq 'Microsoft Corporation'}).Count)

Sign-In Audience Types:
  - AzureADMyOrg: $(($apps | Where-Object {$_.'Sign-In Audience' -eq 'AzureADMyOrg'}).Count)
  - AzureADMultipleOrgs: $(($apps | Where-Object {$_.'Sign-In Audience' -eq 'AzureADMultipleOrgs'}).Count)
  - AzureADandPersonalMicrosoftAccount: $(($apps | Where-Object {$_.'Sign-In Audience' -eq 'AzureADandPersonalMicrosoftAccount'}).Count)

Top Third-Party Publishers:
$(($thirdPartyApps | Group-Object 'Publisher' | Sort-Object Count -Descending | Select-Object -First 10 | ForEach-Object {"  - $($_.Name): $($_.Count) apps"}) -join "`n")
"@
        
        $summary | Out-File -FilePath "$OutputPath\EntraID_Summary.txt"
    }
    catch {
        Write-Host "  ✗ Error gathering Entra ID apps: $_" -ForegroundColor Red
    }
}

# Function to generate executive summary
function Generate-ExecutiveSummary {
    param($OutputPath)
    
    Write-Host "`nGenerating Executive Summary..." -ForegroundColor Cyan
    
    # Get organization details
    $org = Get-MgOrganization
    
    $summary = @"
=================================================================
           MICROSOFT 365 TENANT ASSESSMENT REPORT
                    GRAPH API ONLY VERSION
=================================================================

Assessment Date: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
Tenant ID: $((Get-MgContext).TenantId)
Organization: $($org.DisplayName)
Domains: $(($org.VerifiedDomains | Where-Object {$_.IsDefault -eq $true}).Name)

=================================================================
                      ASSESSMENT COMPLETE
=================================================================

The following reports have been generated:

1. LICENSING
   - M365_Licenses.csv
   - M365_Licenses_Summary.txt

2. SECURITY
   - SecureScore_Current.csv
   - SecureScore_Recommendations.csv
   - SecureScore_Summary.txt

3. EXCHANGE ONLINE
   - Exchange_Mailboxes.csv
   - Exchange_SharedMailboxes.csv (if applicable)
   - Exchange_Groups.csv
   - Security_Groups.csv
   - Exchange_Summary.txt

4. SHAREPOINT ONLINE
   - SharePoint_Sites.csv (if accessible)
   - SharePoint_Sites_Usage.csv
   - SharePoint_Summary.txt

5. ONEDRIVE FOR BUSINESS
   - OneDrive_Usage.csv
   - OneDrive_Summary.txt

6. MICROSOFT TEAMS
   - Teams_List.csv
   - Teams_Details.csv
   - Teams_Usage.csv (if available)
   - Teams_Summary.txt

7. ENTRA ID APPLICATIONS
   - EntraID_AppRegistrations.csv
   - EntraID_ServicePrincipals.csv
   - EntraID_Summary.txt

=================================================================

All reports have been saved to: $OutputPath

Next Steps:
1. Review the executive summaries for each service
2. Analyze the CSV files for detailed information
3. Focus on security recommendations in SecureScore_Recommendations.csv
4. Review third-party app integrations in EntraID reports
5. Check for unused licenses and optimize licensing costs

=================================================================

Note: This assessment used Microsoft Graph API exclusively.
Some Exchange-specific details may require Exchange Online PowerShell
for complete information.

=================================================================
"@
    
    $summary | Out-File -FilePath "$OutputPath\EXECUTIVE_SUMMARY.txt"
    Write-Host $summary -ForegroundColor Yellow
}

# Main execution function
function Start-M365TenantAssessment {
    Write-Host @"
=================================================================
           MICROSOFT 365 TENANT ASSESSMENT TOOL
                    GRAPH API ONLY VERSION
                         Version 2.0
=================================================================
This script will gather comprehensive information about your 
Microsoft 365 tenant for migration and security assessments.

Using Microsoft Graph API exclusively

Required Permissions: Global Reader or equivalent
=================================================================
"@ -ForegroundColor Cyan
    
    # Initialize modules
    Initialize-RequiredModules
    
    # Create output directory
    $outputPath = Initialize-OutputDirectory
    
    # Connect to services
    $graphConnected = Connect-MicrosoftGraphServices
    if (-not $graphConnected) {
        Write-Host "`nFailed to connect to Microsoft Graph. Exiting." -ForegroundColor Red
        return
    }
    
    Write-Host "`n=== Starting Data Collection ===" -ForegroundColor Cyan
    
    # Run assessment functions
    Get-M365LicensingInfo -OutputPath $outputPath
    Get-SecureScoreInfo -OutputPath $outputPath
    Get-ExchangeOnlineInfo -OutputPath $outputPath
    Get-SharePointSites -OutputPath $outputPath
    Get-OneDriveUsage -OutputPath $outputPath
    Get-TeamsInfo -OutputPath $outputPath
    Get-EntraIDApps -OutputPath $outputPath
    
    # Generate executive summary
    Generate-ExecutiveSummary -OutputPath $outputPath
    
    # Disconnect from services
    Write-Host "`nDisconnecting from services..." -ForegroundColor Cyan
    Disconnect-MgGraph -ErrorAction SilentlyContinue
    
    Write-Host "`n✓ Assessment completed successfully!" -ForegroundColor Green
    Write-Host "Results saved to: $outputPath" -ForegroundColor Green
    
    # Open the output folder
    Start-Process explorer.exe $outputPath
}

# Execute the assessment
Start-M365TenantAssessment
