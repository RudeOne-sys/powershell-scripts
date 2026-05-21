# Microsoft 365 Tenant Licensing & Secure Score Script - Graph API
# Version: 2.0-Lite
# Purpose: Only collect Microsoft 365 licensing (friendly names) and Secure Score info
# Required: Global Reader permissions

#Requires -Version 5.1

function Initialize-RequiredModules {
    Write-Host "Checking and installing required modules..." -ForegroundColor Cyan
    $requiredModules = @(
        'Microsoft.Graph.Authentication',
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

function Connect-MicrosoftGraphServices {
    Write-Host "`nConnecting to Microsoft Graph..." -ForegroundColor Cyan
    try {
        $scopes = @(
            "Organization.Read.All",
            "Directory.Read.All",
            "Reports.Read.All",
            "SecurityEvents.Read.All"
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

# Friendly name mapping for common SKUs
$SkuFriendlyNames = @{
    "ENTERPRISEPREMIUM" = "Microsoft 365 E5"
    "ENTERPRISEPACK"    = "Microsoft 365 E3"
    "STANDARDPACK"      = "Office 365 E1"
    "O365_BUSINESS"     = "Microsoft 365 Apps for Business"
    "SMB_BUSINESS_PREMIUM" = "Microsoft 365 Business Premium"
    "EXCHANGESTANDARD"  = "Exchange Online (Plan 1)"
    "EXCHANGEENTERPRISE"= "Exchange Online (Plan 2)"
    "EMS"               = "Enterprise Mobility + Security E3"
    "EMSPREMIUM"        = "Enterprise Mobility + Security E5"
    "POWER_BI_STANDARD" = "Power BI (Free)"
    "POWER_BI_PRO"      = "Power BI Pro"
    # Add more mappings as needed
}

function Get-M365LicensingInfo {
    param($OutputPath)
    Write-Host "`n[1/2] Gathering Microsoft 365 Licensing Information..." -ForegroundColor Cyan
    try {
        $licenses = Get-MgSubscribedSku | Select-Object @(
            @{Name='License Name'; Expression={ $SkuFriendlyNames[$_.SkuPartNumber] ?? $_.SkuPartNumber }},
            @{Name='SKU ID'; Expression={ $_.SkuId }},
            @{Name='Total Licenses'; Expression={ $_.PrepaidUnits.Enabled }},
            @{Name='Consumed Licenses'; Expression={ $_.ConsumedUnits }},
            @{Name='Available Licenses'; Expression={ $_.PrepaidUnits.Enabled - $_.ConsumedUnits }}
        )
        $licenses | Export-Csv -Path "$OutputPath\M365_Licenses.csv" -NoTypeInformation
        Write-Host "  ✓ Exported $($licenses.Count) license SKUs" -ForegroundColor Green

        $summary = @"
Microsoft 365 Licensing Summary
Generated: $(Get-Date)
================================

Total License SKUs: $($licenses.Count)
Total Licenses Purchased: $(($licenses | Measure-Object -Property 'Total Licenses' -Sum).Sum)
Total Licenses Consumed: $(($licenses | Measure-Object -Property 'Consumed Licenses' -Sum).Sum)
Total Licenses Available: $(($licenses | Measure-Object -Property 'Available Licenses' -Sum).Sum)

Top 5 Most Consumed Licenses:
$($licenses | Sort-Object 'Consumed Licenses' -Descending | Select-Object -First 5 | ForEach-Object {"  - $($_.'License Name'): $($_.'Consumed Licenses')/$($_.'Total Licenses')"} | Out-String)
"@
        $summary | Out-File -FilePath "$OutputPath\M365_Licenses_Summary.txt"
    }
    catch {
        Write-Host "  ✗ Error gathering licensing info: $_" -ForegroundColor Red
    }
}

function Get-SecureScoreInfo {
    param($OutputPath)
    Write-Host "`n[2/2] Gathering Secure Score Information..." -ForegroundColor Cyan
    try {
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

        $controlProfiles = Get-MgSecuritySecureScoreControlProfile | Select-Object @(
            @{Name='ID'; Expression={ $_.Id }},
            @{Name='Control Name'; Expression={ $_.Title }},
            @{Name='Description'; Expression={ $_.Description }},
            @{Name='Product'; Expression={ $_.ServiceName }},
            @{Name='Control Category'; Expression={ $_.ControlCategory }},
            @{Name='Max Score'; Expression={ $_.MaxScore }},
            @{Name='Current Score'; Expression={ $_.CurrentScore }},
            @{Name='Implementation Status'; Expression={ $_.ImplementationStatus }}
        )
        # Add 'Have License' column manually with default $true (adjust logic here if needed)
        $controlProfiles | ForEach-Object { $_ | Add-Member -MemberType NoteProperty -Name "Have License" -Value $true -PassThru } |
            Export-Csv -Path "$OutputPath\SecureScore_Recommendations.csv" -NoTypeInformation

        $notImplemented = $controlProfiles | Where-Object { $_.'Implementation Status' -ne 'Completed' }
        $highImpact = $notImplemented | Where-Object { $_.'Max Score' -ge 10 }

        $summary = @"
Secure Score Recommendations Summary
=====================================

Total Recommendations: $($controlProfiles.Count)
Implemented: $(($controlProfiles | Where-Object { $_.'Implementation Status' -eq 'Completed' }).Count)
Not Implemented: $($notImplemented.Count)
High Impact (10+ points) Not Implemented: $($highImpact.Count)

Top 10 High-Value Improvements:
$($highImpact | Sort-Object 'Max Score' -Descending | Select-Object -First 10 | ForEach-Object {"  - $($_.'Control Name'): $($_.'Max Score') points"} | Out-String)
"@
        $summary | Out-File -FilePath "$OutputPath\SecureScore_Summary.txt"
    }
    catch {
        Write-Host "  ✗ Error gathering secure score: $_" -ForegroundColor Red
    }
}

# --- Main Execution ---
Initialize-RequiredModules
$outputPath = Initialize-OutputDirectory
if (Connect-MicrosoftGraphServices) {
    Get-M365LicensingInfo -OutputPath $outputPath
    Get-SecureScoreInfo -OutputPath $outputPath
}
