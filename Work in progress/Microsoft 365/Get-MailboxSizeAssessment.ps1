# Microsoft 365 Mailbox Size Assessment Script
# Purpose: Gather mailbox sizes to inform licensing decisions
# Author: Netsurit Professional Services
# Date: November 2024

<#
.SYNOPSIS
    Analyzes Exchange Online mailbox sizes and archive usage to inform M365 licensing decisions.

.DESCRIPTION
    This script connects to Exchange Online and generates reports on:
    - Primary mailbox sizes
    - Archive mailbox usage
    - Users approaching quota limits
    - Recommendations for license upgrades based on mailbox needs

.NOTES
    Requires:
    - Exchange Online Management PowerShell module
    - Global Admin or Exchange Admin permissions
    
    Run this command first to install the module if needed:
    Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber
#>

# Function to convert bytes to GB for readability
function Convert-BytesToGB {
    param([string]$Size)
    
    if ($Size -match "(\d+\.?\d*)\s*([KMGT]?B)") {
        $value = [double]$matches[1]
        $unit = $matches[2]
        
        switch ($unit) {
            "KB" { $value = $value / 1MB }
            "MB" { $value = $value / 1KB }
            "GB" { $value = $value }
            "TB" { $value = $value * 1024 }
            default { $value = $value / 1GB }
        }
        
        return [math]::Round($value, 2)
    }
    return 0
}

# Main script execution
try {
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host "Microsoft 365 Mailbox Size Assessment" -ForegroundColor Cyan
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host ""
    
    # Connect to Exchange Online
    Write-Host "Connecting to Exchange Online..." -ForegroundColor Yellow
    Connect-ExchangeOnline -ShowBanner:$false
    
    Write-Host "Connected successfully!" -ForegroundColor Green
    Write-Host ""
    
    # Get all mailboxes
    Write-Host "Retrieving mailbox information (this may take a few minutes)..." -ForegroundColor Yellow
    $mailboxes = Get-Mailbox -ResultSize Unlimited | Where-Object {$_.RecipientTypeDetails -eq "UserMailbox"}
    $totalMailboxes = $mailboxes.Count
    
    Write-Host "Found $totalMailboxes user mailboxes" -ForegroundColor Green
    Write-Host ""
    
    # Initialize results array
    $results = @()
    $counter = 0
    
    # Process each mailbox
    foreach ($mailbox in $mailboxes) {
        $counter++
        Write-Progress -Activity "Analyzing mailboxes" -Status "Processing $($mailbox.DisplayName)" -PercentComplete (($counter / $totalMailboxes) * 100)
        
        # Get mailbox statistics
        $stats = Get-MailboxStatistics -Identity $mailbox.Identity
        
        # Get archive statistics if archive is enabled
        $archiveSize = 0
        $archiveItemCount = 0
        if ($mailbox.ArchiveStatus -eq "Active") {
            try {
                $archiveStats = Get-MailboxStatistics -Identity $mailbox.Identity -Archive -ErrorAction SilentlyContinue
                if ($archiveStats) {
                    $archiveSize = Convert-BytesToGB -Size $archiveStats.TotalItemSize.ToString()
                    $archiveItemCount = $archiveStats.ItemCount
                }
            } catch {
                # Archive may not be accessible, skip
            }
        }
        
        # Convert sizes
        $mailboxSizeGB = Convert-BytesToGB -Size $stats.TotalItemSize.ToString()
        
        # Determine current license type based on quota
        $currentLicense = "Unknown"
        if ($mailbox.ProhibitSendReceiveQuota -match "50 GB") {
            $currentLicense = "Exchange Online P1 / Business Premium"
            $quotaGB = 50
        } elseif ($mailbox.ProhibitSendReceiveQuota -match "100 GB") {
            $currentLicense = "Office 365 E3 / M365 E3/E5"
            $quotaGB = 100
        } else {
            $currentLicense = "Custom/Unknown"
            $quotaGB = 0
        }
        
        # Calculate percentage used
        $percentUsed = if ($quotaGB -gt 0) { [math]::Round(($mailboxSizeGB / $quotaGB) * 100, 1) } else { 0 }
        
        # Determine recommendation
        $recommendation = ""
        if ($mailboxSizeGB -gt 90) {
            $recommendation = "CRITICAL: Needs E3/E5 immediately (100GB quota)"
        } elseif ($mailboxSizeGB -gt 45 -and $currentLicense -like "*P1*") {
            $recommendation = "WARNING: Consider E3/E5 upgrade (approaching 50GB limit)"
        } elseif ($archiveSize -gt 10) {
            $recommendation = "INFO: Active archive usage - E3/E5 provides unlimited archive"
        } elseif ($mailboxSizeGB -lt 25) {
            $recommendation = "OK: Current license sufficient"
        } else {
            $recommendation = "MONITOR: Track growth"
        }
        
        # Add to results
        $results += [PSCustomObject]@{
            DisplayName = $mailbox.DisplayName
            UserPrincipalName = $mailbox.UserPrincipalName
            MailboxSizeGB = $mailboxSizeGB
            ItemCount = $stats.ItemCount
            QuotaGB = $quotaGB
            PercentUsed = $percentUsed
            ArchiveEnabled = ($mailbox.ArchiveStatus -eq "Active")
            ArchiveSizeGB = $archiveSize
            ArchiveItemCount = $archiveItemCount
            CurrentLicense = $currentLicense
            Recommendation = $recommendation
        }
    }
    
    Write-Progress -Activity "Analyzing mailboxes" -Completed
    
    # Generate reports
    Write-Host ""
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host "ANALYSIS COMPLETE - GENERATING REPORTS" -ForegroundColor Cyan
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host ""
    
    # Export detailed results
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $detailedReport = "M365_Mailbox_Analysis_$timestamp.csv"
    $results | Export-Csv -Path $detailedReport -NoTypeInformation
    Write-Host "✓ Detailed report exported: $detailedReport" -ForegroundColor Green
    
    # Summary statistics
    $totalMailboxSizeGB = ($results | Measure-Object -Property MailboxSizeGB -Sum).Sum
    $totalArchiveSizeGB = ($results | Measure-Object -Property ArchiveSizeGB -Sum).Sum
    $avgMailboxSizeGB = [math]::Round(($results | Measure-Object -Property MailboxSizeGB -Average).Average, 2)
    
    $over50GB = ($results | Where-Object {$_.MailboxSizeGB -gt 50}).Count
    $over45GB = ($results | Where-Object {$_.MailboxSizeGB -gt 45}).Count
    $over80Percent = ($results | Where-Object {$_.PercentUsed -gt 80}).Count
    
    $activeArchives = ($results | Where-Object {$_.ArchiveEnabled -eq $true}).Count
    $archivesWith10GB = ($results | Where-Object {$_.ArchiveSizeGB -gt 10}).Count
    
    # Display summary
    Write-Host ""
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host "SUMMARY STATISTICS" -ForegroundColor Cyan
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host ""
    
    Write-Host "Total Mailboxes: $totalMailboxes" -ForegroundColor White
    Write-Host "Total Mailbox Data: $([math]::Round($totalMailboxSizeGB, 2)) GB" -ForegroundColor White
    Write-Host "Total Archive Data: $([math]::Round($totalArchiveSizeGB, 2)) GB" -ForegroundColor White
    Write-Host "Average Mailbox Size: $avgMailboxSizeGB GB" -ForegroundColor White
    Write-Host ""
    
    Write-Host "CAPACITY ANALYSIS:" -ForegroundColor Yellow
    Write-Host "  Mailboxes over 50GB: $over50GB" -ForegroundColor $(if ($over50GB -gt 0) {"Red"} else {"Green"})
    Write-Host "  Mailboxes over 45GB: $over45GB" -ForegroundColor $(if ($over45GB -gt 0) {"Yellow"} else {"Green"})
    Write-Host "  Mailboxes over 80% quota: $over80Percent" -ForegroundColor $(if ($over80Percent -gt 0) {"Yellow"} else {"Green"})
    Write-Host ""
    
    Write-Host "ARCHIVE USAGE:" -ForegroundColor Yellow
    Write-Host "  Active Archives: $activeArchives" -ForegroundColor White
    Write-Host "  Archives over 10GB: $archivesWith10GB" -ForegroundColor $(if ($archivesWith10GB -gt 0) {"Yellow"} else {"Green"})
    Write-Host ""
    
    # Licensing recommendation
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host "LICENSING RECOMMENDATION" -ForegroundColor Cyan
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host ""
    
    if ($over50GB -gt 0) {
        Write-Host "❌ CRITICAL: $over50GB users REQUIRE immediate upgrade to E3/E5" -ForegroundColor Red
        Write-Host "   These users cannot receive mail with 50GB mailboxes" -ForegroundColor Red
        Write-Host ""
    }
    
    if ($over45GB -gt 5 -or $over80Percent -gt 5) {
        Write-Host "⚠️  WARNING: Significant number of users approaching limits" -ForegroundColor Yellow
        Write-Host "   Consider E3/E5 upgrade to provide 100GB mailboxes" -ForegroundColor Yellow
        Write-Host ""
    }
    
    if ($archivesWith10GB -gt 10) {
        Write-Host "💡 INFO: Heavy archive usage detected" -ForegroundColor Cyan
        Write-Host "   E3/E5 provides unlimited archive storage" -ForegroundColor Cyan
        Write-Host ""
    }
    
    if ($over50GB -eq 0 -and $over45GB -lt 3 -and $archivesWith10GB -lt 5) {
        Write-Host "✅ ACCEPTABLE: Current mailbox sizes manageable with existing licenses" -ForegroundColor Green
        Write-Host "   However, E3/E5 still recommended for:" -ForegroundColor White
        Write-Host "   - Future growth headroom" -ForegroundColor White
        Write-Host "   - Unlimited archive capability" -ForegroundColor White
        Write-Host "   - Enhanced security features" -ForegroundColor White
        Write-Host ""
    }
    
    # Top 10 largest mailboxes
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host "TOP 10 LARGEST MAILBOXES" -ForegroundColor Cyan
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host ""
    
    $top10 = $results | Sort-Object -Property MailboxSizeGB -Descending | Select-Object -First 10
    $top10 | Format-Table DisplayName, MailboxSizeGB, PercentUsed, CurrentLicense, Recommendation -AutoSize
    
    # Users needing attention
    Write-Host ""
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host "USERS REQUIRING ATTENTION" -ForegroundColor Cyan
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host ""
    
    $needsAttention = $results | Where-Object {$_.Recommendation -like "CRITICAL*" -or $_.Recommendation -like "WARNING*"} | Sort-Object -Property MailboxSizeGB -Descending
    
    if ($needsAttention.Count -gt 0) {
        $needsAttention | Format-Table DisplayName, MailboxSizeGB, PercentUsed, Recommendation -AutoSize
    } else {
        Write-Host "✅ No users requiring immediate attention" -ForegroundColor Green
    }
    
    # Final recommendation summary
    Write-Host ""
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host "FINAL RECOMMENDATION" -ForegroundColor Cyan
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host ""
    
    $e3Recommended = $false
    
    if ($over50GB -gt 0 -or $over45GB -gt 5) {
        Write-Host "✓ Microsoft 365 E3/E5 is STRONGLY RECOMMENDED" -ForegroundColor Green
        Write-Host "  Reason: Mailbox capacity requirements" -ForegroundColor White
        $e3Recommended = $true
    }
    
    if ($archivesWith10GB -gt 10) {
        Write-Host "✓ Microsoft 365 E3/E5 is RECOMMENDED" -ForegroundColor Green
        Write-Host "  Reason: Unlimited archive benefits" -ForegroundColor White
        $e3Recommended = $true
    }
    
    if (-not $e3Recommended) {
        Write-Host "⚠️  Microsoft 365 E3/E5 still recommended for:" -ForegroundColor Yellow
        Write-Host "  - Growth headroom (100GB vs 50GB mailboxes)" -ForegroundColor White
        Write-Host "  - Unlimited archive capability" -ForegroundColor White
        Write-Host "  - Enterprise security features" -ForegroundColor White
        Write-Host "  - Enhanced compliance capabilities" -ForegroundColor White
    }
    
    Write-Host ""
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Report saved to: $detailedReport" -ForegroundColor Green
    Write-Host "Share this report with your Microsoft 365 consultant." -ForegroundColor White
    Write-Host ""
    
    # Disconnect
    Disconnect-ExchangeOnline -Confirm:$false
    Write-Host "Disconnected from Exchange Online" -ForegroundColor Green
    
} catch {
    Write-Host ""
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ""
    Write-Host "Common issues:" -ForegroundColor Yellow
    Write-Host "  1. Exchange Online Management module not installed" -ForegroundColor White
    Write-Host "     Solution: Install-Module -Name ExchangeOnlineManagement" -ForegroundColor White
    Write-Host "  2. Insufficient permissions" -ForegroundColor White
    Write-Host "     Solution: Must be Global Admin or Exchange Admin" -ForegroundColor White
    Write-Host "  3. Modern authentication not enabled" -ForegroundColor White
    Write-Host "     Solution: Enable modern auth in Microsoft 365 admin center" -ForegroundColor White
    Write-Host ""
}

# End of script
