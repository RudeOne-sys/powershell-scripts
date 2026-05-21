# Get-TenantIDs-NoAuth.ps1
# Script to retrieve Microsoft tenant IDs for multiple domains without requiring authentication

# Function to get tenant ID for a domain using OpenID configuration
function Get-AzureADTenantId {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Domain
    )
    
    try {
        $response = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$Domain/.well-known/openid-configuration" -Method Get -ErrorAction Stop
        
        # Extract tenant ID from the issuer URL (format: https://sts.windows.net/{tenant-id}/)
        if ($response.issuer -match 'https://sts.windows.net/([^/]+)/') {
            return $matches[1]
        } else {
            return "Error: Could not extract tenant ID from response"
        }
    }
    catch {
        return "Error: $($_.Exception.Message)"
    }
}

# Script starts here
Write-Host "===================================================" -ForegroundColor Cyan
Write-Host "  Microsoft Tenant ID Finder (No Authentication)" -ForegroundColor Cyan
Write-Host "===================================================" -ForegroundColor Cyan

# Ask for input file
$inputFile = Read-Host("Enter full path to source CSV file")

# Verify the file exists
if (-not (Test-Path $inputFile)) {
    Write-Host "File not found: $inputFile" -ForegroundColor Red
    Write-Host "Would you like to create a template file? (Y/N)" -ForegroundColor Yellow
    $createTemplate = Read-Host
    
    if ($createTemplate -eq "Y" -or $createTemplate -eq "y") {
        $templatePath = "$env:USERPROFILE\Documents\TenantID_Template.csv"
        @"
EmailDomain
contoso.com
fabrikam.com
"@ | Out-File -FilePath $templatePath -Encoding utf8
        
        Write-Host "Template created at: $templatePath" -ForegroundColor Green
        Write-Host "Please fill in your domains and run the script again." -ForegroundColor Yellow
        exit
    } else {
        exit
    }
}

# Create output file name with timestamp
$fileTimestamp = (Get-Date -Format "yyyy-MM-dd_HH-mm-ss")
$outputFile = "$env:USERPROFILE\Downloads\TenantID_Results_$fileTimestamp.csv"

# Import CSV
$domainList = Import-Csv -Path $inputFile

# Check if CSV has the required column
$firstRow = $domainList | Select-Object -First 1
if (-not ($firstRow.PSObject.Properties.Name -contains "EmailDomain")) {
    Write-Host "Error: CSV file must contain an 'EmailDomain' column" -ForegroundColor Red
    exit
}

# Create results array
$outputs = @()

# Process each domain
$totalDomains = $domainList.Count
$currentDomain = 0

foreach ($row in $domainList) {
    $currentDomain++
    $domain = $row.EmailDomain

    # Skip empty domains
    if ([string]::IsNullOrWhiteSpace($domain)) {
        Write-Host "Skipping empty domain at row $currentDomain" -ForegroundColor Yellow
        continue
    }

    # Show progress
    Write-Progress -Activity "Retrieving Tenant IDs" -Status "Processing $domain" -PercentComplete (($currentDomain / $totalDomains) * 100)
    Write-Host "Processing ($currentDomain of $totalDomains): $domain" -ForegroundColor Cyan
    
    # Get tenant ID
    $tenantId = Get-AzureADTenantId -Domain $domain
    
    # Create result object
    $obj = New-Object PSObject
    $obj | Add-Member -MemberType NoteProperty -Name "Domain" -Value $domain
    $obj | Add-Member -MemberType NoteProperty -Name "TenantID" -Value $tenantId
    $obj | Add-Member -MemberType NoteProperty -Name "ProcessedTime" -Value (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
    
    # Add additional columns from original CSV (if any)
    foreach ($property in $row.PSObject.Properties.Name) {
        if ($property -ne "EmailDomain" -and $property -ne "Domain" -and $property -ne "TenantID" -and $property -ne "ProcessedTime") {
            $obj | Add-Member -MemberType NoteProperty -Name $property -Value $row.$property
        }
    }
    
    $outputs += $obj
}

Write-Progress -Activity "Retrieving Tenant IDs" -Completed

# Export results
$outputs | Export-Csv -Path $outputFile -NoTypeInformation

# Display results summary
Write-Host "`nProcessing complete! Summary:" -ForegroundColor Green
$outputs | Format-Table -AutoSize

Write-Host "`nResults saved to: $outputFile" -ForegroundColor Green