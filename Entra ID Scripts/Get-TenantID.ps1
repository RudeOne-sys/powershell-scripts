$domain = "Tshipientle.co.za"
$uri = "https://login.microsoftonline.com/$domain/.well-known/openid-configuration"
$response = Invoke-RestMethod -Uri $uri -Method Get
# The tenant ID is in the 'issuer' URL
if ($response.issuer -match "https://sts.windows.net/([a-f0-9-]+)/") {
    $tenantId = $matches[1]
    Write-Output $tenantId
} else {
    Write-Error "Tenant ID not found."
}