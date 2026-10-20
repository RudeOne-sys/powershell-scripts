param(
    [Parameter(Mandatory, ValueFromPipeline)]
    [string[]]$Domain
)

process {
    foreach ($d in $Domain) {
        $uri = "https://login.microsoftonline.com/$d/.well-known/openid-configuration"
        try {
            $response = Invoke-RestMethod -Uri $uri -Method Get -ErrorAction Stop
        }
        catch {
            Write-Error "No Microsoft 365 / Entra ID tenant found for '$d'."
            continue
        }
        # The tenant ID is in the 'issuer' URL
        if ($response.issuer -match "https://sts.windows.net/([a-f0-9-]+)/") {
            Write-Output $matches[1]
        } else {
            Write-Error "Tenant ID not found for '$d'."
        }
    }
}
