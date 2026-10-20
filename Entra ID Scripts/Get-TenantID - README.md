# Get-TenantID

Looks up the Microsoft Entra ID (Azure AD) tenant ID for a domain. No sign-in or modules needed. It reads Microsoft's public OpenID configuration endpoint.

## Usage

```powershell
.\Get-TenantID.ps1 -Domain contoso.com
```

Several domains at once:

```powershell
.\Get-TenantID.ps1 -Domain contoso.com, fabrikam.com
"contoso.com", "fabrikam.com" | .\Get-TenantID.ps1
```

If you leave out `-Domain`, PowerShell prompts you for it.

## Output

The tenant ID (a GUID), one per domain:

```
de19de3c-fc60-416f-903b-8b391dd48b9e
```

If a domain isn't linked to a Microsoft 365 or Entra ID tenant, the script writes an error and moves on to the next domain.

## Related

For looking up a large list of domains in one go, see `Get-TenantIDs Bulk-NoAuth.ps1` in the same folder.
