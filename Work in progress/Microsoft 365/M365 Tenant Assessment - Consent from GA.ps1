# Global Admin runs this once for the entire organization
Connect-MgGraph -Scopes @(
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
    "Organization.Read.All"
) -TenantId 2ec583fa-1eff-49d9-97a8-693a88b6b7ec

# After authentication, disconnect
Disconnect-MgGraph

# Now Global Readers can use the script without admin consent