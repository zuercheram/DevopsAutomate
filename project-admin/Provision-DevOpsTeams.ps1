#Requires -Version 5.1

<#
.SYNOPSIS
    Provisions or cleans up Teams, Area Paths, Iterations, and Team Members in Azure DevOps.

.DESCRIPTION
    Reads a CSV template defining the desired organisational structure of an Azure DevOps project.
    The team hierarchy in the CSV is used to derive Area Paths automatically.
    The script compares desired state with the current project state and only creates resources
    that do not already exist (idempotent).  A -Cleanup switch reverses every change.

    CSV columns:
      Id          (optional)  – GUID of an existing team.  When present the script
                                 updates the team (rename, area, members) instead of
                                 creating a new one.  If the Id does not match any
                                 existing team the row is treated as a new team.
      TeamName    (required)  – Name of the team.
      ParentTeam  (required)  – Name of the parent team (empty for root-level teams).
      Description (optional)  – Team description.
      Members     (optional)  – Semicolon-separated email addresses of team members.
      Iterations  (optional)  – Semicolon-separated iteration paths to create and subscribe to.
                                 Use backslash for nested iterations, e.g. "Release 1\Sprint 1".
      AreaPaths   (optional)  – Semicolon-separated custom area path names.
                                 A simple name (e.g. "CustomArea") is placed under the parent
                                 team's hierarchy area.  A path with "/" (e.g. "root/child")
                                 is created directly from the project root.
                                 If omitted the team name is used as area path node.

    After provisioning the script writes back the CSV so that every row contains the
    Azure DevOps team Id.  This keeps the CSV consistent for subsequent runs.

.PARAMETER Organization
    Azure DevOps organization name (e.g. "my-org").

.PARAMETER Project
    Azure DevOps project name (e.g. "MyProject").

.PARAMETER CsvPath
    Path to the CSV template file.

.PARAMETER PAT
    Personal Access Token with full access.  Falls back to the AZURE_DEVOPS_PAT env var.

.PARAMETER Cleanup
    Removes all resources defined in the CSV from the project (reverse of provisioning).

.EXAMPLE
    .\Provision-DevOpsTeams.ps1 -Organization "my-org" -Project "MyProject" `
        -CsvPath ".\teams-template.csv" -PAT "your-pat"

.EXAMPLE
    .\Provision-DevOpsTeams.ps1 -Organization "my-org" -Project "MyProject" `
        -CsvPath ".\teams-template.csv" -PAT "your-pat" -Cleanup
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$Organization,

    [Parameter(Mandatory)]
    [string]$Project,

    [Parameter(Mandatory)]
    [ValidateScript({ Test-Path $_ })]
    [string]$CsvPath,

    [Parameter()]
    [string]$PAT = $env:AZURE_DEVOPS_PAT,

    [switch]$Cleanup
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ═══════════════════════════════════════════════════════════════════════════════
#  CONFIGURATION
# ═══════════════════════════════════════════════════════════════════════════════
if (-not $PAT) {
    Write-Error "No PAT provided. Use -PAT or set the AZURE_DEVOPS_PAT environment variable."
    return
}

$script:AuthHeaders = @{
    Authorization = "Basic " + [Convert]::ToBase64String(
        [Text.Encoding]::ASCII.GetBytes(":$PAT")
    )
}

$script:CoreUrl  = "https://dev.azure.com/$Organization"
$script:VsspsUrl = "https://vssps.dev.azure.com/$Organization"

# Caches refreshed at runtime
$script:KnownAreaPaths      = @()
$script:KnownIterationPaths = @()

# ═══════════════════════════════════════════════════════════════════════════════
#  LOGGING HELPERS
# ═══════════════════════════════════════════════════════════════════════════════
function Write-Banner ([string]$Text) {
    $line = [string]::new([char]0x2500, 72)
    Write-Host "`n$line" -ForegroundColor Cyan
    Write-Host "  $Text"  -ForegroundColor Cyan
    Write-Host "$line"    -ForegroundColor Cyan
}

function Write-Step {
    param(
        [string]$Text,
        [ValidateSet("CREATE","SKIP","DELETE","INFO","WARN","ERROR")]
        [string]$Status = "INFO"
    )
    $colors = @{
        CREATE = "Green"; SKIP = "DarkYellow"; DELETE = "Red"
        INFO   = "Gray";  WARN = "Yellow";     ERROR  = "Magenta"
    }
    Write-Host "  [$Status] $Text" -ForegroundColor $colors[$Status]
}

# ═══════════════════════════════════════════════════════════════════════════════
#  REST API WRAPPER
# ═══════════════════════════════════════════════════════════════════════════════
function Invoke-AdoApi {
    param(
        [Parameter(Mandatory)][string]$Uri,
        [string]$Method  = "GET",
        [object]$Body     = $null
    )

    $splat = @{
        Uri         = $Uri
        Method      = $Method
        Headers     = $script:AuthHeaders
        ContentType = "application/json"
    }

    if ($null -ne $Body) {
        $json = if ($Body -is [string]) { $Body }
                else { $Body | ConvertTo-Json -Depth 10 -Compress }
        $splat.Body = [Text.Encoding]::UTF8.GetBytes($json)
    }

    try {
        return Invoke-RestMethod @splat
    }
    catch {
        $detail = ""
        try { $detail = $_.ErrorDetails.Message } catch {}
        Write-Step "API $Method $Uri  ->  $detail" -Status ERROR
        throw
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
#  CSV PARSING & TEAM HIERARCHY
# ═══════════════════════════════════════════════════════════════════════════════
function Import-TeamsCsv ([string]$Path) {
    $rows = Import-Csv -Path $Path -Encoding UTF8

    # Validate required columns
    foreach ($col in @("TeamName","ParentTeam")) {
        if ($col -notin $rows[0].PSObject.Properties.Name) {
            Write-Error "CSV is missing the required column '$col'."
            return
        }
    }

    # Trim whitespace
    foreach ($r in $rows) {
        $r.TeamName   = $r.TeamName.Trim()
        $r.ParentTeam = $r.ParentTeam.Trim()
        foreach ($opt in @("Id","Description","Members","Iterations","AreaPaths")) {
            if ($r.PSObject.Properties[$opt]) { $r.$opt = $r.$opt.Trim() }
        }
    }

    return $rows
}

function Test-CsvValid ([array]$Rows) {
    $names = $Rows.TeamName

    # Check for duplicates
    $dupes = $names | Group-Object | Where-Object Count -gt 1
    if ($dupes) {
        Write-Error "Duplicate TeamName(s): $($dupes.Name -join ', ')"
        return $false
    }

    # Check parent references
    foreach ($r in $Rows) {
        if ($r.ParentTeam -and $r.ParentTeam -notin $names) {
            Write-Error "Team '$($r.TeamName)' references unknown parent '$($r.ParentTeam)'."
            return $false
        }
    }

    return $true
}

function Get-OrderedTeams ([array]$Rows) {
    # Topological sort – parents before children
    $ordered = [System.Collections.Generic.List[object]]::new()
    $added   = [System.Collections.Generic.HashSet[string]]::new(
                   [StringComparer]::OrdinalIgnoreCase)

    function Add-TeamRec ([string]$Name) {
        if ($added.Contains($Name)) { return }
        $row = $Rows | Where-Object { $_.TeamName -eq $Name }
        if (-not $row) { return }
        if ($row.ParentTeam) { Add-TeamRec $row.ParentTeam }
        [void]$added.Add($Name)
        $ordered.Add($row)
    }

    foreach ($r in $Rows) { Add-TeamRec $r.TeamName }
    return $ordered.ToArray()
}

function Resolve-AreaPath ([string]$TeamName, [array]$Rows) {
    # Returns the area path relative to the project root, e.g. "Platform\Backend"
    $row = $Rows | Where-Object { $_.TeamName -eq $TeamName }
    if ($row.ParentTeam) {
        return "$(Resolve-AreaPath $row.ParentTeam $Rows)\$TeamName"
    }
    return $TeamName
}

function Resolve-TeamAreaPaths ([string]$TeamName, [array]$Rows) {
    # Returns an array of area paths for a team, honouring the AreaPaths CSV column.
    # - Empty AreaPaths  → default hierarchy-derived path (team name chain).
    # - Simple name      → placed under parent team's hierarchy area.
    # - Path with "/" or "\" → absolute from project root (ignores parent hierarchy).
    $row = $Rows | Where-Object { $_.TeamName -eq $TeamName }

    $hasCustom = $row.PSObject.Properties["AreaPaths"] -and $row.AreaPaths
    if (-not $hasCustom) {
        return @(Resolve-AreaPath -TeamName $TeamName -Rows $Rows)
    }

    $paths    = ($row.AreaPaths -split ';') | ForEach-Object { $_.Trim() } | Where-Object { $_ }
    $resolved = [System.Collections.Generic.List[string]]::new()

    foreach ($p in $paths) {
        if ($p.Contains('/') -or $p.Contains('\')) {
            # Full path from project root – normalise to \
            $resolved.Add(($p -replace '/', '\'))
        }
        else {
            # Simple node name – place under parent's hierarchy area
            if ($row.ParentTeam) {
                $parentArea = Resolve-AreaPath -TeamName $row.ParentTeam -Rows $Rows
                $resolved.Add("$parentArea\$p")
            }
            else {
                $resolved.Add($p)
            }
        }
    }

    return $resolved.ToArray()
}

# ═══════════════════════════════════════════════════════════════════════════════
#  TEAM CRUD
# ═══════════════════════════════════════════════════════════════════════════════
function Get-ExistingTeams {
    (Invoke-AdoApi "$script:CoreUrl/_apis/projects/$Project/teams?`$top=500&api-version=7.1").value
}

function New-DevOpsTeam ([string]$Name, [string]$Description = "") {
    Invoke-AdoApi `
        "$script:CoreUrl/_apis/projects/$Project/teams?api-version=7.1" `
        -Method POST `
        -Body @{ name = $Name; description = $Description }
}

function Remove-DevOpsTeam ([string]$TeamId) {
    Invoke-AdoApi `
        "$script:CoreUrl/_apis/projects/$Project/teams/${TeamId}?api-version=7.1" `
        -Method DELETE
}

function Rename-DevOpsTeam ([string]$TeamId, [string]$NewName) {
    Invoke-AdoApi `
        "$script:CoreUrl/_apis/projects/$Project/teams/${TeamId}?api-version=7.1" `
        -Method PATCH `
        -Body @{ name = $NewName } | Out-Null
}

function Update-DevOpsTeam ([string]$TeamId, [string]$Name, [string]$Description = "") {
    # Updates both name and description of an existing team
    Invoke-AdoApi `
        "$script:CoreUrl/_apis/projects/$Project/teams/${TeamId}?api-version=7.1" `
        -Method PATCH `
        -Body @{ name = $Name; description = $Description } | Out-Null
}

function Get-DefaultTeam {
    # Returns the project's default team object.
    # The default team ID is embedded in the project properties.
    $projInfo = Invoke-AdoApi "$script:CoreUrl/_apis/projects/${Project}?api-version=7.1"
    $defaultTeamId = $null
    if ($projInfo.PSObject.Properties["defaultTeam"]) {
        $defaultTeamId = $projInfo.defaultTeam.id
    }
    if ($defaultTeamId) {
        return Invoke-AdoApi "$script:CoreUrl/_apis/projects/$Project/teams/${defaultTeamId}?api-version=7.1"
    }
    return $null
}

function Get-ProjectOwner {
    # Returns the authenticated user (PAT owner) as the project owner stand-in
    $connInfo = Invoke-AdoApi "$script:CoreUrl/_apis/connectiondata?api-version=7.1-preview.1"
    if ($connInfo -and $connInfo.authenticatedUser) {
        return $connInfo.authenticatedUser
    }
    return $null
}

function Get-TeamMemberDescriptors ([string]$TeamId, [string]$ScopeDescriptor) {
    # Returns the member descriptors of a team
    $members = (Invoke-AdoApi `
        "$script:CoreUrl/_apis/projects/$Project/teams/$TeamId/members?`$top=500&api-version=7.1").value
    $descriptors = @()
    foreach ($m in $members) {
        if ($m.identity.PSObject.Properties["uniqueName"]) {
            $user = Find-UserByEmail -Email $m.identity.uniqueName
            if ($user) { $descriptors += @{ Email = $m.identity.uniqueName; Descriptor = $user.descriptor } }
        }
    }
    return $descriptors
}

# ═══════════════════════════════════════════════════════════════════════════════
#  AREA PATH HELPERS
# ═══════════════════════════════════════════════════════════════════════════════
function Get-AreaTree {
    Invoke-AdoApi "$script:CoreUrl/$Project/_apis/wit/classificationnodes/areas?`$depth=14&api-version=7.1"
}

function Get-FlatPaths ([object]$Node, [string]$Prefix = "") {
    # Recursively flattens a classification tree into relative path strings.
    $list = [System.Collections.Generic.List[string]]::new()
    $children = $null
    if ($Node.PSObject.Properties["children"]) { $children = $Node.children }
    if ($children) {
        foreach ($child in $children) {
            $p = if ($Prefix) { "$Prefix\$($child.name)" } else { $child.name }
            $list.Add($p)
            (Get-FlatPaths $child $p) | ForEach-Object { $list.Add($_) }
        }
    }
    return $list
}

function ConvertTo-UrlPath ([string]$BackslashPath) {
    # "Platform\Backend" → "Platform/Backend" with each segment URL-encoded
    ($BackslashPath -split '\\' | ForEach-Object { [Uri]::EscapeDataString($_) }) -join '/'
}

function New-AreaNode ([string]$RelativePath) {
    $segments = $RelativePath -split '\\'

    for ($i = 0; $i -lt $segments.Count; $i++) {
        $seg   = $segments[$i]
        $built = ($segments[0..$i] -join '\')

        if ($built -in $script:KnownAreaPaths) {
            Write-Step "Area '$script:ProjectName\$built' exists" -Status SKIP
            continue
        }

        $parentUrl = ""
        if ($i -gt 0) {
            $parentUrl = "/" + (ConvertTo-UrlPath ($segments[0..($i-1)] -join '\'))
        }

        $uri = "$script:CoreUrl/$Project/_apis/wit/classificationnodes/areas${parentUrl}?api-version=7.1"

        try {
            Invoke-AdoApi $uri -Method POST -Body @{ name = $seg } | Out-Null
            $script:KnownAreaPaths += $built
            Write-Step "Created area '$script:ProjectName\$built'" -Status CREATE
        }
        catch {
            # Node may already exist (race condition or stale cache)
            $script:KnownAreaPaths += $built
            Write-Step "Area '$script:ProjectName\$built' may already exist – continuing" -Status SKIP
        }
    }
}

function Remove-AreaNode ([string]$RelativePath, [int]$ReclassifyId) {
    $urlPath = ConvertTo-UrlPath $RelativePath
    $uri = "$script:CoreUrl/$Project/_apis/wit/classificationnodes/areas/${urlPath}?`$reclassifyId=$ReclassifyId&api-version=7.1"
    Invoke-AdoApi $uri -Method DELETE | Out-Null
}

# ═══════════════════════════════════════════════════════════════════════════════
#  ITERATION HELPERS
# ═══════════════════════════════════════════════════════════════════════════════
function Get-IterationTree {
    Invoke-AdoApi "$script:CoreUrl/$Project/_apis/wit/classificationnodes/iterations?`$depth=14&api-version=7.1"
}

function New-IterationNode ([string]$RelativePath) {
    $segments = $RelativePath -split '\\'

    for ($i = 0; $i -lt $segments.Count; $i++) {
        $seg   = $segments[$i]
        $built = ($segments[0..$i] -join '\')

        if ($built -in $script:KnownIterationPaths) {
            Write-Step "Iteration '$script:ProjectName\$built' exists" -Status SKIP
            continue
        }

        $parentUrl = ""
        if ($i -gt 0) {
            $parentUrl = "/" + (ConvertTo-UrlPath ($segments[0..($i-1)] -join '\'))
        }

        $uri = "$script:CoreUrl/$Project/_apis/wit/classificationnodes/iterations${parentUrl}?api-version=7.1"

        try {
            Invoke-AdoApi $uri -Method POST -Body @{ name = $seg } | Out-Null
            $script:KnownIterationPaths += $built
            Write-Step "Created iteration '$script:ProjectName\$built'" -Status CREATE
        }
        catch {
            $script:KnownIterationPaths += $built
            Write-Step "Iteration '$script:ProjectName\$built' may already exist – continuing" -Status SKIP
        }
    }
}

function Remove-IterationNode ([string]$RelativePath, [int]$ReclassifyId) {
    $urlPath = ConvertTo-UrlPath $RelativePath
    $uri = "$script:CoreUrl/$Project/_apis/wit/classificationnodes/iterations/${urlPath}?`$reclassifyId=$ReclassifyId&api-version=7.1"
    Invoke-AdoApi $uri -Method DELETE | Out-Null
}

function Get-IterationNodeByPath ([string]$RelativePath) {
    $urlPath = ConvertTo-UrlPath $RelativePath
    Invoke-AdoApi "$script:CoreUrl/$Project/_apis/wit/classificationnodes/iterations/${urlPath}?api-version=7.1"
}

# ═══════════════════════════════════════════════════════════════════════════════
#  TEAM SETTINGS (AREA & ITERATIONS)
# ═══════════════════════════════════════════════════════════════════════════════
function Set-TeamArea ([string]$TeamId, [string[]]$FullAreaPaths) {
    $values = @($FullAreaPaths | ForEach-Object {
        @{ value = $_; includeChildren = $false }
    })
    $body = @{
        defaultValue = $FullAreaPaths[0]
        values       = $values
    }
    Invoke-AdoApi `
        "$script:CoreUrl/$Project/$TeamId/_apis/work/teamsettings/teamfieldvalues?api-version=7.1" `
        -Method PATCH -Body $body | Out-Null
}

function Add-TeamIteration ([string]$TeamId, [string]$IterationGuid) {
    Invoke-AdoApi `
        "$script:CoreUrl/$Project/$TeamId/_apis/work/teamsettings/iterations?api-version=7.1" `
        -Method POST -Body @{ id = $IterationGuid } | Out-Null
}

function Set-TeamBacklogIteration ([string]$TeamId, [string]$IterationGuid) {
    # Sets the backlog iteration for a team – required before subscribing to sprints
    Invoke-AdoApi `
        "$script:CoreUrl/$Project/$TeamId/_apis/work/teamsettings?api-version=7.1" `
        -Method PATCH -Body @{ backlogIteration = $IterationGuid } | Out-Null
}

# ═══════════════════════════════════════════════════════════════════════════════
#  TEAM MEMBERSHIP (GRAPH API)
# ═══════════════════════════════════════════════════════════════════════════════
function Get-ProjectScopeDescriptor {
    $proj = Invoke-AdoApi "$script:CoreUrl/_apis/projects/${Project}?api-version=7.1"
    $desc = Invoke-AdoApi "$script:VsspsUrl/_apis/graph/descriptors/$($proj.id)?api-version=7.1-preview.1"
    return $desc.value
}

function Get-ProjectGroups ([string]$ScopeDescriptor) {
    (Invoke-AdoApi `
        "$script:VsspsUrl/_apis/graph/groups?scopeDescriptor=$ScopeDescriptor&api-version=7.1-preview.1"
    ).value
}

function Find-TeamGroupDescriptor ([string]$TeamName, [array]$Groups) {
    $group = $Groups | Where-Object {
        $_.displayName -eq $TeamName -and $_.origin -eq "vsts"
    }
    if ($group) { return $group.descriptor }
    return $null
}

function Find-UserByEmail ([string]$Email) {
    $result = Invoke-AdoApi `
        "$script:VsspsUrl/_apis/graph/subjectquery?api-version=7.1-preview.1" `
        -Method POST `
        -Body @{ query = $Email; subjectKind = @("User") }
    if ($result.value -and $result.value.Count -gt 0) {
        return $result.value[0]
    }
    return $null
}

function Add-GroupMember ([string]$MemberDescriptor, [string]$GroupDescriptor) {
    $uri = "$script:VsspsUrl/_apis/graph/memberships/${MemberDescriptor}/${GroupDescriptor}?api-version=7.1-preview.1"
    Invoke-RestMethod -Uri $uri -Method Put -Headers $script:AuthHeaders -ContentType "application/json" | Out-Null
}

function Remove-GroupMember ([string]$MemberDescriptor, [string]$GroupDescriptor) {
    $uri = "$script:VsspsUrl/_apis/graph/memberships/${MemberDescriptor}/${GroupDescriptor}?api-version=7.1-preview.1"
    Invoke-RestMethod -Uri $uri -Method Delete -Headers $script:AuthHeaders -ContentType "application/json" | Out-Null
}

# ═══════════════════════════════════════════════════════════════════════════════
#  PROJECT-SCOPED GROUP CRUD
# ═══════════════════════════════════════════════════════════════════════════════
function New-DevOpsGroup ([string]$DisplayName, [string]$Description = "", [string]$ScopeDescriptor) {
    Invoke-AdoApi `
        "$script:VsspsUrl/_apis/graph/groups?scopeDescriptor=$ScopeDescriptor&api-version=7.1-preview.1" `
        -Method POST `
        -Body @{ displayName = $DisplayName; description = $Description }
}

function Remove-DevOpsGroup ([string]$GroupDescriptor) {
    $uri = "$script:VsspsUrl/_apis/graph/groups/${GroupDescriptor}?api-version=7.1-preview.1"
    Invoke-RestMethod -Uri $uri -Method Delete -Headers $script:AuthHeaders -ContentType "application/json" | Out-Null
}

function Find-GroupByDisplayName ([string]$DisplayName, [array]$Groups) {
    $match = $Groups | Where-Object {
        $_.displayName -eq $DisplayName -and $_.origin -eq "vsts"
    }
    return $match
}

# ═══════════════════════════════════════════════════════════════════════════════
#  AREA PATH PERMISSION HELPERS
# ═══════════════════════════════════════════════════════════════════════════════
$script:CssNamespaceId = "83e28ad4-2d72-4ceb-97b0-c7726d5502c3"

function Get-AreaNodeByPath ([string]$RelativePath) {
    # Returns the full node object (including identifier GUID) for an area path
    $urlPath = ConvertTo-UrlPath $RelativePath
    Invoke-AdoApi "$script:CoreUrl/$Project/_apis/wit/classificationnodes/areas/${urlPath}?api-version=7.1"
}

function Build-AreaSecurityToken ([string]$RelativePath) {
    # Builds the security token for an area path.
    # Token format: vstfs:///Classification/Node/<root-guid>:vstfs:///Classification/Node/<child-guid>:...
    $segments = $RelativePath -split '\\'
    $tokenParts = @()

    # Start from the root area node
    $rootNode = Invoke-AdoApi "$script:CoreUrl/$Project/_apis/wit/classificationnodes/areas?api-version=7.1"
    $tokenParts += "vstfs:///Classification/Node/$($rootNode.identifier)"

    # Walk down the tree segment by segment
    $pathSoFar = ""
    foreach ($seg in $segments) {
        $pathSoFar = if ($pathSoFar) { "$pathSoFar\$seg" } else { $seg }
        $node = Get-AreaNodeByPath -RelativePath $pathSoFar
        $tokenParts += "vstfs:///Classification/Node/$($node.identifier)"
    }

    return ($tokenParts -join ':')
}

function Resolve-IdentityDescriptor ([string]$SubjectDescriptor) {
    # Resolves a Graph API subject descriptor (vssgp.xxx) to an Identity descriptor
    # (Microsoft.TeamFoundation.Identity;S-1-9-...) required by the ACE/ACL APIs.
    $result = Invoke-AdoApi `
        "$script:VsspsUrl/_apis/identities?subjectDescriptors=$SubjectDescriptor&api-version=7.1"
    if ($result.value -and $result.value.Count -gt 0) {
        return $result.value[0].descriptor
    }
    return $null
}

function Set-AreaPathPermission ([string]$SecurityToken, [string]$IdentityDescriptor, [int]$AllowBits) {
    # Sets allow permissions on an area path for an identity descriptor.
    # The descriptor must be in 'Microsoft.TeamFoundation.Identity;S-1-9-...' format.
    $body = @{
        token                = $SecurityToken
        merge                = $true
        accessControlEntries = @(
            @{
                descriptor   = $IdentityDescriptor
                allow        = $AllowBits
                deny         = 0
                extendedInfo = @{}
            }
        )
    }
    Invoke-AdoApi `
        "$script:CoreUrl/_apis/accesscontrolentries/$($script:CssNamespaceId)?api-version=7.1" `
        -Method POST -Body $body | Out-Null
}

function Deny-AreaPathPermission ([string]$SecurityToken, [string]$IdentityDescriptor, [int]$DenyBits) {
    # Sets deny permissions on an area path for an identity descriptor.
    $body = @{
        token                = $SecurityToken
        merge                = $true
        accessControlEntries = @(
            @{
                descriptor   = $IdentityDescriptor
                allow        = 0
                deny         = $DenyBits
                extendedInfo = @{}
            }
        )
    }
    Invoke-AdoApi `
        "$script:CoreUrl/_apis/accesscontrolentries/$($script:CssNamespaceId)?api-version=7.1" `
        -Method POST -Body $body | Out-Null
}

function Remove-AreaPathPermissions ([string]$SecurityToken, [string]$IdentityDescriptor) {
    # Removes all ACEs for an identity descriptor on a specific area path token.
    $encodedDesc = [Uri]::EscapeDataString($IdentityDescriptor)
    $encodedToken = [Uri]::EscapeDataString($SecurityToken)
    $uri = "$script:CoreUrl/_apis/accesscontrolentries/$($script:CssNamespaceId)?token=$encodedToken&descriptors=$encodedDesc&api-version=7.1"
    try {
        Invoke-RestMethod -Uri $uri -Method Delete -Headers $script:AuthHeaders -ContentType "application/json" | Out-Null
    }
    catch {
        # Swallow 404 – ACE may not exist
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
#  PROVISIONING
# ═══════════════════════════════════════════════════════════════════════════════
function Invoke-Provision ([array]$OrderedTeams, [array]$AllRows) {

    # ── 1. Current state ────────────────────────────────────────────────────
    Write-Banner "Fetching current project state"
    $existingTeams              = Get-ExistingTeams
    $areaTree                   = Get-AreaTree
    $iterTree                   = Get-IterationTree
    $script:KnownAreaPaths      = @(Get-FlatPaths $areaTree)
    $script:KnownIterationPaths = @(Get-FlatPaths $iterTree)

    Write-Step ("Teams: $($existingTeams.Count)  |  " +
                "Areas: $($script:KnownAreaPaths.Count)  |  " +
                "Iterations: $($script:KnownIterationPaths.Count)") -Status INFO

    # Build a quick-lookup hashtable of existing teams by Id
    $existingById = @{}
    foreach ($t in $existingTeams) { $existingById[$t.id] = $t }

    # Track teams that were renamed (old name → new name) for resource cleanup
    $renamedTeams = [System.Collections.Generic.List[hashtable]]::new()

    # ── 2. Create / Update Teams ────────────────────────────────────────────
    Write-Banner "Creating / Updating Teams"
    foreach ($team in $OrderedTeams) {
        $hasId = $team.PSObject.Properties["Id"] -and $team.Id

        if ($hasId) {
            # CSV row carries an Id – try to find the existing team by Id
            $byId = if ($existingById.ContainsKey($team.Id)) { $existingById[$team.Id] } else { $null }

            if ($byId) {
                # Team exists – update name & description if needed
                $desc = if ($team.PSObject.Properties["Description"]) { $team.Description } else { "" }
                $nameChanged = $byId.name -ne $team.TeamName
                $descChanged = ($byId.description -replace '\s+$','') -ne ($desc -replace '\s+$','')

                if ($nameChanged -or $descChanged) {
                    Update-DevOpsTeam -TeamId $team.Id -Name $team.TeamName -Description $desc
                    if ($nameChanged) {
                        Write-Step "Renamed team '$($byId.name)' -> '$($team.TeamName)' (id $($team.Id))" -Status CREATE
                        $renamedTeams.Add(@{ OldName = $byId.name; NewName = $team.TeamName; Row = $team })
                    }
                    if ($descChanged) {
                        Write-Step "Updated description for '$($team.TeamName)' (id $($team.Id))" -Status CREATE
                    }
                }
                else {
                    Write-Step "Team '$($team.TeamName)' up to date (id $($team.Id))" -Status SKIP
                }
            }
            else {
                # Id does not match any existing team – treat as new
                Write-Step "Id '$($team.Id)' not found – creating team '$($team.TeamName)' as new" -Status WARN
                $desc    = if ($team.PSObject.Properties["Description"]) { $team.Description } else { "" }
                $created = New-DevOpsTeam -Name $team.TeamName -Description $desc
                $team.Id = $created.id
                Write-Step "Created team '$($team.TeamName)' (id $($created.id))" -Status CREATE
            }
        }
        else {
            # No Id in CSV – match by name (original behaviour) or create
            $existing = $existingTeams | Where-Object { $_.name -eq $team.TeamName }
            if ($existing) {
                Write-Step "Team '$($team.TeamName)' already exists (id $($existing.id))" -Status SKIP
                if ($team.PSObject.Properties["Id"]) { $team.Id = $existing.id }
            }
            else {
                $desc    = if ($team.PSObject.Properties["Description"]) { $team.Description } else { "" }
                $created = New-DevOpsTeam -Name $team.TeamName -Description $desc
                Write-Step "Created team '$($team.TeamName)' (id $($created.id))" -Status CREATE
                if ($team.PSObject.Properties["Id"]) { $team.Id = $created.id }
            }
        }
    }

    # Refresh team list so we have IDs for all teams
    $existingTeams = Get-ExistingTeams

    # ── Sync Ids back to AllRows (AllRows and OrderedTeams share the same objects) ──
    foreach ($row in $AllRows) {
        if ($row.PSObject.Properties["Id"] -and -not $row.Id) {
            $match = $existingTeams | Where-Object { $_.name -eq $row.TeamName }
            if ($match) { $row.Id = $match.id }
        }
    }

    # ── 3. Create Area Paths ────────────────────────────────────────────────
    Write-Banner "Creating Area Paths"
    foreach ($team in $OrderedTeams) {
        $teamAreaPaths = Resolve-TeamAreaPaths -TeamName $team.TeamName -Rows $AllRows
        foreach ($areaRel in $teamAreaPaths) {
            New-AreaNode -RelativePath $areaRel
        }
    }

    # ── 3b. Remove unused default hierarchy areas ───────────────────────────
    #  When a team has custom AreaPaths, the hierarchy-derived default area path
    #  (based on team names) may exist from a previous run and should be cleaned
    #  up if no team still uses it.
    $allUsedPaths = @()
    foreach ($t in $OrderedTeams) {
        $allUsedPaths += Resolve-TeamAreaPaths -TeamName $t.TeamName -Rows $AllRows
    }
    $allUsedPaths = $allUsedPaths | Select-Object -Unique

    foreach ($team in $OrderedTeams) {
        $hasCustom = $team.PSObject.Properties["AreaPaths"] -and $team.AreaPaths
        if (-not $hasCustom) { continue }

        $defaultArea = Resolve-AreaPath -TeamName $team.TeamName -Rows $AllRows

        # Skip if the default path is still in use by some team
        if ($defaultArea -in $allUsedPaths) { continue }

        # Skip if it is a structural parent of any used path
        $isParent = $allUsedPaths | Where-Object {
            $_.StartsWith("$defaultArea\", [StringComparison]::OrdinalIgnoreCase)
        }
        if ($isParent) { continue }

        # Remove if it exists
        if ($defaultArea -in $script:KnownAreaPaths) {
            try {
                $rootAreaIdForClean = [int]$areaTree.id
                Remove-AreaNode -RelativePath $defaultArea -ReclassifyId $rootAreaIdForClean
                Write-Step "Removed unused default area '$script:ProjectName\$defaultArea'" -Status DELETE
            }
            catch {
                Write-Step "Could not remove default area '$script:ProjectName\$defaultArea' (may have children): $_" -Status WARN
            }
        }
    }

    # ── 4. Assign Area to each Team ─────────────────────────────────────────
    Write-Banner "Configuring Team Area Settings"
    foreach ($team in $OrderedTeams) {
        $teamAreaPaths = Resolve-TeamAreaPaths -TeamName $team.TeamName -Rows $AllRows
        $fullPaths     = $teamAreaPaths | ForEach-Object { "$script:ProjectName\$_" }
        $teamObj       = $existingTeams | Where-Object { $_.name -eq $team.TeamName }

        try {
            Set-TeamArea -TeamId $teamObj.id -FullAreaPaths $fullPaths
            Write-Step "'$($team.TeamName)' -> area(s): $($teamAreaPaths -join '; ')" -Status CREATE
        }
        catch {
            Write-Step "Could not set area for '$($team.TeamName)': $_" -Status ERROR
        }
    }

    # ── 5. Create Iterations ────────────────────────────────────────────────
    $hasIterations = $AllRows[0].PSObject.Properties["Iterations"]
    if ($hasIterations) {
        Write-Banner "Creating Iterations"
        $allIters = $AllRows |
            Where-Object  { $_.Iterations } |
            ForEach-Object { $_.Iterations -split ';' } |
            ForEach-Object { $_.Trim() } |
            Where-Object  { $_ } |
            Sort-Object -Unique

        foreach ($iter in $allIters) {
            New-IterationNode -RelativePath $iter
        }

        # ── 6. Set Backlog Iteration & Subscribe Teams ─────────────────────
        Write-Banner "Subscribing Teams to Iterations"
        $rootIterTree = Get-IterationTree
        $rootIterGuid = $rootIterTree.identifier

        foreach ($team in $OrderedTeams) {
            if (-not $team.Iterations) { continue }
            $teamObj = $existingTeams | Where-Object { $_.name -eq $team.TeamName }

            # Ensure the team has a valid backlog iteration before subscribing
            try {
                Set-TeamBacklogIteration -TeamId $teamObj.id -IterationGuid $rootIterGuid
                Write-Step "Set backlog iteration for '$($team.TeamName)'" -Status INFO
            }
            catch {
                Write-Step "Could not set backlog iteration for '$($team.TeamName)': $_" -Status WARN
            }

            $iters   = ($team.Iterations -split ';') |
                        ForEach-Object { $_.Trim() } |
                        Where-Object  { $_ }

            foreach ($iterPath in $iters) {
                try {
                    $node = Get-IterationNodeByPath -RelativePath $iterPath
                    Add-TeamIteration -TeamId $teamObj.id -IterationGuid $node.identifier
                    Write-Step "'$($team.TeamName)' subscribed to '$iterPath'" -Status CREATE
                }
                catch {
                    Write-Step "Could not subscribe '$($team.TeamName)' to '$iterPath': $_" -Status WARN
                }
            }
        }
    }

    # ── 7. Add / Sync Members ───────────────────────────────────────────────
    $hasMembers = $AllRows[0].PSObject.Properties["Members"]
    $anyMembers = $AllRows | Where-Object { $_.PSObject.Properties["Members"] -and $_.Members }

    if ($hasMembers -and $anyMembers) {
        Write-Banner "Syncing Team Members"
        $scopeDesc = Get-ProjectScopeDescriptor
        $groups    = Get-ProjectGroups $scopeDesc

        # Resolve the built-in "Contributors" group so team members can contribute to boards
        $contributorsGroup = Find-GroupByDisplayName -DisplayName "Contributors" -Groups $groups
        $contributorsDesc  = $null
        if ($contributorsGroup) {
            $contributorsDesc = $contributorsGroup.descriptor
            Write-Step "Found built-in 'Contributors' group" -Status INFO
        }
        else {
            Write-Step "Built-in 'Contributors' group not found – members will not be added to Contributors" -Status WARN
        }

        foreach ($team in $OrderedTeams) {
            $desiredEmails = @()
            if ($team.Members) {
                $desiredEmails = ($team.Members -split ';') |
                    ForEach-Object { $_.Trim() } | Where-Object { $_ }
            }
            if (-not $desiredEmails) { continue }

            $groupDesc = Find-TeamGroupDescriptor -TeamName $team.TeamName -Groups $groups

            if (-not $groupDesc) {
                # Refresh groups (team was just created or renamed)
                $groups    = Get-ProjectGroups $scopeDesc
                $groupDesc = Find-TeamGroupDescriptor -TeamName $team.TeamName -Groups $groups
            }
            if (-not $groupDesc) {
                Write-Step "Cannot resolve group for '$($team.TeamName)' – skipping members" -Status WARN
                continue
            }

            # Get current members for diff
            $teamObj = $existingTeams | Where-Object { $_.name -eq $team.TeamName }
            $currentDescriptors = Get-TeamMemberDescriptors -TeamId $teamObj.id -ScopeDescriptor $scopeDesc
            $currentEmails = $currentDescriptors | ForEach-Object { $_.Email.ToLower() }

            # Add missing members
            foreach ($email in $desiredEmails) {
                if ($currentEmails -contains $email.ToLower()) {
                    Write-Step "'$email' already in '$($team.TeamName)'" -Status SKIP
                    continue
                }
                try {
                    $user = Find-UserByEmail -Email $email
                    if (-not $user) {
                        Write-Step "User '$email' not found in org '$Organization'" -Status WARN
                        continue
                    }
                    Add-GroupMember -MemberDescriptor $user.descriptor -GroupDescriptor $groupDesc
                    Write-Step "Added '$email' to '$($team.TeamName)'" -Status CREATE

                    # Also add to the built-in Contributors group for board access
                    if ($contributorsDesc) {
                        try {
                            Add-GroupMember -MemberDescriptor $user.descriptor -GroupDescriptor $contributorsDesc
                            Write-Step "Added '$email' to 'Contributors'" -Status CREATE
                        }
                        catch {
                            Write-Step "Could not add '$email' to 'Contributors': $_" -Status WARN
                        }
                    }
                }
                catch {
                    Write-Step "Failed to add '$email' to '$($team.TeamName)': $_" -Status ERROR
                }
            }

            # Remove members not in desired list
            $desiredLower = $desiredEmails | ForEach-Object { $_.ToLower() }
            foreach ($cd in $currentDescriptors) {
                if ($desiredLower -contains $cd.Email.ToLower()) { continue }
                try {
                    Remove-GroupMember -MemberDescriptor $cd.Descriptor -GroupDescriptor $groupDesc
                    Write-Step "Removed '$($cd.Email)' from '$($team.TeamName)' (not in CSV)" -Status DELETE
                }
                catch {
                    Write-Step "Could not remove '$($cd.Email)' from '$($team.TeamName)': $_" -Status WARN
                }
            }
        }
    }

    # ── 8. Create Role / Permission Groups & Set Area Permissions ───────────
    Write-Banner "Creating Role & Permission Groups"
    $scopeDescForGroups = Get-ProjectScopeDescriptor
    $allGroupsNow       = Get-ProjectGroups $scopeDescForGroups

    foreach ($team in $OrderedTeams) {
        $tn = $team.TeamName.ToLower() -replace '\s+', '-'

        $roleGroupName   = "role-external-contributor-$tn"
        $readerGroupName = "perm-item-reader-$tn"
        $writerGroupName = "perm-item-writer-$tn"

        # ── Create role-external-contributor group ──────────────────────────
        $roleGroup = Find-GroupByDisplayName -DisplayName $roleGroupName -Groups $allGroupsNow
        if ($roleGroup) {
            Write-Step "Group '$roleGroupName' already exists" -Status SKIP
        }
        else {
            $roleGroup = New-DevOpsGroup `
                -DisplayName $roleGroupName `
                -Description "External contributor role for team '$($team.TeamName)'" `
                -ScopeDescriptor $scopeDescForGroups
            Write-Step "Created group '$roleGroupName'" -Status CREATE
        }

        # ── Create perm-item-reader group ───────────────────────────────────
        $readerGroup = Find-GroupByDisplayName -DisplayName $readerGroupName -Groups $allGroupsNow
        if ($readerGroup) {
            Write-Step "Group '$readerGroupName' already exists" -Status SKIP
        }
        else {
            $readerGroup = New-DevOpsGroup `
                -DisplayName $readerGroupName `
                -Description "Work item reader permission group for team '$($team.TeamName)'" `
                -ScopeDescriptor $scopeDescForGroups
            Write-Step "Created group '$readerGroupName'" -Status CREATE
        }

        # Add role group as member of reader group
        try {
            Add-GroupMember -MemberDescriptor $roleGroup.descriptor -GroupDescriptor $readerGroup.descriptor
            Write-Step "Added '$roleGroupName' as member of '$readerGroupName'" -Status CREATE
        }
        catch {
            Write-Step "Could not add '$roleGroupName' to '$readerGroupName': $_" -Status WARN
        }

        # ── Create perm-item-writer group ───────────────────────────────────
        $writerGroup = Find-GroupByDisplayName -DisplayName $writerGroupName -Groups $allGroupsNow
        if ($writerGroup) {
            Write-Step "Group '$writerGroupName' already exists" -Status SKIP
        }
        else {
            $writerGroup = New-DevOpsGroup `
                -DisplayName $writerGroupName `
                -Description "Work item writer permission group for team '$($team.TeamName)'" `
                -ScopeDescriptor $scopeDescForGroups
            Write-Step "Created group '$writerGroupName'" -Status CREATE
        }

        # Add role group as member of writer group
        try {
            Add-GroupMember -MemberDescriptor $roleGroup.descriptor -GroupDescriptor $writerGroup.descriptor
            Write-Step "Added '$roleGroupName' as member of '$writerGroupName'" -Status CREATE
        }
        catch {
            Write-Step "Could not add '$roleGroupName' to '$writerGroupName': $_" -Status WARN
        }

        # ── Set Area Path Permissions ───────────────────────────────────────
        $teamAreaPaths = Resolve-TeamAreaPaths -TeamName $team.TeamName -Rows $AllRows
        $readerIdentity = Resolve-IdentityDescriptor -SubjectDescriptor $readerGroup.descriptor
        $writerIdentity = Resolve-IdentityDescriptor -SubjectDescriptor $writerGroup.descriptor

        foreach ($areaRel in $teamAreaPaths) {
            try {
                $secToken = Build-AreaSecurityToken -RelativePath $areaRel

                # perm-item-reader: View work items in this node (bit 16)
                Set-AreaPathPermission -SecurityToken $secToken `
                    -IdentityDescriptor $readerIdentity -AllowBits 16
                Write-Step "Set 'View work items' for '$readerGroupName' on '$areaRel'" -Status CREATE

                # perm-item-writer: Edit work items (bit 32) + Edit work item comments (bit 512)
                Set-AreaPathPermission -SecurityToken $secToken `
                    -IdentityDescriptor $writerIdentity -AllowBits (32 -bor 512)
                Write-Step "Set 'Edit work items + comments' for '$writerGroupName' on '$areaRel'" -Status CREATE
            }
            catch {
                Write-Step "Could not set area permissions for '$($team.TeamName)' on '$areaRel': $_" -Status ERROR
            }
        }

        # Refresh group list so next iteration sees the newly created groups
        $allGroupsNow = Get-ProjectGroups $scopeDescForGroups
    }

    # ── 9. Deny inherited permissions on child area paths ────────────────────
    #  For each team that has children, explicitly DENY the parent's reader/writer
    #  permissions on every child area path so inherited ALLOWs do not leak down.
    Write-Banner "Denying inherited permissions on child area paths"
    $allGroupsNow = Get-ProjectGroups $scopeDescForGroups

    # Build a complete map of "team → area paths" for deny lookups
    $allTeamAreaMap = @{}
    foreach ($t in $OrderedTeams) {
        $allTeamAreaMap[$t.TeamName] = @(Resolve-TeamAreaPaths -TeamName $t.TeamName -Rows $AllRows)
    }

    foreach ($team in $OrderedTeams) {
        $tn = $team.TeamName.ToLower() -replace '\s+', '-'
        $readerGroupName = "perm-item-reader-$tn"
        $writerGroupName = "perm-item-writer-$tn"

        $readerGroup = Find-GroupByDisplayName -DisplayName $readerGroupName -Groups $allGroupsNow
        $writerGroup = Find-GroupByDisplayName -DisplayName $writerGroupName -Groups $allGroupsNow
        if (-not $readerGroup -and -not $writerGroup) { continue }

        $teamPaths = $allTeamAreaMap[$team.TeamName]

        # For each of this team's area paths, find all OTHER teams' paths nested underneath
        $denyTargets = [System.Collections.Generic.HashSet[string]]::new(
                           [StringComparer]::OrdinalIgnoreCase)
        foreach ($tp in $teamPaths) {
            foreach ($otherTeam in $AllRows) {
                if ($otherTeam.TeamName -eq $team.TeamName) { continue }
                foreach ($op in $allTeamAreaMap[$otherTeam.TeamName]) {
                    if ($op.StartsWith("$tp\", [StringComparison]::OrdinalIgnoreCase)) {
                        [void]$denyTargets.Add($op)
                    }
                }
            }
        }

        if ($denyTargets.Count -eq 0) { continue }

        foreach ($childAreaRel in $denyTargets) {
            try {
                $childSecToken = Build-AreaSecurityToken -RelativePath $childAreaRel

                if ($readerGroup) {
                    $readerIdentity = Resolve-IdentityDescriptor -SubjectDescriptor $readerGroup.descriptor
                    Deny-AreaPathPermission -SecurityToken $childSecToken `
                        -IdentityDescriptor $readerIdentity -DenyBits 16
                    Write-Step "Deny 'View work items' for '$readerGroupName' on '$childAreaRel'" -Status CREATE
                }

                if ($writerGroup) {
                    $writerIdentity = Resolve-IdentityDescriptor -SubjectDescriptor $writerGroup.descriptor
                    Deny-AreaPathPermission -SecurityToken $childSecToken `
                        -IdentityDescriptor $writerIdentity -DenyBits (32 -bor 512)
                    Write-Step "Deny 'Edit work items + comments' for '$writerGroupName' on '$childAreaRel'" -Status CREATE
                }
            }
            catch {
                Write-Step "Could not deny permissions for '$($team.TeamName)' on '$childAreaRel': $_" -Status ERROR
            }
        }
    }
    # ── 10. Clean up old-name resources after team renames ───────────────────
    if ($renamedTeams.Count -gt 0) {
        Write-Banner "Cleaning up old-name resources (renames)"
        $allGroupsNow  = Get-ProjectGroups $scopeDescForGroups
        $areaTree      = Get-AreaTree
        $currentAreas  = @(Get-FlatPaths $areaTree)
        $rootAreaId    = [int]$areaTree.id

        # Collect all area paths that are still in use by any team
        $allUsedPaths = [System.Collections.Generic.HashSet[string]]::new(
                            [StringComparer]::OrdinalIgnoreCase)
        foreach ($t in $OrderedTeams) {
            foreach ($ap in (Resolve-TeamAreaPaths -TeamName $t.TeamName -Rows $AllRows)) {
                [void]$allUsedPaths.Add($ap)
            }
        }

        foreach ($rename in $renamedTeams) {
            $oldName = $rename.OldName
            $newName = $rename.NewName
            $oldTn   = $oldName.ToLower() -replace '\s+', '-'

            $oldRoleGroupName   = "role-external-contributor-$oldTn"
            $oldReaderGroupName = "perm-item-reader-$oldTn"
            $oldWriterGroupName = "perm-item-writer-$oldTn"

            $oldReaderGroup = Find-GroupByDisplayName -DisplayName $oldReaderGroupName -Groups $allGroupsNow
            $oldWriterGroup = Find-GroupByDisplayName -DisplayName $oldWriterGroupName -Groups $allGroupsNow
            $oldRoleGroup   = Find-GroupByDisplayName -DisplayName $oldRoleGroupName   -Groups $allGroupsNow

            # ── Remove permissions using old group identities ───────────────
            #  Build the old area paths by temporarily resolving with the old name
            #  Old hierarchy-derived default: walk the parent chain, but the old team
            #  name sits where the new name is now, so replicate the old default path.
            $oldDefaultArea = (Resolve-AreaPath -TeamName $newName -Rows $AllRows) -replace
                [regex]::Escape($newName), $oldName

            # Collect the old area paths that may still have ACLs to clean
            $oldAreaPaths = @($oldDefaultArea)

            if ($oldReaderGroup -or $oldWriterGroup) {
                $oldReaderIdentity = $null
                $oldWriterIdentity = $null
                if ($oldReaderGroup) {
                    $oldReaderIdentity = Resolve-IdentityDescriptor -SubjectDescriptor $oldReaderGroup.descriptor
                }
                if ($oldWriterGroup) {
                    $oldWriterIdentity = Resolve-IdentityDescriptor -SubjectDescriptor $oldWriterGroup.descriptor
                }

                # Remove ALLOW and DENY permissions from old area paths
                foreach ($oap in $oldAreaPaths) {
                    if ($oap -notin $currentAreas) { continue }
                    try {
                        $secToken = Build-AreaSecurityToken -RelativePath $oap
                        if ($oldReaderIdentity) {
                            Remove-AreaPathPermissions -SecurityToken $secToken -IdentityDescriptor $oldReaderIdentity
                            Write-Step "Removed old permissions for '$oldReaderGroupName' on '$oap'" -Status DELETE
                        }
                        if ($oldWriterIdentity) {
                            Remove-AreaPathPermissions -SecurityToken $secToken -IdentityDescriptor $oldWriterIdentity
                            Write-Step "Removed old permissions for '$oldWriterGroupName' on '$oap'" -Status DELETE
                        }
                    }
                    catch { Write-Step "Could not remove old permissions on '$oap': $_" -Status WARN }
                }
            }

            # ── Delete old groups ──────────────────────────────────────────
            foreach ($g in @(
                @{ Name = $oldReaderGroupName; Obj = $oldReaderGroup },
                @{ Name = $oldWriterGroupName; Obj = $oldWriterGroup },
                @{ Name = $oldRoleGroupName;   Obj = $oldRoleGroup }
            )) {
                if ($g.Obj) {
                    try {
                        Remove-DevOpsGroup -GroupDescriptor $g.Obj.descriptor
                        Write-Step "Deleted old group '$($g.Name)'" -Status DELETE
                    }
                    catch { Write-Step "Could not delete old group '$($g.Name)': $_" -Status WARN }
                }
            }

            # ── Delete old area paths ──────────────────────────────────────
            #  Only remove old hierarchy-derived areas that are no longer used
            foreach ($oap in ($oldAreaPaths | Sort-Object { ($_ -split '\\').Count } -Descending)) {
                if ($oap -notin $currentAreas) { continue }
                if ($allUsedPaths.Contains($oap)) { continue }
                # Skip if it is a structural parent of any used path
                $isParent = $allUsedPaths | Where-Object {
                    $_.StartsWith("$oap\", [StringComparison]::OrdinalIgnoreCase)
                }
                if ($isParent) { continue }
                try {
                    Remove-AreaNode -RelativePath $oap -ReclassifyId $rootAreaId
                    Write-Step "Deleted old area '$script:ProjectName\$oap'" -Status DELETE
                }
                catch { Write-Step "Could not delete old area '$script:ProjectName\$oap': $_" -Status WARN }
            }

            Write-Step "Rename cleanup complete: '$oldName' -> '$newName'" -Status INFO
        }
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
#  CLEANUP
# ═══════════════════════════════════════════════════════════════════════════════
function Invoke-Cleanup ([array]$OrderedTeams, [array]$AllRows) {

    Write-Banner "CLEANUP MODE – Removing provisioned resources"

    $existingTeams = Get-ExistingTeams
    $areaTree      = Get-AreaTree
    $iterTree      = Get-IterationTree
    $rootAreaId    = $areaTree.id
    $rootIterId    = $iterTree.id

    # ── 0. Remove Role / Permission Groups & Area Permissions ───────────────
    Write-Banner "Removing Role & Permission Groups"
    $scopeDescClean = Get-ProjectScopeDescriptor
    $allGroupsClean = Get-ProjectGroups $scopeDescClean

    foreach ($team in $OrderedTeams) {
        $tn = $team.TeamName.ToLower() -replace '\s+', '-'

        $roleGroupName   = "role-external-contributor-$tn"
        $readerGroupName = "perm-item-reader-$tn"
        $writerGroupName = "perm-item-writer-$tn"

        $readerGroup = Find-GroupByDisplayName -DisplayName $readerGroupName -Groups $allGroupsClean
        $writerGroup = Find-GroupByDisplayName -DisplayName $writerGroupName -Groups $allGroupsClean
        $roleGroup   = Find-GroupByDisplayName -DisplayName $roleGroupName   -Groups $allGroupsClean

        # Resolve identity descriptors for permission removal
        $readerIdentity = $null
        $writerIdentity = $null
        if ($readerGroup) {
            $readerIdentity = Resolve-IdentityDescriptor -SubjectDescriptor $readerGroup.descriptor
        }
        if ($writerGroup) {
            $writerIdentity = Resolve-IdentityDescriptor -SubjectDescriptor $writerGroup.descriptor
        }

        # Remove ALLOW permissions from all of the team's area paths
        $teamAreaPaths = Resolve-TeamAreaPaths -TeamName $team.TeamName -Rows $AllRows
        foreach ($areaRel in $teamAreaPaths) {
            try {
                $secToken = Build-AreaSecurityToken -RelativePath $areaRel

                if ($readerIdentity) {
                    Remove-AreaPathPermissions -SecurityToken $secToken -IdentityDescriptor $readerIdentity
                    Write-Step "Removed area permissions for '$readerGroupName' on '$areaRel'" -Status DELETE
                }
                if ($writerIdentity) {
                    Remove-AreaPathPermissions -SecurityToken $secToken -IdentityDescriptor $writerIdentity
                    Write-Step "Removed area permissions for '$writerGroupName' on '$areaRel'" -Status DELETE
                }
            }
            catch { Write-Step "Could not remove area permissions on '$areaRel': $_" -Status WARN }
        }

        # Remove DENY permissions from child area paths
        $denyCleanTargets = [System.Collections.Generic.HashSet[string]]::new(
                                [StringComparer]::OrdinalIgnoreCase)
        foreach ($tp in $teamAreaPaths) {
            foreach ($otherTeam in $AllRows) {
                if ($otherTeam.TeamName -eq $team.TeamName) { continue }
                $otherPaths = Resolve-TeamAreaPaths -TeamName $otherTeam.TeamName -Rows $AllRows
                foreach ($op in $otherPaths) {
                    if ($op.StartsWith("$tp\", [StringComparison]::OrdinalIgnoreCase)) {
                        [void]$denyCleanTargets.Add($op)
                    }
                }
            }
        }
        foreach ($childAreaRel in $denyCleanTargets) {
            try {
                $childSecToken = Build-AreaSecurityToken -RelativePath $childAreaRel
                if ($readerIdentity) {
                    Remove-AreaPathPermissions -SecurityToken $childSecToken -IdentityDescriptor $readerIdentity
                    Write-Step "Removed deny for '$readerGroupName' on '$childAreaRel'" -Status DELETE
                }
                if ($writerIdentity) {
                    Remove-AreaPathPermissions -SecurityToken $childSecToken -IdentityDescriptor $writerIdentity
                    Write-Step "Removed deny for '$writerGroupName' on '$childAreaRel'" -Status DELETE
                }
            }
            catch { Write-Step "Could not remove deny on '$childAreaRel': $_" -Status WARN }
        }

        # Delete groups (reader, writer, then role)
        foreach ($g in @(
            @{ Name = $readerGroupName; Obj = $readerGroup },
            @{ Name = $writerGroupName; Obj = $writerGroup },
            @{ Name = $roleGroupName;   Obj = $roleGroup }
        )) {
            if ($g.Obj) {
                try {
                    Remove-DevOpsGroup -GroupDescriptor $g.Obj.descriptor
                    Write-Step "Deleted group '$($g.Name)'" -Status DELETE
                }
                catch { Write-Step "Could not delete group '$($g.Name)': $_" -Status WARN }
            }
            else {
                Write-Step "Group '$($g.Name)' not found – nothing to delete" -Status SKIP
            }
        }
    }

    # ── 1. Remove Team Members ──────────────────────────────────────────────
    $hasMembers = $AllRows[0].PSObject.Properties["Members"]
    $anyMembers = $AllRows | Where-Object { $_.PSObject.Properties["Members"] -and $_.Members }

    if ($hasMembers -and $anyMembers) {
        Write-Banner "Removing Team Members"
        $scopeDesc = Get-ProjectScopeDescriptor
        $groups    = Get-ProjectGroups $scopeDesc

        foreach ($team in $OrderedTeams) {
            if (-not $team.Members) { continue }
            $groupDesc = Find-TeamGroupDescriptor -TeamName $team.TeamName -Groups $groups
            if (-not $groupDesc) { continue }

            $emails = ($team.Members -split ';') | ForEach-Object { $_.Trim() } | Where-Object { $_ }
            foreach ($email in $emails) {
                try {
                    $user = Find-UserByEmail -Email $email
                    if ($user) {
                        Remove-GroupMember -MemberDescriptor $user.descriptor -GroupDescriptor $groupDesc
                        Write-Step "Removed '$email' from '$($team.TeamName)'" -Status DELETE
                    }
                }
                catch { Write-Step "Could not remove '$email': $_" -Status WARN }
            }
        }
    }

    # ── 2. Reset Team Area Settings to project root ─────────────────────────
    Write-Banner "Resetting Team Area Settings"
    foreach ($team in $OrderedTeams) {
        $existing = $existingTeams | Where-Object { $_.name -eq $team.TeamName }
        if (-not $existing) { continue }
        try {
            Set-TeamArea -TeamId $existing.id -FullAreaPaths @($script:ProjectName)
            Write-Step "Reset area for '$($team.TeamName)'" -Status DELETE
        }
        catch { Write-Step "Could not reset area for '$($team.TeamName)': $_" -Status WARN }
    }

    # ── 3. Delete Teams (reverse order – children first) ────────────────────
    Write-Banner "Deleting Teams"
    $reversed = @($OrderedTeams)
    [Array]::Reverse($reversed)

    # Identify the default team so we can skip its deletion
    $defaultTeam = Get-DefaultTeam
    $defaultTeamId = if ($defaultTeam) { $defaultTeam.id } else { $null }
    Write-Step "Default team: '$($defaultTeam.name)' (id $defaultTeamId)" -Status INFO

    foreach ($team in $reversed) {
        $existing = $existingTeams | Where-Object { $_.name -eq $team.TeamName }
        if (-not $existing) {
            Write-Step "Team '$($team.TeamName)' not found – nothing to delete" -Status SKIP
            continue
        }

        # Check if this is the default team
        if ($defaultTeamId -and $existing.id -eq $defaultTeamId) {
            Write-Step "'$($team.TeamName)' is the default team – renaming to 'Default Team' instead of deleting" -Status INFO

            # Rename to "Default Team" (skip if already named that)
            if ($existing.name -ne "Default Team") {
                # If a team named "Default Team" already exists (stale from prior run), remove it first
                $staleDefault = $existingTeams | Where-Object { $_.name -eq "Default Team" -and $_.id -ne $existing.id }
                if ($staleDefault) {
                    try {
                        Remove-DevOpsTeam -TeamId $staleDefault.id
                        Write-Step "Removed stale 'Default Team' (id $($staleDefault.id))" -Status DELETE
                    }
                    catch { Write-Step "Could not remove stale 'Default Team': $_" -Status WARN }
                }

                try {
                    Rename-DevOpsTeam -TeamId $existing.id -NewName "Default Team"
                    Write-Step "Renamed '$($team.TeamName)' -> 'Default Team'" -Status CREATE
                }
                catch { Write-Step "Could not rename default team: $_" -Status WARN }
            }
            else {
                Write-Step "Default team is already named 'Default Team'" -Status SKIP
            }

            # Reset area to project root
            try {
                Set-TeamArea -TeamId $existing.id -FullAreaPaths @($script:ProjectName)
                Write-Step "Reset area for default team to project root" -Status DELETE
            }
            catch { Write-Step "Could not reset area for default team: $_" -Status WARN }

            # Replace members: remove all current, add only the project owner
            $scopeDescForDefault = Get-ProjectScopeDescriptor
            $groupsForDefault    = Get-ProjectGroups $scopeDescForDefault
            $groupDesc = Find-TeamGroupDescriptor -TeamName $team.TeamName -Groups $groupsForDefault
            if (-not $groupDesc) {
                # Team was just renamed, try with new name
                $groupsForDefault = Get-ProjectGroups $scopeDescForDefault
                $groupDesc = Find-TeamGroupDescriptor -TeamName "Default Team" -Groups $groupsForDefault
            }

            if ($groupDesc) {
                # Remove existing members
                $currentMembers = Get-TeamMemberDescriptors -TeamId $existing.id -ScopeDescriptor $scopeDescForDefault
                foreach ($m in $currentMembers) {
                    try {
                        Remove-GroupMember -MemberDescriptor $m.Descriptor -GroupDescriptor $groupDesc
                        Write-Step "Removed '$($m.Email)' from default team" -Status DELETE
                    }
                    catch { Write-Step "Could not remove '$($m.Email)' from default team: $_" -Status WARN }
                }

                # Add project owner
                $owner = Get-ProjectOwner
                if ($owner) {
                    $ownerEmail = ""
                    if ($owner.PSObject.Properties["properties"] -and $owner.properties.PSObject.Properties["Account"]) {
                        $ownerEmail = $owner.properties.Account.'$value'
                    }
                    if (-not $ownerEmail -and $owner.PSObject.Properties["providerDisplayName"]) {
                        $ownerEmail = $owner.providerDisplayName
                    }

                    if ($ownerEmail) {
                        $ownerUser = Find-UserByEmail -Email $ownerEmail
                        if ($ownerUser) {
                            try {
                                Add-GroupMember -MemberDescriptor $ownerUser.descriptor -GroupDescriptor $groupDesc
                                Write-Step "Added project owner '$ownerEmail' to 'Default Team'" -Status CREATE
                            }
                            catch { Write-Step "Could not add owner to default team: $_" -Status WARN }
                        }
                        else {
                            Write-Step "Could not find owner '$ownerEmail' in org" -Status WARN
                        }
                    }
                    else {
                        Write-Step "Could not determine project owner email" -Status WARN
                    }
                }
            }

            continue
        }

        try {
            Remove-DevOpsTeam -TeamId $existing.id
            Write-Step "Deleted team '$($team.TeamName)'" -Status DELETE
        }
        catch { Write-Step "Could not delete team '$($team.TeamName)': $_" -Status ERROR }
    }

    # ── 4. Delete Area Paths (deepest first) ────────────────────────────────
    Write-Banner "Deleting Area Paths"
    $areaPaths = foreach ($team in $OrderedTeams) {
        Resolve-TeamAreaPaths -TeamName $team.TeamName -Rows $AllRows
        # Also include the hierarchy-derived default path in case it still exists
        Resolve-AreaPath -TeamName $team.TeamName -Rows $AllRows
    }
    # Sort by depth descending so leaves are deleted first
    $currentAreaPaths = @(Get-FlatPaths (Get-AreaTree))
    $areaPaths = $areaPaths |
        Sort-Object   { ($_ -split '\\').Count } -Descending |
        Select-Object -Unique |
        Where-Object  { $_ -in $currentAreaPaths }   # only attempt existing paths

    foreach ($ap in $areaPaths) {
        try {
            Remove-AreaNode -RelativePath $ap -ReclassifyId $rootAreaId
            Write-Step "Deleted area '$script:ProjectName\$ap'" -Status DELETE
        }
        catch { Write-Step "Could not delete area '$script:ProjectName\$ap': $_" -Status WARN }
    }

    # Also try to remove parent area segments that may now be empty
    $parentSegs = $areaPaths | ForEach-Object {
        $parts = $_ -split '\\'
        for ($i = $parts.Count - 2; $i -ge 0; $i--) {
            ($parts[0..$i] -join '\')
        }
    } | Sort-Object { ($_ -split '\\').Count } -Descending |
        Select-Object -Unique |
        Where-Object  { $_ -notin $areaPaths }

    foreach ($ps in $parentSegs) {
        try {
            Remove-AreaNode -RelativePath $ps -ReclassifyId $rootAreaId
            Write-Step "Deleted parent area '$script:ProjectName\$ps'" -Status DELETE
        }
        catch { <# parent may still have other children – that is fine #> }
    }

    # ── 5. Delete Iterations ────────────────────────────────────────────────
    $hasIterations = $AllRows[0].PSObject.Properties["Iterations"]
    if ($hasIterations) {
        Write-Banner "Deleting Iterations"
        $allIters = $AllRows |
            Where-Object  { $_.Iterations } |
            ForEach-Object { $_.Iterations -split ';' } |
            ForEach-Object { $_.Trim() } |
            Where-Object  { $_ } |
            Sort-Object    { ($_ -split '\\').Count } -Descending |
            Select-Object  -Unique

        foreach ($iter in $allIters) {
            try {
                Remove-IterationNode -RelativePath $iter -ReclassifyId $rootIterId
                Write-Step "Deleted iteration '$script:ProjectName\$iter'" -Status DELETE
            }
            catch { Write-Step "Could not delete iteration '$script:ProjectName\$iter': $_" -Status WARN }
        }
    }

    Write-Banner "Cleanup complete"
}

# ═══════════════════════════════════════════════════════════════════════════════
#  MAIN
# ═══════════════════════════════════════════════════════════════════════════════
Write-Banner "Azure DevOps Team Provisioner"
Write-Step "Organization : $Organization" -Status INFO
Write-Step "Project      : $Project"      -Status INFO
Write-Step "CSV          : $CsvPath"      -Status INFO
Write-Step "Mode         : $(if ($Cleanup) {'CLEANUP'} else {'PROVISION'})" -Status INFO

# Validate connectivity and cache the real project name (for correct casing)
try {
    $projInfo = Invoke-AdoApi "$script:CoreUrl/_apis/projects/${Project}?api-version=7.1"
    $script:ProjectName = $projInfo.name
    Write-Step "Connected – project '$($script:ProjectName)' (id $($projInfo.id))" -Status INFO
}
catch {
    Write-Error "Cannot reach Azure DevOps. Verify organisation, project name, and PAT."
    return
}

# Parse CSV
$csvRows = Import-TeamsCsv -Path $CsvPath
if (-not (Test-CsvValid $csvRows)) { return }
$orderedTeams = Get-OrderedTeams -Rows $csvRows

Write-Step "Teams in CSV : $($csvRows.Count)" -Status INFO
Write-Step "Process order: $($orderedTeams.ForEach({ $_.TeamName }) -join ' -> ')" -Status INFO

# Execute
if ($Cleanup) {
    Invoke-Cleanup -OrderedTeams $orderedTeams -AllRows $csvRows
}
else {
    Invoke-Provision -OrderedTeams $orderedTeams -AllRows $csvRows

    # ── Write back CSV with updated Ids ─────────────────────────────────────
    if ($csvRows[0].PSObject.Properties["Id"]) {
        Write-Banner "Updating CSV with Team Ids"
        $csvRows | Select-Object Id, TeamName, ParentTeam, Description, AreaPaths, Members, Iterations |
            Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8
        Write-Step "CSV '$CsvPath' updated with team Ids" -Status CREATE
    }
}

Write-Banner "Done!"
