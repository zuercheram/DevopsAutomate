#Requires -Version 5.1

<#
.SYNOPSIS
    Exports the current Teams, Area Paths, Iterations, and Team Members from an
    Azure DevOps project into a CSV file compatible with Provision-DevOpsTeams.ps1.

.DESCRIPTION
    Connects to an Azure DevOps project and reads:
      - All teams (excluding the default project team)
      - Each team's configured area path (to derive the parent-child hierarchy)
      - Each team's subscribed iterations
      - Each team's members (email addresses)

    The output CSV has the same structure as teams-template.csv:
      Id, TeamName, ParentTeam, Description, Members, Iterations

.PARAMETER Organization
    Azure DevOps organization name (e.g. "my-org").

.PARAMETER Project
    Azure DevOps project name (e.g. "MyProject").

.PARAMETER OutputPath
    Path for the generated CSV file.  Defaults to .\<Project>-export.csv.

.PARAMETER PAT
    Personal Access Token with full access.  Falls back to the AZURE_DEVOPS_PAT env var.

.PARAMETER IncludeDefaultTeam
    By default the project's default team (same name as the project) is excluded
    because Provision-DevOpsTeams.ps1 never creates it.  Use this switch to
    include it in the export.

.EXAMPLE
    .\Export-DevOpsTeams.ps1 -Organization "my-org" -Project "MyProject" -PAT "your-pat"

.EXAMPLE
    .\Export-DevOpsTeams.ps1 -Organization "my-org" -Project "MyProject" `
        -OutputPath ".\current-state.csv" -PAT "your-pat" -IncludeDefaultTeam
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$Organization,

    [Parameter(Mandatory)]
    [string]$Project,

    [Parameter()]
    [string]$OutputPath,

    [Parameter()]
    [string]$PAT = $env:AZURE_DEVOPS_PAT,

    [switch]$IncludeDefaultTeam
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

if (-not $OutputPath) {
    $OutputPath = ".\${Project}-export.csv"
}

$script:AuthHeaders = @{
    Authorization = "Basic " + [Convert]::ToBase64String(
        [Text.Encoding]::ASCII.GetBytes(":$PAT")
    )
}

$script:CoreUrl  = "https://dev.azure.com/$Organization"
$script:VsspsUrl = "https://vssps.dev.azure.com/$Organization"

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
#  DATA RETRIEVAL FUNCTIONS
# ═══════════════════════════════════════════════════════════════════════════════

function Get-AllTeams {
    (Invoke-AdoApi "$script:CoreUrl/_apis/projects/$Project/teams?`$top=500&api-version=7.1").value
}

function Get-TeamFieldValues ([string]$TeamId) {
    # Returns the team's area path configuration
    try {
        return Invoke-AdoApi `
            "$script:CoreUrl/$Project/$TeamId/_apis/work/teamsettings/teamfieldvalues?api-version=7.1"
    }
    catch { return $null }
}

function Get-TeamIterations ([string]$TeamId) {
    try {
        return (Invoke-AdoApi `
            "$script:CoreUrl/$Project/$TeamId/_apis/work/teamsettings/iterations?api-version=7.1").value
    }
    catch { return @() }
}

function Get-TeamMembers ([string]$TeamId) {
    try {
        return (Invoke-AdoApi `
            "$script:CoreUrl/_apis/projects/$Project/teams/$TeamId/members?`$top=500&api-version=7.1").value
    }
    catch { return @() }
}

# ═══════════════════════════════════════════════════════════════════════════════
#  AREA PATH → TEAM HIERARCHY RESOLUTION
# ═══════════════════════════════════════════════════════════════════════════════

function ConvertTo-RelativeAreaPath ([string]$FullPath, [string]$ProjectName) {
    # Converts "MyProject\Platform\Backend" → "Platform\Backend"
    $prefix = "$ProjectName\"
    if ($FullPath.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase)) {
        return $FullPath.Substring($prefix.Length)
    }
    # If it's just the project name itself, return empty
    if ($FullPath -eq $ProjectName) { return "" }
    return $FullPath
}

function Resolve-ParentTeam {
    param(
        [string]$TeamName,
        [string]$RelativeAreaPath,
        [hashtable]$AreaToTeamMap,
        [System.Collections.Generic.HashSet[string]]$AllTeamNames
    )
    # Given a team's area path like "Platform\Backend", the parent is the
    # team that owns "Platform" (the parent area segment).
    if (-not $RelativeAreaPath -or -not $RelativeAreaPath.Contains('\')) {
        return ""
    }

    $parentArea = $RelativeAreaPath.Substring(0, $RelativeAreaPath.LastIndexOf('\'))

    # Walk up until we find a team that owns this area
    while ($parentArea) {
        if ($AreaToTeamMap.ContainsKey($parentArea)) {
            return $AreaToTeamMap[$parentArea]
        }
        # Fallback: if the leaf segment of this area matches a known team name,
        # that team is the parent (handles custom-area teams whose hierarchy
        # path isn't in the area-to-team map).
        $leaf = if ($parentArea.Contains('\')) {
            $parentArea.Substring($parentArea.LastIndexOf('\') + 1)
        } else { $parentArea }
        if ($leaf -ne $TeamName -and $AllTeamNames.Contains($leaf)) {
            return $leaf
        }

        if ($parentArea.Contains('\')) {
            $parentArea = $parentArea.Substring(0, $parentArea.LastIndexOf('\'))
        }
        else {
            break
        }
    }

    return ""
}

# ═══════════════════════════════════════════════════════════════════════════════
#  MAIN
# ═══════════════════════════════════════════════════════════════════════════════

Write-Banner "Azure DevOps Team Exporter"
Write-Step "Organization : $Organization" -Status INFO
Write-Step "Project      : $Project"      -Status INFO
Write-Step "Output       : $OutputPath"   -Status INFO

# ── 1. Connect & validate ──────────────────────────────────────────────────
try {
    $projInfo    = Invoke-AdoApi "$script:CoreUrl/_apis/projects/${Project}?api-version=7.1"
    $projectName = $projInfo.name
    Write-Step "Connected – project '$projectName' (id $($projInfo.id))" -Status INFO
}
catch {
    Write-Error "Cannot reach Azure DevOps. Verify organisation, project name, and PAT."
    return
}

# ── 2. Fetch all teams ─────────────────────────────────────────────────────
Write-Banner "Fetching Teams"
$allTeams = Get-AllTeams

# Find the default team (matching the project name) so we can optionally exclude it
$defaultTeamName = $projectName
$teams = if ($IncludeDefaultTeam) {
    $allTeams
}
else {
    $allTeams | Where-Object { $_.name -ne $defaultTeamName }
}

Write-Step "Found $($allTeams.Count) team(s) total, processing $($teams.Count) ($(if ($IncludeDefaultTeam) {'including'} else {'excluding'}) default team '$defaultTeamName')" -Status INFO

# ── 3. For each team, gather area path, iterations, and members ────────────
Write-Banner "Reading Team Settings"

# Build a map: relative area path → team name (for parent resolution)
$areaToTeamMap = @{}
$teamData      = [System.Collections.Generic.List[hashtable]]::new()

foreach ($team in $teams) {
    Write-Step "Processing '$($team.name)' ..." -Status INFO

    # Area paths
    $fieldValues = Get-TeamFieldValues -TeamId $team.id
    $defaultArea = ""
    $allAreas    = @()
    if ($fieldValues) {
        $defaultArea = $fieldValues.defaultValue
        if ($fieldValues.PSObject.Properties["values"]) {
            $allAreas = @($fieldValues.values | ForEach-Object { $_.value })
        }
    }
    if (-not $allAreas -and $defaultArea) { $allAreas = @($defaultArea) }

    $relArea     = ConvertTo-RelativeAreaPath -FullPath $defaultArea -ProjectName $projectName
    $allRelAreas = @($allAreas | ForEach-Object {
        ConvertTo-RelativeAreaPath -FullPath $_ -ProjectName $projectName
    } | Where-Object { $_ })

    # Map the primary area to the team (for parent resolution)
    if ($relArea) {
        $areaToTeamMap[$relArea] = $team.name
    }
    # Also map additional areas
    foreach ($ar in $allRelAreas) {
        if (-not $areaToTeamMap.ContainsKey($ar)) {
            $areaToTeamMap[$ar] = $team.name
        }
    }

    # Iterations
    $iters = Get-TeamIterations -TeamId $team.id
    $iterNames = @()
    foreach ($it in $iters) {
        # The iteration path comes back as "<Project>\Sprint 1" – strip the project prefix
        $iterRel = ConvertTo-RelativeAreaPath -FullPath $it.path -ProjectName $projectName
        # Also strip the leading backslash if present (API sometimes returns "\<Project>\Sprint 1")
        $iterRel = $iterRel.TrimStart('\')
        if ($iterRel) { $iterNames += $iterRel }
    }

    # Members
    $members    = Get-TeamMembers -TeamId $team.id
    $emailList  = @()
    foreach ($m in $members) {
        $email = ""
        if ($m.identity.PSObject.Properties["uniqueName"]) {
            $email = $m.identity.uniqueName
        }
        if ($email) { $emailList += $email }
    }

    $teamData.Add(@{
        Id          = $team.id
        Name        = $team.name
        Description = if ($team.PSObject.Properties["description"]) { $team.description } else { "" }
        RelArea     = $relArea
        AllRelAreas = $allRelAreas
        Iterations  = $iterNames
        Members     = $emailList
    })
}

# ── 4. Resolve parent teams from area hierarchy ────────────────────────────
Write-Banner "Resolving Team Hierarchy"

$allTeamNames = [System.Collections.Generic.HashSet[string]]::new(
    [StringComparer]::OrdinalIgnoreCase)
foreach ($td in $teamData) { [void]$allTeamNames.Add($td.Name) }

$csvRows = [System.Collections.Generic.List[PSObject]]::new()

foreach ($td in $teamData) {
    $parentTeam = Resolve-ParentTeam `
        -TeamName     $td.Name `
        -RelativeAreaPath $td.RelArea `
        -AreaToTeamMap $areaToTeamMap `
        -AllTeamNames $allTeamNames

    if ($parentTeam) {
        Write-Step "'$($td.Name)' → parent '$parentTeam'" -Status INFO
    }
    else {
        Write-Step "'$($td.Name)' → root-level team" -Status INFO
    }

    $row = [PSCustomObject]@{
        Id         = $td.Id
        TeamName   = $td.Name
        ParentTeam = $parentTeam
        Description = $td.Description
        AreaPaths  = ""
        Members    = ($td.Members -join ';')
        Iterations = ($td.Iterations -join ';')
    }

    $csvRows.Add($row)
}

# ── 4b. Compute AreaPaths column ───────────────────────────────────────────
#  Walk the ParentTeam chain (using team names) to compute the hierarchy-
#  derived default area path.  If the team's actual area paths differ from
#  that default, encode them in the AreaPaths column.

function Get-HierarchyDefaultArea {
    param(
        [string]$TeamName,
        [System.Collections.Generic.IEnumerable[PSObject]]$Rows
    )
    $row = $Rows | Where-Object { $_.TeamName -eq $TeamName }
    if (-not $row) { return $TeamName }
    if ($row.ParentTeam) {
        return "$(Get-HierarchyDefaultArea -TeamName $row.ParentTeam -Rows $Rows)\$TeamName"
    }
    return $TeamName
}

function Convert-AreaToExportFormat {
    param(
        [string]$RelativeArea,
        [string]$ParentHierarchyArea
    )
    if ($ParentHierarchyArea) {
        $prefix = "$ParentHierarchyArea\"
        if ($RelativeArea.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase)) {
            $remainder = $RelativeArea.Substring($prefix.Length)
            if (-not $remainder.Contains('\')) {
                return $remainder   # Simple name under parent
            }
        }
    }
    else {
        # Root-level team: if area is a single segment, it's a simple name
        if (-not $RelativeArea.Contains('\')) {
            return $RelativeArea
        }
    }
    # Full path – convert \ to /
    return ($RelativeArea -replace '\\', '/')
}

foreach ($row in $csvRows) {
    $td = $teamData | Where-Object { $_.Name -eq $row.TeamName }
    $hierarchyDefault = Get-HierarchyDefaultArea -TeamName $row.TeamName -Rows $csvRows

    # If the team has exactly one area that matches the hierarchy default, leave AreaPaths empty
    if ($td.AllRelAreas.Count -le 1 -and $td.AllRelAreas.Count -gt 0 -and $td.AllRelAreas[0] -eq $hierarchyDefault) {
        continue
    }

    $parentHArea = if ($row.ParentTeam) {
        Get-HierarchyDefaultArea -TeamName $row.ParentTeam -Rows $csvRows
    } else { "" }

    $parts = @()
    foreach ($area in $td.AllRelAreas) {
        $parts += Convert-AreaToExportFormat -RelativeArea $area -ParentHierarchyArea $parentHArea
    }
    $row.AreaPaths = $parts -join ';'
}

# ── 5. Topological sort – parents first (same order as provisioning) ───────
$sorted = [System.Collections.Generic.List[PSObject]]::new()
$added  = [System.Collections.Generic.HashSet[string]]::new(
              [StringComparer]::OrdinalIgnoreCase)

function Add-SortedRec ([string]$Name) {
    if ($added.Contains($Name)) { return }
    $row = $csvRows | Where-Object { $_.TeamName -eq $Name }
    if (-not $row) { return }
    if ($row.ParentTeam) { Add-SortedRec $row.ParentTeam }
    [void]$added.Add($Name)
    $sorted.Add($row)
}

foreach ($r in $csvRows) { Add-SortedRec $r.TeamName }

# ── 6. Write CSV ───────────────────────────────────────────────────────────
Write-Banner "Writing CSV"

$sorted | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
Write-Step "Exported $($sorted.Count) team(s) to '$OutputPath'" -Status CREATE

# Preview
Write-Banner "Preview"
$sorted | Format-Table -AutoSize | Out-String | Write-Host

Write-Banner "Done!"
