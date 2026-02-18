# Azure DevOps Team Automater

Automate the provisioning and export of **Teams**, **Area Paths**, **Iterations**, and **Team Members** in an Azure DevOps project using PowerShell and a simple CSV file.

## Prerequisites

- **PowerShell 5.1+** (Windows PowerShell or PowerShell 7+)
- An **Azure DevOps Personal Access Token (PAT)** with _Full access_ scope  
  (or at minimum: Work Items Read/Write, Project & Team Read/Write, Graph Read/Write, Member Entitlement Management Read/Write)
- The PAT can be passed via the `-PAT` parameter or the `AZURE_DEVOPS_PAT` environment variable

## Files

| File                        | Purpose                                                               |
| --------------------------- | --------------------------------------------------------------------- |
| `Provision-DevOpsTeams.ps1` | Creates teams, area paths, iterations, and assigns members from a CSV |
| `Export-DevOpsTeams.ps1`    | Reads the current project state and exports it to a CSV               |
| `teams-template.csv`        | Example CSV template to get started                                   |

---

## CSV Template Structure

The CSV defines the desired state of your Azure DevOps project. Both scripts use the same format:

```
Id,TeamName,ParentTeam,Description,AreaPaths,Members,Iterations
```

| Column        | Required | Description                                                                                    |
| ------------- | -------- | ---------------------------------------------------------------------------------------------- |
| `Id`          | No       | GUID of an existing team. When present the team is **updated** instead of created.             |
| `TeamName`    | **Yes**  | Name of the team                                                                               |
| `ParentTeam`  | **Yes**  | Name of the parent team (leave empty for root-level teams)                                     |
| `Description` | No       | Team description text                                                                          |
| `AreaPaths`   | No       | Semicolon-separated custom area path names (see [Custom Area Paths](#custom-area-paths) below) |
| `Members`     | No       | Semicolon-separated email addresses of team members                                            |
| `Iterations`  | No       | Semicolon-separated iteration paths. Use `\` for nesting (e.g. `Release 1\Sprint 1`)           |

### Example

```csv
Id,TeamName,ParentTeam,Description,AreaPaths,Members,Iterations
,Platform,,Platform Engineering Team,,user1@company.com;user2@company.com,Sprint 1;Sprint 2;Sprint 3
,Backend,Platform,Backend Services Team,,dev1@company.com;dev2@company.com,Sprint 1;Sprint 2
,Frontend,Platform,Frontend UI Team,,dev3@company.com,Sprint 1;Sprint 2
,QA,,Quality Assurance Team,,qa1@company.com;qa2@company.com,Sprint 1;Sprint 2;Sprint 3
```

After the first provisioning run, the script writes team IDs back into the CSV:

```csv
Id,TeamName,ParentTeam,Description,AreaPaths,Members,Iterations
a1b2c3d4-...,Platform,,Platform Engineering Team,,...
```

### How the hierarchy works

The **ParentTeam** column drives the Area Path tree automatically:

```
Project Root
├── Platform                ← root team (no parent)
│   ├── Backend             ← child of Platform → area: Project\Platform\Backend
│   └── Frontend            ← child of Platform → area: Project\Platform\Frontend
└── QA                      ← root team (no parent)  → area: Project\QA
```

### Custom Area Paths

By default the team name is used as the area path node. The `AreaPaths` column overrides this:

| Format         | Behaviour                                   | Example                                   |
| -------------- | ------------------------------------------- | ----------------------------------------- |
| _(empty)_      | Uses the team name as area node (default)   | Backend → `Project\Platform\Backend`      |
| Simple name    | Placed under parent's hierarchy area        | `Services` → `Project\Platform\Services`  |
| Path with `/`  | Absolute from project root (ignores parent) | `Shared/Common` → `Project\Shared\Common` |
| Multiple (`;`) | Each entry is resolved independently        | `Services;Shared/Common`                  |

When a custom area path is set, the default hierarchy-derived area path (team name based) is **not** created. If it existed from a previous run it is automatically removed.

#### Example with custom area paths

```csv
Id,TeamName,ParentTeam,Description,AreaPaths,Members,Iterations
,Platform,,Platform Team,,user1@company.com,Sprint 1
,Backend,Platform,Backend Team,Services;Shared/Infra,dev1@company.com,Sprint 1
,Frontend,Platform,Frontend Team,,dev2@company.com,Sprint 1
```

Resulting area tree:

```
Project Root
├── Platform
│   ├── Services            ← Backend team (custom simple name)
│   └── Frontend            ← Frontend team (default = team name)
└── Shared
    └── Infra               ← Backend team (custom absolute path)
```

---

## Provision-DevOpsTeams.ps1

Reads a CSV and provisions the Azure DevOps project to match the desired state.

### What it does

1. **Creates or updates teams** — new teams (no Id) are created; existing teams (Id present) are updated (rename, description)
2. **Creates area paths** — derived from the team hierarchy by default, or using custom paths from the `AreaPaths` column. Removes stale default area paths when custom paths are set.
3. **Assigns each team** to its corresponding area path(s) — supports multiple area paths per team
4. **Creates iterations** listed in the CSV
5. **Subscribes teams** to their specified iterations
6. **Syncs members** — adds missing members and removes members not listed in the CSV
7. **Creates security groups** per team:
   - `role-external-contributor-<team-name>` — role group (no automatic members)
   - `perm-item-reader-<team-name>` — grants "View work items in this node" on the team's area path; the role group is added as a member
   - `perm-item-writer-<team-name>` — grants "Edit work items in this node" and "Edit work item comments in this node" on the team's area path; the role group is added as a member
8. **Blocks permission inheritance** — for every team that has child teams, explicit **Deny** ACEs are set on all child area paths for the parent's reader and writer groups, preventing inherited permissions from leaking down the hierarchy
9. **Writes back team Ids** into the CSV so subsequent runs can track teams by Id

The script is **idempotent** — running it multiple times won't duplicate resources.

### Permission inheritance blocking

Azure DevOps area path permissions **inherit** from parent to child nodes by default. This means if `perm-item-reader-platform` has "Allow – View work items" on the `Platform` area, that permission automatically flows down to `Platform\Backend` and `Platform\Frontend`.

To ensure external contributors assigned to a parent team can **only** see and edit items in that team's own area path (and not in child team areas), the script sets explicit **Deny** ACEs on every child area path:

```
Platform (Allow: reader-platform, writer-platform)
├── Backend (Deny: reader-platform, writer-platform)  ← blocks inheritance
│                (Allow: reader-backend, writer-backend)
└── Frontend (Deny: reader-platform, writer-platform) ← blocks inheritance
                  (Allow: reader-frontend, writer-frontend)
```

In a deeper hierarchy (e.g. `Alpha → Bravo → Charlie`), the deny entries are set transitively:

- Alpha's groups are denied on both Bravo **and** Charlie
- Bravo's groups are denied on Charlie

### Usage

```powershell
# Provision (create everything defined in the CSV)
.\Provision-DevOpsTeams.ps1 `
    -Organization "myorg" `
    -Project "MyProject" `
    -CsvPath ".\teams-template.csv" `
    -PAT "your-personal-access-token"

# Cleanup (remove everything defined in the CSV)
.\Provision-DevOpsTeams.ps1 `
    -Organization "myorg" `
    -Project "MyProject" `
    -CsvPath ".\teams-template.csv" `
    -PAT "your-personal-access-token" `
    -Cleanup
```

### Parameters

| Parameter       | Required | Description                                                                     |
| --------------- | -------- | ------------------------------------------------------------------------------- |
| `-Organization` | Yes      | Azure DevOps organization name (e.g. `my-org`)                                  |
| `-Project`      | Yes      | Azure DevOps project name (e.g. `MyProject`)                                    |
| `-CsvPath`      | Yes      | Path to the CSV template file                                                   |
| `-PAT`          | No       | Personal Access Token. Falls back to `$env:AZURE_DEVOPS_PAT`                    |
| `-Cleanup`      | No       | Switch to reverse all changes (delete teams, areas, iterations, remove members) |

### Cleanup mode

The `-Cleanup` switch reverses provisioning in the correct order:

1. Removes area path permissions (both Allow and Deny ACEs) and deletes security groups (`role-*`, `perm-*`)
2. Removes members from teams
3. Resets team area settings to the project root
4. Deletes teams (children first)
5. Deletes area paths (deepest first)
6. Deletes iterations

> **Note:** The project's default team cannot be deleted by Azure DevOps. During cleanup it is renamed to "Default Team" with only the project owner as a member.

### Updating existing teams

When a CSV row has an `Id`, the script looks up the team by that GUID:

- **Found** — the team's name, description, area, iterations, and members are updated to match the CSV
- **Not found** — the row is treated as a new team, created fresh, and the new Id is written back

This means you can **rename** a team simply by changing the `TeamName` while keeping the `Id`.
Members are **synced**: users not in the CSV are removed, users missing are added.

---

## Export-DevOpsTeams.ps1

Connects to an Azure DevOps project and exports the current teams, area paths, iterations, and members into a CSV file with the **same format** as the provisioning template.

### What it does

1. Fetches all teams from the project
2. Reads each team's configured area path(s) to derive the parent-child hierarchy
3. Detects custom area paths and populates the `AreaPaths` column (left empty when the default hierarchy path is used)
4. Collects each team's subscribed iterations
5. Lists each team's members (email addresses)
6. Outputs a sorted CSV (parents before children)

### Usage

```powershell
# Export to default file (<Project>-export.csv)
.\Export-DevOpsTeams.ps1 `
    -Organization "myorg" `
    -Project "MyProject" `
    -PAT "your-personal-access-token"

# Export to a specific file
.\Export-DevOpsTeams.ps1 `
    -Organization "myorg" `
    -Project "MyProject" `
    -OutputPath ".\current-state.csv" `
    -PAT "your-personal-access-token"

# Include the project's default team in the export
.\Export-DevOpsTeams.ps1 `
    -Organization "myorg" `
    -Project "MyProject" `
    -PAT "your-personal-access-token" `
    -IncludeDefaultTeam
```

### Parameters

| Parameter             | Required | Description                                                            |
| --------------------- | -------- | ---------------------------------------------------------------------- |
| `-Organization`       | Yes      | Azure DevOps organization name                                         |
| `-Project`            | Yes      | Azure DevOps project name                                              |
| `-OutputPath`         | No       | Output CSV path. Defaults to `.\<Project>-export.csv`                  |
| `-PAT`                | No       | Personal Access Token. Falls back to `$env:AZURE_DEVOPS_PAT`           |
| `-IncludeDefaultTeam` | No       | Include the project's default team in the export (excluded by default) |

---

## Typical Workflow

### 1. Bootstrap from an existing project

```powershell
# Export current state
.\Export-DevOpsTeams.ps1 -Organization "myorg" -Project "MyProject" -PAT $pat

# Edit the exported CSV to add/remove teams, members, iterations
# Then apply changes
.\Provision-DevOpsTeams.ps1 -Organization "myorg" -Project "MyProject" `
    -CsvPath ".\MyProject-export.csv" -PAT $pat
```

### 2. Start from scratch with a template

```powershell
# Edit teams-template.csv with your desired structure
# Then provision
.\Provision-DevOpsTeams.ps1 -Organization "myorg" -Project "MyProject" `
    -CsvPath ".\teams-template.csv" -PAT $pat
```

### 3. Replicate structure across projects

```powershell
# Export from source project
.\Export-DevOpsTeams.ps1 -Organization "myorg" -Project "SourceProject" -PAT $pat

# Provision into target project
.\Provision-DevOpsTeams.ps1 -Organization "myorg" -Project "TargetProject" `
    -CsvPath ".\SourceProject-export.csv" -PAT $pat
```
