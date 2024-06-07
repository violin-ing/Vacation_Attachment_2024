# Install the required module to read Excel files
Install-Module -Name ImportExcel -Force -Scope CurrentUser

# Paths to the Excel files
$pastGroupsPath = "C:\path\to\pastGroups.xlsx"
$groupRequirementsPath = "C:\path\to\groupRequirements.xlsx"
$completedModulesPath = "C:\path\to\completedModules.xlsx"

# Import the Excel data
$pastGroups = Import-Excel -Path $pastGroupsPath
$groupRequirements = Import-Excel -Path $groupRequirementsPath
$completedModules = Import-Excel -Path $completedModulesPath

# Convert data to more usable formats
$pastGroupsDict = @{}
foreach ($row in $pastGroups) {
    if (-not $pastGroupsDict.ContainsKey($row.name)) {
        $pastGroupsDict[$row.name] = @()
    }
    $pastGroupsDict[$row.name] += $row.pastGroup
}

$groupRequirementsDict = @{}
foreach ($row in $groupRequirements) {
    $groupRequirementsDict[$row.groupName] = @{
        modulesRequested = $row.modulesRequested.Split(',')
        quota = $row.quota
        members = @()
    }
}

$completedModulesDict = @{}
foreach ($row in $completedModules) {
    if (-not $completedModulesDict.ContainsKey($row.name)) {
        $completedModulesDict[$row.name] = @()
    }
    $completedModulesDict[$row.name] += $row.moduleName
}

# Function to check if a person qualifies for a group
function QualifiesForGroup {
    param ($person, $group)
    $modulesRequested = $groupRequirementsDict[$group].modulesRequested
    $completedModules = $completedModulesDict[$person]
    foreach ($module in $modulesRequested) {
        if ($completedModules -notcontains $module) {
            return $false
        }
    }
    return $true
}

# Sort people into new groups
$assigned = @{}
$remaining = @{}

foreach ($person in $completedModulesDict.Keys) {
    $assignedToGroup = $false
    foreach ($group in $groupRequirementsDict.Keys) {
        if (($groupRequirementsDict[$group].members.Count -lt $groupRequirementsDict[$group].quota) -and
            ($person -notin $pastGroupsDict[$person]) -and
            (QualifiesForGroup $person $group)) {
            $groupRequirementsDict[$group].members += $person
            $assigned[$person] = $group
            $assignedToGroup = $true
            break
        }
    }
    if (-not $assignedToGroup) {
        $remaining[$person] = $true
    }
}

# Randomly assign remaining people
foreach ($person in $remaining.Keys) {
    foreach ($group in $groupRequirementsDict.Keys) {
        if ($groupRequirementsDict[$group].members.Count -lt $groupRequirementsDict[$group].quota) {
            $groupRequirementsDict[$group].members += $person
            $assigned[$person] = $group
            break
        }
    }
}

# Output results
$results = @()
foreach ($group in $groupRequirementsDict.Keys) {
    foreach ($member in $groupRequirementsDict[$group].members) {
        $results += [pscustomobject]@{
            Name = $member
            Group = $group
        }
    }
}

# Export results to Excel
$results | Export-Excel -Path "C:\path\to\sortedGroups.xlsx" -WorksheetName "SortedGroups" -Force

Write-Output "People have been successfully sorted into new groups."
