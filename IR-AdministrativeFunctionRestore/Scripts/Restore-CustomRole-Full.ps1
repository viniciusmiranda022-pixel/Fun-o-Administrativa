param(
    [Parameter(Mandatory = $true)]
    [string]$TenantId,

    [Parameter(Mandatory = $true)]
    [string]$ClientId,

    [Parameter(Mandatory = $true)]
    [string]$CertificateThumbprint,

    [Parameter(Mandatory = $true)]
    [string]$SnapshotFolder,

    [Parameter(Mandatory = $true)]
    [string]$RoleName
)

$ErrorActionPreference = "Stop"

Import-Module Microsoft.Graph.Authentication
Import-Module Microsoft.Graph.Identity.Governance

function Write-Log {
    param([string]$Message)
    Write-Host ("[{0}] {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Message)
}


function Connect-AppOnlyGraphWithRetry {
    param(
        [string]$TenantId,
        [string]$ClientId,
        [string]$CertificateThumbprint,
        [int]$MaxAttempts = 8,
        [int]$DelaySeconds = 15
    )

    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null

        try {
            Write-Log "App Registration replication in progress - attempt $attempt of $MaxAttempts."
            Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -NoWelcome -ContextScope Process | Out-Null
            Get-MgRoleManagementDirectoryRoleDefinition -Top 1 | Out-Null
            Write-Log "App-only connection established on attempt $attempt."
            return
        }
        catch {
            $errorMessage = $_.Exception.Message
            Write-Log "Attempt $attempt failed: $errorMessage"

            if ($attempt -eq $MaxAttempts) {
                throw "Failed to connect to Graph with app-only after 2 minutes (8 attempts every 15s). Last error: $errorMessage"
            }

            Write-Log "Waiting $DelaySeconds seconds before next attempt..."
            Start-Sleep -Seconds $DelaySeconds
        }
    }
}

function Get-AssignmentKey {
    param($Assignment)

    $principalId = [string]$Assignment.principalId
    $directoryScopeId = [string]$Assignment.directoryScopeId
    $appScopeId = [string]$Assignment.appScopeId
    $condition = [string]$Assignment.condition
    $conditionVersion = [string]$Assignment.conditionVersion

    return "$principalId|$directoryScopeId|$appScopeId|$condition|$conditionVersion"
}

try {
    Write-Log "Connecting to Microsoft Graph (with retry for App Registration replication)"
    Connect-AppOnlyGraphWithRetry -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint

    $roleDefinitionsPath = Join-Path $SnapshotFolder "roleDefinitions.json"
    $roleAssignmentsPath = Join-Path $SnapshotFolder "roleAssignments.json"

    if (-not (Test-Path $roleDefinitionsPath)) {
        throw "roleDefinitions.json was not found in $SnapshotFolder"
    }

    if (-not (Test-Path $roleAssignmentsPath)) {
        throw "roleAssignments.json was not found in $SnapshotFolder"
    }

    Write-Log "Loading snapshot files"
    $definitions = Get-Content $roleDefinitionsPath -Raw | ConvertFrom-Json
    $assignments = Get-Content $roleAssignmentsPath -Raw | ConvertFrom-Json

    $desiredRole = $definitions | Where-Object { $_.displayName -eq $RoleName } | Select-Object -First 1

    if (-not $desiredRole) {
        throw "Role '$RoleName' was not found in roleDefinitions.json"
    }

    Write-Log "Role found in snapshot: $($desiredRole.displayName)"

    $currentRole = Get-MgRoleManagementDirectoryRoleDefinition -Filter "displayName eq '$RoleName'" | Select-Object -First 1

    if ($currentRole -and $currentRole.IsBuiltIn) {
        throw "The existing role '$RoleName' in the tenant is built-in. This restore only supports custom roles."
    }

    if (-not $currentRole) {
        Write-Log "Role does not exist currently. Recreating from snapshot"

        $newRoleParams = @{
            DisplayName     = $desiredRole.displayName
            Description     = $desiredRole.description
            IsEnabled       = [bool]$desiredRole.isEnabled
            RolePermissions = @()
        }

        foreach ($rp in $desiredRole.rolePermissions) {
            $newRoleParams.RolePermissions += @{
                allowedResourceActions = @($rp.allowedResourceActions)
            }
        }

        if ($desiredRole.templateId) {
            $newRoleParams.TemplateId = [string]$desiredRole.templateId
        }

        $currentRole = New-MgRoleManagementDirectoryRoleDefinition @newRoleParams
        Write-Log "Role recreated successfully. Current Id: $($currentRole.Id)"
    }
    else {
        Write-Log "Role exists currently. Overwriting full definition"

        $updateParams = @{
            displayName     = $desiredRole.displayName
            description     = $desiredRole.description
            isEnabled       = [bool]$desiredRole.isEnabled
            rolePermissions = @()
        }

        foreach ($rp in $desiredRole.rolePermissions) {
            $updateParams.rolePermissions += @{
                allowedResourceActions = @($rp.allowedResourceActions)
            }
        }

        Update-MgRoleManagementDirectoryRoleDefinition `
            -UnifiedRoleDefinitionId $currentRole.Id `
            -BodyParameter $updateParams

        $currentRole = Get-MgRoleManagementDirectoryRoleDefinition -UnifiedRoleDefinitionId $currentRole.Id
        Write-Log "Role definition updated"
    }

    Write-Log "Processing assignments from snapshot"

    $desiredAssignments = @($assignments | Where-Object { $_.roleDefinitionId -eq $desiredRole.id })
    $currentAssignments = @(Get-MgRoleManagementDirectoryRoleAssignment -All | Where-Object { $_.roleDefinitionId -eq $currentRole.Id })

    $desiredMap = @{}
    foreach ($a in $desiredAssignments) {
        $key = Get-AssignmentKey -Assignment $a
        $desiredMap[$key] = $a
    }

    $currentMap = @{}
    foreach ($a in $currentAssignments) {
        $key = Get-AssignmentKey -Assignment $a
        $currentMap[$key] = $a
    }

    foreach ($key in $desiredMap.Keys) {
        if (-not $currentMap.ContainsKey($key)) {
            $a = $desiredMap[$key]

            $body = @{
                "@odata.type"    = "#microsoft.graph.unifiedRoleAssignment"
                principalId      = [string]$a.principalId
                roleDefinitionId = [string]$currentRole.Id
            }

            if ($a.directoryScopeId) {
                $body.directoryScopeId = [string]$a.directoryScopeId
            }

            if ($a.appScopeId) {
                $body.appScopeId = [string]$a.appScopeId
            }

            if ($a.condition) {
                $body.condition = [string]$a.condition
            }

            if ($a.conditionVersion) {
                $body.conditionVersion = [string]$a.conditionVersion
            }

            New-MgRoleManagementDirectoryRoleAssignment -BodyParameter $body | Out-Null
            Write-Log "Assignment recreated: $key"
        }
    }

    $currentAssignments = @(Get-MgRoleManagementDirectoryRoleAssignment -All | Where-Object { $_.roleDefinitionId -eq $currentRole.Id })

    foreach ($a in $currentAssignments) {
        $key = Get-AssignmentKey -Assignment $a
        if (-not $desiredMap.ContainsKey($key)) {
            Remove-MgRoleManagementDirectoryRoleAssignment -UnifiedRoleAssignmentId $a.Id -Confirm:$false
            Write-Log "Extra assignment removed: $key"
        }
    }

    Write-Log "Final validation"
    $finalRole = Get-MgRoleManagementDirectoryRoleDefinition -UnifiedRoleDefinitionId $currentRole.Id
    $finalAssignments = @(Get-MgRoleManagementDirectoryRoleAssignment -All | Where-Object { $_.roleDefinitionId -eq $currentRole.Id })

    Write-Log "Role restored: $($finalRole.DisplayName)"
    Write-Log "Current description: $($finalRole.Description)"
    Write-Log "Current IsEnabled: $($finalRole.IsEnabled)"
    Write-Log "Final assignment count: $($finalAssignments.Count)"

    $finalRole | Select-Object Id, DisplayName, Description, IsEnabled
    $finalRole.RolePermissions | ConvertTo-Json -Depth 20
    $finalAssignments | Select-Object Id, PrincipalId, RoleDefinitionId, DirectoryScopeId, AppScopeId, Condition, ConditionVersion
}
finally {
    Disconnect-MgGraph | Out-Null
}