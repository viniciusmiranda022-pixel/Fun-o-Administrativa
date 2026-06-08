[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$TenantId,

    [Parameter(Mandatory = $true)]
    [string]$ClientId,

    [Parameter(Mandatory = $false)]
    [string]$CertificateThumbprint = "",

    [Parameter(Mandatory = $false)]
    [string]$ClientSecret = "",

    [Parameter(Mandatory = $true)]
    [string]$SnapshotFolder,

    [Parameter(Mandatory = $true)]
    [string]$RoleName,

    [Parameter(Mandatory = $false)]
    [string]$CurrentRoleName = "",

    [Parameter()]
    [ValidateSet("WhatIf", "Apply")]
    [string]$Mode = "WhatIf",

    [Parameter()]
    [switch]$RemoveExtraAssignments,

    [Parameter()]
    [switch]$SkipConfirmation
)
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

try {
    chcp 65001 | Out-Null
} catch {}

$ErrorActionPreference = "Stop"

try {
    Import-Module Microsoft.PowerShell.Utility

    # Expand PSModulePath to include user-profile paths that may be missing
    # when running under a Windows Service account
    $extraPaths = @(
        "$env:USERPROFILE\Documents\WindowsPowerShell\Modules",
        "$env:USERPROFILE\OneDrive - Corporativo\Documentos\WindowsPowerShell\Modules",
        "$env:USERPROFILE\OneDrive\Documentos\WindowsPowerShell\Modules",
        "$env:ProgramFiles\WindowsPowerShell\Modules",
        "$env:ProgramFiles\PowerShell\Modules"
    )
    foreach ($p in $extraPaths) {
        if ((Test-Path $p) -and ($env:PSModulePath -notlike "*$p*")) {
            $env:PSModulePath = "$p;$env:PSModulePath"
        }
    }

    Import-Module Microsoft.Graph.Authentication -Force

    # Try the sub-module directly first; if not found, import the full
    # Microsoft.Graph meta-module which contains all sub-modules
    if (Get-Module -Name 'Microsoft.Graph.Identity.Governance' -ListAvailable) {
        Import-Module Microsoft.Graph.Identity.Governance -Force
    } else {
        Import-Module Microsoft.Graph -Force
    }

    if (-not (Get-Command -Name 'Get-MgRoleManagementDirectoryRoleDefinition' -ErrorAction SilentlyContinue)) {
        throw "Cmdlet 'Get-MgRoleManagementDirectoryRoleDefinition' not found after importing Microsoft.Graph. Run: Install-Module Microsoft.Graph -Scope AllUsers"
    }

    $sharedModulePath = Join-Path $PSScriptRoot "..\..\IR-AdministrativeFunctionBackup\Scripts\IR-AdministrativeFunctions.psm1"
    Import-Module $sharedModulePath -Force
} catch {
    Write-Error "Failed to import required modules: $($_.Exception.Message)"
    exit 1
}

$RestoreInstallRoot = "C:\ProgramData\Quest\IR-AdministrativeFunctionRestore"

function Initialize-StructuredLog {
    param(
        [string]$TenantId,
        [string]$RoleName,
        [string]$SnapshotFolder
    )

    $logDirectory = Join-Path $RestoreInstallRoot "Logs"
    New-Item -Path $logDirectory -ItemType Directory -Force | Out-Null

    $logFileName = "restore-{0}.log" -f (Get-Date -Format "yyyy-MM-dd_HH-mm-ss")
    $script:StructuredLogPath = Join-Path $logDirectory $logFileName
    $script:RestoreOperator = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
    $script:RestoreTenant = $TenantId
    $script:RestoreRoleTarget = $RoleName
    $script:RestoreSnapshot = (Resolve-Path $SnapshotFolder).Path
}

function Write-StructuredLog {
    param(
        [string]$Message,
        [string]$Result = "INFO"
    )

    $entry = [ordered]@{
        timestamp    = (Get-Date).ToString("o")
        roleTarget   = $script:RestoreRoleTarget
        snapshotUsed = $script:RestoreSnapshot
        operator     = $script:RestoreOperator
        tenant       = $script:RestoreTenant
        result       = $Result
        message      = $Message
    }

    $entry | ConvertTo-Json -Compress | Add-Content -Path $script:StructuredLogPath -Encoding UTF8
}

function Write-Log {
    param([string]$Message)
    Write-Host ("[{0}] {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Message)
    Write-StructuredLog -Message $Message
}


function Test-SnapshotIntegrity {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SnapshotFolder
    )

    $manifestPath = Join-Path $SnapshotFolder "manifest.json"
    $roleDefinitionsPath = Join-Path $SnapshotFolder "roleDefinitions.json"
    $roleAssignmentsPath = Join-Path $SnapshotFolder "roleAssignments.json"

    if (-not (Test-Path $manifestPath)) {
        throw "manifest.json was not found in $SnapshotFolder"
    }

    $manifest = Get-Content $manifestPath -Raw | ConvertFrom-Json

    if (-not $manifest.RoleDefinitionsHash) {
        throw "RoleDefinitionsHash was not found in manifest.json"
    }

    $sha256 = [System.Security.Cryptography.SHA256]::Create()
    try {
        $roleDefinitionsHashCurrent = [System.BitConverter]::ToString(
            $sha256.ComputeHash([System.IO.File]::ReadAllBytes($roleDefinitionsPath))
        ).Replace('-', '')
        if ($roleDefinitionsHashCurrent -ne [string]$manifest.RoleDefinitionsHash) {
            throw "Integrity failed: roleDefinitions.json was changed."
        }

        if ($manifest.RoleAssignmentsHash) {
            $roleAssignmentsHashCurrent = [System.BitConverter]::ToString(
                $sha256.ComputeHash([System.IO.File]::ReadAllBytes($roleAssignmentsPath))
            ).Replace('-', '')
            if ($roleAssignmentsHashCurrent -ne [string]$manifest.RoleAssignmentsHash) {
                throw "Integrity failed: roleAssignments.json was changed."
            }
        }
    } finally {
        $sha256.Dispose()
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

function Get-PermissionList {
    param($RoleDefinition)

    $actions = @()
    if ($RoleDefinition -and $RoleDefinition.rolePermissions) {
        foreach ($rp in $RoleDefinition.rolePermissions) {
            if ($rp.allowedResourceActions) {
                $actions += @($rp.allowedResourceActions)
            }
        }
    }
    return @($actions | Sort-Object -Unique)
}

function Export-PreRestoreReport {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SnapshotFolder,
        [Parameter(Mandatory = $true)]
        [string]$RoleName,
        [Parameter(Mandatory = $true)]
        [string]$Mode,
        [Parameter(Mandatory = $true)]
        [bool]$RemoveExtraAssignments,
        $DesiredRole,
        $CurrentRole,
        [array]$DesiredAssignments,
        [array]$CurrentAssignments
    )

    $desiredPermissions = Get-PermissionList -RoleDefinition $DesiredRole
    $currentPermissions = Get-PermissionList -RoleDefinition $CurrentRole

    $permissionsToAdd = @($desiredPermissions | Where-Object { $_ -notin $currentPermissions })
    $permissionsToRemove = @($currentPermissions | Where-Object { $_ -notin $desiredPermissions })

    $desiredMap = @{}
    foreach ($a in $DesiredAssignments) {
        $key = Get-AssignmentKey -Assignment $a
        $desiredMap[$key] = $a
    }

    $currentMap = @{}
    foreach ($a in $CurrentAssignments) {
        $key = Get-AssignmentKey -Assignment $a
        $currentMap[$key] = $a
    }

    $assignmentsToCreate = @()
    foreach ($key in $desiredMap.Keys) {
        if (-not $currentMap.ContainsKey($key)) {
            $a = $desiredMap[$key]
            $assignmentsToCreate += [ordered]@{
                key              = $key
                principalId      = [string]$a.principalId
                directoryScopeId = [string]$a.directoryScopeId
                appScopeId       = [string]$a.appScopeId
                condition        = [string]$a.condition
                conditionVersion = [string]$a.conditionVersion
            }
        }
    }

    $assignmentsToRemove = @()
    foreach ($key in $currentMap.Keys) {
        if (-not $desiredMap.ContainsKey($key)) {
            $a = $currentMap[$key]
            $assignmentsToRemove += [ordered]@{
                key              = $key
                id               = [string]$a.Id
                principalId      = [string]$a.principalId
                directoryScopeId = [string]$a.directoryScopeId
                appScopeId       = [string]$a.appScopeId
                condition        = [string]$a.condition
                conditionVersion = [string]$a.conditionVersion
                willBeRemoved    = ($Mode -eq "Apply" -and $RemoveExtraAssignments)
            }
        }
    }

    $report = [ordered]@{
        generatedAt             = (Get-Date).ToString("o")
        roleTarget              = $RoleName
        mode                    = $Mode
        removeExtraAssignments  = $RemoveExtraAssignments
        roleExistsInTenant      = [bool]($null -ne $CurrentRole)
        snapshotFolder          = (Resolve-Path $SnapshotFolder).Path
        permissions             = [ordered]@{
            toAdd    = $permissionsToAdd
            toRemove = $permissionsToRemove
        }
        assignments             = [ordered]@{
            toCreate = $assignmentsToCreate
            toRemove = $assignmentsToRemove
        }
    }

    $reportDirectory = Join-Path $SnapshotFolder "restore-preview"
    New-Item -Path $reportDirectory -ItemType Directory -Force | Out-Null
    $reportPath = Join-Path $reportDirectory ("pre-restore-report-{0}.json" -f (Get-Date -Format "yyyyMMdd-HHmmss"))
    $report | ConvertTo-Json -Depth 20 | Out-File -FilePath $reportPath -Encoding utf8
    return $reportPath
}

try {
    Initialize-StructuredLog -TenantId $TenantId -RoleName $RoleName -SnapshotFolder $SnapshotFolder
    Write-Log "Connecting to Microsoft Graph (with retry for App Registration replication)"
    Connect-AppOnlyGraphWithRetry -TenantId $TenantId -ClientId $ClientId -Thumbprint $CertificateThumbprint -ClientSecret $ClientSecret -VerboseOutput -FailureMessage 'Failed to connect to Graph in app-only mode after {0} attempts. Last error: {1}'

    $roleDefinitionsPath = Join-Path $SnapshotFolder "roleDefinitions.json"
    $roleAssignmentsPath = Join-Path $SnapshotFolder "roleAssignments.json"

    if (-not (Test-Path $roleDefinitionsPath)) {
        throw "roleDefinitions.json was not found in $SnapshotFolder"
    }

    if (-not (Test-Path $roleAssignmentsPath)) {
        throw "roleAssignments.json was not found in $SnapshotFolder"
    }

    Write-Log "Validating snapshot integrity"
    Test-SnapshotIntegrity -SnapshotFolder $SnapshotFolder

    Write-Log "Loading snapshot files"
    $definitions = Get-Content $roleDefinitionsPath -Raw | ConvertFrom-Json
    $assignments = Get-Content $roleAssignmentsPath -Raw | ConvertFrom-Json

    $desiredRole = $definitions | Where-Object { $_.displayName -eq $RoleName } | Select-Object -First 1

    if (-not $desiredRole) {
        throw "The role '$RoleName' was not found in roleDefinitions.json"
    }

    Write-Log "Role found in snapshot: $($desiredRole.displayName)"
    Write-Log "Execution mode: $Mode"

    if ($RemoveExtraAssignments -and $Mode -eq "WhatIf") {
        Write-Log "-RemoveExtraAssignments was specified with -Mode WhatIf. No assignment removal will be executed in preview mode."
    }

    # Client-side enumeration: the server-side filter ("displayName eq ...") suffers from
    # replication latency and may return empty even when the role already exists,
    # leading the script down the CREATE path and to a 400 "conflicting object" error.
    $allRoleDefinitions = @(Get-MgRoleManagementDirectoryRoleDefinition -All)

    # When -CurrentRoleName is provided, the target role in the tenant has a different name
    # than in the backup (it was renamed). Locate it by that name; the UPDATE renames it back
    # to the displayName from the backup.
    if (-not [string]::IsNullOrWhiteSpace($CurrentRoleName)) {
        $lookupName = $CurrentRoleName
        $currentRole = $allRoleDefinitions | Where-Object { $_.DisplayName -eq $CurrentRoleName } | Select-Object -First 1
        if (-not $currentRole) {
            throw "The target role '$CurrentRoleName' specified for overwrite was not found in the tenant. Run a new compare and try again."
        }
    }
    else {
        $lookupName = $RoleName
        $currentRole = $allRoleDefinitions | Where-Object { $_.DisplayName -eq $RoleName } | Select-Object -First 1
    }
    Write-Output "DIAG: roleDefinitions in tenant=$($allRoleDefinitions.Count) lookupName='$lookupName' currentRole found=$([bool]$currentRole)"

    if ($currentRole -and $currentRole.IsBuiltIn) {
        throw "The existing role '$RoleName' in the tenant is built-in. This restore only supports custom roles."
    }

    if (-not $currentRole) {
        Write-Output "DIAG: path=CREATE"
        Write-Log "The role does not currently exist. The snapshot requires creating the role."
        Write-Output "DIAG: snapshot displayName='$($desiredRole.displayName)' isEnabled=$($desiredRole.isEnabled)"

        if (-not [bool]$desiredRole.isEnabled) {
            Write-Output "DIAG: WARNING isEnabled=false in snapshot. Role will be created DISABLED."
            Write-Log "WARNING: isEnabled=false in snapshot. The role will be created as DISABLED and will not appear as active in the Azure portal."
        }

        if ($Mode -eq "Apply") {
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

            # Pass the original TemplateId so that the Azure portal displays the role correctly.
            # If there is a 400 conflict (role in soft-delete still occupying the templateId), fall back without TemplateId.
            $templateIdFromSnapshot = if ($desiredRole.templateId) { [string]$desiredRole.templateId } else { $null }
            if ($templateIdFromSnapshot) {
                $newRoleParams.TemplateId = $templateIdFromSnapshot
            }

            Write-Output "DIAG: chamando New-MgRoleManagementDirectoryRoleDefinition IsEnabled=$($newRoleParams.IsEnabled) TemplateId=$templateIdFromSnapshot"

            # -ErrorAction Stop is mandatory: Graph SDK cmdlets emit HTTP API errors
            # as non-terminating, and the global $ErrorActionPreference does not promote them
            # reliably. Without this, try/catch does not catch the 400 and the fallback does not run.
            try {
                $currentRole = New-MgRoleManagementDirectoryRoleDefinition @newRoleParams -ErrorAction Stop
            }
            catch {
                $createError = $_.Exception.Message
                if ($templateIdFromSnapshot -and ($createError -match '400|BadRequest|[Cc]onflict|soft.deleted|already.*exist')) {
                    Write-Log "Creation with TemplateId=$templateIdFromSnapshot failed (templateId already occupied in directory). Error: $createError. Retrying without TemplateId..."
                    Write-Output "DIAG: retry CREATE without TemplateId"
                    $newRoleParams.Remove('TemplateId')
                    $currentRole = New-MgRoleManagementDirectoryRoleDefinition @newRoleParams -ErrorAction Stop
                }
                else {
                    throw
                }
            }

            if (-not $currentRole -or -not $currentRole.Id) {
                throw "New-MgRoleManagementDirectoryRoleDefinition returned empty without throwing an exception. Verify that the App Registration has the 'RoleManagement.ReadWrite.Directory' permission and that the tenant has an Entra ID P1/P2 license."
            }
            Write-Output "DIAG: role created Id=$($currentRole.Id) DisplayName='$($currentRole.DisplayName)' IsEnabled=$($currentRole.IsEnabled) TemplateId=$($currentRole.TemplateId)"
            Write-Log "Role created: Id=$($currentRole.Id), DisplayName='$($currentRole.DisplayName)', IsEnabled=$($currentRole.IsEnabled), TemplateId=$($currentRole.TemplateId)"
        }
        else {
            Write-Log "[PREVIEW] The role would be created based on the snapshot definition."
        }
    }
    else {
        Write-Output "DIAG: path=UPDATE existingId=$($currentRole.Id)"
        if ($currentRole.DisplayName -ne $desiredRole.displayName) {
            Write-Log "The target role '$($currentRole.DisplayName)' will be overwritten and renamed to '$($desiredRole.displayName)' as per the backup."
        }
        else {
            Write-Log "The role already exists. The snapshot requires a full overwrite of the definition."
        }

        if ($Mode -eq "Apply") {
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
            if (-not $currentRole -or -not $currentRole.Id) {
                throw "Update-MgRoleManagementDirectoryRoleDefinition failed silently. Verify the 'RoleManagement.ReadWrite.Directory' permission on the App Registration."
            }
            Write-Log "Role definition updated"
        }
        else {
            Write-Log "[PREVIEW] The role definition would be updated to the snapshot values."
        }
    }

    Write-Log "Processing snapshot assignments"

    $desiredAssignments = @($assignments | Where-Object { $_.roleDefinitionId -eq $desiredRole.id })

    if ($currentRole) {
        $currentAssignments = @(Get-MgRoleManagementDirectoryRoleAssignment -All | Where-Object { $_.roleDefinitionId -eq $currentRole.Id })
    }
    else {
        $currentAssignments = @()
    }

    $preRestoreReportPath = Export-PreRestoreReport `
        -SnapshotFolder $SnapshotFolder `
        -RoleName $RoleName `
        -Mode $Mode `
        -RemoveExtraAssignments ([bool]$RemoveExtraAssignments) `
        -DesiredRole $desiredRole `
        -CurrentRole $currentRole `
        -DesiredAssignments $desiredAssignments `
        -CurrentAssignments $currentAssignments

    Write-Log "Pre-restore report exported: $preRestoreReportPath"

    if ($Mode -eq "Apply") {
        Write-Host ""
        Write-Host "Pre-restore report generated at: $preRestoreReportPath" -ForegroundColor Yellow

        if (-not $SkipConfirmation) {
            $confirmation = Read-Host "Type CONFIRM to proceed with the restore changes"
            if ($confirmation -ne "CONFIRM") {
                throw "Restore cancelled by operator. Review the report and run again when ready."
            }
            Write-Log "Operator confirmed restore execution after reviewing the report."
        }
        else {
            Write-Log "Interactive confirmation skipped by -SkipConfirmation parameter."
        }
    }

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

            if ($Mode -eq "Apply") {
                New-MgRoleManagementDirectoryRoleAssignment -BodyParameter $body | Out-Null
                Write-Log "Assignment recreated: $key"
            }
            else {
                Write-Log "[PREVIEW] The assignment would be created: $key"
            }
        }
    }

    if ($currentRole) {
        $currentAssignments = @(Get-MgRoleManagementDirectoryRoleAssignment -All | Where-Object { $_.roleDefinitionId -eq $currentRole.Id })
    }

    foreach ($a in $currentAssignments) {
        $key = Get-AssignmentKey -Assignment $a
        if (-not $desiredMap.ContainsKey($key)) {
            if ($Mode -eq "Apply" -and $RemoveExtraAssignments) {
                Remove-MgRoleManagementDirectoryRoleAssignment -UnifiedRoleAssignmentId $a.Id -Confirm:$false
                Write-Log "Extra assignment removed: $key"
            }
            else {
                Write-Log "[PREVIEW] Extra assignment detected (not removed): $key"
            }
        }
    }

    Write-Log "Final validation"

    if ($currentRole) {
        $finalRole = Get-MgRoleManagementDirectoryRoleDefinition -UnifiedRoleDefinitionId $currentRole.Id
        $finalAssignments = @(Get-MgRoleManagementDirectoryRoleAssignment -All | Where-Object { $_.roleDefinitionId -eq $currentRole.Id })

        if ($Mode -eq "Apply") {
            Write-Log "Role restored: $($finalRole.DisplayName)"
        }
        else {
            Write-Log "Preview completed for role: $($finalRole.DisplayName)"
        }

        Write-Output "DIAG: final validation Id=$($finalRole.Id) DisplayName='$($finalRole.DisplayName)' IsEnabled=$($finalRole.IsEnabled) assignments=$($finalAssignments.Count)"
        Write-Log "Current description: $($finalRole.Description)"
        Write-Log "Current IsEnabled: $($finalRole.IsEnabled)"
        Write-Log "Current assignment count: $($finalAssignments.Count)"

        $finalRole | Select-Object Id, DisplayName, Description, IsEnabled
        $finalRole.RolePermissions | ConvertTo-Json -Depth 20
        $finalAssignments | Select-Object Id, PrincipalId, RoleDefinitionId, DirectoryScopeId, AppScopeId, Condition, ConditionVersion
    }
    else {
        Write-Log "Preview completed. The role does not currently exist and would be created in Apply mode."
    }

    Write-StructuredLog -Message "Restore flow completed successfully." -Result "SUCCESS"
    exit 0
}
catch {
    $errorMsg = $_.Exception.Message
    if ($script:StructuredLogPath) {
        Write-StructuredLog -Message ("Restore flow failed: {0}" -f $errorMsg) -Result "FAILED"
    }
    Write-Error "Restore failed: $errorMsg"
    exit 1
}
finally {
    try { Disconnect-MgGraph | Out-Null } catch {}
}
