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
    Write-Log "Conectando no Microsoft Graph"
    Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -NoWelcome

    $roleDefinitionsPath = Join-Path $SnapshotFolder "roleDefinitions.json"
    $roleAssignmentsPath = Join-Path $SnapshotFolder "roleAssignments.json"

    if (-not (Test-Path $roleDefinitionsPath)) {
        throw "Arquivo roleDefinitions.json não encontrado em $SnapshotFolder"
    }

    if (-not (Test-Path $roleAssignmentsPath)) {
        throw "Arquivo roleAssignments.json não encontrado em $SnapshotFolder"
    }

    Write-Log "Carregando arquivos do snapshot"
    $definitions = Get-Content $roleDefinitionsPath -Raw | ConvertFrom-Json
    $assignments = Get-Content $roleAssignmentsPath -Raw | ConvertFrom-Json

    $desiredRole = $definitions | Where-Object { $_.displayName -eq $RoleName } | Select-Object -First 1

    if (-not $desiredRole) {
        throw "Função '$RoleName' não encontrada no roleDefinitions.json"
    }

    Write-Log "Função encontrada no snapshot: $($desiredRole.displayName)"

    $currentRole = Get-MgRoleManagementDirectoryRoleDefinition -Filter "displayName eq '$RoleName'" | Select-Object -First 1

    if ($currentRole -and $currentRole.IsBuiltIn) {
        throw "A função '$RoleName' existente no tenant é built-in. Esse restore só funciona para custom role."
    }

    if (-not $currentRole) {
        Write-Log "Função não existe hoje. Criando novamente a partir do snapshot"

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
        Write-Log "Função recriada com sucesso. Id atual: $($currentRole.Id)"
    }
    else {
        Write-Log "Função existe hoje. Sobrescrevendo definição completa"

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
        Write-Log "Definição da função atualizada"
    }

    Write-Log "Processando atribuições do snapshot"

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
            Write-Log "Atribuição recriada: $key"
        }
    }

    $currentAssignments = @(Get-MgRoleManagementDirectoryRoleAssignment -All | Where-Object { $_.roleDefinitionId -eq $currentRole.Id })

    foreach ($a in $currentAssignments) {
        $key = Get-AssignmentKey -Assignment $a
        if (-not $desiredMap.ContainsKey($key)) {
            Remove-MgRoleManagementDirectoryRoleAssignment -UnifiedRoleAssignmentId $a.Id -Confirm:$false
            Write-Log "Atribuição extra removida: $key"
        }
    }

    Write-Log "Validação final"
    $finalRole = Get-MgRoleManagementDirectoryRoleDefinition -UnifiedRoleDefinitionId $currentRole.Id
    $finalAssignments = @(Get-MgRoleManagementDirectoryRoleAssignment -All | Where-Object { $_.roleDefinitionId -eq $currentRole.Id })

    Write-Log "Role restaurada: $($finalRole.DisplayName)"
    Write-Log "Descrição atual: $($finalRole.Description)"
    Write-Log "IsEnabled atual: $($finalRole.IsEnabled)"
    Write-Log "Quantidade final de atribuições: $($finalAssignments.Count)"

    $finalRole | Select-Object Id, DisplayName, Description, IsEnabled
    $finalRole.RolePermissions | ConvertTo-Json -Depth 20
    $finalAssignments | Select-Object Id, PrincipalId, RoleDefinitionId, DirectoryScopeId, AppScopeId, Condition, ConditionVersion
}
finally {
    Disconnect-MgGraph | Out-Null
}