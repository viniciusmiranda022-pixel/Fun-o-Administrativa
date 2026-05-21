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
    Import-Module Microsoft.Graph.Authentication
    Import-Module Microsoft.Graph.Identity.Governance
    $sharedModulePath = Join-Path $PSScriptRoot "..\..\IR-AdministrativeFunctionBackup\Scripts\IR-AdministrativeFunctions.psm1"
    Import-Module $sharedModulePath -Force
} catch {
    Write-Error "Falha ao importar modulos necessarios: $($_.Exception.Message)"
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
    Write-Log "Conectando ao Microsoft Graph (com retentativa para replicação do App Registration)"
    Connect-AppOnlyGraphWithRetry -TenantId $TenantId -ClientId $ClientId -Thumbprint $CertificateThumbprint -ClientSecret $ClientSecret -VerboseOutput -FailureMessage 'Falha ao conectar ao Graph em modo app-only após {0} tentativas. Último erro: {1}'

    $roleDefinitionsPath = Join-Path $SnapshotFolder "roleDefinitions.json"
    $roleAssignmentsPath = Join-Path $SnapshotFolder "roleAssignments.json"

    if (-not (Test-Path $roleDefinitionsPath)) {
        throw "roleDefinitions.json não foi encontrado em $SnapshotFolder"
    }

    if (-not (Test-Path $roleAssignmentsPath)) {
        throw "roleAssignments.json não foi encontrado em $SnapshotFolder"
    }

    Write-Log "Validando integridade do snapshot"
    Test-SnapshotIntegrity -SnapshotFolder $SnapshotFolder

    Write-Log "Carregando arquivos do snapshot"
    $definitions = Get-Content $roleDefinitionsPath -Raw | ConvertFrom-Json
    $assignments = Get-Content $roleAssignmentsPath -Raw | ConvertFrom-Json

    $desiredRole = $definitions | Where-Object { $_.displayName -eq $RoleName } | Select-Object -First 1

    if (-not $desiredRole) {
        throw "A função '$RoleName' não foi encontrada em roleDefinitions.json"
    }

    Write-Log "Função encontrada no snapshot: $($desiredRole.displayName)"
    Write-Log "Modo de execução: $Mode"

    if ($RemoveExtraAssignments -and $Mode -eq "WhatIf") {
        Write-Log "-RemoveExtraAssignments foi informado com -Mode WhatIf. Nenhuma remoção de atribuição será executada no modo de pré-visualização."
    }

    $currentRole = Get-MgRoleManagementDirectoryRoleDefinition -Filter "displayName eq '$RoleName'" | Select-Object -First 1

    if ($currentRole -and $currentRole.IsBuiltIn) {
        throw "A função existente '$RoleName' no tenant é built-in. Este restore suporta apenas funções personalizadas."
    }

    if (-not $currentRole) {
        Write-Output "DIAG: path=CREATE"
        Write-Log "A função não existe atualmente. O snapshot exige criação da função."
        Write-Output "DIAG: snapshot displayName='$($desiredRole.displayName)' isEnabled=$($desiredRole.isEnabled)"

        if (-not [bool]$desiredRole.isEnabled) {
            Write-Output "DIAG: ATENCAO isEnabled=false no snapshot. Role sera criada DESATIVADA."
            Write-Log "ATENCAO: isEnabled=false no snapshot. A funcao sera criada como DESATIVADA e nao aparecera como ativa no portal Azure."
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

            Write-Output "DIAG: chamando New-MgRoleManagementDirectoryRoleDefinition IsEnabled=$($newRoleParams.IsEnabled)"

            # Nao passa TemplateId: roles soft-deleted no Entra ID ocupam o templateId antigo,
            # causando conflito 400. O Entra atribui um novo GUID automaticamente.

            $currentRole = New-MgRoleManagementDirectoryRoleDefinition @newRoleParams
            if (-not $currentRole -or -not $currentRole.Id) {
                throw "New-MgRoleManagementDirectoryRoleDefinition retornou vazio sem lancai excecao. Verifique se o App Registration tem a permissao 'RoleManagement.ReadWrite.Directory' e se o tenant possui licenca Entra ID P1/P2."
            }
            Write-Output "DIAG: role criada Id=$($currentRole.Id) DisplayName='$($currentRole.DisplayName)' IsEnabled=$($currentRole.IsEnabled)"
            Write-Log "Funcao criada: Id=$($currentRole.Id), DisplayName='$($currentRole.DisplayName)', IsEnabled=$($currentRole.IsEnabled)"
        }
        else {
            Write-Log "[PRÉVIA] A função seria criada com base na definição do snapshot."
        }
    }
    else {
        Write-Output "DIAG: path=UPDATE existingId=$($currentRole.Id)"
        Write-Log "A função já existe. O snapshot exige sobrescrita completa da definição."

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
                throw "Update-MgRoleManagementDirectoryRoleDefinition falhou silenciosamente. Verifique a permissao 'RoleManagement.ReadWrite.Directory' no App Registration."
            }
            Write-Log "Definição da função atualizada"
        }
        else {
            Write-Log "[PRÉVIA] A definição da função seria atualizada para os valores do snapshot."
        }
    }

    Write-Log "Processando atribuições do snapshot"

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

    Write-Log "Relatório pré-restore exportado: $preRestoreReportPath"

    if ($Mode -eq "Apply") {
        Write-Host ""
        Write-Host "Relatório pré-restore gerado em: $preRestoreReportPath" -ForegroundColor Yellow

        if (-not $SkipConfirmation) {
            $confirmation = Read-Host "Digite CONFIRM para continuar com as alterações de restore"
            if ($confirmation -ne "CONFIRM") {
                throw "Restore cancelado pelo operador. Revise o relatório e execute novamente quando estiver pronto."
            }
            Write-Log "Operador confirmou a execução do restore após revisar o relatório."
        }
        else {
            Write-Log "Confirmação interativa ignorada por parâmetro -SkipConfirmation."
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
                Write-Log "Atribuição recriada: $key"
            }
            else {
                Write-Log "[PRÉVIA] A atribuição seria criada: $key"
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
                Write-Log "Atribuição extra removida: $key"
            }
            else {
                Write-Log "[PRÉVIA] Atribuição extra detectada (não removida): $key"
            }
        }
    }

    Write-Log "Validação final"

    if ($currentRole) {
        $finalRole = Get-MgRoleManagementDirectoryRoleDefinition -UnifiedRoleDefinitionId $currentRole.Id
        $finalAssignments = @(Get-MgRoleManagementDirectoryRoleAssignment -All | Where-Object { $_.roleDefinitionId -eq $currentRole.Id })

        if ($Mode -eq "Apply") {
            Write-Log "Função restaurada: $($finalRole.DisplayName)"
        }
        else {
            Write-Log "Prévia concluída para a função: $($finalRole.DisplayName)"
        }

        Write-Output "DIAG: validacao final Id=$($finalRole.Id) DisplayName='$($finalRole.DisplayName)' IsEnabled=$($finalRole.IsEnabled) assignments=$($finalAssignments.Count)"
        Write-Log "Descrição atual: $($finalRole.Description)"
        Write-Log "IsEnabled atual: $($finalRole.IsEnabled)"
        Write-Log "Quantidade atual de atribuições: $($finalAssignments.Count)"

        $finalRole | Select-Object Id, DisplayName, Description, IsEnabled
        $finalRole.RolePermissions | ConvertTo-Json -Depth 20
        $finalAssignments | Select-Object Id, PrincipalId, RoleDefinitionId, DirectoryScopeId, AppScopeId, Condition, ConditionVersion
    }
    else {
        Write-Log "Prévia concluída. A função não existe atualmente e seria criada no modo Apply."
    }

    Write-StructuredLog -Message "Fluxo de restore concluído com sucesso." -Result "SUCCESS"
    exit 0
}
catch {
    $errorMsg = $_.Exception.Message
    if ($script:StructuredLogPath) {
        Write-StructuredLog -Message ("Fluxo de restore falhou: {0}" -f $errorMsg) -Result "FAILED"
    }
    Write-Error "Restore falhou: $errorMsg"
    exit 1
}
finally {
    try { Disconnect-MgGraph | Out-Null } catch {}
}
