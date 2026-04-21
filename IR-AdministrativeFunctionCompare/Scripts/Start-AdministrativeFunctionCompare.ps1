Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

Import-Module Microsoft.Graph.Authentication
Import-Module Microsoft.Graph.Identity.Governance

$BasePath = "C:\ProgramData\Quest\IR-AdministrativeFunctionCompare"
$XamlPath = Join-Path $BasePath "Xaml\MainWindow.xaml"
$LogPath = Join-Path $BasePath "Logs\Compare-$(Get-Date -Format 'yyyy-MM-dd_HH-mm-ss').log"

$DefaultTenantId = "ca9b03ea-578e-4277-b684-969fa2a34a9a"
$DefaultClientId = "9a682c54-d080-46b9-b217-8ffeda5dfa3e"
$DefaultThumbprint = "8668EA996EB3D6A9467B0AFF3658E219C95969D6"

# AJUSTE ESTES 3 VALORES PARA O MÓDULO DE RESTORE
$RestoreScriptPath = "C:\ProgramData\Quest\IR-AdministrativeFunctionRestore\Scripts\Restore-CustomRole-Full.ps1"
$RestoreClientId = "3c1e0342-a7c1-434e-bd70-04ace3dfd88d"
$RestoreThumbprint = "FDD65BA99EB803D11BFBDD1F02424EBA39CBC91B"

function Write-Log {
    param([string]$Message)
    $line = "[{0}] {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Message
    $line | Tee-Object -FilePath $LogPath -Append
}

function Normalize-RoleDefinition {
    param($Role)

    $actions = @()

    if ($Role -and $Role.rolePermissions) {
        foreach ($rp in $Role.rolePermissions) {
            if ($rp.allowedResourceActions) {
                foreach ($a in $rp.allowedResourceActions) {
                    $actions += [string]$a
                }
            }
        }
    }

    [PSCustomObject]@{
        Id          = [string]$Role.id
        DisplayName = [string]$Role.displayName
        Description = [string]$Role.description
        IsEnabled   = [string]$Role.isEnabled
        Actions     = @($actions | Sort-Object -Unique)
        IsBuiltIn   = [string]$Role.isBuiltIn
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

function Normalize-AssignmentsForRole {
    param(
        [array]$Assignments,
        [string]$RoleDefinitionId
    )

    $items = @($Assignments | Where-Object { [string]$_.roleDefinitionId -eq $RoleDefinitionId })
    $result = @()

    foreach ($a in $items) {
        $result += [PSCustomObject]@{
            Key              = Get-AssignmentKey -Assignment $a
            PrincipalId      = [string]$a.principalId
            DirectoryScopeId = [string]$a.directoryScopeId
            AppScopeId       = [string]$a.appScopeId
            Condition        = [string]$a.condition
            ConditionVersion = [string]$a.conditionVersion
        }
    }

    return $result
}

function Load-Snapshot {
    param([string]$Folder)

    $roleDefinitionsPath = Join-Path $Folder "roleDefinitions.json"
    $roleAssignmentsPath = Join-Path $Folder "roleAssignments.json"
    $manifestPath = Join-Path $Folder "manifest.json"

    if (-not (Test-Path $roleDefinitionsPath)) {
        throw "Arquivo roleDefinitions.json não encontrado na pasta selecionada."
    }

    if (-not (Test-Path $roleAssignmentsPath)) {
        throw "Arquivo roleAssignments.json não encontrado na pasta selecionada."
    }

    return [PSCustomObject]@{
        Definitions = Get-Content $roleDefinitionsPath -Raw | ConvertFrom-Json
        Assignments = Get-Content $roleAssignmentsPath -Raw | ConvertFrom-Json
        Manifest    = if (Test-Path $manifestPath) { Get-Content $manifestPath -Raw | ConvertFrom-Json } else { $null }
    }
}

function Load-CurrentTenantData {
    param(
        [string]$TenantId,
        [string]$ClientId,
        [string]$Thumbprint
    )

    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $Thumbprint -NoWelcome

    $definitions = Get-MgRoleManagementDirectoryRoleDefinition -All
    $assignments = Get-MgRoleManagementDirectoryRoleAssignment -All

    return [PSCustomObject]@{
        Definitions = $definitions
        Assignments = $assignments
    }
}

function Format-RoleDetails {
    param(
        $RoleObject,
        [array]$Assignments,
        [string]$Label
    )

    if (-not $RoleObject) {
        return "$Label`r`n`r`nObjeto não existe."
    }

    $normalized = Normalize-RoleDefinition -Role $RoleObject
    $lines = New-Object System.Collections.Generic.List[string]

    $lines.Add("$Label")
    $lines.Add("")
    $lines.Add("ID: $($normalized.Id)")
    $lines.Add("Nome: $($normalized.DisplayName)")
    $lines.Add("Descrição: $($normalized.Description)")
    $lines.Add("IsEnabled: $($normalized.IsEnabled)")
    $lines.Add("IsBuiltIn: $($normalized.IsBuiltIn)")
    $lines.Add("")
    $lines.Add("Permissões:")

    if ($normalized.Actions.Count -eq 0) {
        $lines.Add("  Nenhuma")
    }
    else {
        foreach ($a in $normalized.Actions) {
            $lines.Add("  $a")
        }
    }

    $lines.Add("")
    $lines.Add("Atribuições:")

    if (-not $Assignments -or $Assignments.Count -eq 0) {
        $lines.Add("  Nenhuma")
    }
    else {
        foreach ($a in $Assignments) {
            $lines.Add("  PrincipalId: $($a.PrincipalId)")
            $lines.Add("  DirectoryScopeId: $($a.DirectoryScopeId)")
            $lines.Add("  AppScopeId: $($a.AppScopeId)")
            $lines.Add("  Condition: $($a.Condition)")
            $lines.Add("  ConditionVersion: $($a.ConditionVersion)")
            $lines.Add("")
        }
    }

    return ($lines -join "`r`n")
}

function Compare-Roles {
    param(
        $SnapshotData,
        $CurrentData
    )

    $results = @()

    $snapshotRoles = @{}
    foreach ($r in $SnapshotData.Definitions) {
        $snapshotRoles[[string]$r.displayName] = $r
    }

    $currentRoles = @{}
    foreach ($r in $CurrentData.Definitions) {
        $currentRoles[[string]$r.DisplayName] = $r
    }

    $allNames = @($snapshotRoles.Keys + $currentRoles.Keys | Sort-Object -Unique)

    foreach ($name in $allNames) {
        $snapRaw = $snapshotRoles[$name]
        $currRaw = $currentRoles[$name]

        if ($snapRaw -and -not $currRaw) {
            $snapAssignments = Normalize-AssignmentsForRole -Assignments $SnapshotData.Assignments -RoleDefinitionId ([string]$snapRaw.id)

            $results += [PSCustomObject]@{
                RoleName           = $name
                RoleType           = if ($snapRaw.isBuiltIn) { "Built-in" } else { "Custom" }
                Status             = "Removida"
                DefinitionChanged  = "Sim"
                AssignmentsChanged = "Sim"
                DifferenceCount    = 1
                Summary            = "Função existia no backup e não existe no tenant atual."
                BackupObject       = $snapRaw
                CurrentObject      = $null
                BackupAssignments  = $snapAssignments
                CurrentAssignments = @()
            }
            continue
        }

        if (-not $snapRaw -and $currRaw) {
            $currAssignments = Normalize-AssignmentsForRole -Assignments $CurrentData.Assignments -RoleDefinitionId ([string]$currRaw.Id)

            $results += [PSCustomObject]@{
                RoleName           = $name
                RoleType           = if ($currRaw.IsBuiltIn) { "Built-in" } else { "Custom" }
                Status             = "Nova"
                DefinitionChanged  = "Sim"
                AssignmentsChanged = "Sim"
                DifferenceCount    = 1
                Summary            = "Função não existia no backup e existe no tenant atual."
                BackupObject       = $null
                CurrentObject      = $currRaw
                BackupAssignments  = @()
                CurrentAssignments = $currAssignments
            }
            continue
        }

        $snapNorm = Normalize-RoleDefinition -Role $snapRaw
        $currNorm = Normalize-RoleDefinition -Role $currRaw

        $definitionDiffs = @()

        if ($snapNorm.Description -ne $currNorm.Description) {
            $definitionDiffs += "Description"
        }

        if ($snapNorm.IsEnabled -ne $currNorm.IsEnabled) {
            $definitionDiffs += "IsEnabled"
        }

        $missingActions = @($snapNorm.Actions | Where-Object { $_ -notin $currNorm.Actions })
        $newActions = @($currNorm.Actions | Where-Object { $_ -notin $snapNorm.Actions })

        if ($missingActions.Count -gt 0 -or $newActions.Count -gt 0) {
            $definitionDiffs += "AllowedResourceActions"
        }

        $snapAssignments = Normalize-AssignmentsForRole -Assignments $SnapshotData.Assignments -RoleDefinitionId ([string]$snapRaw.id)
        $currAssignments = Normalize-AssignmentsForRole -Assignments $CurrentData.Assignments -RoleDefinitionId ([string]$currRaw.Id)

        $snapKeys = @($snapAssignments.Key)
        $currKeys = @($currAssignments.Key)

        $removedAssignments = @($snapKeys | Where-Object { $_ -notin $currKeys })
        $addedAssignments = @($currKeys | Where-Object { $_ -notin $snapKeys })

        $assignmentsChanged = ($removedAssignments.Count -gt 0 -or $addedAssignments.Count -gt 0)
        $definitionChanged = ($definitionDiffs.Count -gt 0)

        $status = if (-not $definitionChanged -and -not $assignmentsChanged) { "Igual" } else { "Alterada" }
        $diffCount = $definitionDiffs.Count + $removedAssignments.Count + $addedAssignments.Count

        $summaryParts = @()

        if ($definitionDiffs.Count -gt 0) {
            $summaryParts += ("Definição: " + ($definitionDiffs -join ", "))
        }

        if ($removedAssignments.Count -gt 0) {
            $summaryParts += ("Atribuições removidas: " + $removedAssignments.Count)
        }

        if ($addedAssignments.Count -gt 0) {
            $summaryParts += ("Atribuições adicionadas: " + $addedAssignments.Count)
        }

        if ($summaryParts.Count -eq 0) {
            $summaryParts += "Sem diferenças"
        }

        $results += [PSCustomObject]@{
            RoleName           = $name
            RoleType           = if ($currNorm.IsBuiltIn -eq "True") { "Built-in" } else { "Custom" }
            Status             = $status
            DefinitionChanged  = if ($definitionChanged) { "Sim" } else { "Não" }
            AssignmentsChanged = if ($assignmentsChanged) { "Sim" } else { "Não" }
            DifferenceCount    = $diffCount
            Summary            = ($summaryParts -join " | ")
            BackupObject       = $snapRaw
            CurrentObject      = $currRaw
            BackupAssignments  = $snapAssignments
            CurrentAssignments = $currAssignments
        }
    }

    return @(
        $results | Sort-Object `
            @{ Expression = {
                    switch ($_.Status) {
                        "Alterada" { 0 }
                        "Nova"     { 1 }
                        "Removida" { 2 }
                        "Igual"    { 3 }
                        default    { 9 }
                    }
                }
            }, RoleName
    )
}

function Export-Diff {
    param(
        [array]$Data,
        [string]$BasePath
    )

    $stamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
    $jsonPath = Join-Path $BasePath "Exports\diff-report-$stamp.json"
    $csvPath = Join-Path $BasePath "Exports\diff-report-$stamp.csv"

    $Data | ConvertTo-Json -Depth 20 | Set-Content -Path $jsonPath -Encoding UTF8

    $Data |
        Select-Object RoleName, RoleType, Status, DefinitionChanged, AssignmentsChanged, DifferenceCount, Summary |
        Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

    return [PSCustomObject]@{
        JsonPath = $jsonPath
        CsvPath  = $csvPath
    }
}

function Apply-RoleFilter {
    param(
        [array]$Data,
        [string]$NameFilter,
        [string]$StatusFilter
    )

    $filtered = $Data

    if ($NameFilter) {
        $filtered = @($filtered | Where-Object { $_.RoleName -like "*$NameFilter*" })
    }

    if ($StatusFilter -and $StatusFilter -ne "Todos") {
        $filtered = @($filtered | Where-Object { $_.Status -eq $StatusFilter })
    }

    return @($filtered)
}

function Update-StatusText {
    param([array]$Data)

    if (-not $Data -or $Data.Count -eq 0) {
        return "Status: nenhuma função listada."
    }

    $total = $Data.Count
    $iguais = @($Data | Where-Object { $_.Status -eq "Igual" }).Count
    $alteradas = @($Data | Where-Object { $_.Status -eq "Alterada" }).Count
    $novas = @($Data | Where-Object { $_.Status -eq "Nova" }).Count
    $removidas = @($Data | Where-Object { $_.Status -eq "Removida" }).Count

    return "Status: comparação concluída. Total: $total | Iguais: $iguais | Alteradas: $alteradas | Novas: $novas | Removidas: $removidas"
}

function Refresh-RestoreButtonState {
    param($SelectedItem)

    if (-not $SelectedItem) {
        $BtnRestoreSelected.IsEnabled = $false
        return
    }

    if ($SelectedItem.RoleType -eq "Built-in") {
        $BtnRestoreSelected.IsEnabled = $false
        return
    }

    if ($SelectedItem.Status -eq "Igual") {
        $BtnRestoreSelected.IsEnabled = $false
        return
    }

    $BtnRestoreSelected.IsEnabled = $true
}

function Run-ComparisonWorkflow {
    if (-not $TxtSnapshotFolder.Text) {
        throw "Selecione primeiro uma pasta de snapshot."
    }

    Write-Log "Carregando snapshot"
    $script:SnapshotData = Load-Snapshot -Folder $TxtSnapshotFolder.Text

    Write-Log "Lendo tenant atual"
    $script:CurrentData = Load-CurrentTenantData -TenantId $TxtTenantId.Text -ClientId $TxtClientId.Text -Thumbprint $TxtThumbprint.Text

    Write-Log "Comparando backup com tenant atual"
    $script:CompareResults = Compare-Roles -SnapshotData $script:SnapshotData -CurrentData $script:CurrentData
    $script:FilteredResults = Apply-RoleFilter -Data $script:CompareResults -NameFilter $TxtFilterRoleName.Text -StatusFilter $CmbStatusFilter.Text

    $GridRoles.ItemsSource = $null
    $GridRoles.ItemsSource = $script:FilteredResults

    $TxtStatus.Text = Update-StatusText -Data $script:FilteredResults
    $TxtBackupDetails.Text = ""
    $TxtCurrentDetails.Text = ""

    Refresh-RestoreButtonState -SelectedItem $null

    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
}

function Invoke-ExternalRestore {
    param(
        [string]$TenantId,
        [string]$RoleName,
        [string]$SnapshotFolder
    )

    if (-not (Test-Path $RestoreScriptPath)) {
        throw "Script de restore não encontrado em: $RestoreScriptPath"
    }

    if ([string]::IsNullOrWhiteSpace($RestoreClientId) -or $RestoreClientId -like "COLE_AQUI*") {
        throw "Preencha a variável `$RestoreClientId no topo do script da Fase 3."
    }

    if ([string]::IsNullOrWhiteSpace($RestoreThumbprint) -or $RestoreThumbprint -like "COLE_AQUI*") {
        throw "Preencha a variável `$RestoreThumbprint no topo do script da Fase 3."
    }

    Write-Log "Chamando restore externo para a função '$RoleName'"

    $restoreOutput = & powershell.exe `
        -ExecutionPolicy Bypass `
        -File $RestoreScriptPath `
        -TenantId $TenantId `
        -ClientId $RestoreClientId `
        -CertificateThumbprint $RestoreThumbprint `
        -SnapshotFolder $SnapshotFolder `
        -RoleName $RoleName 2>&1 | Out-String

    $exitCode = $LASTEXITCODE
    Write-Log "Saída do restore externo:"
    Write-Log $restoreOutput

    if ($exitCode -ne 0) {
        throw "O script de restore retornou exit code $exitCode. Verifique o log."
    }
}

[xml]$xaml = Get-Content $XamlPath -Raw
$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)

$TxtStatus = $window.FindName("TxtStatus")
$TxtTenantId = $window.FindName("TxtTenantId")
$TxtClientId = $window.FindName("TxtClientId")
$TxtThumbprint = $window.FindName("TxtThumbprint")
$TxtSnapshotFolder = $window.FindName("TxtSnapshotFolder")
$BtnBrowseSnapshot = $window.FindName("BtnBrowseSnapshot")
$BtnConnect = $window.FindName("BtnConnect")
$BtnCompare = $window.FindName("BtnCompare")
$BtnRestoreSelected = $window.FindName("BtnRestoreSelected")
$BtnExport = $window.FindName("BtnExport")
$TxtFilterRoleName = $window.FindName("TxtFilterRoleName")
$CmbStatusFilter = $window.FindName("CmbStatusFilter")
$BtnApplyFilter = $window.FindName("BtnApplyFilter")
$GridRoles = $window.FindName("GridRoles")
$TxtBackupDetails = $window.FindName("TxtBackupDetails")
$TxtCurrentDetails = $window.FindName("TxtCurrentDetails")

$TxtTenantId.Text = $DefaultTenantId
$TxtClientId.Text = $DefaultClientId
$TxtThumbprint.Text = $DefaultThumbprint
$BtnRestoreSelected.IsEnabled = $false

$script:SnapshotData = $null
$script:CurrentData = $null
$script:CompareResults = @()
$script:FilteredResults = @()

$BtnBrowseSnapshot.Add_Click({
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = "Selecione a pasta do snapshot da Fase 1"

    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $TxtSnapshotFolder.Text = $dialog.SelectedPath
        $TxtStatus.Text = "Status: snapshot selecionado"
        Write-Log "Snapshot selecionado: $($dialog.SelectedPath)"
    }
})

$BtnConnect.Add_Click({
    try {
        Write-Log "Iniciando conexão e leitura do tenant atual"
        $tenantData = Load-CurrentTenantData -TenantId $TxtTenantId.Text -ClientId $TxtClientId.Text -Thumbprint $TxtThumbprint.Text

        $definitionCount = @($tenantData.Definitions).Count
        $assignmentCount = @($tenantData.Assignments).Count

        $TxtStatus.Text = "Status: conexão validada. RoleDefinitions: $definitionCount | RoleAssignments: $assignmentCount"
        Write-Log "Conexão validada com sucesso"

        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    }
    catch {
        $TxtStatus.Text = "Status: erro ao conectar"
        [System.Windows.MessageBox]::Show($_.Exception.Message, "Erro")
        Write-Log ("Erro ao conectar: " + $_.Exception.Message)
    }
})

$BtnCompare.Add_Click({
    try {
        Run-ComparisonWorkflow
        Write-Log "Comparação concluída com sucesso"
    }
    catch {
        $TxtStatus.Text = "Status: erro na comparação"
        [System.Windows.MessageBox]::Show($_.Exception.Message, "Erro")
        Write-Log ("Erro na comparação: " + $_.Exception.Message)
    }
})

$BtnApplyFilter.Add_Click({
    try {
        $script:FilteredResults = Apply-RoleFilter -Data $script:CompareResults -NameFilter $TxtFilterRoleName.Text -StatusFilter $CmbStatusFilter.Text

        $GridRoles.ItemsSource = $null
        $GridRoles.ItemsSource = $script:FilteredResults

        $TxtStatus.Text = Update-StatusText -Data $script:FilteredResults
        Refresh-RestoreButtonState -SelectedItem $GridRoles.SelectedItem
        Write-Log "Filtro aplicado com sucesso"
    }
    catch {
        [System.Windows.MessageBox]::Show($_.Exception.Message, "Erro")
        Write-Log ("Erro ao aplicar filtro: " + $_.Exception.Message)
    }
})

$GridRoles.Add_SelectionChanged({
    $item = $GridRoles.SelectedItem

    if ($null -ne $item) {
        $TxtBackupDetails.Text = Format-RoleDetails -RoleObject $item.BackupObject -Assignments $item.BackupAssignments -Label "BACKUP"
        $TxtCurrentDetails.Text = Format-RoleDetails -RoleObject $item.CurrentObject -Assignments $item.CurrentAssignments -Label "ATUAL"
    }
    else {
        $TxtBackupDetails.Text = ""
        $TxtCurrentDetails.Text = ""
    }

    Refresh-RestoreButtonState -SelectedItem $item
})

$BtnRestoreSelected.Add_Click({
    try {
        $item = $GridRoles.SelectedItem

        if (-not $item) {
            throw "Selecione uma função antes de restaurar."
        }

        if ($item.RoleType -eq "Built-in") {
            throw "Restore por este fluxo só é suportado para custom role."
        }

        if ($item.Status -eq "Igual") {
            throw "A função selecionada já está igual ao snapshot."
        }

        if (-not $TxtSnapshotFolder.Text) {
            throw "Selecione primeiro uma pasta de snapshot."
        }

        $confirmationMessage = "Deseja restaurar a função abaixo usando o snapshot selecionado?`r`n`r`nFunção: $($item.RoleName)`r`nStatus atual: $($item.Status)`r`nSnapshot: $($TxtSnapshotFolder.Text)`r`n`r`nA restauração chamará o script externo da Fase 2 e reconciliará a definição e as atribuições da função."
        $confirmation = [System.Windows.MessageBox]::Show(
            $confirmationMessage,
            "Confirmar restore",
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Warning
        )

        if ($confirmation -ne [System.Windows.MessageBoxResult]::Yes) {
            Write-Log "Restore cancelado pelo operador para a função '$($item.RoleName)'"
            return
        }

        $TxtStatus.Text = "Status: executando restore da função '$($item.RoleName)'"
        Write-Log "Iniciando restore da função '$($item.RoleName)'"

        Invoke-ExternalRestore -TenantId $TxtTenantId.Text -RoleName $item.RoleName -SnapshotFolder $TxtSnapshotFolder.Text

        Start-Sleep -Seconds 5
        Run-ComparisonWorkflow

        [System.Windows.MessageBox]::Show("Restore concluído com sucesso para a função '$($item.RoleName)'.", "Restore concluído")
        Write-Log "Restore concluído com sucesso para a função '$($item.RoleName)'"
    }
    catch {
        $TxtStatus.Text = "Status: erro no restore"
        [System.Windows.MessageBox]::Show($_.Exception.Message, "Erro no restore")
        Write-Log ("Erro no restore: " + $_.Exception.Message)
    }
})

$BtnExport.Add_Click({
    try {
        if (-not $script:FilteredResults -or $script:FilteredResults.Count -eq 0) {
            throw "Não há resultados para exportar."
        }

        $paths = Export-Diff -Data $script:FilteredResults -BasePath $BasePath
        $TxtStatus.Text = "Status: diff exportado com sucesso"
        [System.Windows.MessageBox]::Show("Arquivos gerados:`r`n$($paths.JsonPath)`r`n$($paths.CsvPath)", "Export concluído")
        Write-Log "Diff exportado com sucesso"
    }
    catch {
        $TxtStatus.Text = "Status: erro ao exportar"
        [System.Windows.MessageBox]::Show($_.Exception.Message, "Erro")
        Write-Log ("Erro ao exportar: " + $_.Exception.Message)
    }
})

$window.ShowDialog() | Out-Null