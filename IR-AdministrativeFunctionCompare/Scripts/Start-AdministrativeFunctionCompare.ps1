Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

Import-Module Microsoft.Graph.Authentication
Import-Module Microsoft.Graph.Identity.Governance

$BasePath = "C:\ProgramData\Quest\IR-AdministrativeFunctionCompare"
$XamlPath = Join-Path $BasePath "Xaml\MainWindow.xaml"
$LogDirectory = Join-Path $BasePath "Logs"
if (-not (Test-Path $LogDirectory)) { New-Item -Path $LogDirectory -ItemType Directory -Force | Out-Null }
$LogPath = Join-Path $LogDirectory "Compare-$(Get-Date -Format 'yyyy-MM-dd_HH-mm-ss').log"
$BackupExportsPath = "C:\ProgramData\Quest\IR-AdministrativeFunctionBackup\Exports"
$RestoreScriptPath = "C:\ProgramData\Quest\IR-AdministrativeFunctionRestore\Scripts\Restore-CustomRole-Full.ps1"

$DefaultTenantId = ""
$DefaultClientId = ""
$DefaultThumbprint = ""

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO", "WARN", "ERROR")]
        [string]$Severity = "INFO"
    )

    $line = "[{0}] {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Message
    $line | Tee-Object -FilePath $LogPath -Append

    if ($script:EventEntries -and $GridEvents) {
        $script:EventEntries.Insert(0, [PSCustomObject]@{
                Time     = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                Severity = $Severity
                Message  = $Message
            })
        $GridEvents.ItemsSource = $null
        $GridEvents.ItemsSource = $script:EventEntries
    }
}

function Set-UiState {
    param(
        [string]$StatusText,
        [string]$ReplicationText,
        [bool]$Busy = $false
    )

    if ($StatusText) { $TxtStatus.Text = $StatusText }
    if ($ReplicationText) { $TxtReplicationStatus.Text = $ReplicationText }

    $BtnConnect.IsEnabled = -not $Busy
    if ($BtnHamburger) { $BtnHamburger.IsEnabled = -not $Busy }
    $BtnCompare.IsEnabled = -not $Busy
    $BtnBrowseSnapshot.IsEnabled = -not $Busy
    $BtnLoadBackups.IsEnabled = -not $Busy
    $BtnRefreshBackups.IsEnabled = -not $Busy
    $BtnApplyFilter.IsEnabled = -not $Busy
    $BtnExport.IsEnabled = -not $Busy
    $GridBackups.IsEnabled = -not $Busy
    $GridRoles.IsEnabled = -not $Busy

    if ($Busy) {
        $BtnRestoreSelected.IsEnabled = $false
    }
    else {
        Refresh-RestoreButtonState -SelectedItem $GridRoles.SelectedItem
    }
}

function Reset-GraphContext {
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null

    $clearContextCommand = Get-Command -Name Clear-MgContext -ErrorAction SilentlyContinue
    if ($clearContextCommand) {
        Clear-MgContext -ErrorAction SilentlyContinue | Out-Null
    }

    $script:CurrentData = $null
}

function Test-IsElevated {
    try {
        $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object Security.Principal.WindowsPrincipal($identity)
        return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    }
    catch {
        return $false
    }
}

function Connect-OperatorTenantSession {
    param([string]$TenantId)

    Reset-GraphContext
    if (Test-IsElevated) {
        throw "Esta aplicação deve ser executada em PowerShell não elevado para abrir o pop-up interativo da Microsoft (WAM). Feche e reabra sem 'Executar como administrador'."
    }

    $connectParams = @{
        TenantId     = $TenantId
        Scopes       = @("RoleManagement.Read.Directory")
        NoWelcome    = $true
        ContextScope = "Process"
    }

    Connect-MgGraph @connectParams | Out-Null

    $context = Get-MgContext
    if (-not $context -or $context.TenantId -ne $TenantId) {
        throw "A autenticação retornou tenant inesperado. Tenant solicitado: $TenantId | Tenant autenticado: $($context.TenantId)"
    }

    Write-Log "Autenticação do operador concluída no tenant correto: $TenantId"
}

function Add-TaskEntry {
    param(
        [string]$Task,
        [string]$Status,
        [string]$Details
    )

    if (-not $script:TaskEntries) {
        return
    }

    $script:TaskEntries.Insert(0, [PSCustomObject]@{
            Time    = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            Task    = $Task
            Status  = $Status
            Details = $Details
        })
    $GridTasks.ItemsSource = $null
    $GridTasks.ItemsSource = $script:TaskEntries
}

function Connect-AppOnlyWithRetry {
    param(
        [string]$TenantId,
        [string]$ClientId,
        [string]$Thumbprint,
        [string]$Operation = "Validação",
        [scriptblock]$StatusCallback
    )

    $maxAttempts = 8
    $delaySeconds = 15

    for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {
        Reset-GraphContext

        try {
            if ($StatusCallback) {
                & $StatusCallback "Replicação de App Registration: tentativa $attempt de $maxAttempts (janela máxima de 2 minutos)." 0
            }

            Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $Thumbprint -NoWelcome -ContextScope Process | Out-Null
            Get-MgRoleManagementDirectoryRoleDefinition -Top 1 | Out-Null
            Write-Log "${Operation}: conexão app-only bem-sucedida na tentativa $attempt."

            if ($StatusCallback) {
                & $StatusCallback "Replicação de App Registration: concluída na tentativa $attempt." 0
            }

            return
        }
        catch {
            $message = $_.Exception.Message
            Write-Log "${Operation}: tentativa $attempt falhou - $message"

            if ($attempt -eq $maxAttempts) {
                throw "Falha na conexão app-only após 2 minutos (8 tentativas com espera de 15s). Último erro: $message"
            }

            for ($remaining = $delaySeconds; $remaining -gt 0; $remaining--) {
                if ($StatusCallback) {
                    & $StatusCallback "Replicação de App Registration: aguardando próxima tentativa ($attempt/$maxAttempts)." $remaining
                }

                Start-Sleep -Seconds 1
            }
        }
        finally {
            Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        }
    }
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

    Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $Thumbprint -NoWelcome -ContextScope Process | Out-Null

    $definitions = Get-MgRoleManagementDirectoryRoleDefinition -All
    $assignments = Get-MgRoleManagementDirectoryRoleAssignment -All

    return [PSCustomObject]@{
        Definitions = $definitions
        Assignments = $assignments
    }
}

function Get-BackupSnapshots {
    if (-not (Test-Path $BackupExportsPath)) {
        return @()
    }

    $items = Get-ChildItem -Path $BackupExportsPath -Directory -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending |
        ForEach-Object {
            $manifestPath = Join-Path $_.FullName "manifest.json"
            $collectedAt = "-"

            if (Test-Path $manifestPath) {
                try {
                    $manifest = Get-Content $manifestPath -Raw | ConvertFrom-Json
                    if ($manifest.CollectedAt) {
                        $collectedAt = [string]$manifest.CollectedAt
                    }
                }
                catch {
                    $collectedAt = "manifest inválido"
                }
            }

            [PSCustomObject]@{
                Timestamp   = $_.Name
                CollectedAt = $collectedAt
                FullName    = $_.FullName
            }
        }

    return @($items)
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

function ConvertTo-DisplayValue {
    param($Value)

    if ($null -eq $Value) {
        return ""
    }

    if ($Value -is [array]) {
        return ($Value | ForEach-Object { [string]$_ } | Sort-Object -Unique) -join "; "
    }

    return [string]$Value
}

function Update-UnpackedObjectsGrid {
    param($Item)

    $rows = New-Object System.Collections.Generic.List[object]
    if (-not $Item) {
        $GridUnpackedObjects.ItemsSource = $null
        return
    }

    $backupNorm = if ($Item.BackupObject) { Normalize-RoleDefinition -Role $Item.BackupObject } else { $null }
    $currentNorm = if ($Item.CurrentObject) { Normalize-RoleDefinition -Role $Item.CurrentObject } else { $null }

    $fields = @("Id", "DisplayName", "Description", "IsEnabled", "IsBuiltIn", "Actions")
    foreach ($field in $fields) {
        $rows.Add([PSCustomObject]@{
                Section      = "Definition"
                Field        = $field
                BackupValue  = ConvertTo-DisplayValue -Value ($backupNorm.$field)
                CurrentValue = ConvertTo-DisplayValue -Value ($currentNorm.$field)
            })
    }

    $rows.Add([PSCustomObject]@{
            Section      = "Assignments"
            Field        = "Total"
            BackupValue  = [string](@($Item.BackupAssignments).Count)
            CurrentValue = [string](@($Item.CurrentAssignments).Count)
        })

    $GridUnpackedObjects.ItemsSource = $null
    $GridUnpackedObjects.ItemsSource = $rows
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


function Get-SelectedStatusFilter {
    $selected = $CmbStatusFilter.SelectedItem

    if ($selected -is [System.Windows.Controls.ComboBoxItem]) {
        return [string]$selected.Content
    }

    return [string]$CmbStatusFilter.Text
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

function Ensure-ConnectionInputs {
    $missingFields = New-Object System.Collections.Generic.List[string]

    if ([string]::IsNullOrWhiteSpace($TxtTenantId.Text)) {
        $missingFields.Add("Tenant ID")
    }

    if ([string]::IsNullOrWhiteSpace($TxtClientId.Text)) {
        $missingFields.Add("Client ID")
    }

    if ([string]::IsNullOrWhiteSpace($TxtThumbprint.Text)) {
        $missingFields.Add("Certificate Thumbprint")
    }

    if ($missingFields.Count -gt 0) {
        $MainTabs.SelectedIndex = 2
        $TxtStatus.Text = "Status: preencha os dados obrigatórios na aba TENANTS."
        $fields = [string]::Join(", ", $missingFields.ToArray())
        throw "Preencha os campos obrigatórios na aba TENANTS: $fields."
    }
}

function Run-ComparisonWorkflow {
    Ensure-ConnectionInputs

    if (-not $TxtSnapshotFolder.Text) {
        throw "Selecione primeiro uma pasta de snapshot."
    }

    Set-UiState -StatusText "Status: preparando comparação..." -Busy $true

    try {
        Write-Log "Carregando snapshot"
        $script:SnapshotData = Load-Snapshot -Folder $TxtSnapshotFolder.Text

        Connect-AppOnlyWithRetry -TenantId $TxtTenantId.Text -ClientId $TxtClientId.Text -Thumbprint $TxtThumbprint.Text -Operation "Comparação" -StatusCallback {
            param($text, $remainingSeconds)
            if ($remainingSeconds -gt 0) {
                $TxtReplicationStatus.Text = "$text Aguarde $remainingSeconds s..."
            }
            else {
                $TxtReplicationStatus.Text = $text
            }
        }

        Write-Log "Lendo tenant atual"
        $script:CurrentData = Load-CurrentTenantData -TenantId $TxtTenantId.Text -ClientId $TxtClientId.Text -Thumbprint $TxtThumbprint.Text

        Write-Log "Comparando backup com tenant atual"
        $script:CompareResults = Compare-Roles -SnapshotData $script:SnapshotData -CurrentData $script:CurrentData
        $script:FilteredResults = Apply-RoleFilter -Data $script:CompareResults -NameFilter $TxtFilterRoleName.Text -StatusFilter (Get-SelectedStatusFilter)

        $GridRoles.ItemsSource = $null
        $GridRoles.ItemsSource = $script:FilteredResults

        $TxtStatus.Text = Update-StatusText -Data $script:FilteredResults
        $TxtBackupDetails.Text = ""
        $TxtCurrentDetails.Text = ""
        $MainTabs.SelectedIndex = 1

        Refresh-RestoreButtonState -SelectedItem $null

        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    }
    finally {
        Set-UiState -ReplicationText "Replicação de App Registration: pronta para próximo fluxo." -Busy $false
    }
}

function Invoke-ExternalRestore {
    param(
        [string]$TenantId,
        [string]$ClientId,
        [string]$Thumbprint,
        [string]$RoleName,
        [string]$SnapshotFolder
    )

    if (-not (Test-Path $RestoreScriptPath)) {
        throw "Script de restore não encontrado em: $RestoreScriptPath"
    }

    Write-Log "Chamando restore externo para a função '$RoleName'"

    $restoreOutput = & powershell.exe `
        -ExecutionPolicy Bypass `
        -File $RestoreScriptPath `
        -TenantId $TenantId `
        -ClientId $ClientId `
        -CertificateThumbprint $Thumbprint `
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
$TxtReplicationStatus = $window.FindName("TxtReplicationStatus")
$TxtTenantId = $window.FindName("TxtTenantId")
$TxtClientId = $window.FindName("TxtClientId")
$TxtThumbprint = $window.FindName("TxtThumbprint")
$TxtSnapshotFolder = $window.FindName("TxtSnapshotFolder")
$BtnBrowseSnapshot = $window.FindName("BtnBrowseSnapshot")
$BtnConnect = $window.FindName("BtnConnect")
$BtnLoadBackups = $window.FindName("BtnLoadBackups")
$BtnRefreshBackups = $window.FindName("BtnRefreshBackups")
$BtnCompare = $window.FindName("BtnCompare")
$BtnRestoreSelected = $window.FindName("BtnRestoreSelected")
$BtnExport = $window.FindName("BtnExport")
$BtnHamburger = $window.FindName("BtnHamburger")
$TxtFilterRoleName = $window.FindName("TxtFilterRoleName")
$CmbStatusFilter = $window.FindName("CmbStatusFilter")
$BtnApplyFilter = $window.FindName("BtnApplyFilter")
$MainTabs = $window.FindName("MainTabs")
$GridBackups = $window.FindName("GridBackups")
$GridRoles = $window.FindName("GridRoles")
$TxtBackupDetails = $window.FindName("TxtBackupDetails")
$TxtCurrentDetails = $window.FindName("TxtCurrentDetails")
$GridUnpackedObjects = $window.FindName("GridUnpackedObjects")
$GridEvents = $window.FindName("GridEvents")
$TxtEventDetails = $window.FindName("TxtEventDetails")
$GridTasks = $window.FindName("GridTasks")
$TxtTenantNotes = $window.FindName("TxtTenantNotes")
$MnuConnectTenant = $window.FindName("MnuConnectTenant")
$MnuOpenTenantsTab = $window.FindName("MnuOpenTenantsTab")
$MnuManageBackups = $window.FindName("MnuManageBackups")
$MnuRunCompare = $window.FindName("MnuRunCompare")
$MnuRestoreSelected = $window.FindName("MnuRestoreSelected")

$TxtTenantId.Text = $DefaultTenantId
$TxtClientId.Text = $DefaultClientId
$TxtThumbprint.Text = $DefaultThumbprint
$BtnRestoreSelected.IsEnabled = $false

$script:SnapshotData = $null
$script:CurrentData = $null
$script:CompareResults = @()
$script:FilteredResults = @()
$script:EventEntries = New-Object 'System.Collections.Generic.List[object]'
$script:TaskEntries = New-Object 'System.Collections.Generic.List[object]'

$TxtTenantNotes.Text = "Segurança: toda operação de compare/restore usa Service Principal + Certificate (app-only)." +
"`r`n`r`nLogin interativo: para setup e validação inicial, o tenant informado é obrigatório e a conexão força Connect-MgGraph -TenantId." +
"`r`n`r`nImportante: execute em PowerShell NÃO elevado para o pop-up WAM da Microsoft funcionar."

if (Test-IsElevated) {
    $TxtStatus.Text = "Status: execução elevada detectada. Reabra sem administrador para autenticação interativa."
    [System.Windows.MessageBox]::Show(
        "Modo administrador detectado. O pop-up interativo da Microsoft (WAM) pode falhar. Reabra em PowerShell comum.",
        "Aviso de execução",
        [System.Windows.MessageBoxButton]::OK,
        [System.Windows.MessageBoxImage]::Warning
    )
}

$loadBackupsAction = {
    $snapshots = Get-BackupSnapshots
    $GridBackups.ItemsSource = $null
    $GridBackups.ItemsSource = $snapshots

    if ($snapshots.Count -eq 0) {
        $TxtStatus.Text = "Status: nenhum backup encontrado em $BackupExportsPath"
    }
    else {
        $TxtStatus.Text = "Status: $($snapshots.Count) backup(s) carregado(s)."
    }

    Write-Log "Backups recarregados da pasta fixa: $BackupExportsPath"
    Add-TaskEntry -Task "Load Backups" -Status "Completed" -Details "$($snapshots.Count) snapshot(s) carregados"
}

$BtnLoadBackups.Add_Click({
    try {
        & $loadBackupsAction
        $MainTabs.SelectedIndex = 0
    }
    catch {
        [System.Windows.MessageBox]::Show($_.Exception.Message, "Erro")
        Write-Log ("Erro ao carregar backups: " + $_.Exception.Message) -Severity "ERROR"
        Add-TaskEntry -Task "Load Backups" -Status "Failed" -Details $_.Exception.Message
    }
})

$BtnRefreshBackups.Add_Click({
    try {
        & $loadBackupsAction
    }
    catch {
        [System.Windows.MessageBox]::Show($_.Exception.Message, "Erro")
        Write-Log ("Erro ao atualizar backups: " + $_.Exception.Message) -Severity "ERROR"
        Add-TaskEntry -Task "Refresh Backups" -Status "Failed" -Details $_.Exception.Message
    }
})

$BtnHamburger.Add_Click({
    $menu = $BtnHamburger.ContextMenu
    $menu.PlacementTarget = $BtnHamburger
    $menu.IsOpen = $true
})

$MnuOpenTenantsTab.Add_Click({
    $MainTabs.SelectedIndex = 2
})

$MnuConnectTenant.Add_Click({
    $BtnConnect.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent)))
})

$MnuManageBackups.Add_Click({
    $BtnLoadBackups.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent)))
})

$MnuRunCompare.Add_Click({
    $BtnCompare.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent)))
})

$MnuRestoreSelected.Add_Click({
    $BtnRestoreSelected.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Button]::ClickEvent)))
})

$GridBackups.Add_SelectionChanged({
    $selected = $GridBackups.SelectedItem
    if ($selected) {
        $TxtSnapshotFolder.Text = $selected.FullName
        $TxtStatus.Text = "Status: backup selecionado ($($selected.Timestamp))."
        $MainTabs.SelectedIndex = 1
        Write-Log "Backup selecionado pelo grid: $($selected.FullName)"
        Add-TaskEntry -Task "Select Snapshot" -Status "Completed" -Details $selected.FullName
    }
})

$BtnBrowseSnapshot.Add_Click({
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = "Selecione a pasta do snapshot"

    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $TxtSnapshotFolder.Text = $dialog.SelectedPath
        $TxtStatus.Text = "Status: snapshot selecionado manualmente."
        Write-Log "Snapshot selecionado manualmente: $($dialog.SelectedPath)"
        Add-TaskEntry -Task "Select Snapshot" -Status "Completed" -Details $dialog.SelectedPath
    }
})

$BtnConnect.Add_Click({
    try {
        Ensure-ConnectionInputs
        Set-UiState -StatusText "Status: autenticando operador no tenant informado..." -Busy $true

        Connect-OperatorTenantSession -TenantId $TxtTenantId.Text

        Set-UiState -StatusText "Status: validando conexão app-only com retry automático..." -Busy $true
        Connect-AppOnlyWithRetry -TenantId $TxtTenantId.Text -ClientId $TxtClientId.Text -Thumbprint $TxtThumbprint.Text -Operation "Validação de conexão" -StatusCallback {
            param($text, $remainingSeconds)
            if ($remainingSeconds -gt 0) {
                $TxtReplicationStatus.Text = "$text Aguarde $remainingSeconds s..."
            }
            else {
                $TxtReplicationStatus.Text = $text
            }
        }

        $TxtStatus.Text = "Status: autenticação e validação concluídas no tenant $($TxtTenantId.Text)."
        Write-Log "Conexão validada com sucesso"
        Add-TaskEntry -Task "Connect Tenant" -Status "Completed" -Details "Tenant $($TxtTenantId.Text)"
    }
    catch {
        $TxtStatus.Text = "Status: erro ao autenticar/conectar"
        [System.Windows.MessageBox]::Show($_.Exception.Message, "Erro")
        Write-Log ("Erro na autenticação/conexão: " + $_.Exception.Message) -Severity "ERROR"
        Add-TaskEntry -Task "Connect Tenant" -Status "Failed" -Details $_.Exception.Message
    }
    finally {
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        Set-UiState -ReplicationText "Replicação de App Registration: pronta para próximo fluxo." -Busy $false
    }
})

$BtnCompare.Add_Click({
    try {
        Run-ComparisonWorkflow
        Write-Log "Comparação concluída com sucesso"
        Add-TaskEntry -Task "Run Compare" -Status "Completed" -Details "$($script:FilteredResults.Count) item(ns) após filtro"
    }
    catch {
        $TxtStatus.Text = "Status: erro na comparação"
        [System.Windows.MessageBox]::Show($_.Exception.Message, "Erro")
        Write-Log ("Erro na comparação: " + $_.Exception.Message) -Severity "ERROR"
        Add-TaskEntry -Task "Run Compare" -Status "Failed" -Details $_.Exception.Message
    }
})

$BtnApplyFilter.Add_Click({
    try {
        $script:FilteredResults = Apply-RoleFilter -Data $script:CompareResults -NameFilter $TxtFilterRoleName.Text -StatusFilter (Get-SelectedStatusFilter)

        $GridRoles.ItemsSource = $null
        $GridRoles.ItemsSource = $script:FilteredResults

        $TxtStatus.Text = Update-StatusText -Data $script:FilteredResults
        Refresh-RestoreButtonState -SelectedItem $GridRoles.SelectedItem
        Write-Log "Filtro aplicado com sucesso"
        Add-TaskEntry -Task "Apply Filter" -Status "Completed" -Details "$($script:FilteredResults.Count) item(ns) filtrados"
    }
    catch {
        [System.Windows.MessageBox]::Show($_.Exception.Message, "Erro")
        Write-Log ("Erro ao aplicar filtro: " + $_.Exception.Message) -Severity "ERROR"
        Add-TaskEntry -Task "Apply Filter" -Status "Failed" -Details $_.Exception.Message
    }
})

$GridRoles.Add_SelectionChanged({
    $item = $GridRoles.SelectedItem

    if ($null -ne $item) {
        $TxtBackupDetails.Text = Format-RoleDetails -RoleObject $item.BackupObject -Assignments $item.BackupAssignments -Label "BACKUP"
        $TxtCurrentDetails.Text = Format-RoleDetails -RoleObject $item.CurrentObject -Assignments $item.CurrentAssignments -Label "ATUAL"
        Update-UnpackedObjectsGrid -Item $item
    }
    else {
        $TxtBackupDetails.Text = ""
        $TxtCurrentDetails.Text = ""
        Update-UnpackedObjectsGrid -Item $null
    }

    Refresh-RestoreButtonState -SelectedItem $item
})

$BtnRestoreSelected.Add_Click({
    try {
        Ensure-ConnectionInputs
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

        $confirmationMessage = "Deseja restaurar a função abaixo usando o snapshot selecionado?`r`n`r`nFunção: $($item.RoleName)`r`nStatus atual: $($item.Status)`r`nSnapshot: $($TxtSnapshotFolder.Text)"
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

        Set-UiState -StatusText "Status: executando restore de '$($item.RoleName)'..." -Busy $true
        Write-Log "Iniciando restore externo da função '$($item.RoleName)'"

        Invoke-ExternalRestore -TenantId $TxtTenantId.Text -ClientId $TxtClientId.Text -Thumbprint $TxtThumbprint.Text -RoleName $item.RoleName -SnapshotFolder $TxtSnapshotFolder.Text

        $TxtStatus.Text = "Status: restore concluído. Execute um novo compare para validar o resultado."
        [System.Windows.MessageBox]::Show("Restore concluído com sucesso para a função '$($item.RoleName)'. Execute Run Compare para atualizar a visão de diferenças.", "Restore concluído")
        Write-Log "Restore externo concluído para a função '$($item.RoleName)'"
        Add-TaskEntry -Task "Restore Selected" -Status "Completed" -Details "Função $($item.RoleName)"
    }
    catch {
        $TxtStatus.Text = "Status: erro no restore"
        [System.Windows.MessageBox]::Show($_.Exception.Message, "Erro no restore")
        Write-Log ("Erro no restore: " + $_.Exception.Message) -Severity "ERROR"
        Add-TaskEntry -Task "Restore Selected" -Status "Failed" -Details $_.Exception.Message
    }
    finally {
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        Set-UiState -ReplicationText "Replicação de App Registration: pronta para próximo fluxo." -Busy $false
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
        Add-TaskEntry -Task "Export Diff" -Status "Completed" -Details $paths.CsvPath
    }
    catch {
        $TxtStatus.Text = "Status: erro ao exportar"
        [System.Windows.MessageBox]::Show($_.Exception.Message, "Erro")
        Write-Log ("Erro ao exportar: " + $_.Exception.Message) -Severity "ERROR"
        Add-TaskEntry -Task "Export Diff" -Status "Failed" -Details $_.Exception.Message
    }
})

if ($GridEvents) {
    $GridEvents.Add_SelectionChanged({
            $selectedEvent = $GridEvents.SelectedItem
            if ($selectedEvent) {
                $TxtEventDetails.Text = "[{0}] {1}`r`n`r`n{2}" -f $selectedEvent.Time, $selectedEvent.Severity, $selectedEvent.Message
            }
            else {
                $TxtEventDetails.Text = ""
            }
        })
}

Write-Log "Painel iniciado. Diretório de backups monitorado: $BackupExportsPath"
Add-TaskEntry -Task "Start Panel" -Status "Completed" -Details "Aplicação iniciada"

& $loadBackupsAction
$window.ShowDialog() | Out-Null
