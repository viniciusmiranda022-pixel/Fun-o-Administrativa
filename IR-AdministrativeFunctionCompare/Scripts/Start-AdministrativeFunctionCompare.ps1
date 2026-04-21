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
        [System.Windows.Input.Mouse]::OverrideCursor = [System.Windows.Input.Cursors]::Wait
    }
    else {
        Refresh-RestoreButtonState -SelectedItem $GridRoles.SelectedItem
        [System.Windows.Input.Mouse]::OverrideCursor = $null
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
        throw "This application must run in non-elevated PowerShell to open the Microsoft interactive sign-in popup (WAM). Close and reopen without 'Run as administrator'."
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
        throw "Authentication returned an unexpected tenant. Requested tenant: $TenantId | Authenticated tenant: $($context.TenantId)"
    }

    Write-Log "Operator authentication completed in the correct tenant: $TenantId"
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
        [string]$Operation = "Validation",
        [scriptblock]$StatusCallback
    )

    $maxAttempts = 8
    $delaySeconds = 15

    for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {
        Reset-GraphContext

        try {
            if ($StatusCallback) {
                & $StatusCallback "App Registration replication: attempt $attempt of $maxAttempts (maximum 2-minute window)." 0
            }

            Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $Thumbprint -NoWelcome -ContextScope Process | Out-Null
            Get-MgRoleManagementDirectoryRoleDefinition -Top 1 | Out-Null
            Write-Log "${Operation}: app-only connection succeeded on attempt $attempt."

            if ($StatusCallback) {
                & $StatusCallback "App Registration replication: completed on attempt $attempt." 0
            }

            return
        }
        catch {
            $message = $_.Exception.Message
            Write-Log "${Operation}: attempt $attempt failed - $message"

            if ($attempt -eq $maxAttempts) {
                throw "App-only connection failed after 2 minutes (8 attempts with 15s wait). Last error: $message"
            }

            for ($remaining = $delaySeconds; $remaining -gt 0; $remaining--) {
                if ($StatusCallback) {
                    & $StatusCallback "App Registration replication: waiting for next attempt ($attempt/$maxAttempts)." $remaining
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
        throw "roleDefinitions.json was not found in the selected folder."
    }

    if (-not (Test-Path $roleAssignmentsPath)) {
        throw "roleAssignments.json was not found in the selected folder."
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
                    $collectedAt = "invalid manifest"
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
        return "$Label`r`n`r`nObject does not exist."
    }

    $normalized = Normalize-RoleDefinition -Role $RoleObject
    $lines = New-Object System.Collections.Generic.List[string]

    $lines.Add("$Label")
    $lines.Add("")
    $lines.Add("ID: $($normalized.Id)")
    $lines.Add("Name: $($normalized.DisplayName)")
    $lines.Add("Description: $($normalized.Description)")
    $lines.Add("IsEnabled: $($normalized.IsEnabled)")
    $lines.Add("IsBuiltIn: $($normalized.IsBuiltIn)")
    $lines.Add("")
    $lines.Add("Permissions:")

    if ($normalized.Actions.Count -eq 0) {
        $lines.Add("  None")
    }
    else {
        foreach ($a in $normalized.Actions) {
            $lines.Add("  $a")
        }
    }

    $lines.Add("")
    $lines.Add("Assignments:")

    if (-not $Assignments -or $Assignments.Count -eq 0) {
        $lines.Add("  None")
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
                Status             = "Removed"
                DefinitionChanged  = "Yes"
                AssignmentsChanged = "Yes"
                DifferenceCount    = 1
                Summary            = "Role existed in the backup and does not exist in the current tenant."
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
                Status             = "New"
                DefinitionChanged  = "Yes"
                AssignmentsChanged = "Yes"
                DifferenceCount    = 1
                Summary            = "Role did not exist in the backup and exists in the current tenant."
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

        $status = if (-not $definitionChanged -and -not $assignmentsChanged) { "Equal" } else { "Changed" }
        $diffCount = $definitionDiffs.Count + $removedAssignments.Count + $addedAssignments.Count

        $summaryParts = @()

        if ($definitionDiffs.Count -gt 0) {
            $summaryParts += ("Definition: " + ($definitionDiffs -join ", "))
        }

        if ($removedAssignments.Count -gt 0) {
            $summaryParts += ("Removed assignments: " + $removedAssignments.Count)
        }

        if ($addedAssignments.Count -gt 0) {
            $summaryParts += ("Added assignments: " + $addedAssignments.Count)
        }

        if ($summaryParts.Count -eq 0) {
            $summaryParts += "No differences"
        }

        $results += [PSCustomObject]@{
            RoleName           = $name
            RoleType           = if ($currNorm.IsBuiltIn -eq "True") { "Built-in" } else { "Custom" }
            Status             = $status
            DefinitionChanged  = if ($definitionChanged) { "Yes" } else { "No" }
            AssignmentsChanged = if ($assignmentsChanged) { "Yes" } else { "No" }
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
                        "Changed" { 0 }
                        "New"     { 1 }
                        "Removed" { 2 }
                        "Equal"   { 3 }
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

    if ($StatusFilter -and $StatusFilter -ne "All") {
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
        return "Status: no roles listed."
    }

    $total = $Data.Count
    $iguais = @($Data | Where-Object { $_.Status -eq "Equal" }).Count
    $alteradas = @($Data | Where-Object { $_.Status -eq "Changed" }).Count
    $novas = @($Data | Where-Object { $_.Status -eq "New" }).Count
    $removidas = @($Data | Where-Object { $_.Status -eq "Removed" }).Count

    return "Status: comparison completed. Total: $total | Equal: $iguais | Changed: $alteradas | New: $novas | Removed: $removidas"
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

    if ($SelectedItem.Status -eq "Equal") {
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
        $TxtStatus.Text = "Status: fill in the required fields on the TENANTS tab."
        $fields = [string]::Join(", ", $missingFields.ToArray())
        throw "Fill in the required fields on the TENANTS tab: $fields."
    }
}

function Run-ComparisonWorkflow {
    Ensure-ConnectionInputs

    if (-not $TxtSnapshotFolder.Text) {
        throw "Select a snapshot folder first."
    }

    Set-UiState -StatusText "Status: preparing comparison..." -Busy $true

    try {
        Write-Log "Loading snapshot"
        $script:SnapshotData = Load-Snapshot -Folder $TxtSnapshotFolder.Text

        Connect-AppOnlyWithRetry -TenantId $TxtTenantId.Text -ClientId $TxtClientId.Text -Thumbprint $TxtThumbprint.Text -Operation "Comparison" -StatusCallback {
            param($text, $remainingSeconds)
            if ($remainingSeconds -gt 0) {
                $TxtReplicationStatus.Text = "$text Please wait $remainingSeconds s..."
            }
            else {
                $TxtReplicationStatus.Text = $text
            }
        }

        Write-Log "Reading current tenant"
        $script:CurrentData = Load-CurrentTenantData -TenantId $TxtTenantId.Text -ClientId $TxtClientId.Text -Thumbprint $TxtThumbprint.Text

        Write-Log "Comparing backup with current tenant"
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
        Set-UiState -ReplicationText "App Registration replication: ready for next workflow." -Busy $false
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
        throw "Restore script not found at: $RestoreScriptPath"
    }

    Write-Log "Calling external restore for role '$RoleName'"

    $restoreOutput = & powershell.exe `
        -ExecutionPolicy Bypass `
        -File $RestoreScriptPath `
        -TenantId $TenantId `
        -ClientId $ClientId `
        -CertificateThumbprint $Thumbprint `
        -SnapshotFolder $SnapshotFolder `
        -RoleName $RoleName 2>&1 | Out-String

    $exitCode = $LASTEXITCODE
    Write-Log "External restore output:"
    Write-Log $restoreOutput

    if ($exitCode -ne 0) {
        throw "The restore script returned exit code $exitCode. Check the log."
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

$TxtTenantNotes.Text = "Security: every compare/restore operation uses Service Principal + Certificate (app-only)." +
"`r`n`r`nInteractive sign-in: for setup and initial validation, the provided tenant is required and the connection enforces Connect-MgGraph -TenantId." +
"`r`n`r`nImportant: run in non-elevated PowerShell so the Microsoft WAM popup can work."

if (Test-IsElevated) {
    $TxtStatus.Text = "Status: elevated execution detected. Reopen without administrator privileges for interactive authentication."
    [System.Windows.MessageBox]::Show(
        "Administrator mode detected. The Microsoft interactive popup (WAM) can fail. Reopen in regular PowerShell.",
        "Execution warning",
        [System.Windows.MessageBoxButton]::OK,
        [System.Windows.MessageBoxImage]::Warning
    )
}

$loadBackupsAction = {
    $snapshots = Get-BackupSnapshots
    $GridBackups.ItemsSource = $null
    $GridBackups.ItemsSource = $snapshots

    if ($snapshots.Count -eq 0) {
        $TxtStatus.Text = "Status: no backup found in $BackupExportsPath"
    }
    else {
        $TxtStatus.Text = "Status: $($snapshots.Count) backup(s) loaded."
    }

    Write-Log "Backups reloaded from fixed folder: $BackupExportsPath"
    Add-TaskEntry -Task "Load Backups" -Status "Completed" -Details "$($snapshots.Count) snapshot(s) loaded"
}

$BtnLoadBackups.Add_Click({
    try {
        & $loadBackupsAction
        $MainTabs.SelectedIndex = 0
    }
    catch {
        [System.Windows.MessageBox]::Show($_.Exception.Message, "Error")
        Write-Log ("Error loading backups: " + $_.Exception.Message) -Severity "ERROR"
        Add-TaskEntry -Task "Load Backups" -Status "Failed" -Details $_.Exception.Message
    }
})

$BtnRefreshBackups.Add_Click({
    try {
        & $loadBackupsAction
    }
    catch {
        [System.Windows.MessageBox]::Show($_.Exception.Message, "Error")
        Write-Log ("Error refreshing backups: " + $_.Exception.Message) -Severity "ERROR"
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
        $TxtStatus.Text = "Status: backup selected ($($selected.Timestamp))."
        $MainTabs.SelectedIndex = 1
        Write-Log "Backup selected from grid: $($selected.FullName)"
        Add-TaskEntry -Task "Select Snapshot" -Status "Completed" -Details $selected.FullName
    }
})

$BtnBrowseSnapshot.Add_Click({
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = "Select snapshot folder"

    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $TxtSnapshotFolder.Text = $dialog.SelectedPath
        $TxtStatus.Text = "Status: snapshot selected manually."
        Write-Log "Snapshot selected manually: $($dialog.SelectedPath)"
        Add-TaskEntry -Task "Select Snapshot" -Status "Completed" -Details $dialog.SelectedPath
    }
})

$BtnConnect.Add_Click({
    try {
        Ensure-ConnectionInputs
        Set-UiState -StatusText "Status: authenticating operator in the provided tenant..." -Busy $true

        Connect-OperatorTenantSession -TenantId $TxtTenantId.Text

        Set-UiState -StatusText "Status: validating app-only connection with automatic retry..." -Busy $true
        Connect-AppOnlyWithRetry -TenantId $TxtTenantId.Text -ClientId $TxtClientId.Text -Thumbprint $TxtThumbprint.Text -Operation "Connection validation" -StatusCallback {
            param($text, $remainingSeconds)
            if ($remainingSeconds -gt 0) {
                $TxtReplicationStatus.Text = "$text Please wait $remainingSeconds s..."
            }
            else {
                $TxtReplicationStatus.Text = $text
            }
        }

        $TxtStatus.Text = "Status: authentication and validation completed in tenant $($TxtTenantId.Text)."
        Write-Log "Connection validated successfully"
        Add-TaskEntry -Task "Connect Tenant" -Status "Completed" -Details "Tenant $($TxtTenantId.Text)"
    }
    catch {
        $TxtStatus.Text = "Status: authentication/connection error"
        [System.Windows.MessageBox]::Show($_.Exception.Message, "Error")
        Write-Log ("Authentication/connection error: " + $_.Exception.Message) -Severity "ERROR"
        Add-TaskEntry -Task "Connect Tenant" -Status "Failed" -Details $_.Exception.Message
    }
    finally {
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        Set-UiState -ReplicationText "App Registration replication: ready for next workflow." -Busy $false
    }
})

$BtnCompare.Add_Click({
    try {
        Run-ComparisonWorkflow
        Write-Log "Comparison completed successfully"
        Add-TaskEntry -Task "Run Compare" -Status "Completed" -Details "$($script:FilteredResults.Count) item(s) after filter"
    }
    catch {
        $TxtStatus.Text = "Status: comparison error"
        [System.Windows.MessageBox]::Show($_.Exception.Message, "Error")
        Write-Log ("Comparison error: " + $_.Exception.Message) -Severity "ERROR"
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
        Write-Log "Filter applied successfully"
        Add-TaskEntry -Task "Apply Filter" -Status "Completed" -Details "$($script:FilteredResults.Count) item(s) filtered"
    }
    catch {
        [System.Windows.MessageBox]::Show($_.Exception.Message, "Error")
        Write-Log ("Error applying filter: " + $_.Exception.Message) -Severity "ERROR"
        Add-TaskEntry -Task "Apply Filter" -Status "Failed" -Details $_.Exception.Message
    }
})

$GridRoles.Add_SelectionChanged({
    $item = $GridRoles.SelectedItem

    if ($null -ne $item) {
        $TxtBackupDetails.Text = Format-RoleDetails -RoleObject $item.BackupObject -Assignments $item.BackupAssignments -Label "BACKUP"
        $TxtCurrentDetails.Text = Format-RoleDetails -RoleObject $item.CurrentObject -Assignments $item.CurrentAssignments -Label "CURRENT"
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
            throw "Select a role before restoring."
        }

        if ($item.RoleType -eq "Built-in") {
            throw "Restore through this workflow is supported only for custom roles."
        }

        if ($item.Status -eq "Equal") {
            throw "The selected role is already equal to the snapshot."
        }

        if (-not $TxtSnapshotFolder.Text) {
            throw "Select a snapshot folder first."
        }

        $confirmationMessage = "Do you want to restore the role below using the selected snapshot?`r`n`r`nRole: $($item.RoleName)`r`nCurrent status: $($item.Status)`r`nSnapshot: $($TxtSnapshotFolder.Text)"
        $confirmation = [System.Windows.MessageBox]::Show(
            $confirmationMessage,
            "Confirm restore",
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Warning
        )

        if ($confirmation -ne [System.Windows.MessageBoxResult]::Yes) {
            Write-Log "Restore cancelled by operator for role '$($item.RoleName)'"
            return
        }

        Set-UiState -StatusText "Status: running restore for '$($item.RoleName)'..." -Busy $true
        Write-Log "Starting external restore for role '$($item.RoleName)'"

        Invoke-ExternalRestore -TenantId $TxtTenantId.Text -ClientId $TxtClientId.Text -Thumbprint $TxtThumbprint.Text -RoleName $item.RoleName -SnapshotFolder $TxtSnapshotFolder.Text

        $TxtStatus.Text = "Status: restore completed. Run a new compare to validate the result."
        [System.Windows.MessageBox]::Show("Restore completed successfully for role '$($item.RoleName)'. Run Compare to refresh the differences view.", "Restore completed")
        Write-Log "External restore completed for role '$($item.RoleName)'"
        Add-TaskEntry -Task "Restore Selected" -Status "Completed" -Details "Role $($item.RoleName)"
    }
    catch {
        $TxtStatus.Text = "Status: restore error"
        [System.Windows.MessageBox]::Show($_.Exception.Message, "Restore error")
        Write-Log ("Restore error: " + $_.Exception.Message) -Severity "ERROR"
        Add-TaskEntry -Task "Restore Selected" -Status "Failed" -Details $_.Exception.Message
    }
    finally {
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        Set-UiState -ReplicationText "App Registration replication: ready for next workflow." -Busy $false
    }
})

$BtnExport.Add_Click({
    try {
        if (-not $script:FilteredResults -or $script:FilteredResults.Count -eq 0) {
            throw "There are no results to export."
        }

        $paths = Export-Diff -Data $script:FilteredResults -BasePath $BasePath
        $TxtStatus.Text = "Status: diff exported successfully"
        [System.Windows.MessageBox]::Show("Files generated:`r`n$($paths.JsonPath)`r`n$($paths.CsvPath)", "Export completed")
        Write-Log "Diff exported successfully"
        Add-TaskEntry -Task "Export Diff" -Status "Completed" -Details $paths.CsvPath
    }
    catch {
        $TxtStatus.Text = "Status: export error"
        [System.Windows.MessageBox]::Show($_.Exception.Message, "Error")
        Write-Log ("Export error: " + $_.Exception.Message) -Severity "ERROR"
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

Write-Log "Panel started. Monitored backup directory: $BackupExportsPath"
Add-TaskEntry -Task "Start Panel" -Status "Completed" -Details "Application started"

& $loadBackupsAction
$window.ShowDialog() | Out-Null
