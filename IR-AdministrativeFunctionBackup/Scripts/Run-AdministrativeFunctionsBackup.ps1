[CmdletBinding()]
param(
    [switch]$RmadMode,
    [string]$ClientSecret = $env:IR_CLIENT_SECRET,

    # Accepted for compatibility with the default invoker (app's PowerShellRunner).
    # The backup reads the effective TenantId/ClientId/CertificateThumbprint from settings.json,
    # synchronized by the app before execution. Without these declared parameters,
    # [CmdletBinding()] rejects the call and the script aborts at binding without running.
    [string]$TenantId = "",
    [string]$ClientId = "",
    [string]$Thumbprint = ""
)

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

try {
    chcp 65001 | Out-Null
} catch {}

$ErrorActionPreference = 'Stop'

$sharedModulePath = Join-Path $PSScriptRoot 'IR-AdministrativeFunctions.psm1'
Import-Module $sharedModulePath -Force

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$basePath = Split-Path -Parent $scriptRoot
$configPath = Join-Path $basePath 'Config\settings.json'
$setupScript = Join-Path $scriptRoot 'Setup-AdministrativeFunctionsBackup.ps1'
$exportScript = Join-Path $scriptRoot 'Export-AdministrativeFunctions.ps1'



function Test-RmadInvocation {
    if ($RmadMode) { return $true }

    try {
        if (-not [Environment]::UserInteractive) { return $true }
    } catch {}

    try {
        $process = Get-CimInstance Win32_Process -Filter "ProcessId = $PID" -ErrorAction Stop
        for ($i = 0; $i -lt 10 -and $process -and $process.ParentProcessId; $i++) {
            $process = Get-CimInstance Win32_Process -Filter "ProcessId = $($process.ParentProcessId)" -ErrorAction SilentlyContinue
            if (-not $process) { break }

            if ($process.Name -ieq 'BackupServer.exe') { return $true }
            if ($process.CommandLine -and ($process.CommandLine -match 'Run-AdministrativeFunctionsBackup-RMAD\.cmd')) { return $true }
        }
    } catch {}

    return $false
}

function Write-RmadLog {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    if (-not $script:RmadLogFile) { return }
    $line = "[{0}] {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message
    Add-Content -Path $script:RmadLogFile -Value $line -Encoding UTF8
}

function Get-MicrosoftGraphAvailability {
    $modules = @(Get-Module -ListAvailable -Name Microsoft.Graph, Microsoft.Graph.Authentication -ErrorAction SilentlyContinue |
        Sort-Object Name, Version -Descending |
        Select-Object Name, Version, Path)
    $connectCommand = Get-Command Connect-MgGraph -ErrorAction SilentlyContinue
    $disconnectCommand = Get-Command Disconnect-MgGraph -ErrorAction SilentlyContinue

    return [pscustomobject]@{
        ModuleAvailable = ($modules.Count -gt 0)
        Modules = $modules
        ConnectMgGraphAvailable = [bool]$connectCommand
        DisconnectMgGraphAvailable = [bool]$disconnectCommand
    }
}

function Write-RmadDiagnostics {
    param(
        [string]$SettingsFile,
        [object]$Settings = $null
    )

    Write-RmadLog "Running user: $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)"
    Write-RmadLog "PowerShell architecture: $([IntPtr]::Size * 8)-bit"
    Write-RmadLog "PowerShell path: $((Get-Process -Id $PID).Path)"
    Write-RmadLog "settings.json exists: $(Test-Path $SettingsFile) ($SettingsFile)"

    $thumbprint = $null
    if ($Settings -and -not [string]::IsNullOrWhiteSpace($Settings.CertificateThumbprint)) {
        $thumbprint = $Settings.CertificateThumbprint
    }

    if ($thumbprint) {
        $cert = Get-Item "Cert:\LocalMachine\My\$thumbprint" -ErrorAction SilentlyContinue
        if (-not $cert) { $cert = Get-Item "Cert:\CurrentUser\My\$thumbprint" -ErrorAction SilentlyContinue }
        $certLocation = if ($cert) { if ((Get-Item "Cert:\LocalMachine\My\$thumbprint" -ErrorAction SilentlyContinue)) { 'LocalMachine\My' } else { 'CurrentUser\My' } } else { 'n/a' }
        Write-RmadLog "Certificate exists: $([bool]$cert) ($certLocation\$thumbprint)"
        if ($cert) {
            Write-RmadLog "HasPrivateKey: $($cert.HasPrivateKey)"
            Write-RmadLog "Certificate expires on: $($cert.NotAfter.ToString('yyyy-MM-dd HH:mm:ss'))"
        } else {
            Write-RmadLog "HasPrivateKey: n/a"
        }
    } else {
        Write-RmadLog 'Certificate exists: n/a (CertificateThumbprint missing in settings.json)'
        Write-RmadLog 'HasPrivateKey: n/a'
    }

    $graph = Get-MicrosoftGraphAvailability
    Write-RmadLog "Microsoft.Graph available: $($graph.ModuleAvailable)"
    Write-RmadLog "Connect-MgGraph available: $($graph.ConnectMgGraphAvailable)"
    Write-RmadLog "Disconnect-MgGraph available: $($graph.DisconnectMgGraphAvailable)"
    foreach ($module in $graph.Modules) {
        Write-RmadLog "Microsoft.Graph module: $($module.Name) $($module.Version) - $($module.Path)"
    }
}

function Get-ValidBackupFolder {
    param(
        [string]$BackupRoot,
        [datetime]$NotBefore
    )

    if (-not (Test-Path $BackupRoot)) { return $null }

    $requiredFiles = @('manifest.json','roleDefinitions.json','roleAssignments.json')
    $folders = @(Get-ChildItem -Path $BackupRoot -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.CreationTime -ge $NotBefore.AddSeconds(-5) -or $_.LastWriteTime -ge $NotBefore.AddSeconds(-5) } |
        Sort-Object LastWriteTime -Descending)

    foreach ($folder in $folders) {
        $missing = @($requiredFiles | Where-Object { -not (Test-Path (Join-Path $folder.FullName $_)) })
        if ($missing.Count -eq 0) { return $folder.FullName }
    }

    return $null
}

function Invoke-RmadBackup {
    Initialize-BackupDirectories -RootPath $basePath

    $logRoot = Join-Path $basePath 'Logs'
    New-Item -ItemType Directory -Path $logRoot -Force | Out-Null
    $script:RmadLogFile = Join-Path $logRoot ("rmad-custom-script-{0}.log" -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
    New-Item -ItemType File -Path $script:RmadLogFile -Force | Out-Null

    Write-RmadLog 'Starting RMAD/non-interactive execution of administrative functions backup.'
    Write-RmadDiagnostics -SettingsFile $configPath

    try {
        Ensure-MicrosoftGraphPowerShell *>&1 | ForEach-Object { Write-RmadLog "GRAPH-SETUP: $_" }
        Write-RmadLog 'Microsoft.Graph availability after setup:'
        $graph = Get-MicrosoftGraphAvailability
        Write-RmadLog "Microsoft.Graph available: $($graph.ModuleAvailable)"
        Write-RmadLog "Connect-MgGraph available: $($graph.ConnectMgGraphAvailable)"
        Write-RmadLog "Disconnect-MgGraph available: $($graph.DisconnectMgGraphAvailable)"

        $validation = Test-BackupPrerequisites -SettingsFile $configPath
        Write-RmadDiagnostics -SettingsFile $configPath -Settings $validation.Settings

        if (-not $validation.IsValid) {
            Write-RmadLog "Pre-validation failed: $($validation.Reasons -join '; ')"
            Write-RmadLog 'Interactive wizard will not be opened in RMAD mode.'
            exit 1
        }

        $effectiveSecret = if (-not [string]::IsNullOrWhiteSpace($env:IR_CLIENT_SECRET)) { $env:IR_CLIENT_SECRET }
                           elseif (-not [string]::IsNullOrWhiteSpace($ClientSecret)) { $ClientSecret }
                           elseif (-not [string]::IsNullOrWhiteSpace($validation.Settings.ClientSecret)) { $validation.Settings.ClientSecret }
                           else { "" }

        Test-GraphConnection -TenantId $validation.Settings.TenantId -ClientId $validation.Settings.ClientId -Thumbprint "$($validation.Settings.CertificateThumbprint)" -ClientSecret $effectiveSecret -VerboseOutput *>&1 | ForEach-Object { Write-RmadLog "GRAPH-TEST: $_" }

        $backupStart = Get-Date
        Write-RmadLog 'Starting backup export.'
        & $exportScript -TenantId $validation.Settings.TenantId -ClientId $validation.Settings.ClientId -CertificateThumbprint "$($validation.Settings.CertificateThumbprint)" -ClientSecret $effectiveSecret -BasePath $basePath *>&1 | ForEach-Object { Write-RmadLog "EXPORT: $_" }

        $backupFolder = Get-ValidBackupFolder -BackupRoot (Join-Path $basePath 'Backups') -NotBefore $backupStart
        if ($backupFolder) {
            Write-RmadLog "Generated backup path: $backupFolder"
            Write-RmadLog 'RMAD backup completed successfully; manifest.json, roleDefinitions.json and roleAssignments.json were created.'
            exit 0
        }

        Write-RmadLog 'Export finished without error, but the required files were not found in a new backup.'
        exit 1
    }
    catch {
        Write-RmadLog "Actual failure in RMAD export/setup: $($_.Exception.Message)"
        if ($_.ScriptStackTrace) { Write-RmadLog "StackTrace: $($_.ScriptStackTrace)" }
        exit 1
    }
}

function Test-BackupPrerequisites {
    param([string]$SettingsFile)

    $reasons = New-Object System.Collections.Generic.List[string]
    $settings = $null

    if (-not (Test-Path $SettingsFile)) {
        $reasons.Add("Configuration file missing: $SettingsFile")
        return [pscustomobject]@{ IsValid = $false; Reasons = $reasons; Settings = $null }
    }

    try {
        $settings = Get-Content -Path $SettingsFile -Raw | ConvertFrom-Json
    }
    catch {
        $reasons.Add("Unable to read configuration JSON: $($_.Exception.Message)")
        return [pscustomobject]@{ IsValid = $false; Reasons = $reasons; Settings = $null }
    }

    foreach ($required in @('TenantId','ClientId')) {
        if ([string]::IsNullOrWhiteSpace($settings.$required)) {
            $reasons.Add("Required field missing/empty: $required")
        }
    }

    $hasThumbprint = -not [string]::IsNullOrWhiteSpace($settings.CertificateThumbprint)
    $hasSecret = (-not [string]::IsNullOrWhiteSpace($ClientSecret)) -or (-not [string]::IsNullOrWhiteSpace($settings.ClientSecret))
    if (-not $hasThumbprint -and -not $hasSecret) {
        $reasons.Add("CertificateThumbprint or ClientSecret is required to authenticate to Microsoft Graph.")
    }

    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
        $reasons.Add('Module Microsoft.Graph.Authentication is not installed.')
    }

    if (-not (Get-Command Connect-MgGraph -ErrorAction SilentlyContinue)) {
        $reasons.Add('Cmdlet Connect-MgGraph is not available.')
    }

    if (-not (Get-Command Disconnect-MgGraph -ErrorAction SilentlyContinue)) {
        $reasons.Add('Cmdlet Disconnect-MgGraph is not available.')
    }

    if ($hasThumbprint -and -not $hasSecret) {
        $cert = Get-Item "Cert:\LocalMachine\My\$($settings.CertificateThumbprint)" -ErrorAction SilentlyContinue
        if (-not $cert) { $cert = Get-Item "Cert:\CurrentUser\My\$($settings.CertificateThumbprint)" -ErrorAction SilentlyContinue }
        if (-not $cert) {
            $reasons.Add('Configured certificate does not exist in Cert:\LocalMachine\My or Cert:\CurrentUser\My. Use ClientSecret as an alternative or import a PFX with a private key.')
        }
        else {
            if (-not $cert.HasPrivateKey) { $reasons.Add('Certificate has no private key (HasPrivateKey=False).') }
            if ($cert.NotAfter -le (Get-Date)) { $reasons.Add("Certificate expired on $($cert.NotAfter.ToString('yyyy-MM-dd HH:mm:ss')).") }
        }
    }

    return [pscustomobject]@{ IsValid = ($reasons.Count -eq 0); Reasons = $reasons; Settings = $settings }
}

if (Test-RmadInvocation) {
    Invoke-RmadBackup
}

Write-Host "Running scripts at: $basePath" -ForegroundColor Cyan
Initialize-BackupDirectories -RootPath $basePath
Ensure-MicrosoftGraphPowerShell

$validation = Test-BackupPrerequisites -SettingsFile $configPath
if (-not $validation.IsValid) {
    Write-Host 'Pre-validation failed. Opening initial configuration wizard...' -ForegroundColor Yellow
    $validation.Reasons | ForEach-Object { Write-Host " - $_" -ForegroundColor Yellow }

    if (-not (Test-Path $setupScript)) {
        throw "Setup script not found at: $setupScript"
    }

    & $setupScript -SettingsPath $configPath

    $validation = Test-BackupPrerequisites -SettingsFile $configPath
    if (-not $validation.IsValid) {
        throw "Initial configuration did not complete successfully. Pending issues: $($validation.Reasons -join '; ')"
    }
}

Test-GraphConnection -TenantId $validation.Settings.TenantId -ClientId $validation.Settings.ClientId -Thumbprint $validation.Settings.CertificateThumbprint -VerboseOutput

& $exportScript -TenantId $validation.Settings.TenantId -ClientId $validation.Settings.ClientId -CertificateThumbprint $validation.Settings.CertificateThumbprint -BasePath $basePath
