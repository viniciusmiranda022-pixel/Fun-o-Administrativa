[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$RoleName,

    [Parameter(Mandatory = $false)]
    [string]$SnapshotFolder
)

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

try {
    chcp 65001 | Out-Null
} catch {}

$ProgramDataQuestRoot = "C:\ProgramData\Quest"
$BackupInstallRoot = Join-Path $ProgramDataQuestRoot "IR-AdministrativeFunctionBackup"
$RestoreInstallRoot = Join-Path $ProgramDataQuestRoot "IR-AdministrativeFunctionRestore"

$SettingsPath = Join-Path $BackupInstallRoot "Config\settings.json"
$DefaultBackupRoot = Join-Path $BackupInstallRoot "Backups"
$RestoreScriptPath = Join-Path $RestoreInstallRoot "Scripts\Restore-CustomRole-Full.ps1"

function Get-ValidBackupFolder {
    param(
        [string]$BackupRoot
    )

    if (-not (Test-Path $BackupRoot)) { return $null }

    $requiredFiles = @('manifest.json','roleDefinitions.json','roleAssignments.json')
    $folders = @(Get-ChildItem -Path $BackupRoot -Directory -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending)

    foreach ($folder in $folders) {
        $missing = @($requiredFiles | Where-Object { -not (Test-Path (Join-Path $folder.FullName $_)) })
        if ($missing.Count -eq 0) { return $folder.FullName }
    }

    return $null
}

if (-not (Test-Path $SettingsPath)) {
    throw "settings.json not found at $SettingsPath. Run backup setup first."
}

$settings = Get-Content $SettingsPath -Raw | ConvertFrom-Json
$tenantId = [string]$settings.TenantId
$clientId = [string]$settings.ClientId
$thumb = [string]$settings.CertificateThumbprint
$backupRoot = if (-not [string]::IsNullOrWhiteSpace([string]$settings.BackupRoot)) { [string]$settings.BackupRoot } else { $DefaultBackupRoot }

if ([string]::IsNullOrWhiteSpace($tenantId)) { throw "TenantId vazio. Verifique settings.json." }
if ([string]::IsNullOrWhiteSpace($clientId)) { throw "ClientId vazio. Verifique settings.json." }
if ([string]::IsNullOrWhiteSpace($thumb)) { throw "CertificateThumbprint vazio. Verifique settings.json." }

if ([string]::IsNullOrWhiteSpace($SnapshotFolder)) {
    if (-not (Test-Path $backupRoot)) {
        throw "BackupRoot does not exist: $backupRoot"
    }

    $SnapshotFolder = Get-ValidBackupFolder -BackupRoot $backupRoot

    if (-not $SnapshotFolder) {
        throw "No valid backup snapshots found under $backupRoot (missing manifest.json, roleDefinitions.json, or roleAssignments.json)."
    }
}

if (-not (Test-Path $RestoreScriptPath)) {
    throw "Restore script not found at $RestoreScriptPath"
}

& powershell.exe -ExecutionPolicy Bypass -File $RestoreScriptPath `
    -TenantId $tenantId `
    -ClientId $clientId `
    -CertificateThumbprint $thumb `
    -SnapshotFolder $SnapshotFolder `
    -RoleName $RoleName
