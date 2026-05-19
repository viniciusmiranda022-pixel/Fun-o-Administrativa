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

$SettingsPath = "C:\ProgramData\Quest\IR-AdministrativeFunctionBackup\Config\settings.json"
$DefaultBackupRoot = "C:\ProgramData\Quest\IR-AdministrativeFunctionBackup\Backups"
$RestoreScriptPath = "C:\ProgramData\Quest\IR-AdministrativeFunctionRestore\Scripts\Restore-CustomRole-Full.ps1"

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

    $latest = Get-ChildItem -Path $backupRoot -Directory -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1

    if (-not $latest) {
        throw "No backup snapshots found under $backupRoot"
    }

    $SnapshotFolder = $latest.FullName
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
