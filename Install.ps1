[CmdletBinding()]
param(
    [string]$PackageRoot = $PSScriptRoot,
    [string]$ProgramDataRoot = 'C:\ProgramData\Quest'
)

$ErrorActionPreference = 'Stop'

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

try {
    chcp 65001 | Out-Null
} catch {}

$BackupRoot = Join-Path $ProgramDataRoot 'IR-AdministrativeFunctionBackup'
$CompareRoot = Join-Path $ProgramDataRoot 'IR-AdministrativeFunctionCompare'
$RestoreRoot = Join-Path $ProgramDataRoot 'IR-AdministrativeFunctionRestore'

$BackupPackage = Join-Path $PackageRoot 'IR-AdministrativeFunctionBackup'
$ComparePackage = Join-Path $PackageRoot 'IR-AdministrativeFunctionCompare'
$RestorePackage = Join-Path $PackageRoot 'IR-AdministrativeFunctionRestore'

function Write-Step { param([string]$Message) Write-Host "[INSTALL] $Message" -ForegroundColor Cyan }
function Write-Ok { param([string]$Message) Write-Host "[OK] $Message" -ForegroundColor Green }

function Assert-Admin {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        throw 'This installer must be run as a local administrator.'
    }
}

function Assert-PowerShell51 {
    if ($PSVersionTable.PSVersion.Major -ne 5 -or $PSVersionTable.PSVersion.Minor -lt 1) {
        throw "PowerShell 5.1 is required. Current version: $($PSVersionTable.PSVersion)"
    }
}

function Ensure-Directory {
    param([Parameter(Mandatory=$true)][string]$Path)
    if (-not (Test-Path $Path)) { New-Item -Path $Path -ItemType Directory -Force | Out-Null }
}

function Ensure-OperationalDirectories {
    foreach ($path in @(
        (Join-Path $BackupRoot 'Config'),
        (Join-Path $BackupRoot 'Logs'),
        (Join-Path $BackupRoot 'Backups'),
        (Join-Path $BackupRoot 'Certificates'),
        (Join-Path $BackupRoot 'Assets'),
        (Join-Path $BackupRoot 'Scripts'),
        (Join-Path $CompareRoot 'Logs'),
        (Join-Path $CompareRoot 'Scripts'),
        (Join-Path $CompareRoot 'Xaml'),
        (Join-Path $RestoreRoot 'Logs'),
        (Join-Path $RestoreRoot 'Scripts')
    )) { Ensure-Directory -Path $path }
}

function Install-GraphPrerequisites {
    Write-Step 'Installing Microsoft.Graph prerequisites...'
    Set-ExecutionPolicy Bypass -Scope Process -Force
    Install-PackageProvider NuGet -Force | Out-Null
    Install-Module Microsoft.Graph -MinimumVersion "2.0.0" -Scope AllUsers -Force -AllowClobber
    Import-Module Microsoft.Graph.Authentication -Force

    Get-Command Connect-MgGraph -ErrorAction Stop | Out-Null
    Get-Command Disconnect-MgGraph -ErrorAction Stop | Out-Null
    Get-Module Microsoft.Graph -ListAvailable -ErrorAction Stop | Out-Null
    Get-Module Microsoft.Graph.Authentication -ListAvailable -ErrorAction Stop | Out-Null
    Write-Ok 'Microsoft.Graph modules validated.'
}

function Copy-PackageContent {
    param(
        [string]$Source,
        [string]$Target
    )

    if (-not (Test-Path $Source)) { throw "Source directory not found: $Source" }

    foreach ($item in Get-ChildItem -Path $Source -Force) {
        $destination = Join-Path $Target $item.Name
        Copy-Item -Path $item.FullName -Destination $destination -Recurse -Force
    }
}

function Assert-WriteAccess {
    Ensure-Directory -Path $ProgramDataRoot
    $probe = Join-Path $ProgramDataRoot ("write-test-{0}.tmp" -f ([guid]::NewGuid().Guid))
    'ok' | Set-Content -Path $probe -Encoding UTF8
    Remove-Item -Path $probe -Force
}

function Run-PostInstallValidations {
    $rmadWrapper = Join-Path $BackupRoot 'Scripts\Run-AdministrativeFunctionsBackup-RMAD.cmd'
    $backupScript = Join-Path $BackupRoot 'Scripts\Run-AdministrativeFunctionsBackup.ps1'
    $compareScript = Join-Path $CompareRoot 'Scripts\Start-AdministrativeFunctionCompare.ps1'
    $restoreScript = Join-Path $RestoreRoot 'Scripts\Run-AdministrativeFunctionsRestore.ps1'
    $restoreRoleScript = Join-Path $RestoreRoot 'Scripts\Restore-CustomRole-Full.ps1'
    $compareXaml = Join-Path $CompareRoot 'Xaml\MainWindow.xaml'

    foreach ($required in @($rmadWrapper,$backupScript,$compareScript,$restoreScript,$restoreRoleScript,$compareXaml)) {
        if (-not (Test-Path $required)) { throw "Required file not found: $required" }
    }

    $settingsPath = Join-Path $BackupRoot 'Config\settings.json'
    if (Test-Path $settingsPath) {
        try {
            $settings = Get-Content $settingsPath -Raw | ConvertFrom-Json
            if ($settings.BackupRoot -and -not (Test-Path $settings.BackupRoot)) {
                throw "Configured backup path does not exist: $($settings.BackupRoot)"
            }
        }
        catch {
            throw "Failed to validate settings.json: $($_.Exception.Message)"
        }
    }

    Write-Ok 'Post-installation validations completed.'
}

Write-Step 'Validating environment...'
Assert-PowerShell51
Assert-Admin
Assert-WriteAccess
Ensure-OperationalDirectories
Install-GraphPrerequisites

Write-Step 'Copying package files (preserving operational data)...'
Copy-PackageContent -Source $BackupPackage -Target $BackupRoot
Copy-PackageContent -Source $ComparePackage -Target $CompareRoot
Copy-PackageContent -Source $RestorePackage -Target $RestoreRoot
Ensure-OperationalDirectories

Run-PostInstallValidations

Write-Ok 'Installation completed successfully.'
Write-Host "Backup UI: powershell -ExecutionPolicy Bypass -STA -File `"$(Join-Path $BackupRoot 'Scripts\Run-AdministrativeFunctionsBackup.ps1')`"" -ForegroundColor Green
Write-Host "Compare UI: powershell -ExecutionPolicy Bypass -STA -File `"$(Join-Path $CompareRoot 'Scripts\Start-AdministrativeFunctionCompare.ps1')`"" -ForegroundColor Green
Write-Host "Restore UI: powershell -ExecutionPolicy Bypass -STA -File `"$(Join-Path $RestoreRoot 'Scripts\Run-AdministrativeFunctionsRestore.ps1')`"" -ForegroundColor Green
