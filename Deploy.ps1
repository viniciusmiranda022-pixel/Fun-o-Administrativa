[CmdletBinding()]
param(
    [string]$SourcePath = (Join-Path $PSScriptRoot 'IR-AdministrativeFunctionBackup'),
    [string]$InstallRoot = (Join-Path 'C:\ProgramData\Quest' 'IR-AdministrativeFunctionBackup')
)

$ProgramDataQuestRoot = 'C:\ProgramData\Quest'

$ErrorActionPreference = 'Stop'

$installer = Join-Path $PSScriptRoot 'Install.ps1'
if (-not (Test-Path $installer)) {
    throw "Install.ps1 not found at: $installer"
}

& $installer -PackageRoot $SourcePath -ProgramDataRoot $InstallRoot
