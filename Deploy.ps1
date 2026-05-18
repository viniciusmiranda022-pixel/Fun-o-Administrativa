[CmdletBinding()]
param(
    [string]$SourcePath = (Join-Path $PSScriptRoot 'IR-AdministrativeFunctionBackup'),
    [string]$InstallRoot = 'C:\ProgramData\Quest\IR-AdministrativeFunctionBackup'
)

$ErrorActionPreference = 'Stop'

$installer = Join-Path $PSScriptRoot 'Install.ps1'
if (-not (Test-Path $installer)) {
    throw "Install.ps1 não encontrado em: $installer"
}

& $installer -SourcePath $SourcePath -InstallRoot $InstallRoot
