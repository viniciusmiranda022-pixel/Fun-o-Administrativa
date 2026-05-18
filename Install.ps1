[CmdletBinding()]
param(
    [string]$SourcePath = (Join-Path $PSScriptRoot 'IR-AdministrativeFunctionBackup'),
    [string]$InstallRoot = 'C:\ProgramData\Quest\IR-AdministrativeFunctionBackup'
)

$ErrorActionPreference = 'Stop'

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

try {
    chcp 65001 | Out-Null
} catch {}

function Initialize-InstallDirectories {
    param([string]$RootPath)

    foreach ($folder in @('Config','Logs','Backups','Certificates','Assets','Scripts')) {
        $path = Join-Path $RootPath $folder
        if (-not (Test-Path $path)) {
            New-Item -ItemType Directory -Path $path -Force | Out-Null
        }
    }
}

if (-not (Test-Path $SourcePath)) {
    throw "Diretório de origem não encontrado: $SourcePath"
}

$resolvedSource = (Resolve-Path $SourcePath).Path
$resolvedTargetParent = Split-Path -Parent $InstallRoot
New-Item -ItemType Directory -Path $resolvedTargetParent -Force | Out-Null
Initialize-InstallDirectories -RootPath $InstallRoot

Write-Host "Instalando IR-AdministrativeFunctionBackup de '$resolvedSource' para '$InstallRoot'..." -ForegroundColor Cyan

foreach ($item in Get-ChildItem -Path $resolvedSource -Force) {
    $destination = Join-Path $InstallRoot $item.Name
    Copy-Item -Path $item.FullName -Destination $destination -Recurse -Force
}

Initialize-InstallDirectories -RootPath $InstallRoot

$runScript = Join-Path $InstallRoot 'Scripts\Run-AdministrativeFunctionsBackup.ps1'
if (-not (Test-Path $runScript)) {
    throw "Instalação incompleta: script de execução não encontrado em $runScript"
}

$rmadWrapper = Join-Path $InstallRoot 'Scripts\Run-AdministrativeFunctionsBackup-RMAD.cmd'
if (-not (Test-Path $rmadWrapper)) {
    throw "Instalação incompleta: wrapper RMAD não encontrado em $rmadWrapper"
}

Write-Host 'Instalação concluída.' -ForegroundColor Green
Write-Host "Execute: powershell -ExecutionPolicy Bypass -STA -File `"$runScript`"" -ForegroundColor Green
Write-Host "RMAD Custom Script: $rmadWrapper" -ForegroundColor Green
