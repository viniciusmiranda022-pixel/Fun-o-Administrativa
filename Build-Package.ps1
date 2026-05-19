[CmdletBinding()]
param(
    [string]$SourceRoot = $PSScriptRoot,
    [string]$OutputRoot = (Join-Path $PSScriptRoot 'dist')
)

$ErrorActionPreference = 'Stop'

$packageFolders = @(
    'IR-AdministrativeFunctionBackup',
    'IR-AdministrativeFunctionCompare',
    'IR-AdministrativeFunctionRestore'
)

$requiredFiles = @(
    'Install.ps1',
    'IR-AdministrativeFunctionBackup/Scripts/Run-AdministrativeFunctionsBackup-RMAD.cmd',
    'IR-AdministrativeFunctionBackup/Scripts/Run-AdministrativeFunctionsBackup.ps1',
    'IR-AdministrativeFunctionCompare/Scripts/Start-AdministrativeFunctionCompare.ps1',
    'IR-AdministrativeFunctionCompare/Xaml/MainWindow.xaml',
    'IR-AdministrativeFunctionRestore/Scripts/Run-AdministrativeFunctionsRestore.ps1',
    'IR-AdministrativeFunctionRestore/Scripts/Restore-CustomRole-Full.ps1'
)

function Write-Step { param([string]$Message) Write-Host "[PACKAGE] $Message" -ForegroundColor Cyan }
function Write-Ok { param([string]$Message) Write-Host "[OK] $Message" -ForegroundColor Green }

function Assert-PathExists {
    param([Parameter(Mandatory=$true)][string]$RelativePath)
    $fullPath = Join-Path $SourceRoot $RelativePath
    if (-not (Test-Path $fullPath)) {
        throw "Caminho obrigatório ausente no repositório: $RelativePath"
    }
}

function Copy-DirectoryContent {
    param(
        [Parameter(Mandatory=$true)][string]$Source,
        [Parameter(Mandatory=$true)][string]$Destination
    )

    if (-not (Test-Path $Destination)) {
        New-Item -Path $Destination -ItemType Directory -Force | Out-Null
    }

    foreach ($item in Get-ChildItem -Path $Source -Force) {
        $target = Join-Path $Destination $item.Name
        Copy-Item -Path $item.FullName -Destination $target -Recurse -Force
    }
}

Write-Step 'Validando estrutura mínima do pacote...'
foreach ($folder in $packageFolders) { Assert-PathExists -RelativePath $folder }
foreach ($file in $requiredFiles) { Assert-PathExists -RelativePath $file }
Write-Ok 'Estrutura mínima validada.'

Write-Step "Gerando pacote em: $OutputRoot"
if (Test-Path $OutputRoot) {
    Remove-Item -Path $OutputRoot -Recurse -Force
}
New-Item -Path $OutputRoot -ItemType Directory -Force | Out-Null

Copy-Item -Path (Join-Path $SourceRoot 'Install.ps1') -Destination (Join-Path $OutputRoot 'Install.ps1') -Force
if (Test-Path (Join-Path $SourceRoot 'Deploy.ps1')) {
    Copy-Item -Path (Join-Path $SourceRoot 'Deploy.ps1') -Destination (Join-Path $OutputRoot 'Deploy.ps1') -Force
}

foreach ($folder in $packageFolders) {
    $source = Join-Path $SourceRoot $folder
    $destination = Join-Path $OutputRoot $folder
    Copy-DirectoryContent -Source $source -Destination $destination
}

Write-Ok 'Pacote reproduzível gerado com sucesso.'
Write-Host "Use o conteúdo de '$OutputRoot' como raiz do pacote de instalação." -ForegroundColor Yellow
