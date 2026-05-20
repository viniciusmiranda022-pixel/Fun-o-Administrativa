[CmdletBinding()]
param(
    [string]$Configuration = 'Release',
    [string]$Runtime = 'win-x64',
    [string]$OutputDir = (Join-Path $PSScriptRoot 'publish')
)

$ErrorActionPreference = 'Stop'

$root = $PSScriptRoot
$ui = Join-Path $root 'ui'
$webProject = Join-Path $root 'src\IR.AdminFunctions.Web\IR.AdminFunctions.Web.csproj'

Write-Host "[1/4] Build do frontend (Vite)..." -ForegroundColor Cyan
Push-Location $ui
try {
    if (-not (Test-Path (Join-Path $ui 'node_modules'))) {
        npm install
    }
    npm run build
} finally {
    Pop-Location
}

Write-Host "[2/4] Limpando publish anterior..." -ForegroundColor Cyan
if (Test-Path $OutputDir) {
    Remove-Item -Path $OutputDir -Recurse -Force
}

Write-Host "[3/4] dotnet publish ($Configuration / $Runtime)..." -ForegroundColor Cyan
dotnet publish $webProject `
    -c $Configuration `
    -r $Runtime `
    --self-contained true `
    -p:PublishSingleFile=true `
    -p:IncludeNativeLibrariesForSelfExtract=true `
    -o $OutputDir

Write-Host "[4/4] Copiando modulos PowerShell para o pacote..." -ForegroundColor Cyan
$psModulesDest = Join-Path $OutputDir 'PowerShellModules'
$repoRoot = (Resolve-Path (Join-Path $root '..')).Path
foreach ($mod in 'IR-AdministrativeFunctionBackup','IR-AdministrativeFunctionCompare','IR-AdministrativeFunctionRestore') {
    $src = Join-Path $repoRoot $mod
    if (Test-Path $src) {
        $dst = Join-Path $psModulesDest $mod
        New-Item -ItemType Directory -Path $dst -Force | Out-Null
        Copy-Item -Path (Join-Path $src '*') -Destination $dst -Recurse -Force
        Write-Host "  $mod copiado" -ForegroundColor Green
    } else {
        Write-Warning "Modulo $mod nao encontrado em $src"
    }
}

Write-Host ""
Write-Host "Pacote gerado em: $OutputDir" -ForegroundColor Green
