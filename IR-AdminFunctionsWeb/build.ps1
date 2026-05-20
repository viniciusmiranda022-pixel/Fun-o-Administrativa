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

Write-Host "[1/3] Build do frontend (Vite)..." -ForegroundColor Cyan
Push-Location $ui
try {
    if (-not (Test-Path (Join-Path $ui 'node_modules'))) {
        npm install
    }
    npm run build
} finally {
    Pop-Location
}

Write-Host "[2/3] Limpando publish anterior..." -ForegroundColor Cyan
if (Test-Path $OutputDir) {
    Remove-Item -Path $OutputDir -Recurse -Force
}

Write-Host "[3/3] dotnet publish ($Configuration / $Runtime)..." -ForegroundColor Cyan
dotnet publish $webProject `
    -c $Configuration `
    -r $Runtime `
    --self-contained true `
    -p:PublishSingleFile=true `
    -p:IncludeNativeLibrariesForSelfExtract=true `
    -o $OutputDir

Write-Host ""
Write-Host "Pacote gerado em: $OutputDir" -ForegroundColor Green
