[CmdletBinding()]
param(
    [string]$ProgramDataRoot = 'C:\ProgramData\Quest',
    [string]$RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path,
    [switch]$Force
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

function Copy-ModuleFiles {
    param(
        [string]$SourceDir,
        [string]$DestinationDir,
        [string]$Label
    )

    if (-not (Test-Path $SourceDir)) {
        Write-Warning "Origem '$Label' não encontrada: $SourceDir (pulando)"
        return
    }

    New-Item -ItemType Directory -Path $DestinationDir -Force | Out-Null
    $files = Get-ChildItem -Path $SourceDir -File
    foreach ($file in $files) {
        $target = Join-Path $DestinationDir $file.Name
        if ((Test-Path $target) -and -not $Force) {
            Write-Host "  [skip] $($file.Name) já existe em $Label (use -Force para sobrescrever)" -ForegroundColor DarkGray
        } else {
            Copy-Item -Path $file.FullName -Destination $target -Force
            Write-Host "  [copy] $($file.Name) → $Label" -ForegroundColor Green
        }
    }
}

Write-Host "Provisionando motor PowerShell em $ProgramDataRoot" -ForegroundColor Cyan
Write-Host "Repo: $RepoRoot" -ForegroundColor DarkGray
Write-Host ""

$modules = @(
    @{
        Name = 'IR-AdministrativeFunctionBackup'
        Source = Join-Path $RepoRoot 'IR-AdministrativeFunctionBackup\Scripts'
        SubFolders = @('Scripts','Config','Logs','Backups')
    },
    @{
        Name = 'IR-AdministrativeFunctionCompare'
        Source = Join-Path $RepoRoot 'IR-AdministrativeFunctionCompare\Scripts'
        SubFolders = @('Scripts','Xaml','Logs','Output')
        Extra = @{
            'Xaml' = Join-Path $RepoRoot 'IR-AdministrativeFunctionCompare\Xaml'
        }
    },
    @{
        Name = 'IR-AdministrativeFunctionRestore'
        Source = Join-Path $RepoRoot 'IR-AdministrativeFunctionRestore\Scripts'
        SubFolders = @('Scripts','Logs')
    }
)

foreach ($mod in $modules) {
    $modRoot = Join-Path $ProgramDataRoot $mod.Name
    Write-Host "[$($mod.Name)]" -ForegroundColor Yellow

    foreach ($sub in $mod.SubFolders) {
        $subPath = Join-Path $modRoot $sub
        New-Item -ItemType Directory -Path $subPath -Force | Out-Null
    }

    Copy-ModuleFiles -SourceDir $mod.Source -DestinationDir (Join-Path $modRoot 'Scripts') -Label "$($mod.Name)\Scripts"

    if ($mod.Extra) {
        foreach ($key in $mod.Extra.Keys) {
            Copy-ModuleFiles -SourceDir $mod.Extra[$key] -DestinationDir (Join-Path $modRoot $key) -Label "$($mod.Name)\$key"
        }
    }
    Write-Host ""
}

$settingsPath = Join-Path $ProgramDataRoot 'IR-AdministrativeFunctionBackup\Config\settings.json'
if ((Test-Path $settingsPath) -and -not $Force) {
    Write-Host "settings.json já existe em $settingsPath (preservado)" -ForegroundColor DarkGray
} else {
    $template = [ordered]@{
        Tenants = @(
            [ordered]@{
                Name = 'PREENCHER-NomeDoTenant'
                TenantId = 'PREENCHER-TenantId'
                ClientId = 'PREENCHER-ClientId'
                Thumbprint = 'PREENCHER-Thumbprint'
            }
        )
        Defaults = [ordered]@{
            RetentionDays = 30
            ParallelDegree = 4
        }
    }
    $template | ConvertTo-Json -Depth 5 | Set-Content -Path $settingsPath -Encoding UTF8
    Write-Host "Template settings.json criado em $settingsPath" -ForegroundColor Green
    Write-Host "  → edite os campos 'PREENCHER-*' com os dados reais do seu Entra ID." -ForegroundColor Yellow
}

Write-Host ""
Write-Host "Provisionamento concluído." -ForegroundColor Green
Write-Host "Próximos passos:" -ForegroundColor Cyan
Write-Host "  1. Edite $settingsPath com TenantId, ClientId e Thumbprint reais." -ForegroundColor White
Write-Host "  2. Importe o certificado correspondente em Cert:\LocalMachine\My\." -ForegroundColor White
Write-Host "  3. Reinicie a aplicação web (Ctrl+C + dotnet run, ou Restart-Service IR-AdminFunctionsWeb)." -ForegroundColor White
