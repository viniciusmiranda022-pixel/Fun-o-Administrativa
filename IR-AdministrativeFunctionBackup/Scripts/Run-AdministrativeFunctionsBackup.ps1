[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

try {
    chcp 65001 | Out-Null
} catch {}

$ErrorActionPreference = 'Stop'

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$basePath = Split-Path -Parent $scriptRoot
$configPath = Join-Path $basePath 'Config\settings.json'
Write-Host "Executando scripts em: $basePath" -ForegroundColor Cyan
$setupScript = Join-Path $scriptRoot 'Setup-AdministrativeFunctionsBackup.ps1'
$exportScript = Join-Path $scriptRoot 'Export-AdministrativeFunctions.ps1'


function Ensure-MicrosoftGraphPowerShell {
    $requiredCommands = @('Connect-MgGraph','Disconnect-MgGraph','Get-MgApplication','New-MgApplication','Update-MgApplication','Get-MgServicePrincipal','New-MgServicePrincipal','Get-MgRoleManagementDirectoryRoleDefinition')
    $missingCommands = @($requiredCommands | Where-Object { -not (Get-Command $_ -ErrorAction SilentlyContinue) })
    if ($missingCommands.Count -eq 0) { return }

    Write-Host "Microsoft Graph PowerShell não está disponível. Instalando módulo Microsoft.Graph automaticamente..." -ForegroundColor Yellow
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    } catch {}

    try {
        if (-not (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue)) {
            Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser | Out-Null
        }
    } catch {
        Write-Host "WARNING: não foi possível preparar o provider NuGet automaticamente: $($_.Exception.Message)" -ForegroundColor Yellow
    }

    try {
        Install-Module Microsoft.Graph -Scope AllUsers -Force -AllowClobber -Repository PSGallery
    } catch {
        Write-Host "Instalação em AllUsers falhou; tentando CurrentUser. Detalhes: $($_.Exception.Message)" -ForegroundColor Yellow
        Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber -Repository PSGallery
    }

    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    Import-Module Microsoft.Graph.Applications -ErrorAction Stop
    Import-Module Microsoft.Graph.Identity.DirectoryManagement -ErrorAction SilentlyContinue
    Import-Module Microsoft.Graph.Identity.Governance -ErrorAction SilentlyContinue
}

function Initialize-BackupDirectories {
    param([string]$RootPath)

    foreach ($folder in @('Config','Logs','Backups','Certificates','Assets')) {
        $path = Join-Path $RootPath $folder
        if (-not (Test-Path $path)) {
            New-Item -ItemType Directory -Path $path -Force | Out-Null
        }
    }
}

function Test-BackupPrerequisites {
    param([string]$SettingsFile)

    $reasons = New-Object System.Collections.Generic.List[string]
    $settings = $null

    if (-not (Test-Path $SettingsFile)) {
        $reasons.Add("Arquivo de configuração ausente: $SettingsFile")
        return [pscustomobject]@{ IsValid = $false; Reasons = $reasons; Settings = $null }
    }

    try {
        $settings = Get-Content -Path $SettingsFile -Raw | ConvertFrom-Json
    }
    catch {
        $reasons.Add("Não foi possível ler o JSON de configuração: $($_.Exception.Message)")
        return [pscustomobject]@{ IsValid = $false; Reasons = $reasons; Settings = $null }
    }

    foreach ($required in @('TenantId','ClientId','CertificateThumbprint')) {
        if ([string]::IsNullOrWhiteSpace($settings.$required)) {
            $reasons.Add("Campo obrigatório ausente/vazio: $required")
        }
    }

    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
        $reasons.Add('Módulo Microsoft.Graph.Authentication não está instalado.')
    }

    if (-not (Get-Command Connect-MgGraph -ErrorAction SilentlyContinue)) {
        $reasons.Add('Cmdlet Connect-MgGraph não está disponível.')
    }

    if (-not (Get-Command Disconnect-MgGraph -ErrorAction SilentlyContinue)) {
        $reasons.Add('Cmdlet Disconnect-MgGraph não está disponível.')
    }

    if (-not [string]::IsNullOrWhiteSpace($settings.CertificateThumbprint)) {
        $cert = Get-Item "Cert:\LocalMachine\My\$($settings.CertificateThumbprint)" -ErrorAction SilentlyContinue
        if (-not $cert) {
            $reasons.Add('Certificado configurado não existe em Cert:\LocalMachine\My. Use o assistente para recriar automaticamente ou importar um PFX com chave privada.')
        }
        else {
            if (-not $cert.HasPrivateKey) { $reasons.Add('Certificado sem chave privada (HasPrivateKey=False).') }
            if ($cert.NotAfter -le (Get-Date)) { $reasons.Add("Certificado expirado em $($cert.NotAfter.ToString('yyyy-MM-dd HH:mm:ss')).") }
        }
    }

    return [pscustomobject]@{ IsValid = ($reasons.Count -eq 0); Reasons = $reasons; Settings = $settings }
}

function Test-GraphConnection {
    param(
        [string]$TenantId,
        [string]$ClientId,
        [string]$Thumbprint,
        [int]$MaxAttempts = 8,
        [int]$DelaySeconds = 15
    )

    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        try {
            Write-Host "Validando conexão app-only com Microsoft Graph (tentativa $attempt de $MaxAttempts)..." -ForegroundColor Cyan
            Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $Thumbprint -NoWelcome -ContextScope Process | Out-Null
            Get-MgRoleManagementDirectoryRoleDefinition -All | Select-Object -First 1 | Out-Null
            Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
            return
        } catch {
            $lastError = $_.Exception.Message
            Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
            if ($attempt -eq $MaxAttempts) {
                throw "Falha ao validar app-only após $MaxAttempts tentativas. Último erro: $lastError"
            }
            Write-Host "Validação app-only ainda não disponível: $lastError" -ForegroundColor Yellow
            Write-Host "Aguardando $DelaySeconds segundos para possível propagação no Microsoft Entra ID..." -ForegroundColor Yellow
            Start-Sleep -Seconds $DelaySeconds
        }
    }
}

Initialize-BackupDirectories -RootPath $basePath
Ensure-MicrosoftGraphPowerShell

$validation = Test-BackupPrerequisites -SettingsFile $configPath
if (-not $validation.IsValid) {
    Write-Host 'Pré-validação falhou. Abrindo assistente de configuração inicial...' -ForegroundColor Yellow
    $validation.Reasons | ForEach-Object { Write-Host " - $_" -ForegroundColor Yellow }

    if (-not (Test-Path $setupScript)) {
        throw "Setup script não encontrado em: $setupScript"
    }

    & $setupScript -SettingsPath $configPath

    $validation = Test-BackupPrerequisites -SettingsFile $configPath
    if (-not $validation.IsValid) {
        throw "A configuração inicial não concluiu com sucesso. Pendências: $($validation.Reasons -join '; ')"
    }
}

Test-GraphConnection -TenantId $validation.Settings.TenantId -ClientId $validation.Settings.ClientId -Thumbprint $validation.Settings.CertificateThumbprint

& $exportScript -TenantId $validation.Settings.TenantId -ClientId $validation.Settings.ClientId -CertificateThumbprint $validation.Settings.CertificateThumbprint -BasePath $basePath
