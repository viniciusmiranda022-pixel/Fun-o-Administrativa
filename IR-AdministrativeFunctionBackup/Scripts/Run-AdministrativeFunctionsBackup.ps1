[CmdletBinding()]
param(
    [switch]$RmadMode
)

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

try {
    chcp 65001 | Out-Null
} catch {}

$ErrorActionPreference = 'Stop'

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$basePath = Split-Path -Parent $scriptRoot
$configPath = Join-Path $basePath 'Config\settings.json'
$setupScript = Join-Path $scriptRoot 'Setup-AdministrativeFunctionsBackup.ps1'
$exportScript = Join-Path $scriptRoot 'Export-AdministrativeFunctions.ps1'



function Test-RmadInvocation {
    if ($RmadMode) { return $true }

    try {
        if (-not [Environment]::UserInteractive) { return $true }
    } catch {}

    try {
        $process = Get-CimInstance Win32_Process -Filter "ProcessId = $PID" -ErrorAction Stop
        for ($i = 0; $i -lt 10 -and $process -and $process.ParentProcessId; $i++) {
            $process = Get-CimInstance Win32_Process -Filter "ProcessId = $($process.ParentProcessId)" -ErrorAction SilentlyContinue
            if (-not $process) { break }

            if ($process.Name -ieq 'BackupServer.exe') { return $true }
            if ($process.CommandLine -and ($process.CommandLine -match 'Run-AdministrativeFunctionsBackup-RMAD\.cmd')) { return $true }
        }
    } catch {}

    return $false
}

function Write-RmadLog {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    if (-not $script:RmadLogFile) { return }
    $line = "[{0}] {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message
    Add-Content -Path $script:RmadLogFile -Value $line -Encoding UTF8
}

function Get-MicrosoftGraphAvailability {
    $modules = @(Get-Module -ListAvailable -Name Microsoft.Graph, Microsoft.Graph.Authentication -ErrorAction SilentlyContinue |
        Sort-Object Name, Version -Descending |
        Select-Object Name, Version, Path)
    $connectCommand = Get-Command Connect-MgGraph -ErrorAction SilentlyContinue
    $disconnectCommand = Get-Command Disconnect-MgGraph -ErrorAction SilentlyContinue

    return [pscustomobject]@{
        ModuleAvailable = ($modules.Count -gt 0)
        Modules = $modules
        ConnectMgGraphAvailable = [bool]$connectCommand
        DisconnectMgGraphAvailable = [bool]$disconnectCommand
    }
}

function Write-RmadDiagnostics {
    param(
        [string]$SettingsFile,
        [object]$Settings = $null
    )

    Write-RmadLog "Usuário de execução: $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)"
    Write-RmadLog "Arquitetura do PowerShell: $([IntPtr]::Size * 8)-bit"
    Write-RmadLog "Caminho do PowerShell: $((Get-Process -Id $PID).Path)"
    Write-RmadLog "settings.json existe: $(Test-Path $SettingsFile) ($SettingsFile)"

    $thumbprint = $null
    if ($Settings -and -not [string]::IsNullOrWhiteSpace($Settings.CertificateThumbprint)) {
        $thumbprint = $Settings.CertificateThumbprint
    }

    if ($thumbprint) {
        $cert = Get-Item "Cert:\LocalMachine\My\$thumbprint" -ErrorAction SilentlyContinue
        Write-RmadLog "Certificado existe: $([bool]$cert) (Cert:\LocalMachine\My\$thumbprint)"
        if ($cert) {
            Write-RmadLog "HasPrivateKey: $($cert.HasPrivateKey)"
            Write-RmadLog "Certificado expira em: $($cert.NotAfter.ToString('yyyy-MM-dd HH:mm:ss'))"
        } else {
            Write-RmadLog "HasPrivateKey: n/a"
        }
    } else {
        Write-RmadLog 'Certificado existe: n/a (CertificateThumbprint ausente no settings.json)'
        Write-RmadLog 'HasPrivateKey: n/a'
    }

    $graph = Get-MicrosoftGraphAvailability
    Write-RmadLog "Microsoft.Graph disponível: $($graph.ModuleAvailable)"
    Write-RmadLog "Connect-MgGraph disponível: $($graph.ConnectMgGraphAvailable)"
    Write-RmadLog "Disconnect-MgGraph disponível: $($graph.DisconnectMgGraphAvailable)"
    foreach ($module in $graph.Modules) {
        Write-RmadLog "Microsoft.Graph módulo: $($module.Name) $($module.Version) - $($module.Path)"
    }
}

function Get-ValidBackupFolder {
    param(
        [string]$BackupRoot,
        [datetime]$NotBefore
    )

    if (-not (Test-Path $BackupRoot)) { return $null }

    $requiredFiles = @('manifest.json','roleDefinitions.json','roleAssignments.json')
    $folders = @(Get-ChildItem -Path $BackupRoot -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.CreationTime -ge $NotBefore.AddSeconds(-5) -or $_.LastWriteTime -ge $NotBefore.AddSeconds(-5) } |
        Sort-Object LastWriteTime -Descending)

    foreach ($folder in $folders) {
        $missing = @($requiredFiles | Where-Object { -not (Test-Path (Join-Path $folder.FullName $_)) })
        if ($missing.Count -eq 0) { return $folder.FullName }
    }

    return $null
}

function Invoke-RmadBackup {
    Initialize-BackupDirectories -RootPath $basePath

    $logRoot = Join-Path $basePath 'Logs'
    New-Item -ItemType Directory -Path $logRoot -Force | Out-Null
    $script:RmadLogFile = Join-Path $logRoot ("rmad-custom-script-{0}.log" -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
    New-Item -ItemType File -Path $script:RmadLogFile -Force | Out-Null

    Write-RmadLog 'Iniciando execução RMAD/non-interactive do backup de funções administrativas.'
    Write-RmadDiagnostics -SettingsFile $configPath

    try {
        Ensure-MicrosoftGraphPowerShell *>&1 | ForEach-Object { Write-RmadLog "GRAPH-SETUP: $_" }
        Write-RmadLog 'Disponibilidade do Microsoft.Graph após preparação:'
        $graph = Get-MicrosoftGraphAvailability
        Write-RmadLog "Microsoft.Graph disponível: $($graph.ModuleAvailable)"
        Write-RmadLog "Connect-MgGraph disponível: $($graph.ConnectMgGraphAvailable)"
        Write-RmadLog "Disconnect-MgGraph disponível: $($graph.DisconnectMgGraphAvailable)"

        $validation = Test-BackupPrerequisites -SettingsFile $configPath
        Write-RmadDiagnostics -SettingsFile $configPath -Settings $validation.Settings

        if (-not $validation.IsValid) {
            Write-RmadLog "Pré-validação falhou: $($validation.Reasons -join '; ')"
            Write-RmadLog 'Assistente interativo não será aberto em modo RMAD.'
            exit 1
        }

        Test-GraphConnection -TenantId $validation.Settings.TenantId -ClientId $validation.Settings.ClientId -Thumbprint $validation.Settings.CertificateThumbprint *>&1 | ForEach-Object { Write-RmadLog "GRAPH-TEST: $_" }

        $backupStart = Get-Date
        Write-RmadLog 'Iniciando export do backup.'
        & $exportScript -TenantId $validation.Settings.TenantId -ClientId $validation.Settings.ClientId -CertificateThumbprint $validation.Settings.CertificateThumbprint -BasePath $basePath *>&1 | ForEach-Object { Write-RmadLog "EXPORT: $_" }

        $backupFolder = Get-ValidBackupFolder -BackupRoot (Join-Path $basePath 'Backups') -NotBefore $backupStart
        if ($backupFolder) {
            Write-RmadLog "Caminho do backup gerado: $backupFolder"
            Write-RmadLog 'Backup RMAD concluído com sucesso; manifest.json, roleDefinitions.json e roleAssignments.json foram criados.'
            exit 0
        }

        Write-RmadLog 'Export terminou sem erro, mas os arquivos obrigatórios não foram encontrados em um novo backup.'
        exit 1
    }
    catch {
        Write-RmadLog "Falha real no export/preparação RMAD: $($_.Exception.Message)"
        if ($_.ScriptStackTrace) { Write-RmadLog "StackTrace: $($_.ScriptStackTrace)" }
        exit 1
    }
}

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

if (Test-RmadInvocation) {
    Invoke-RmadBackup
}

Write-Host "Executando scripts em: $basePath" -ForegroundColor Cyan
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
