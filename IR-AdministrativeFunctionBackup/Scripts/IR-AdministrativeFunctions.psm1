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

function Connect-AppOnlyGraphWithRetry {
    param(
        [string]$TenantId,
        [string]$ClientId,
        [string]$Thumbprint,
        [int]$MaxAttempts = 8,
        [int]$DelaySeconds = 15,
        [string]$FailureMessage = 'Falha ao validar app-only após {0} tentativas. Último erro: {1}',
        [switch]$VerboseOutput
    )

    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        try {
            if ($VerboseOutput) {
                Write-Host "Validando conexão app-only com Microsoft Graph (tentativa $attempt de $MaxAttempts)..." -ForegroundColor Cyan
            }
            Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $Thumbprint -NoWelcome -ContextScope Process | Out-Null
            Get-MgRoleManagementDirectoryRoleDefinition -All | Select-Object -First 1 | Out-Null
            return
        } catch {
            $lastError = $_.Exception.Message
            Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
            if ($attempt -eq $MaxAttempts) {
                throw ($FailureMessage -f $MaxAttempts, $lastError)
            }
            if ($VerboseOutput) {
                Write-Host "Validação app-only ainda não disponível: $lastError" -ForegroundColor Yellow
                Write-Host "Aguardando $DelaySeconds segundos para possível propagação no Microsoft Entra ID..." -ForegroundColor Yellow
            }
            Start-Sleep -Seconds $DelaySeconds
        }
    }
}

function Test-GraphConnection {
    param(
        [string]$TenantId,
        [string]$ClientId,
        [string]$Thumbprint,
        [int]$MaxAttempts = 8,
        [int]$DelaySeconds = 15,
        [string]$FailureMessage = 'Falha ao validar app-only após {0} tentativas. Último erro: {1}',
        [switch]$VerboseOutput
    )

    try {
        Connect-AppOnlyGraphWithRetry -TenantId $TenantId -ClientId $ClientId -Thumbprint $Thumbprint -MaxAttempts $MaxAttempts -DelaySeconds $DelaySeconds -FailureMessage $FailureMessage -VerboseOutput:$VerboseOutput
    } finally {
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    }
}

Export-ModuleMember -Function Ensure-MicrosoftGraphPowerShell, Initialize-BackupDirectories, Connect-AppOnlyGraphWithRetry, Test-GraphConnection
