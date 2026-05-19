[CmdletBinding()]
param(
    [string]$SettingsPath = 'C:\ProgramData\Quest\IR-AdministrativeFunctionBackup\Config\settings.json',
    [int]$CertificateValidityMonths = 24,
    [string]$AppDisplayName = 'Quest Recovery Function - Administrative Roles Backup',
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$TenantId
)

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

try {
    chcp 65001 | Out-Null
} catch {}

$ErrorActionPreference = 'Stop'

$sharedModulePath = Join-Path $PSScriptRoot 'IR-AdministrativeFunctions.psm1'
Import-Module $sharedModulePath -Force
$GraphAppId = '00000003-0000-0000-c000-000000000000'
$InstallRoot = 'C:\ProgramData\Quest\IR-AdministrativeFunctionBackup'
$BootstrapLogPath = Join-Path $InstallRoot 'Logs\tenant-bootstrap.log'
$LogoPath = Join-Path $InstallRoot 'Assets\quest-logo.png'
$RequiredInteractiveScopes = @(
    'Application.ReadWrite.All'
    'AppRoleAssignment.ReadWrite.All'
    'Directory.ReadWrite.All'
    'RoleManagement.ReadWrite.Directory'
)
$RequiredApplicationPermissions = @(
    'RoleManagement.Read.Directory'
    'RoleManagement.ReadWrite.Directory'
    'Directory.Read.All'
    'Directory.ReadWrite.All'
)
$OptionalApplicationPermissions = @(
    'Application.Read.All'
    'AppRoleAssignment.ReadWrite.All'
)



function Initialize-BootstrapDirectories {
    foreach ($folder in @('Config','Logs','Backups','Certificates','Assets')) {
        $path = Join-Path $InstallRoot $folder
        if (-not (Test-Path $path)) {
            New-Item -ItemType Directory -Path $path -Force | Out-Null
        }
    }
}

function Initialize-BootstrapLog {
    Initialize-BootstrapDirectories
    $logDir = Split-Path -Parent $BootstrapLogPath
    New-Item -ItemType Directory -Path $logDir -Force | Out-Null
    if (-not (Test-Path $BootstrapLogPath)) {
        New-Item -ItemType File -Path $BootstrapLogPath -Force | Out-Null
    }
}

function Write-BootstrapLog {
    param([string]$Message)

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "[$timestamp] $Message"

    Write-Host $line
    $dir = Split-Path -Parent $BootstrapLogPath
    New-Item -ItemType Directory -Path $dir -Force | Out-Null
    Add-Content -Path $BootstrapLogPath -Value $line -Encoding UTF8
}

function Get-DefaultSettings {
    [ordered]@{
        TenantId = ''
        ClientId = ''
        AppDisplayName = $AppDisplayName
        CertificateThumbprint = ''
        CertificateStore = 'LocalMachine\My'
        BackupRoot = (Join-Path $InstallRoot 'Backups')
        LogRoot = (Join-Path $InstallRoot 'Logs')
        LogoPath = $LogoPath
    }
}

function Save-Settings {
    param([hashtable]$Settings,[string]$Path)
    $dir = Split-Path -Parent $Path
    New-Item -ItemType Directory -Path $dir -Force | Out-Null
    ($Settings | ConvertTo-Json -Depth 10) | Set-Content -Path $Path -Encoding UTF8
}

function Export-PublicCertificate {
    param([string]$Thumbprint)
    $certDir = (Join-Path $InstallRoot 'Certificates')
    $cerPath = Join-Path $certDir 'public-certificate.cer'
    New-Item -ItemType Directory -Path $certDir -Force | Out-Null
    $cert = Get-Item "Cert:\LocalMachine\My\$Thumbprint"
    Export-Certificate -Cert $cert -FilePath $cerPath -Force | Out-Null
    return $cerPath
}

function Add-CertificateToApplication {
    param(
        [Microsoft.Graph.PowerShell.Models.IMicrosoftGraphApplication]$Application,
        [System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate,
        [string]$CertificatePath
    )

    $app = Get-MgApplication -ApplicationId $Application.Id -Property Id,AppId,DisplayName,KeyCredentials
    Write-BootstrapLog "Application ObjectId: $($app.Id)"
    Write-BootstrapLog "Application AppId/ClientId: $($app.AppId)"

    $match = $app.KeyCredentials | Where-Object { $_.CustomKeyIdentifier -and ([System.Convert]::ToBase64String($_.CustomKeyIdentifier) -eq [System.Convert]::ToBase64String($Certificate.GetCertHash())) }
    if ($match) { return $false }

    $certBytes = [System.IO.File]::ReadAllBytes($CertificatePath)
    $newKeyCredential = @{
        Type = 'AsymmetricX509Cert'
        Usage = 'Verify'
        Key = $certBytes
        DisplayName = 'Quest Recovery Function - Administrative Roles Backup'
        StartDateTime = $Certificate.NotBefore.ToUniversalTime()
        EndDateTime = $Certificate.NotAfter.ToUniversalTime()
    }

    $existingKeys = @()
    if ($app.KeyCredentials) { $existingKeys = @($app.KeyCredentials) }
    $updatedKeys = @($existingKeys + $newKeyCredential)

    Write-BootstrapLog "Atualizando certificado no Application ObjectId: $($app.Id)"
    Update-MgApplication -ApplicationId $app.Id -KeyCredentials $updatedKeys
    return $true
}

function Ensure-GraphApplicationPermissions {
    param([string]$ApplicationId,[string]$ServicePrincipalId)

    $graphSp = Get-MgServicePrincipal -Filter "appId eq '$GraphAppId'"
    $currentApp = Get-MgApplication -ApplicationId $ApplicationId -Property RequiredResourceAccess
    $currentAccess = @($currentApp.RequiredResourceAccess | Where-Object { $_.ResourceAppId -eq $GraphAppId } | ForEach-Object { $_.ResourceAccess })

    $requiredRoles = @($RequiredApplicationPermissions + $OptionalApplicationPermissions)
    $roleMap = @{}
    foreach ($value in $requiredRoles) {
        $role = $graphSp.AppRoles | Where-Object { $_.Value -eq $value -and $_.AllowedMemberTypes -contains 'Application' }
        if ($role) { $roleMap[$value] = $role.Id } else { Write-BootstrapLog "Permissão $value não encontrada nos appRoles do Graph." }
    }

    $merged = @{}
    foreach ($entry in $currentAccess) { $merged[$entry.Id] = 'Role' }
    foreach ($roleId in $roleMap.Values) { $merged[$roleId] = 'Role' }

    $resourceAccess = @()
    foreach ($k in $merged.Keys) { $resourceAccess += @{ Id = $k; Type = 'Role' } }
    Update-MgApplication -ApplicationId $ApplicationId -RequiredResourceAccess @(@{ ResourceAppId = $GraphAppId; ResourceAccess = $resourceAccess })
    Write-BootstrapLog 'Permissões de aplicação do Microsoft Graph configuradas no App Registration.'

    $existingAssignments = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $ServicePrincipalId -All
    foreach ($pair in $roleMap.GetEnumerator()) {
        $already = $existingAssignments | Where-Object { $_.ResourceId -eq $graphSp.Id -and $_.AppRoleId -eq $pair.Value }
        if (-not $already) {
            try {
                New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $ServicePrincipalId -PrincipalId $ServicePrincipalId -ResourceId $graphSp.Id -AppRoleId $pair.Value | Out-Null
                Write-BootstrapLog "Admin consent concedido automaticamente para: $($pair.Key)"
            } catch {
                $consentMsg = $_.Exception.Message
                Write-BootstrapLog "falha ao conceder admin consent: $consentMsg"
                throw
            }
        }
    }
}


function New-InsufficientPrivilegeErrorMessage {
    return 'A autenticação foi concluída, mas a operação foi negada pelo Microsoft Graph (Authorization_RequestDenied/Insufficient privileges). Valide as permissões efetivas no tenant, políticas de acesso e restrições administrativas para atualização do objeto Application.'
}


function Test-GraphContextScopes {
    param([string[]]$ExpectedScopes)
    $ctx = Get-MgContext
    if (-not $ctx.Account) { throw 'Falha na validação pós-login: conta autenticada não identificada.' }
    if (-not $ctx.TenantId) { throw 'Falha na validação pós-login: tenant autenticado não identificado.' }
    if ($ctx.TenantId -ne $TenantId) { throw "Falha na validação pós-login: tenant autenticado ($($ctx.TenantId)) difere do TenantId informado ($TenantId)." }

    $missing = @($ExpectedScopes | Where-Object { $_ -notin $ctx.Scopes })
    if ($missing.Count -gt 0) { throw "Falha na validação pós-login: scopes ausentes: $($missing -join ', ')." }

    Write-BootstrapLog "Conta autenticada: $($ctx.Account)"
    Write-BootstrapLog "Tenant autenticado: $($ctx.TenantId)"
    Write-BootstrapLog "Scopes concedidos: $($ctx.Scopes -join ', ')"
}

function Set-AppLogoIfAvailable {
    param([string]$ApplicationId)
    if (-not (Test-Path $LogoPath)) {
        Write-BootstrapLog "Logo não encontrado em $LogoPath. Prosseguindo sem aplicar identidade visual."
        return
    }

    $file = Get-Item $LogoPath
    $allowedExtensions = @('.png','.jpg','.jpeg','.gif','.bmp')
    if ($file.Extension.ToLowerInvariant() -notin $allowedExtensions) {
        Write-BootstrapLog "WARNING: Formato de logo não suportado para upload ($($file.Extension))."
        return
    }

    if ($file.Length -gt 102400KB) {
        Write-BootstrapLog 'WARNING: Logo excede tamanho máximo esperado para upload (100 MB). Ignorando upload.'
        return
    }

    try {
        Set-MgApplicationLogo -ApplicationId $ApplicationId -InFile $LogoPath
        Write-BootstrapLog 'Logo aplicado com sucesso no App Registration.'
    } catch {
        Write-BootstrapLog "WARNING: Falha ao aplicar logo no App Registration: $($_.Exception.Message)"
    }
}


function Test-AppOnlyGraphWithRetry {
    param(
        [string]$TenantId,
        [string]$ClientId,
        [string]$CertificateThumbprint,
        [int]$MaxAttempts = 8,
        [int]$DelaySeconds = 15
    )

    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        try {
            Write-BootstrapLog "Validação app-only final: tentativa $attempt de $MaxAttempts."
            Connect-MgGraph -ClientId $ClientId -TenantId $TenantId -CertificateThumbprint $CertificateThumbprint -NoWelcome -ContextScope Process | Out-Null
            Get-MgRoleManagementDirectoryRoleDefinition -All | Select-Object -First 1 | Out-Null
            Write-BootstrapLog "Teste app-only concluído com sucesso na tentativa $attempt."
            return
        } catch {
            $lastError = $_.Exception.Message
            Write-BootstrapLog "WARNING: tentativa $attempt de validação app-only falhou: $lastError"
            if ($attempt -eq $MaxAttempts) {
                throw "Configuração salva. A validação app-only final falhou, possivelmente por propagação do certificado no Microsoft Entra ID. Aguarde alguns minutos e teste novamente. Último erro: $lastError"
            }
            Write-BootstrapLog "Aguardando $DelaySeconds segundos antes de nova tentativa de validação app-only."
            Start-Sleep -Seconds $DelaySeconds
        }
    }
}

function Test-IsAuthorizationDenied {
    param([string]$Message)
    if ([string]::IsNullOrWhiteSpace($Message)) { return $false }
    return ($Message -match 'Authorization_RequestDenied' -or $Message -match 'Insufficient privileges' -or $Message -match 'insufficient privileges')
}

try {
    Initialize-BootstrapLog
    Write-BootstrapLog 'Log de bootstrap inicializado.'
    Ensure-MicrosoftGraphPowerShell
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    Write-BootstrapLog 'Iniciando autenticação interativa (Device Code) no Microsoft Graph...'
    Connect-MgGraph -TenantId $TenantId -Scopes $RequiredInteractiveScopes -UseDeviceCode -NoWelcome
    Test-GraphContextScopes -ExpectedScopes $RequiredInteractiveScopes
    $tenantId = $TenantId
    Write-BootstrapLog "Autenticação concluída. Tenant: $tenantId"

    $existingApps = @(Get-MgApplication -Filter "displayName eq '$AppDisplayName'" -ConsistencyLevel eventual)
    $selectedApp = $null
    if ($existingApps.Count -gt 1) {
        $appList = ($existingApps | ForEach-Object { "$($_.DisplayName) [ObjectId=$($_.Id); AppId=$($_.AppId)]" }) -join '; '
        $multiMsg = "Foram encontrados múltiplos App Registrations com o mesmo displayName '$AppDisplayName'. Não é seguro reutilizar automaticamente. Selecione explicitamente o App Registration correto ou crie um novo com sufixo único. Encontrados: $appList"
        Write-BootstrapLog $multiMsg
        throw $multiMsg
    }
    if ($existingApps.Count -eq 1) {
        $selectedApp = $existingApps[0]
        $null = Get-MgApplication -ApplicationId $selectedApp.Id -Property Id,AppId,DisplayName
        Write-BootstrapLog "App Registration existente localizado e reutilizado. ObjectId: $($selectedApp.Id) | AppId: $($selectedApp.AppId)"
    }
    if (-not $selectedApp) {
        $selectedApp = New-MgApplication -DisplayName $AppDisplayName -SignInAudience 'AzureADMyOrg'
        Write-BootstrapLog "Novo App Registration criado. AppId: $($selectedApp.AppId)"
    }

    $sp = Get-MgServicePrincipal -Filter "appId eq '$($selectedApp.AppId)'"
    if (-not $sp) { $sp = New-MgServicePrincipal -AppId $selectedApp.AppId; Write-BootstrapLog 'Service Principal criado no tenant.' } else { Write-BootstrapLog 'Service Principal existente localizado no tenant.' }

    $friendlyName = 'Quest Recovery Function - Administrative Roles Backup'
    $cert = New-SelfSignedCertificate -DnsName 'IR-AdministrativeFunctionBackup' -FriendlyName $friendlyName -CertStoreLocation 'Cert:\LocalMachine\My' -NotAfter (Get-Date).AddMonths([int]$CertificateValidityMonths) -KeyExportPolicy Exportable -KeyAlgorithm RSA -KeyLength 2048 -HashAlgorithm SHA256
    Write-BootstrapLog "Certificado self-signed criado. Thumbprint: $($cert.Thumbprint)"
    $cerPath = Export-PublicCertificate -Thumbprint $cert.Thumbprint
    Write-BootstrapLog "Certificado público exportado para: $cerPath"

    try {
        $added = Add-CertificateToApplication -Application $selectedApp -Certificate $cert -CertificatePath $cerPath
        if ($added) { Write-BootstrapLog 'Certificado público adicionado ao keyCredentials do App Registration.' } else { Write-BootstrapLog 'Certificado já estava presente no keyCredentials; nenhuma alteração necessária.' }
    } catch {
        $certMsg = $_.Exception.Message
        Write-BootstrapLog "falha ao adicionar certificado ao App Registration (comando: Update-MgApplication -ApplicationId $($selectedApp.Id) -KeyCredentials ... ; ObjectId=$($selectedApp.Id)): $certMsg"
        if (Test-IsAuthorizationDenied -Message $certMsg) { throw (New-InsufficientPrivilegeErrorMessage) }
        throw
    }

    try {
        Ensure-GraphApplicationPermissions -ApplicationId $selectedApp.Id -ServicePrincipalId $sp.Id
        Write-BootstrapLog 'Permissões Graph adicionadas ao App Registration.'
    } catch {
        $permMsg = $_.Exception.Message
        Write-BootstrapLog "falha ao adicionar permissões Graph: $permMsg"
        if (Test-IsAuthorizationDenied -Message $permMsg) { throw (New-InsufficientPrivilegeErrorMessage) }
        throw
    }

    Set-AppLogoIfAvailable -ApplicationId $selectedApp.Id

    $settingsPath = $SettingsPath
    $configRoot = Split-Path -Parent $settingsPath

    if (-not (Test-Path $configRoot)) {
        New-Item -Path $configRoot -ItemType Directory -Force | Out-Null
    }

    $settings = [ordered]@{
        TenantId = $tenantId
        ClientId = $selectedApp.AppId
        AppDisplayName = "Quest Recovery Function - Administrative Roles Backup"
        CertificateThumbprint = $cert.Thumbprint
        CertificateStore = "LocalMachine\My"
        BackupRoot = (Join-Path $InstallRoot 'Backups')
        LogRoot = (Join-Path $InstallRoot 'Logs')
        LogoPath = $LogoPath
    }

    $settings | ConvertTo-Json -Depth 10 | Set-Content -Path $settingsPath -Encoding UTF8

    if (-not (Test-Path $settingsPath)) {
        throw "Falha ao criar settings.json em $settingsPath"
    }

    Write-BootstrapLog "Configuração salva em: $settingsPath"

    Write-BootstrapLog 'Iniciando validação app-only final com retry...'
    try {
        Test-AppOnlyGraphWithRetry -ClientId $settings.ClientId -TenantId $settings.TenantId -CertificateThumbprint $settings.CertificateThumbprint
        exit 0
    } catch {
        Write-BootstrapLog "WARNING: $($_.Exception.Message)"
        exit 0
    }
} catch {
    $errorMsg = $_.Exception.Message
    if (Test-IsAuthorizationDenied -Message $errorMsg) {
        $friendlyMsg = New-InsufficientPrivilegeErrorMessage
        try { Write-BootstrapLog "ERRO: $friendlyMsg" } catch { Write-Error "O bootstrap falhou antes de inicializar o log. Verifique o caminho do script Start-TenantBootstrapInteractive.ps1 e os parâmetros usados pelo wizard. Erro original: $friendlyMsg" }
        throw $friendlyMsg
    }

    try {
        Write-BootstrapLog "ERRO: $errorMsg"
    } catch {
        Write-Error "O bootstrap falhou antes de inicializar o log. Verifique o caminho do script Start-TenantBootstrapInteractive.ps1 e os parâmetros usados pelo wizard. Erro original: $errorMsg"
    }
    throw
} finally {
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
}
