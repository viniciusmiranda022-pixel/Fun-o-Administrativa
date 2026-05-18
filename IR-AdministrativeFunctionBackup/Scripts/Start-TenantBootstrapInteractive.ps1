[CmdletBinding()]
param(
    [string]$SettingsPath = 'C:\ProgramData\Quest\IR-AdministrativeFunctionBackup\Config\settings.json',
    [int]$CertificateValidityMonths = 24,
    [string]$AppDisplayName = 'Quest Recovery Function - Administrative Roles Backup'
)

$ErrorActionPreference = 'Stop'
$GraphAppId = '00000003-0000-0000-c000-000000000000'
$BootstrapLogPath = 'C:\ProgramData\Quest\IR-AdministrativeFunctionBackup\Logs\tenant-bootstrap.log'
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
        BackupRoot = 'C:\ProgramData\Quest\IR-AdministrativeFunctionBackup\Backups'
        LogRoot = 'C:\ProgramData\Quest\IR-AdministrativeFunctionBackup\Logs'
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
    $certDir = 'C:\ProgramData\Quest\IR-AdministrativeFunctionBackup\Certificates'
    $cerPath = Join-Path $certDir 'public-certificate.cer'
    New-Item -ItemType Directory -Path $certDir -Force | Out-Null
    $cert = Get-Item "Cert:\LocalMachine\My\$Thumbprint"
    Export-Certificate -Cert $cert -FilePath $cerPath -Force | Out-Null
    return $cerPath
}

function Add-CertificateToApplication {
    param(
        [string]$ApplicationId,
        [System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate
    )

    $existing = Get-MgApplication -ApplicationId $ApplicationId -Property KeyCredentials
    $match = $existing.KeyCredentials | Where-Object { $_.CustomKeyIdentifier -and ([System.Convert]::ToBase64String($_.CustomKeyIdentifier) -eq [System.Convert]::ToBase64String($Certificate.GetCertHash())) }
    if ($match) { return $false }

    $params = @{ keyCredential = @{ type = 'AsymmetricX509Cert'; usage = 'Verify'; key = [System.Convert]::ToBase64String($Certificate.RawData); displayName = 'Primary certificate' } }
    Add-MgApplicationKey -ApplicationId $ApplicationId -BodyParameter $params | Out-Null
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
            New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $ServicePrincipalId -PrincipalId $ServicePrincipalId -ResourceId $graphSp.Id -AppRoleId $pair.Value | Out-Null
            Write-BootstrapLog "Admin consent concedido automaticamente para: $($pair.Key)"
        }
    }
}

try {
    Write-BootstrapLog 'Iniciando autenticação interativa (Device Code) no Microsoft Graph...'
    Connect-MgGraph -Scopes $RequiredInteractiveScopes -UseDeviceCode -NoWelcome
    $ctx = Get-MgContext
    $tenantId = $ctx.TenantId
    Write-BootstrapLog "Autenticação concluída. Tenant: $tenantId"

    $existingApps = Get-MgApplication -Filter "displayName eq '$AppDisplayName'" -ConsistencyLevel eventual
    $selectedApp = $null
    if ($existingApps) {
        $selectedApp = $existingApps | Select-Object -First 1
        Write-BootstrapLog "Reutilizando App Registration existente. AppId: $($selectedApp.AppId)"
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

    $added = Add-CertificateToApplication -ApplicationId $selectedApp.Id -Certificate $cert
    if ($added) { Write-BootstrapLog 'Certificado público adicionado ao keyCredentials do App Registration.' } else { Write-BootstrapLog 'Certificado já estava presente no keyCredentials; nenhuma alteração necessária.' }

    Ensure-GraphApplicationPermissions -ApplicationId $selectedApp.Id -ServicePrincipalId $sp.Id

    $configRoot = "C:\ProgramData\Quest\IR-AdministrativeFunctionBackup\Config"

    if (-not (Test-Path $configRoot)) {
        New-Item -Path $configRoot -ItemType Directory -Force | Out-Null
    }

    $settingsPath = Join-Path $configRoot "settings.json"

    $settings = [ordered]@{
        TenantId = $tenantId
        ClientId = $selectedApp.AppId
        AppDisplayName = "Quest Recovery Function - Administrative Roles Backup"
        CertificateThumbprint = $cert.Thumbprint
        CertificateStore = "LocalMachine\My"
        BackupRoot = "C:\ProgramData\Quest\IR-AdministrativeFunctionBackup\Backups"
        LogRoot = "C:\ProgramData\Quest\IR-AdministrativeFunctionBackup\Logs"
    }

    $settings | ConvertTo-Json -Depth 10 | Set-Content -Path $settingsPath -Encoding UTF8

    if (-not (Test-Path $settingsPath)) {
        throw "Falha ao criar settings.json em $settingsPath"
    }

    Write-BootstrapLog "Configuração salva em: $settingsPath"

    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    Connect-MgGraph -ClientId $settings.ClientId -TenantId $settings.TenantId -CertificateThumbprint $settings.CertificateThumbprint -NoWelcome -ContextScope Process | Out-Null
    Get-MgRoleManagementDirectoryRoleDefinition -Top 1 | Out-Null
    Write-BootstrapLog 'Teste app-only concluído com sucesso.'
    exit 0
} catch {
    Write-BootstrapLog "ERRO: $($_.Exception.Message)"
    throw
} finally {
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
}
