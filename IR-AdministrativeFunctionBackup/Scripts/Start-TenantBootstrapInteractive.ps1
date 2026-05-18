[CmdletBinding()]
param(
    [string]$SettingsPath = 'C:\ProgramData\Quest\IR-AdministrativeFunctionBackup\Config\settings.json',
    [int]$CertificateValidityMonths = 24,
    [string]$AppDisplayName = 'Quest Recovery Function - Administrative Roles Backup'
)

$ErrorActionPreference = 'Stop'
$GraphAppId = '00000003-0000-0000-c000-000000000000'
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

function Write-Step { param([string]$Message) Write-Host "[$(Get-Date -Format 'HH:mm:ss')] $Message" }

function Get-DefaultSettings {
    [ordered]@{
        TenantId = ''
        ClientId = ''
        AppDisplayName = $AppDisplayName
        CertificateThumbprint = ''
        CertificateStore = 'LocalMachine\\My'
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
        if ($role) { $roleMap[$value] = $role.Id } else { Write-Step "Permissão $value não encontrada nos appRoles do Graph." }
    }

    $merged = @{}
    foreach ($entry in $currentAccess) { $merged[$entry.Id] = 'Role' }
    foreach ($roleId in $roleMap.Values) { $merged[$roleId] = 'Role' }

    $resourceAccess = @()
    foreach ($k in $merged.Keys) { $resourceAccess += @{ Id = $k; Type = 'Role' } }
    Update-MgApplication -ApplicationId $ApplicationId -RequiredResourceAccess @(@{ ResourceAppId = $GraphAppId; ResourceAccess = $resourceAccess })
    Write-Step 'Permissões de aplicação do Microsoft Graph configuradas no App Registration.'

    $existingAssignments = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $ServicePrincipalId -All
    foreach ($pair in $roleMap.GetEnumerator()) {
        $already = $existingAssignments | Where-Object { $_.ResourceId -eq $graphSp.Id -and $_.AppRoleId -eq $pair.Value }
        if (-not $already) {
            New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $ServicePrincipalId -PrincipalId $ServicePrincipalId -ResourceId $graphSp.Id -AppRoleId $pair.Value | Out-Null
            Write-Step "Admin consent concedido automaticamente para: $($pair.Key)"
        }
    }
}

try {
    Write-Step 'Iniciando autenticação interativa (Device Code) no Microsoft Graph...'
    Connect-MgGraph -Scopes $RequiredInteractiveScopes -UseDeviceCode -NoWelcome
    $ctx = Get-MgContext
    $tenantId = $ctx.TenantId
    Write-Step "Autenticação concluída. Tenant: $tenantId"

    $existingApps = Get-MgApplication -Filter "displayName eq '$AppDisplayName'" -ConsistencyLevel eventual
    $selectedApp = $null
    if ($existingApps) {
        $selectedApp = $existingApps | Select-Object -First 1
        Write-Step "Reutilizando App Registration existente. AppId: $($selectedApp.AppId)"
    }
    if (-not $selectedApp) {
        $selectedApp = New-MgApplication -DisplayName $AppDisplayName -SignInAudience 'AzureADMyOrg'
        Write-Step "Novo App Registration criado. AppId: $($selectedApp.AppId)"
    }

    $sp = Get-MgServicePrincipal -Filter "appId eq '$($selectedApp.AppId)'"
    if (-not $sp) { $sp = New-MgServicePrincipal -AppId $selectedApp.AppId; Write-Step 'Service Principal criado no tenant.' } else { Write-Step 'Service Principal existente localizado no tenant.' }

    $friendlyName = 'Quest Recovery Function - Administrative Roles Backup'
    $cert = New-SelfSignedCertificate -DnsName 'IR-AdministrativeFunctionBackup' -FriendlyName $friendlyName -CertStoreLocation 'Cert:\LocalMachine\My' -NotAfter (Get-Date).AddMonths([int]$CertificateValidityMonths) -KeyExportPolicy Exportable -KeyAlgorithm RSA -KeyLength 2048 -HashAlgorithm SHA256
    Write-Step "Certificado self-signed criado. Thumbprint: $($cert.Thumbprint)"
    $cerPath = Export-PublicCertificate -Thumbprint $cert.Thumbprint
    Write-Step "Certificado público exportado para: $cerPath"

    $added = Add-CertificateToApplication -ApplicationId $selectedApp.Id -Certificate $cert
    if ($added) { Write-Step 'Certificado público adicionado ao keyCredentials do App Registration.' } else { Write-Step 'Certificado já estava presente no keyCredentials; nenhuma alteração necessária.' }

    try { Ensure-GraphApplicationPermissions -ApplicationId $selectedApp.Id -ServicePrincipalId $sp.Id }
    catch { Write-Step "Falha ao conceder admin consent automático: $($_.Exception.Message)" }

    $settings = Get-DefaultSettings
    $settings.TenantId = $tenantId
    $settings.ClientId = $selectedApp.AppId
    $settings.CertificateThumbprint = $cert.Thumbprint
    Save-Settings -Settings $settings -Path $SettingsPath
    Write-Step "Configuração salva em: $SettingsPath"

    Connect-MgGraph -ClientId $settings.ClientId -TenantId $settings.TenantId -CertificateThumbprint $settings.CertificateThumbprint -NoWelcome -ContextScope Process | Out-Null
    Get-MgRoleManagementDirectoryRoleDefinition -Top 1 | Out-Null
    Write-Step 'Teste app-only concluído com sucesso.'
    exit 0
} catch {
    Write-Error "Falha na configuração automática: $($_.Exception.Message)"
    exit 1
} finally {
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
}
