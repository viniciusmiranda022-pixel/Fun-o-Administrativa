[CmdletBinding()]
param(
    [string]$SettingsPath = (Join-Path 'C:\ProgramData\Quest\IR-AdministrativeFunctionBackup' 'Config\settings.json'),
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

$ProgramDataQuestRoot = 'C:\ProgramData\Quest'
$BackupInstallRoot = Join-Path $ProgramDataQuestRoot 'IR-AdministrativeFunctionBackup'

$ErrorActionPreference = 'Stop'

$sharedModulePath = Join-Path $PSScriptRoot 'IR-AdministrativeFunctions.psm1'
Import-Module $sharedModulePath -Force
$GraphAppId = '00000003-0000-0000-c000-000000000000'
$InstallRoot = $BackupInstallRoot
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

    Write-BootstrapLog "Updating certificate on Application ObjectId: $($app.Id)"
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
        if ($role) { $roleMap[$value] = $role.Id } else { Write-BootstrapLog "Permission $value not found in Graph appRoles." }
    }

    $merged = @{}
    foreach ($entry in $currentAccess) { $merged[$entry.Id] = 'Role' }
    foreach ($roleId in $roleMap.Values) { $merged[$roleId] = 'Role' }

    $resourceAccess = @()
    foreach ($k in $merged.Keys) { $resourceAccess += @{ Id = $k; Type = 'Role' } }
    Update-MgApplication -ApplicationId $ApplicationId -RequiredResourceAccess @(@{ ResourceAppId = $GraphAppId; ResourceAccess = $resourceAccess })
    Write-BootstrapLog 'Microsoft Graph application permissions configured in App Registration.'

    $existingAssignments = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $ServicePrincipalId -All
    foreach ($pair in $roleMap.GetEnumerator()) {
        $already = $existingAssignments | Where-Object { $_.ResourceId -eq $graphSp.Id -and $_.AppRoleId -eq $pair.Value }
        if (-not $already) {
            try {
                New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $ServicePrincipalId -PrincipalId $ServicePrincipalId -ResourceId $graphSp.Id -AppRoleId $pair.Value | Out-Null
                Write-BootstrapLog "Admin consent automatically granted for: $($pair.Key)"
            } catch {
                $consentMsg = $_.Exception.Message
                Write-BootstrapLog "Failed to grant admin consent: $consentMsg"
                throw
            }
        }
    }
}


function New-InsufficientPrivilegeErrorMessage {
    return 'Authentication completed, but the operation was denied by Microsoft Graph (Authorization_RequestDenied/Insufficient privileges). Validate the effective permissions in the tenant, access policies, and administrative restrictions for updating the Application object.'
}


function Test-GraphContextScopes {
    param([string[]]$ExpectedScopes)
    $ctx = Get-MgContext
    if (-not $ctx.Account) { throw 'Post-login validation failed: authenticated account not identified.' }
    if (-not $ctx.TenantId) { throw 'Post-login validation failed: authenticated tenant not identified.' }
    if ($ctx.TenantId -ne $TenantId) { throw "Post-login validation failed: authenticated tenant ($($ctx.TenantId)) differs from the provided TenantId ($TenantId)." }

    $missing = @($ExpectedScopes | Where-Object { $_ -notin $ctx.Scopes })
    if ($missing.Count -gt 0) { throw "Post-login validation failed: missing scopes: $($missing -join ', ')." }

    Write-BootstrapLog "Authenticated account: $($ctx.Account)"
    Write-BootstrapLog "Authenticated tenant: $($ctx.TenantId)"
    Write-BootstrapLog "Granted scopes: $($ctx.Scopes -join ', ')"
}

function Set-AppLogoIfAvailable {
    param([string]$ApplicationId)
    if (-not (Test-Path $LogoPath)) {
        Write-BootstrapLog "Logo not found at $LogoPath. Proceeding without applying visual identity."
        return
    }

    $file = Get-Item $LogoPath
    $allowedExtensions = @('.png','.jpg','.jpeg','.gif','.bmp')
    if ($file.Extension.ToLowerInvariant() -notin $allowedExtensions) {
        Write-BootstrapLog "WARNING: Logo format not supported for upload ($($file.Extension))."
        return
    }

    if ($file.Length -gt 102400KB) {
        Write-BootstrapLog 'WARNING: Logo exceeds the maximum expected size for upload (100 MB). Skipping upload.'
        return
    }

    try {
        Set-MgApplicationLogo -ApplicationId $ApplicationId -InFile $LogoPath
        Write-BootstrapLog 'Logo successfully applied to App Registration.'
    } catch {
        Write-BootstrapLog "WARNING: Failed to apply logo to App Registration: $($_.Exception.Message)"
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
            Write-BootstrapLog "Final app-only validation: attempt $attempt of $MaxAttempts."
            Connect-MgGraph -ClientId $ClientId -TenantId $TenantId -CertificateThumbprint $CertificateThumbprint -NoWelcome -ContextScope Process | Out-Null
            Get-MgRoleManagementDirectoryRoleDefinition -All | Select-Object -First 1 | Out-Null
            Write-BootstrapLog "App-only test completed successfully on attempt $attempt."
            return
        } catch {
            $lastError = $_.Exception.Message
            Write-BootstrapLog "WARNING: attempt $attempt of app-only validation failed: $lastError"
            if ($attempt -eq $MaxAttempts) {
                throw "Configuration saved. The final app-only validation failed, possibly due to certificate propagation in Microsoft Entra ID. Wait a few minutes and test again. Last error: $lastError"
            }
            Write-BootstrapLog "Waiting $DelaySeconds seconds before next app-only validation attempt."
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
    Write-BootstrapLog 'Bootstrap log initialized.'
    Ensure-MicrosoftGraphPowerShell
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    Write-BootstrapLog 'Starting interactive authentication (Device Code) with Microsoft Graph...'
    Connect-MgGraph -TenantId $TenantId -Scopes $RequiredInteractiveScopes -UseDeviceCode -NoWelcome
    Test-GraphContextScopes -ExpectedScopes $RequiredInteractiveScopes
    $tenantId = $TenantId
    Write-BootstrapLog "Authentication completed. Tenant: $tenantId"

    $existingApps = @(Get-MgApplication -Filter "displayName eq '$AppDisplayName'" -ConsistencyLevel eventual)
    $selectedApp = $null
    if ($existingApps.Count -gt 1) {
        $appList = ($existingApps | ForEach-Object { "$($_.DisplayName) [ObjectId=$($_.Id); AppId=$($_.AppId)]" }) -join '; '
        $multiMsg = "Multiple App Registrations with the same displayName '$AppDisplayName' were found. It is not safe to reuse one automatically. Explicitly select the correct App Registration or create a new one with a unique suffix. Found: $appList"
        Write-BootstrapLog $multiMsg
        throw $multiMsg
    }
    if ($existingApps.Count -eq 1) {
        $selectedApp = $existingApps[0]
        $null = Get-MgApplication -ApplicationId $selectedApp.Id -Property Id,AppId,DisplayName
        Write-BootstrapLog "Existing App Registration found and reused. ObjectId: $($selectedApp.Id) | AppId: $($selectedApp.AppId)"
    }
    if (-not $selectedApp) {
        $selectedApp = New-MgApplication -DisplayName $AppDisplayName -SignInAudience 'AzureADMyOrg'
        Write-BootstrapLog "New App Registration created. AppId: $($selectedApp.AppId)"
    }

    $sp = Get-MgServicePrincipal -Filter "appId eq '$($selectedApp.AppId)'"
    if (-not $sp) { $sp = New-MgServicePrincipal -AppId $selectedApp.AppId; Write-BootstrapLog 'Service Principal created in tenant.' } else { Write-BootstrapLog 'Existing Service Principal found in tenant.' }

    $friendlyName = 'Quest Recovery Function - Administrative Roles Backup'
    $cert = New-SelfSignedCertificate -DnsName 'IR-AdministrativeFunctionBackup' -FriendlyName $friendlyName -CertStoreLocation 'Cert:\LocalMachine\My' -NotAfter (Get-Date).AddMonths([int]$CertificateValidityMonths) -KeyExportPolicy Exportable -KeyAlgorithm RSA -KeyLength 2048 -HashAlgorithm SHA256
    Write-BootstrapLog "Self-signed certificate created. Thumbprint: $($cert.Thumbprint)"
    $cerPath = Export-PublicCertificate -Thumbprint $cert.Thumbprint
    Write-BootstrapLog "Public certificate exported to: $cerPath"

    try {
        $added = Add-CertificateToApplication -Application $selectedApp -Certificate $cert -CertificatePath $cerPath
        if ($added) { Write-BootstrapLog 'Public certificate added to App Registration keyCredentials.' } else { Write-BootstrapLog 'Certificate was already present in keyCredentials; no changes needed.' }
    } catch {
        $certMsg = $_.Exception.Message
        Write-BootstrapLog "Failed to add certificate to App Registration (command: Update-MgApplication -ApplicationId $($selectedApp.Id) -KeyCredentials ... ; ObjectId=$($selectedApp.Id)): $certMsg"
        if (Test-IsAuthorizationDenied -Message $certMsg) { throw (New-InsufficientPrivilegeErrorMessage) }
        throw
    }

    try {
        Ensure-GraphApplicationPermissions -ApplicationId $selectedApp.Id -ServicePrincipalId $sp.Id
        Write-BootstrapLog 'Graph permissions added to App Registration.'
    } catch {
        $permMsg = $_.Exception.Message
        Write-BootstrapLog "Failed to add Graph permissions: $permMsg"
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
        throw "Failed to create settings.json at $settingsPath"
    }

    Write-BootstrapLog "Configuration saved to: $settingsPath"

    Write-BootstrapLog 'Starting final app-only validation with retry...'
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
        try { Write-BootstrapLog "ERROR: $friendlyMsg" } catch { Write-Error "The bootstrap failed before the log could be initialized. Check the path of the Start-TenantBootstrapInteractive.ps1 script and the parameters used by the wizard. Original error: $friendlyMsg" }
        throw $friendlyMsg
    }

    try {
        Write-BootstrapLog "ERROR: $errorMsg"
    } catch {
        Write-Error "The bootstrap failed before the log could be initialized. Check the path of the Start-TenantBootstrapInteractive.ps1 script and the parameters used by the wizard. Original error: $errorMsg"
    }
    throw
} finally {
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
}
