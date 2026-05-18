[CmdletBinding()]
param(
    [string]$SettingsPath = 'C:\ProgramData\Quest\IR-AdministrativeFunctionBackup\Config\settings.json',
    [int]$CertificateValidityMonths = 24,
    [switch]$SkipGraphTest
)

$ErrorActionPreference = 'Stop'

$AppDisplayName = 'Quest Recovery Function - Administrative Roles Backup'
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

function Get-DefaultSettings {
    [ordered]@{
        TenantId = ''
        ClientId = ''
        AppDisplayName = $AppDisplayName
        CertificateThumbprint = ''
        CertificateStore = 'LocalMachine\\My'
        BackupRoot = 'C:\ProgramData\Quest\IR-AdministrativeFunctionBackup\Backups'
        LogRoot = 'C:\ProgramData\Quest\IR-AdministrativeFunctionBackup\Logs'
        LogoPath = 'C:\ProgramData\Quest\IR-AdministrativeFunctionBackup\Assets\quest-logo.png'
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

function Test-GraphConnection {
    param([string]$TenantId,[string]$ClientId,[string]$Thumbprint)
    Connect-MgGraph -ClientId $ClientId -TenantId $TenantId -CertificateThumbprint $Thumbprint -NoWelcome -ContextScope Process | Out-Null
    Get-MgRoleManagementDirectoryRoleDefinition -Top 1 | Out-Null
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
}

function Add-CertificateToApplication {
    param(
        [string]$ApplicationId,
        [System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate
    )

    $existing = Get-MgApplication -ApplicationId $ApplicationId -Property KeyCredentials
    $base64Value = [System.Convert]::ToBase64String($Certificate.RawData)
    $match = $existing.KeyCredentials | Where-Object { $_.CustomKeyIdentifier -and ([System.Convert]::ToBase64String($_.CustomKeyIdentifier) -eq [System.Convert]::ToBase64String($Certificate.GetCertHash())) }

    if ($match) { return $false }

    $params = @{
        keyCredential = @{
            type = 'AsymmetricX509Cert'
            usage = 'Verify'
            key = $base64Value
            displayName = 'Primary certificate'
        }
    }

    Add-MgApplicationKey -ApplicationId $ApplicationId -BodyParameter $params | Out-Null
    return $true
}

function Ensure-GraphApplicationPermissions {
    param(
        [string]$ApplicationId,
        [string]$ServicePrincipalId,
        [scriptblock]$StatusWriter
    )

    $graphSp = Get-MgServicePrincipal -Filter "appId eq '$GraphAppId'"
    $currentApp = Get-MgApplication -ApplicationId $ApplicationId -Property RequiredResourceAccess
    $currentAccess = @($currentApp.RequiredResourceAccess | Where-Object { $_.ResourceAppId -eq $GraphAppId } | ForEach-Object { $_.ResourceAccess })

    $requiredRoles = @($RequiredApplicationPermissions + $OptionalApplicationPermissions)
    $roleMap = @{}
    foreach ($value in $requiredRoles) {
        $role = $graphSp.AppRoles | Where-Object { $_.Value -eq $value -and $_.AllowedMemberTypes -contains 'Application' }
        if ($role) {
            $roleMap[$value] = $role.Id
        } else {
            & $StatusWriter "Permissão $value não encontrada nos appRoles do Graph."
        }
    }

    $merged = @{}
    foreach ($entry in $currentAccess) { $merged[$entry.Id] = 'Role' }
    foreach ($roleId in $roleMap.Values) { $merged[$roleId] = 'Role' }

    $resourceAccess = @()
    foreach ($k in $merged.Keys) {
        $resourceAccess += @{ Id = $k; Type = 'Role' }
    }

    Update-MgApplication -ApplicationId $ApplicationId -RequiredResourceAccess @(@{ ResourceAppId = $GraphAppId; ResourceAccess = $resourceAccess })
    & $StatusWriter 'Permissões de aplicação do Microsoft Graph configuradas no App Registration.'

    $existingAssignments = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $ServicePrincipalId -All
    foreach ($pair in $roleMap.GetEnumerator()) {
        $already = $existingAssignments | Where-Object { $_.ResourceId -eq $graphSp.Id -and $_.AppRoleId -eq $pair.Value }
        if (-not $already) {
            New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $ServicePrincipalId -PrincipalId $ServicePrincipalId -ResourceId $graphSp.Id -AppRoleId $pair.Value | Out-Null
            & $StatusWriter "Admin consent concedido automaticamente para: $($pair.Key)"
        }
    }
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = 'IR - Administrative Functions Backup | Configuração inicial'
$form.Width = 780
$form.Height = 640
$form.StartPosition = 'CenterScreen'

$lblTenant = New-Object System.Windows.Forms.Label
$lblTenant.Text = 'Tenant ID:'
$lblTenant.Left = 20
$lblTenant.Top = 20
$lblTenant.Width = 150
$form.Controls.Add($lblTenant)

$txtTenant = New-Object System.Windows.Forms.TextBox
$txtTenant.Left = 20
$txtTenant.Top = 45
$txtTenant.Width = 720
$form.Controls.Add($txtTenant)

$lblClient = New-Object System.Windows.Forms.Label
$lblClient.Text = 'Client ID (App Registration):'
$lblClient.Left = 20
$lblClient.Top = 80
$lblClient.Width = 250
$form.Controls.Add($lblClient)

$txtClient = New-Object System.Windows.Forms.TextBox
$txtClient.Left = 20
$txtClient.Top = 105
$txtClient.Width = 720
$form.Controls.Add($txtClient)

$lblMonths = New-Object System.Windows.Forms.Label
$lblMonths.Text = 'Validade do certificado (meses):'
$lblMonths.Left = 20
$lblMonths.Top = 140
$lblMonths.Width = 260
$form.Controls.Add($lblMonths)

$numMonths = New-Object System.Windows.Forms.NumericUpDown
$numMonths.Left = 20
$numMonths.Top = 165
$numMonths.Width = 120
$numMonths.Minimum = 1
$numMonths.Maximum = 120
$numMonths.Value = $CertificateValidityMonths
$form.Controls.Add($numMonths)

$lblThumb = New-Object System.Windows.Forms.Label
$lblThumb.Text = 'Thumbprint selecionado:'
$lblThumb.Left = 20
$lblThumb.Top = 205
$lblThumb.Width = 200
$form.Controls.Add($lblThumb)

$txtThumb = New-Object System.Windows.Forms.TextBox
$txtThumb.Left = 20
$txtThumb.Top = 230
$txtThumb.Width = 720
$form.Controls.Add($txtThumb)

$btnCreate = New-Object System.Windows.Forms.Button
$btnCreate.Text = 'Criar certificado self-signed'
$btnCreate.Left = 20
$btnCreate.Top = 270
$btnCreate.Width = 240
$form.Controls.Add($btnCreate)

$btnImport = New-Object System.Windows.Forms.Button
$btnImport.Text = 'Importar .pfx existente'
$btnImport.Left = 280
$btnImport.Top = 270
$btnImport.Width = 180
$form.Controls.Add($btnImport)

$btnExportCer = New-Object System.Windows.Forms.Button
$btnExportCer.Text = 'Exportar .cer público'
$btnExportCer.Left = 480
$btnExportCer.Top = 270
$btnExportCer.Width = 170
$form.Controls.Add($btnExportCer)

$btnAuto = New-Object System.Windows.Forms.Button
$btnAuto.Text = 'Configurar automaticamente no tenant'
$btnAuto.Left = 20
$btnAuto.Top = 315
$btnAuto.Width = 300
$form.Controls.Add($btnAuto)

$btnTest = New-Object System.Windows.Forms.Button
$btnTest.Text = 'Testar conexão Graph'
$btnTest.Left = 340
$btnTest.Top = 315
$btnTest.Width = 190
$form.Controls.Add($btnTest)

$btnSave = New-Object System.Windows.Forms.Button
$btnSave.Text = 'Salvar configuração'
$btnSave.Left = 20
$btnSave.Top = 355
$btnSave.Width = 190
$form.Controls.Add($btnSave)

$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text = 'Fechar'
$btnClose.Left = 230
$btnClose.Top = 355
$btnClose.Width = 120
$form.Controls.Add($btnClose)

$txtStatus = New-Object System.Windows.Forms.TextBox
$txtStatus.Left = 20
$txtStatus.Top = 400
$txtStatus.Width = 720
$txtStatus.Height = 180
$txtStatus.Multiline = $true
$txtStatus.ScrollBars = 'Vertical'
$txtStatus.ReadOnly = $true
$form.Controls.Add($txtStatus)

function Set-Status { param([string]$Message) $txtStatus.AppendText("$(Get-Date -Format 'HH:mm:ss') - $Message`r`n") }

$default = Get-DefaultSettings
if (Test-Path $SettingsPath) {
    try {
        $existing = Get-Content -Path $SettingsPath -Raw | ConvertFrom-Json
        $txtTenant.Text = $existing.TenantId
        $txtClient.Text = $existing.ClientId
        $txtThumb.Text = $existing.CertificateThumbprint
        Set-Status "Configuração existente carregada: $SettingsPath"
    } catch {
        Set-Status "Falha ao ler configuração existente: $($_.Exception.Message)"
    }
}

$btnCreate.Add_Click({
    try {
        $friendlyName = 'Quest Recovery Function - Administrative Roles Backup'
        $cert = New-SelfSignedCertificate -DnsName 'IR-AdministrativeFunctionBackup' -FriendlyName $friendlyName -CertStoreLocation 'Cert:\LocalMachine\My' -NotAfter (Get-Date).AddMonths([int]$numMonths.Value) -KeyExportPolicy Exportable -KeyAlgorithm RSA -KeyLength 2048 -HashAlgorithm SHA256
        $txtThumb.Text = $cert.Thumbprint
        $cer = Export-PublicCertificate -Thumbprint $cert.Thumbprint
        Set-Status "Certificado criado em Cert:\LocalMachine\My. Thumbprint: $($cert.Thumbprint)"
        Set-Status "Certificado público exportado para: $cer"
    } catch {
        Set-Status "Erro ao criar certificado: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Erro ao criar certificado: $($_.Exception.Message)") | Out-Null
    }
})


function Get-BootstrapInsufficientPrivilegeMessage {
    return 'A autenticação foi concluída, mas a conta usada não possui permissões suficientes para configurar o App Registration. Execute novamente com uma conta Global Administrator ou uma conta com permissões para gerenciar App Registrations, Service Principals e consentimento administrativo do Microsoft Graph.'
}

function Test-IsBootstrapPrivilegeError {
    param([string]$Message)
    if ([string]::IsNullOrWhiteSpace($Message)) { return $false }
    return ($Message -match 'Authorization_RequestDenied' -or $Message -match 'Insufficient privileges' -or $Message -match 'não possui permissões suficientes para configurar o App Registration')
}

function Start-AutomaticTenantSetup {
    try {
        $bootstrapScript = Join-Path $PSScriptRoot 'Start-TenantBootstrapInteractive.ps1'
        if (-not (Test-Path $bootstrapScript)) { throw "Script não encontrado: $bootstrapScript" }

        $tenantId = $txtTenant.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($tenantId)) {
            throw 'Tenant ID é obrigatório para iniciar o bootstrap automático.'
        }

        Set-Status 'Iniciando processo externo para autenticação Device Code e bootstrap do tenant...'

        $proc = Start-Process -FilePath 'powershell.exe' -ArgumentList @(
            '-NoProfile',
            '-ExecutionPolicy','Bypass',
            '-File',$bootstrapScript,
            '-SettingsPath',$SettingsPath,
            '-CertificateValidityMonths',$([int]$numMonths.Value),
            '-TenantId',$tenantId
        ) -PassThru -WindowStyle Normal
        $proc.WaitForExit()

        if ($proc.ExitCode -ne 0) {
            throw "Bootstrap do tenant falhou com código $($proc.ExitCode)."
        }

        Set-Status 'Processo externo finalizado. Recarregando settings.json...'
        if (-not (Test-Path $SettingsPath)) { throw "Arquivo de configuração não encontrado em: $SettingsPath" }

        $settings = Get-Content -Path $SettingsPath -Raw | ConvertFrom-Json
        $txtTenant.Text = $settings.TenantId
        $txtClient.Text = $settings.ClientId
        $txtThumb.Text = $settings.CertificateThumbprint

        Test-GraphConnection -TenantId $settings.TenantId -ClientId $settings.ClientId -Thumbprint $settings.CertificateThumbprint
        Set-Status 'Configuração concluída com sucesso'
        [System.Windows.Forms.MessageBox]::Show('Configuração concluída com sucesso','Sucesso','OK','Information') | Out-Null
    } catch {
        $errorMessage = $_.Exception.Message
        if (Test-IsBootstrapPrivilegeError -Message $errorMessage) {
            $friendly = Get-BootstrapInsufficientPrivilegeMessage
            Set-Status $friendly
            [System.Windows.Forms.MessageBox]::Show($friendly, 'Permissões insuficientes', 'OK', 'Warning') | Out-Null
        } else {
            Set-Status "Erro na configuração automática: $errorMessage"
            [System.Windows.Forms.MessageBox]::Show("Erro na configuração automática: $errorMessage", 'Erro', 'OK', 'Error') | Out-Null
        }
    }
}

$btnAuto.Add_Click({
    Start-AutomaticTenantSetup
})

$btnImport.Add_Click({
    try {
        $dialog = New-Object System.Windows.Forms.OpenFileDialog
        $dialog.Filter = 'PFX files (*.pfx)|*.pfx'
        if ($dialog.ShowDialog() -ne 'OK') { return }
        $secure = Read-Host 'Digite a senha do PFX' -AsSecureString
        $cert = Import-PfxCertificate -FilePath $dialog.FileName -CertStoreLocation 'Cert:\LocalMachine\My' -Password $secure -Exportable
        $txtThumb.Text = $cert.Thumbprint
        $cer = Export-PublicCertificate -Thumbprint $cert.Thumbprint
        Set-Status "PFX importado com sucesso. Thumbprint: $($cert.Thumbprint)"
        Set-Status "Certificado público exportado para: $cer"
    } catch {
        Set-Status "Erro ao importar PFX: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Erro ao importar PFX: $($_.Exception.Message)") | Out-Null
    }
})

$btnExportCer.Add_Click({
    try {
        if (-not $txtThumb.Text) { throw 'Informe/seleciona um thumbprint primeiro.' }
        $cer = Export-PublicCertificate -Thumbprint $txtThumb.Text.Trim()
        Set-Status "Certificado público exportado para: $cer"
    } catch {
        Set-Status "Erro ao exportar CER: $($_.Exception.Message)"
    }
})

$btnTest.Add_Click({
    try {
        Test-GraphConnection -TenantId $txtTenant.Text.Trim() -ClientId $txtClient.Text.Trim() -Thumbprint $txtThumb.Text.Trim()
        Set-Status 'Conexão com Microsoft Graph testada com sucesso.'
        [System.Windows.Forms.MessageBox]::Show('Conexão com Microsoft Graph realizada com sucesso.') | Out-Null
    } catch {
        Set-Status "Falha no teste de conexão: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Falha no teste de conexão: $($_.Exception.Message)") | Out-Null
    }
})

$btnSave.Add_Click({
    try {
        $s = Get-DefaultSettings
        $s.TenantId = $txtTenant.Text.Trim()
        $s.ClientId = $txtClient.Text.Trim()
        $s.CertificateThumbprint = $txtThumb.Text.Trim()
        Save-Settings -Settings $s -Path $SettingsPath
        Set-Status "Configuração salva em: $SettingsPath"
        if (-not $SkipGraphTest) { Test-GraphConnection -TenantId $s.TenantId -ClientId $s.ClientId -Thumbprint $s.CertificateThumbprint; Set-Status 'Validação final de conexão concluída com sucesso.' }
        [System.Windows.Forms.MessageBox]::Show('Configuração concluída com sucesso.','Sucesso', 'OK', 'Information') | Out-Null
    } catch {
        Set-Status "Erro ao salvar/validar configuração: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Erro ao salvar/validar configuração: $($_.Exception.Message)") | Out-Null
    }
})

$btnClose.Add_Click({ $form.Close() })

[void]$form.ShowDialog()
