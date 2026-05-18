[CmdletBinding()]
param(
    [string]$SettingsPath = 'C:\ProgramData\Quest\IR-AdministrativeFunctionBackup\Config\settings.json',
    [int]$CertificateValidityMonths = 24,
    [switch]$SkipGraphTest
)

$ErrorActionPreference = 'Stop'

function Get-DefaultSettings {
    [ordered]@{
        TenantId = ''
        ClientId = ''
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

function Test-GraphConnection {
    param([string]$TenantId,[string]$ClientId,[string]$Thumbprint)
    Connect-MgGraph -ClientId $ClientId -TenantId $TenantId -CertificateThumbprint $Thumbprint -NoWelcome -ContextScope Process | Out-Null
    Get-MgRoleManagementDirectoryRoleDefinition -Top 1 | Out-Null
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = 'IR - Administrative Functions Backup | Configuração inicial'
$form.Width = 780
$form.Height = 600
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

$btnTest = New-Object System.Windows.Forms.Button
$btnTest.Text = 'Testar conexão Graph'
$btnTest.Left = 20
$btnTest.Top = 315
$btnTest.Width = 190
$form.Controls.Add($btnTest)

$btnSave = New-Object System.Windows.Forms.Button
$btnSave.Text = 'Salvar configuração'
$btnSave.Left = 230
$btnSave.Top = 315
$btnSave.Width = 190
$form.Controls.Add($btnSave)

$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text = 'Fechar'
$btnClose.Left = 440
$btnClose.Top = 315
$btnClose.Width = 120
$form.Controls.Add($btnClose)

$txtStatus = New-Object System.Windows.Forms.TextBox
$txtStatus.Left = 20
$txtStatus.Top = 360
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
        $friendlyName = 'IR-AdministrativeFunctionBackup Certificate'
        $cert = New-SelfSignedCertificate -DnsName 'IR-AdministrativeFunctionBackup' -FriendlyName $friendlyName -CertStoreLocation 'Cert:\LocalMachine\My' -NotAfter (Get-Date).AddMonths([int]$numMonths.Value) -KeyExportPolicy Exportable -KeyAlgorithm RSA -KeyLength 2048 -HashAlgorithm SHA256
        $txtThumb.Text = $cert.Thumbprint
        $cer = Export-PublicCertificate -Thumbprint $cert.Thumbprint
        Set-Status "Certificado criado em Cert:\LocalMachine\My. Thumbprint: $($cert.Thumbprint)"
        Set-Status "Certificado público exportado para: $cer"
        [System.Windows.Forms.MessageBox]::Show('Certificado criado com sucesso. Adicione o arquivo .cer na App Registration do Microsoft Entra ID (Certificates & secrets > Certificates), ou execute a automação dessa etapa com autenticação interativa e permissões adequadas.','Ação obrigatória', 'OK', 'Information') | Out-Null
    } catch {
        Set-Status "Erro ao criar certificado: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Erro ao criar certificado: $($_.Exception.Message)") | Out-Null
    }
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
