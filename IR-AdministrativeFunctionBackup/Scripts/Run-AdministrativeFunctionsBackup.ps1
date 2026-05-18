[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

try {
    chcp 65001 | Out-Null
} catch {}

$ErrorActionPreference = 'Stop'

$basePath = 'C:\ProgramData\Quest\IR-AdministrativeFunctionBackup'
$configPath = Join-Path $basePath 'Config\settings.json'
$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$setupScript = Join-Path $scriptRoot 'Setup-AdministrativeFunctionsBackup.ps1'
$exportScript = Join-Path $scriptRoot 'Export-AdministrativeFunctions.ps1'

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
            $reasons.Add('Certificado configurado não existe em Cert:\LocalMachine\My.')
        }
        else {
            if (-not $cert.HasPrivateKey) { $reasons.Add('Certificado sem chave privada (HasPrivateKey=False).') }
            if ($cert.NotAfter -le (Get-Date)) { $reasons.Add("Certificado expirado em $($cert.NotAfter.ToString('yyyy-MM-dd HH:mm:ss')).") }
        }
    }

    return [pscustomobject]@{ IsValid = ($reasons.Count -eq 0); Reasons = $reasons; Settings = $settings }
}

function Test-GraphConnection {
    param([string]$TenantId,[string]$ClientId,[string]$Thumbprint)
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $Thumbprint -NoWelcome | Out-Null
    Get-MgRoleManagementDirectoryRoleDefinition -All | Select-Object -First 1 | Out-Null
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
}

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
