param(
    [Parameter(Mandatory = $true)]
    [string]$TenantId,

    [Parameter(Mandatory = $true)]
    [string]$ClientId,

    [Parameter(Mandatory = $true)]
    [string]$CertificateThumbprint,

    [Parameter(Mandatory = $false)]
    [string]$BasePath = (Join-Path "C:\ProgramData\Quest" "IR-AdministrativeFunctionBackup"),

    [Parameter(Mandatory = $false)]
    [int]$RetentionDays = 30
)

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

try {
    chcp 65001 | Out-Null
} catch {}

$ErrorActionPreference = "Stop"

$modulePath = Join-Path $PSScriptRoot "IR-AdministrativeFunctions.psm1"
Import-Module $modulePath -Force

$logPath = Join-Path $BasePath "Logs"
$exportPath = Join-Path $BasePath "Backups"

New-Item -ItemType Directory -Path $logPath -Force | Out-Null
New-Item -ItemType Directory -Path $exportPath -Force | Out-Null

$runStamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$todayFolder = Join-Path $exportPath $runStamp
New-Item -ItemType Directory -Path $todayFolder -Force | Out-Null

$logFile = Join-Path $logPath "Export-AdministrativeFunctions-$runStamp.log"

function Write-Log {
    param([string]$Message)
    $line = "[{0}] {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Message
    $line | Tee-Object -FilePath $logFile -Append
}

try {
    Write-Log "Iniciando conexão com Microsoft Graph"
    Connect-AppOnlyGraphWithRetry -ClientId $ClientId -TenantId $TenantId -Thumbprint $CertificateThumbprint -VerboseOutput -FailureMessage 'Falha ao conectar ao Graph em modo app-only após {0} tentativas. Último erro: {1}'

    Write-Log "Coletando definições de função"
    $roleDefinitions = Get-MgRoleManagementDirectoryRoleDefinition -All

    Write-Log "Coletando atribuições de função"
    $roleAssignments = Get-MgRoleManagementDirectoryRoleAssignment -All

    $roleDefinitionsPath = Join-Path $todayFolder "roleDefinitions.json"
    $roleAssignmentsPath = Join-Path $todayFolder "roleAssignments.json"

    $roleDefinitions | ConvertTo-Json -Depth 30 | Set-Content -Path $roleDefinitionsPath -Encoding UTF8
    $roleAssignments | ConvertTo-Json -Depth 30 | Set-Content -Path $roleAssignmentsPath -Encoding UTF8

    $manifest = [ordered]@{
        CollectedAt          = (Get-Date).ToString("o")
        TenantId             = $TenantId
        ClientId             = $ClientId
        RoleDefinitionsCount = @($roleDefinitions).Count
        RoleAssignmentsCount = @($roleAssignments).Count
        RoleDefinitionsHash  = (Get-FileHash $roleDefinitionsPath -Algorithm SHA256).Hash
        RoleAssignmentsHash  = (Get-FileHash $roleAssignmentsPath -Algorithm SHA256).Hash
    }

    $manifestPath = Join-Path $todayFolder "manifest.json"
    $manifest | ConvertTo-Json -Depth 10 | Set-Content -Path $manifestPath -Encoding UTF8

    if ($RetentionDays -gt 0) {
        $cutoffDate = (Get-Date).AddDays(-$RetentionDays)
        $oldBackups = Get-ChildItem -Path $exportPath -Directory | Where-Object { $_.LastWriteTime -lt $cutoffDate }
        foreach ($backupFolder in $oldBackups) {
            Remove-Item -Path $backupFolder.FullName -Recurse -Force
            Write-Log "Backup antigo removido por retenção ($RetentionDays dias): $($backupFolder.FullName)"
        }
    }

    Write-Log "Exportação concluída com sucesso"
}
catch {
    Write-Log "ERRO: $($_.Exception.Message)"
    throw
}
finally {
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    Write-Log "Conexão encerrada"
}
