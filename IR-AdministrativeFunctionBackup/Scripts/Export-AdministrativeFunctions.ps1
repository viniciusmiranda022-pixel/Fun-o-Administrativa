param(
    [Parameter(Mandatory = $true)]
    [string]$TenantId,

    [Parameter(Mandatory = $true)]
    [string]$ClientId,

    [Parameter(Mandatory = $true)]
    [string]$CertificateThumbprint,

    [Parameter(Mandatory = $false)]
    [string]$BasePath = "C:\ProgramData\Quest\IR-AdministrativeFunctionBackup"
)

$ErrorActionPreference = "Stop"

$logPath = Join-Path $BasePath "Logs"
$exportPath = Join-Path $BasePath "Exports"

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

function Connect-AppOnlyGraphWithRetry {
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
            Write-Log "App Registration em replicação - tentativa $attempt de $MaxAttempts."
            Connect-MgGraph -ClientId $ClientId -TenantId $TenantId -CertificateThumbprint $CertificateThumbprint -NoWelcome -ContextScope Process | Out-Null
            Get-MgRoleManagementDirectoryRoleDefinition -Top 1 | Out-Null
            Write-Log "Conexão app-only estabelecida na tentativa $attempt."
            return
        }
        catch {
            $errorMessage = $_.Exception.Message
            Write-Log "Tentativa $attempt falhou: $errorMessage"

            if ($attempt -eq $MaxAttempts) {
                throw "Falha ao conectar no Graph com app-only após 2 minutos (8 tentativas a cada 15s). Último erro: $errorMessage"
            }

            Write-Log "Aguardando $DelaySeconds segundos para nova tentativa..."
            Start-Sleep -Seconds $DelaySeconds
        }
    }
}

try {
    Write-Log "Iniciando conexão com Microsoft Graph"
    Connect-AppOnlyGraphWithRetry -ClientId $ClientId -TenantId $TenantId -CertificateThumbprint $CertificateThumbprint

    Write-Log "Coletando role definitions"
    $roleDefinitions = Get-MgRoleManagementDirectoryRoleDefinition -All

    Write-Log "Coletando role assignments"
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

    Write-Log "Export concluído com sucesso"
}
catch {
    Write-Log "ERRO: $($_.Exception.Message)"
    throw
}
finally {
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    Write-Log "Conexão encerrada"
}
