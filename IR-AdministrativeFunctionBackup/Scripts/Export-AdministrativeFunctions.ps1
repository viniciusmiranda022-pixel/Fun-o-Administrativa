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

try {
    Write-Log "Iniciando conexão com Microsoft Graph"
    Connect-MgGraph -ClientId $ClientId -TenantId $TenantId -CertificateThumbprint $CertificateThumbprint -NoWelcome

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
    Disconnect-MgGraph | Out-Null
    Write-Log "Conexão encerrada"
}