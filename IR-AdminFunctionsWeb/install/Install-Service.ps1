[CmdletBinding()]
param(
    [string]$InstallPath = 'C:\ProgramData\Quest\IR-AdministrativeFunctionWeb',
    [string]$PublishPath = (Join-Path $PSScriptRoot '..\publish'),
    [string]$ServiceName = 'IR-AdminFunctionsWeb',
    [string]$ServiceAccount = 'NT AUTHORITY\NetworkService',
    [switch]$SkipProvision
)

$ErrorActionPreference = 'Stop'

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

if (-not (Test-Path $PublishPath)) {
    throw "Pasta publish não encontrada: $PublishPath. Execute 'dotnet publish' antes."
}

if (-not $SkipProvision) {
    Write-Host "Provisionando motor PowerShell..." -ForegroundColor Cyan
    & (Join-Path $PSScriptRoot 'Provision-Scripts.ps1')
}

$exe = Join-Path $InstallPath 'IR.AdminFunctions.Web.exe'

Write-Host "Copiando binários para $InstallPath..." -ForegroundColor Cyan
New-Item -ItemType Directory -Path $InstallPath -Force | Out-Null
Copy-Item -Path (Join-Path $PublishPath '*') -Destination $InstallPath -Recurse -Force

$existing = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
if ($existing) {
    Write-Host "Serviço já existe; parando e removendo para reinstalar..." -ForegroundColor Yellow
    if ($existing.Status -ne 'Stopped') {
        Stop-Service -Name $ServiceName -Force
    }
    sc.exe delete $ServiceName | Out-Null
    Start-Sleep -Seconds 2
}

Write-Host "Criando serviço $ServiceName..." -ForegroundColor Cyan
& sc.exe create $ServiceName binPath= "`"$exe`"" start= auto obj= $ServiceAccount DisplayName= 'IR Administrative Functions Web' | Out-Null
& sc.exe description $ServiceName 'Interface web para backup, comparacao e restore de funcoes administrativas do Microsoft Entra ID.' | Out-Null
& sc.exe failure $ServiceName reset= 86400 actions= restart/5000/restart/10000/restart/30000 | Out-Null

Write-Host "Iniciando serviço..." -ForegroundColor Cyan
Start-Service -Name $ServiceName

Write-Host ""
Write-Host "Instalação concluída." -ForegroundColor Green
Write-Host "URL: http://localhost:8080" -ForegroundColor Green
Write-Host ""
Write-Host "Atenção: a conta '$ServiceAccount' precisa de acesso à chave privada do certificado em Cert:\LocalMachine\My\." -ForegroundColor Yellow
Write-Host "Use certlm.msc -> certificado -> All Tasks -> Manage Private Keys para conceder permissões." -ForegroundColor Yellow
