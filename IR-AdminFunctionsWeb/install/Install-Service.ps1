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
    throw "Publish folder not found: $PublishPath. Run 'dotnet publish' first."
}

if (-not $SkipProvision) {
    Write-Host "Provisioning PowerShell engine..." -ForegroundColor Cyan
    & (Join-Path $PSScriptRoot 'Provision-Scripts.ps1')
}

$exe = Join-Path $InstallPath 'IR.AdminFunctions.Web.exe'

Write-Host "Copying binaries to $InstallPath..." -ForegroundColor Cyan
New-Item -ItemType Directory -Path $InstallPath -Force | Out-Null
Copy-Item -Path (Join-Path $PublishPath '*') -Destination $InstallPath -Recurse -Force

$existing = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
if ($existing) {
    Write-Host "Service already exists; stopping and removing for reinstall..." -ForegroundColor Yellow
    if ($existing.Status -ne 'Stopped') {
        Stop-Service -Name $ServiceName -Force
    }
    sc.exe delete $ServiceName | Out-Null
    Start-Sleep -Seconds 2
}

Write-Host "Creating service $ServiceName..." -ForegroundColor Cyan
& sc.exe create $ServiceName binPath= "`"$exe`"" start= auto obj= $ServiceAccount DisplayName= 'IR Administrative Functions Web' | Out-Null
& sc.exe description $ServiceName 'Interface web para backup, comparacao e restore de funcoes administrativas do Microsoft Entra ID.' | Out-Null
& sc.exe failure $ServiceName reset= 86400 actions= restart/5000/restart/10000/restart/30000 | Out-Null

Write-Host "Starting service..." -ForegroundColor Cyan
Start-Service -Name $ServiceName

Write-Host ""
Write-Host "Installation complete." -ForegroundColor Green
Write-Host "URL: http://localhost:8080" -ForegroundColor Green
Write-Host ""
Write-Host "Note: the account '$ServiceAccount' needs access to the certificate private key at Cert:\LocalMachine\My\." -ForegroundColor Yellow
Write-Host "Use certlm.msc -> certificate -> All Tasks -> Manage Private Keys to grant access." -ForegroundColor Yellow
