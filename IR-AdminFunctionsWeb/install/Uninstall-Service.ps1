[CmdletBinding()]
param(
    [string]$ServiceName = 'IR-AdminFunctionsWeb',
    [string]$InstallPath = 'C:\ProgramData\Quest\IR-AdministrativeFunctionWeb',
    [switch]$KeepFiles
)

$ErrorActionPreference = 'Stop'

$existing = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
if ($existing) {
    if ($existing.Status -ne 'Stopped') {
        Write-Host "Stopping service..." -ForegroundColor Cyan
        Stop-Service -Name $ServiceName -Force
    }
    Write-Host "Removing service..." -ForegroundColor Cyan
    sc.exe delete $ServiceName | Out-Null
} else {
    Write-Host "Service $ServiceName is not installed." -ForegroundColor Yellow
}

if (-not $KeepFiles -and (Test-Path $InstallPath)) {
    Write-Host "Removing binaries at $InstallPath..." -ForegroundColor Cyan
    Remove-Item -Path $InstallPath -Recurse -Force
}

Write-Host "Uninstallation complete." -ForegroundColor Green
