$TenantId = "ca9b03ea-578e-4277-b684-969fa2a34a9a"
$ClientId = "3c1e0342-a7c1-434e-bd70-04ace3dfd88d"
$Thumb = "FDD65BA99EB803D11BFBDD1F02424EBA39CBC91B"
$SnapshotFolder = "C:\ProgramData\Quest\IR-AdministrativeFunctionBackup\Exports\2026-04-18_14-36-05"
$RoleName = "POC App Registration Operator"

powershell.exe -ExecutionPolicy Bypass -File "C:\ProgramData\Quest\IR-AdministrativeFunctionRestore\Restore-CustomRole-Full.ps1" `
  -TenantId $TenantId `
  -ClientId $ClientId `
  -CertificateThumbprint $Thumb `
  -SnapshotFolder $SnapshotFolder `
  -RoleName $RoleName