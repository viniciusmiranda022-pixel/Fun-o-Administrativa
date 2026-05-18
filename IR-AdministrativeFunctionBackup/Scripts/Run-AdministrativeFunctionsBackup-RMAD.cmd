@echo off
set PS64=%windir%\Sysnative\WindowsPowerShell\v1.0\powershell.exe

if not exist "%PS64%" (
    set PS64=%windir%\System32\WindowsPowerShell\v1.0\powershell.exe
)

"%PS64%" -NoProfile -ExecutionPolicy Bypass -File "C:\ProgramData\Quest\IR-AdministrativeFunctionBackup\Scripts\Run-AdministrativeFunctionsBackup.ps1"
exit /b %ERRORLEVEL%
