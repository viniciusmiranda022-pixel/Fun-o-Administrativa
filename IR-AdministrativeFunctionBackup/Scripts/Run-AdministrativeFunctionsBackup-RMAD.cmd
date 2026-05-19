@echo off
set PS64=%windir%\Sysnative\WindowsPowerShell\v1.0\powershell.exe

if not exist "%PS64%" (
    set PS64=%windir%\System32\WindowsPowerShell\v1.0\powershell.exe
)

set QUEST_ROOT=C:\ProgramData\Quest
set BACKUP_ROOT=%QUEST_ROOT%\IR-AdministrativeFunctionBackup

"%PS64%" -NoProfile -ExecutionPolicy Bypass -File "%BACKUP_ROOT%\Scripts\Run-AdministrativeFunctionsBackup.ps1"
exit /b %ERRORLEVEL%
