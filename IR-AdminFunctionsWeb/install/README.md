# Instalação do IR-AdminFunctionsWeb

## Pré-requisitos

1. Windows Server 2019+ (x64).
2. Motor PowerShell de backup já instalado em `C:\ProgramData\Quest\IR-AdministrativeFunction*`.
3. `settings.json` configurado em `C:\ProgramData\Quest\IR-AdministrativeFunctionBackup\Config\settings.json`.
4. Certificado com chave privada em `Cert:\LocalMachine\My\` correspondente ao thumbprint do `settings.json`.

## Passos

1. **Build do pacote** (em qualquer máquina com .NET 8 SDK e Node 18+):

   ```powershell
   .\build.ps1
   ```

   Isso gera `publish\IR.AdminFunctions.Web.exe` self-contained.

2. **Copie a pasta `publish\`** para o servidor de destino (junto com `install\`).

3. **Execute como administrador** no servidor:

   ```powershell
   .\install\Install-Service.ps1
   ```

4. **Conceda acesso à chave privada** do certificado para a conta do serviço (`NT AUTHORITY\NetworkService` por padrão):
   - Abra `certlm.msc`
   - Vá em `Pessoal` → `Certificados`
   - Selecione o certificado da app de backup
   - Clique com botão direito → `Todas as tarefas` → `Gerenciar Chaves Privadas`
   - Adicione `NetworkService` com leitura

5. Acesse `http://localhost:8080` no navegador.

## Desinstalar

```powershell
.\install\Uninstall-Service.ps1
```

## Solução de problemas

- Logs do web app: `C:\ProgramData\Quest\IR-AdministrativeFunctionWeb\Logs\webapp-YYYY-MM-DD.log`
- Status do serviço: `Get-Service IR-AdminFunctionsWeb`
- Healthcheck: `Invoke-RestMethod http://localhost:8080/api/health`
