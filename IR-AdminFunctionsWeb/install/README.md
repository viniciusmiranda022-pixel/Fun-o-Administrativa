# Instalação do IR-AdminFunctionsWeb

## Pré-requisitos

1. Windows Server 2019+ (x64).
2. Certificado com chave privada em `Cert:\LocalMachine\My\` para autenticação no Microsoft Graph.

> O motor PowerShell (`IR-AdministrativeFunction*`) e um `settings.json` template são provisionados automaticamente em `C:\ProgramData\Quest\` pelo `Install-Service.ps1`. Após a instalação, edite `settings.json` com os dados reais do seu Entra ID (TenantId, ClientId, Thumbprint).

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

5. **Edite o `settings.json`** criado em `C:\ProgramData\Quest\IR-AdministrativeFunctionBackup\Config\settings.json` substituindo todos os campos `PREENCHER-*` pelos valores reais do seu Entra ID.

6. Reinicie o serviço: `Restart-Service IR-AdminFunctionsWeb`.

7. Acesse `http://localhost:8080` no navegador.

## Provisionamento manual (modo dev)

Se você estiver rodando em modo desenvolvimento (`dotnet run`) e quiser apenas criar a estrutura de pastas e copiar os scripts, sem instalar o serviço:

```powershell
.\install\Provision-Scripts.ps1
```

Use `-Force` para sobrescrever scripts e settings.json existentes.

## Desinstalar

```powershell
.\install\Uninstall-Service.ps1
```

## Solução de problemas

- Logs do web app: `C:\ProgramData\Quest\IR-AdministrativeFunctionWeb\Logs\webapp-YYYY-MM-DD.log`
- Status do serviço: `Get-Service IR-AdminFunctionsWeb`
- Healthcheck: `Invoke-RestMethod http://localhost:8080/api/health`
