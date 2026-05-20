# IR-AdminFunctionsWeb

Interface web local (ASP.NET Core + React) para o motor PowerShell de backup/comparação/restore de funções administrativas do Microsoft Entra ID. Roda como Windows Service em `http://localhost:8080`.

## Arquitetura

```
Browser
   ↓ http://localhost:8080
IR.AdminFunctions.Web.exe (Windows Service, ASP.NET Core 8)
   ├─ /         → React SPA servido de wwwroot/
   ├─ /api/*    → REST
   └─ chama PowerShell via Microsoft.PowerShell.SDK
        ↓
   Scripts existentes em C:\ProgramData\Quest\
```

**Princípio**: a camada PowerShell em produção (`IR-AdministrativeFunctionBackup`, `IR-AdministrativeFunctionCompare`, `IR-AdministrativeFunctionRestore`) **não é alterada**, exceto pela adição do switch `-Headless` em `Start-AdministrativeFunctionCompare.ps1`. O motor de backup, manifest SHA-256, retry, log estruturado e RMAD mode continuam intactos.

## Estrutura do repositório

```
IR-AdminFunctionsWeb/
├── src/IR.AdminFunctions.Web/    # Projeto ASP.NET Core 8
│   ├── Program.cs
│   ├── Controllers/              # Endpoints REST
│   ├── Services/                 # PowerShellRunner, BackupReader, JobManager, ...
│   ├── Models/
│   ├── appsettings.json
│   └── wwwroot/                  # Onde o build React aterrissa
├── ui/                           # React + Vite + Tailwind
│   ├── src/
│   ├── package.json
│   └── vite.config.js
├── install/
│   ├── Install-Service.ps1
│   ├── Uninstall-Service.ps1
│   └── README.md
├── build.ps1                     # Build UI + publish .NET self-contained
└── IR-AdminFunctionsWeb.sln
```

## Desenvolvimento

```powershell
# Backend (.NET 8 SDK necessário)
dotnet run --project src/IR.AdminFunctions.Web

# Frontend em paralelo (Node 18+)
cd ui
npm install
npm run dev   # Vite em http://localhost:5173 com proxy para /api → :8080
```

## Build e empacotamento

```powershell
.\build.ps1
```

Gera `publish/IR.AdminFunctions.Web.exe` (single-file, self-contained, win-x64).

## Instalação como Windows Service

Ver `install/README.md`.

## Endpoints principais

| Método | Rota | Descrição |
|--------|------|-----------|
| GET    | /api/health                | Healthcheck |
| GET    | /api/settings              | Lê settings.json (thumbprint mascarado) |
| GET    | /api/tenants               | Lista tenants configurados |
| GET    | /api/backups               | Lista snapshots de backup |
| GET    | /api/backups/{id}          | Detalhe de um backup |
| POST   | /api/backups/run           | Dispara backup (job assíncrono) |
| GET    | /api/unpacked-objects?backupId=X | Lê roleDefinitions/roleAssignments |
| POST   | /api/compare/run           | Dispara comparação em modo headless (job) |
| GET    | /api/compare/results       | Último resultado de compare |
| POST   | /api/restore/preview       | WhatIf restore (job) |
| POST   | /api/restore/apply         | Apply restore (job) |
| GET    | /api/tasks                 | Lista tasks (logs + jobs) |
| GET    | /api/events                | Eventos a partir dos logs |
| GET    | /api/jobs/{id}             | Status de job |

Padrão de resposta:

```json
{ "success": true,  "data": { ... } }
{ "success": false, "error": "msg",   "details": "..." }
```

## Logs

`C:\ProgramData\Quest\IR-AdministrativeFunctionWeb\Logs\webapp-YYYY-MM-DD.log`

## Operações longas (backup/compare/restore)

Padrão de job assíncrono: POST retorna `{ jobId, status: "Queued" }` imediatamente; o frontend faz polling em `GET /api/jobs/{id}` a cada 2s; quando o status vira `Completed` ou `Failed`, exibe o resultado.
