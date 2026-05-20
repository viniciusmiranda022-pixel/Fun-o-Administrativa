using System.Collections.Concurrent;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using IR.AdminFunctions.Web.Models;

namespace IR.AdminFunctions.Web.Services;

// Bootstrap auto-registra a aplicação multi-tenant no Azure AD usando
// device code flow (Microsoft Graph PowerShell well-known ClientId).
// O admin global autentica no navegador/celular com o user code e a sessão
// resultante é usada para criar o App Registration definitivo.
public class AppRegistrationService
{
    private readonly AppConfigStore _store;
    private readonly ILogger<AppRegistrationService> _logger;
    private readonly ConcurrentDictionary<string, SetupStatus> _states = new();
    private const string AppDisplayName = "IR Administrative Function Recovery";

    // Microsoft Graph permission IDs (constantes globais do tenant da Microsoft)
    private const string GraphResourceAppId = "00000003-0000-0000-c000-000000000000";
    private const string RoleManagementReadDirectory = "483bed4a-2ad3-4361-a73b-c83ccdbdc53c";       // Application
    private const string RoleManagementReadWriteDirectory = "9e3f62cf-ca93-4989-b6ce-bf83c28f9fe8"; // Application

    public AppRegistrationService(AppConfigStore store, ILogger<AppRegistrationService> logger)
    {
        _store = store;
        _logger = logger;
    }

    public SetupStatus GetStatus(string sessionId) =>
        _states.TryGetValue(sessionId, out var s) ? s : new SetupStatus { Status = "NotStarted" };

    public string StartSetup()
    {
        var sessionId = Guid.NewGuid().ToString("N");
        _states[sessionId] = new SetupStatus { Status = "WaitingForUser", Message = "Aguardando usuário autenticar..." };
        _logger.LogInformation("[Setup:{SessionId}] Setup iniciado. Task em background lançada.", sessionId);
        _ = Task.Run(() => RunSetupAsync(sessionId));
        // Retorna imediatamente — o frontend faz polling via /setup/status/{sessionId}
        return sessionId;
    }

    private async Task RunSetupAsync(string sessionId)
    {
        try
        {
            _logger.LogInformation("[Setup:{SessionId}] Abrindo runspace PowerShell...", sessionId);
            using var rs = RunspaceFactory.CreateRunspace();
            rs.Open();
            _logger.LogInformation("[Setup:{SessionId}] Runspace aberto.", sessionId);

            using var ps = PowerShell.Create();
            ps.Runspace = rs;

            // Captura Information, Warning e Verbose para extrair o device code
            ps.Streams.Information.DataAdded += (s, e) =>
            {
                var record = ((PSDataCollection<InformationRecord>)s!)[e.Index];
                var msg = record?.MessageData?.ToString() ?? "";
                _logger.LogInformation("[Setup:{SessionId}] [PS:Info] {Msg}", sessionId, msg);
                TryExtractUserCode(sessionId, msg);
            };
            ps.Streams.Warning.DataAdded += (s, e) =>
            {
                var record = ((PSDataCollection<WarningRecord>)s!)[e.Index];
                var msg = record?.Message ?? "";
                _logger.LogWarning("[Setup:{SessionId}] [PS:Warn] {Msg}", sessionId, msg);
                TryExtractUserCode(sessionId, msg);
            };
            ps.Streams.Verbose.DataAdded += (s, e) =>
            {
                var record = ((PSDataCollection<VerboseRecord>)s!)[e.Index];
                var msg = record?.Message ?? "";
                _logger.LogDebug("[Setup:{SessionId}] [PS:Verbose] {Msg}", sessionId, msg);
                TryExtractUserCode(sessionId, msg);
            };
            ps.Streams.Error.DataAdded += (s, e) =>
            {
                var record = ((PSDataCollection<ErrorRecord>)s!)[e.Index];
                _logger.LogError("[Setup:{SessionId}] [PS:Error] {Msg}", sessionId, record?.ToString() ?? "");
            };

            _logger.LogInformation("[Setup:{SessionId}] Etapa 1/3 — Importando módulos e iniciando Connect-MgGraph...", sessionId);

            // 1) Connect-MgGraph com device code e scopes necessários para criar o app
            ps.AddScript(@"
                [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
                Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
                Import-Module Microsoft.Graph.Applications -ErrorAction Stop
                Connect-MgGraph -Scopes 'Application.ReadWrite.All','Directory.ReadWrite.All','User.Read' -UseDeviceCode -NoWelcome
                $ctx = Get-MgContext
                Write-Output $ctx.TenantId
                Write-Output $ctx.Account
            ");

            var results = await Task.Run(() => ps.Invoke());
            if (ps.HadErrors)
            {
                var err = string.Join("; ", ps.Streams.Error.Select(e => e.ToString()));
                _logger.LogError("[Setup:{SessionId}] Falha no Connect-MgGraph: {Err}", sessionId, err);
                Fail(sessionId, $"Falha no Connect-MgGraph: {err}");
                return;
            }

            var tenantId = results.FirstOrDefault()?.BaseObject?.ToString();
            var account = results.Skip(1).FirstOrDefault()?.BaseObject?.ToString();
            _logger.LogInformation("[Setup:{SessionId}] Etapa 1/3 concluída. TenantId={TenantId} Account={Account}",
                sessionId, tenantId ?? "(nulo)", account ?? "(nulo)");

            if (string.IsNullOrEmpty(tenantId))
            {
                Fail(sessionId, "Não foi possível obter TenantId após autenticação.");
                return;
            }

            UpdateStatus(sessionId, s =>
            {
                s.Status = "Registering";
                s.Message = $"Conectado como {account}. Criando App Registration...";
                s.TenantId = tenantId;
            });
            _logger.LogInformation("[Setup:{SessionId}] Etapa 2/3 — Criando App Registration...", sessionId);

            // 2) Criar App Registration multi-tenant com as permissões necessárias
            using var ps2 = PowerShell.Create();
            ps2.Runspace = rs;
            ps2.AddScript($@"
                $required = @(
                    @{{
                        ResourceAppId = '{GraphResourceAppId}'
                        ResourceAccess = @(
                            @{{ Id = '{RoleManagementReadDirectory}'; Type = 'Role' }},
                            @{{ Id = '{RoleManagementReadWriteDirectory}'; Type = 'Role' }}
                        )
                    }}
                )
                $existing = Get-MgApplication -Filter ""displayName eq '{AppDisplayName}'"" -ErrorAction SilentlyContinue | Select-Object -First 1
                if ($existing) {{
                    $app = $existing
                }} else {{
                    $app = New-MgApplication -DisplayName '{AppDisplayName}' -SignInAudience 'AzureADMultipleOrgs' -RequiredResourceAccess $required
                }}

                # Garante service principal local no tenant bootstrap
                $sp = Get-MgServicePrincipal -Filter ""appId eq '$($app.AppId)'"" -ErrorAction SilentlyContinue | Select-Object -First 1
                if (-not $sp) {{
                    $sp = New-MgServicePrincipal -AppId $app.AppId
                }}

                # Cria client secret válido por 2 anos
                $passwordCred = @{{
                    DisplayName = 'Auto-generated by IR setup'
                    EndDateTime = (Get-Date).AddYears(2).ToString('o')
                }}
                $secret = Add-MgApplicationPassword -ApplicationId $app.Id -PasswordCredential $passwordCred

                Write-Output $app.AppId
                Write-Output $app.Id
                Write-Output $sp.Id
                Write-Output $secret.SecretText

                Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
            ");

            var appResults = await Task.Run(() => ps2.Invoke());
            if (ps2.HadErrors)
            {
                var err = string.Join("; ", ps2.Streams.Error.Select(e => e.ToString()));
                _logger.LogError("[Setup:{SessionId}] Falha criando App Registration: {Err}", sessionId, err);
                Fail(sessionId, $"Falha criando App Registration: {err}");
                return;
            }

            var list = appResults
                .Select(r => r?.BaseObject?.ToString() ?? "")
                .Where(v => !string.IsNullOrWhiteSpace(v))
                .ToList();

            _logger.LogInformation("[Setup:{SessionId}] Etapa 2/3 concluída. Outputs recebidos: {Count}", sessionId, list.Count);

            if (list.Count < 4)
            {
                var received = string.Join(", ", list.Select((v, i) => $"[{i}]={v}"));
                _logger.LogError("[Setup:{SessionId}] Resposta incompleta do Graph. Recebidos: {Received}", sessionId, received);
                Fail(sessionId, $"Resposta incompleta do Graph (esperado 4 valores, recebidos {list.Count}).");
                return;
            }

            _logger.LogInformation("[Setup:{SessionId}] Etapa 3/3 — Persistindo configuração...", sessionId);
            var config = new AppConfig
            {
                ClientId = list[0],
                ApplicationObjectId = list[1],
                ServicePrincipalId = list[2],
                ClientSecret = list[3],
                DisplayName = AppDisplayName,
                BootstrapTenantId = tenantId,
                CreatedAt = DateTimeOffset.UtcNow
            };
            _store.Write(config);
            _logger.LogInformation("[Setup:{SessionId}] Setup concluído. ClientId={ClientId}", sessionId, config.ClientId);

            UpdateStatus(sessionId, s =>
            {
                s.Status = "Completed";
                s.Message = "App registrado com sucesso. Você já pode adicionar tenants.";
                s.ClientId = config.ClientId;
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "[Setup:{SessionId}] Exceção não tratada no setup.", sessionId);
            Fail(sessionId, ex.Message);
        }
    }

    private void TryExtractUserCode(string sessionId, string text)
    {
        if (string.IsNullOrWhiteSpace(text)) return;

        // Já tem code? Não reprocessa.
        if (_states.TryGetValue(sessionId, out var cur) && !string.IsNullOrEmpty(cur.UserCode)) return;

        // Padrão 1: "...devicelogin...code XXXX-XXXX" (ordem URL antes do código)
        // Padrão 2: "...code XXXX-XXXX...devicelogin" (ordem invertida em algumas versões do módulo)
        // Padrão 3: apenas o código standalone "enter the code XXXX-XXXX"
        var patterns = new[]
        {
            @"https?://[^\s]*devicelogin[^\s]*.*?code\s+([A-Z0-9]{4,}-[A-Z0-9]{4,}(?:-[A-Z0-9]{4,})?)",
            @"code\s+([A-Z0-9]{4,}-[A-Z0-9]{4,}(?:-[A-Z0-9]{4,})?).*?devicelogin",
            @"enter\s+(?:the\s+)?code[:\s]+([A-Z0-9]{4,}-[A-Z0-9]{4,}(?:-[A-Z0-9]{4,})?)",
        };

        foreach (var pattern in patterns)
        {
            var match = System.Text.RegularExpressions.Regex.Match(text, pattern,
                System.Text.RegularExpressions.RegexOptions.Singleline |
                System.Text.RegularExpressions.RegexOptions.IgnoreCase);

            if (match.Success)
            {
                var code = match.Groups[1].Value.ToUpperInvariant();
                _logger.LogInformation("[Setup:{SessionId}] UserCode extraído: {Code}", sessionId, code);
                UpdateStatus(sessionId, s =>
                {
                    s.UserCode = code;
                    s.VerificationUrl = "https://microsoft.com/devicelogin";
                    s.Message = $"Acesse https://microsoft.com/devicelogin e digite o código: {code}";
                });
                return;
            }
        }
    }

    private void UpdateStatus(string sessionId, Action<SetupStatus> mutator)
    {
        _states.AddOrUpdate(sessionId,
            _ => { var s = new SetupStatus(); mutator(s); return s; },
            (_, s) => { mutator(s); return s; });
    }

    private void Fail(string sessionId, string reason)
    {
        UpdateStatus(sessionId, s =>
        {
            s.Status = "Failed";
            s.Message = reason;
        });
    }
}
