using System.Collections.Concurrent;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using Microsoft.Identity.Client;
using IR.AdminFunctions.Web.Models;

namespace IR.AdminFunctions.Web.Services;

// Registra a aplicação multi-tenant no Azure AD usando MSAL device code flow.
// O callback do device code fornece UserCode e VerificationUri diretamente,
// sem depender de parsing de saída do PowerShell.
public class AppRegistrationService
{
    private readonly AppConfigStore _store;
    private readonly ILogger<AppRegistrationService> _logger;
    private readonly IHttpClientFactory _httpClientFactory;
    private readonly ConcurrentDictionary<string, SetupStatus> _states = new();

    private const string AppDisplayName = "IR Administrative Function Recovery";

    // Client ID público do Microsoft Graph PowerShell (bem conhecido, multi-tenant)
    private const string PublicClientId = "14d82eec-204b-4c2f-b7e8-296a70dab67e";
    private const string Authority = "https://login.microsoftonline.com/organizations";

    // IDs de permissão do Microsoft Graph (globais)
    private const string GraphResourceAppId = "00000003-0000-0000-c000-000000000000";
    private const string RoleManagementReadDirectory = "483bed4a-2ad3-4361-a73b-c83ccdbdc53c";
    private const string RoleManagementReadWriteDirectory = "9e3f62cf-ca93-4989-b6ce-bf83c28f9fe8";

    private static readonly string[] GraphScopes =
    [
        "https://graph.microsoft.com/Application.ReadWrite.All",
        "https://graph.microsoft.com/Directory.ReadWrite.All",
        "https://graph.microsoft.com/User.Read",
    ];

    private static readonly JsonSerializerOptions JsonOpts = new()
    {
        PropertyNameCaseInsensitive = true,
        WriteIndented = false,
    };

    public AppRegistrationService(
        AppConfigStore store,
        ILogger<AppRegistrationService> logger,
        IHttpClientFactory httpClientFactory)
    {
        _store = store;
        _logger = logger;
        _httpClientFactory = httpClientFactory;
    }

    public SetupStatus GetStatus(string sessionId) =>
        _states.TryGetValue(sessionId, out var s) ? s : new SetupStatus { Status = "NotStarted" };

    public string StartSetup()
    {
        var sessionId = Guid.NewGuid().ToString("N");
        _states[sessionId] = new SetupStatus { Status = "WaitingForUser", Message = "Aguardando usuário autenticar..." };
        _logger.LogInformation("[Setup:{SessionId}] Setup iniciado.", sessionId);
        _ = Task.Run(() => RunSetupAsync(sessionId));
        return sessionId;
    }

    private async Task RunSetupAsync(string sessionId)
    {
        try
        {
            // Etapa 1 — autenticação via device code (MSAL)
            _logger.LogInformation("[Setup:{SessionId}] Etapa 1/3 — Iniciando autenticação via device code (MSAL)...", sessionId);

            var msalApp = PublicClientApplicationBuilder
                .Create(PublicClientId)
                .WithAuthority(Authority)
                .Build();

            AuthenticationResult tokenResult;
            try
            {
                tokenResult = await msalApp
                    .AcquireTokenWithDeviceCode(GraphScopes, deviceCode =>
                    {
                        _logger.LogInformation(
                            "[Setup:{SessionId}] Device code recebido. UserCode={Code}, Url={Url}",
                            sessionId, deviceCode.UserCode, deviceCode.VerificationUrl);

                        UpdateStatus(sessionId, s =>
                        {
                            s.UserCode = deviceCode.UserCode;
                            s.VerificationUrl = deviceCode.VerificationUrl;
                            s.Message = $"Acesse {deviceCode.VerificationUrl} e digite o código: {deviceCode.UserCode}";
                        });

                        return Task.CompletedTask;
                    })
                    .ExecuteAsync();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "[Setup:{SessionId}] Falha na autenticação MSAL.", sessionId);
                Fail(sessionId, $"Falha na autenticação: {ex.Message}");
                return;
            }

            var tenantId = tokenResult.TenantId;
            var account = tokenResult.Account?.Username;

            _logger.LogInformation(
                "[Setup:{SessionId}] Etapa 1/3 concluída. TenantId={TenantId}, Account={Account}",
                sessionId, tenantId, account);

            if (string.IsNullOrEmpty(tenantId))
            {
                Fail(sessionId, "Não foi possível obter o TenantId do token.");
                return;
            }

            UpdateStatus(sessionId, s =>
            {
                s.Status = "Registering";
                s.Message = $"Conectado como {account}. Criando App Registration...";
                s.TenantId = tenantId;
            });

            // Etapa 2 — criar App Registration via Microsoft Graph REST
            _logger.LogInformation("[Setup:{SessionId}] Etapa 2/3 — Criando App Registration via Graph API...", sessionId);

            using var http = _httpClientFactory.CreateClient();
            http.DefaultRequestHeaders.Authorization =
                new AuthenticationHeaderValue("Bearer", tokenResult.AccessToken);
            http.DefaultRequestHeaders.Accept.Add(
                new MediaTypeWithQualityHeaderValue("application/json"));

            // Verifica se o app já existe
            string appObjectId;
            string appClientId;

            var filterResp = await http.GetAsync(
                $"https://graph.microsoft.com/v1.0/applications?$filter=displayName eq '{Uri.EscapeDataString(AppDisplayName)}'&$select=id,appId");

            filterResp.EnsureSuccessStatusCode();
            var filterJson = await filterResp.Content.ReadAsStringAsync();
            var filterDoc = JsonSerializer.Deserialize<GraphList>(filterJson, JsonOpts);

            if (filterDoc?.Value?.Length > 0)
            {
                appObjectId = filterDoc.Value[0].Id!;
                appClientId = filterDoc.Value[0].AppId!;
                _logger.LogInformation(
                    "[Setup:{SessionId}] App Registration existente encontrado. AppId={AppId}", sessionId, appClientId);
            }
            else
            {
                // Cria o App Registration
                var appPayload = new
                {
                    displayName = AppDisplayName,
                    signInAudience = "AzureADMultipleOrgs",
                    requiredResourceAccess = new[]
                    {
                        new
                        {
                            resourceAppId = GraphResourceAppId,
                            resourceAccess = new[]
                            {
                                new { id = RoleManagementReadDirectory, type = "Role" },
                                new { id = RoleManagementReadWriteDirectory, type = "Role" },
                            },
                        },
                    },
                };

                var createContent = new StringContent(
                    JsonSerializer.Serialize(appPayload, JsonOpts),
                    Encoding.UTF8, "application/json");

                var createResp = await http.PostAsync(
                    "https://graph.microsoft.com/v1.0/applications", createContent);

                if (!createResp.IsSuccessStatusCode)
                {
                    var errBody = await createResp.Content.ReadAsStringAsync();
                    _logger.LogError("[Setup:{SessionId}] Falha ao criar App Registration: {Status} {Body}",
                        sessionId, (int)createResp.StatusCode, errBody);
                    Fail(sessionId, $"Falha ao criar App Registration ({(int)createResp.StatusCode}): {errBody}");
                    return;
                }

                var createJson = await createResp.Content.ReadAsStringAsync();
                var created = JsonSerializer.Deserialize<GraphApp>(createJson, JsonOpts);
                appObjectId = created?.Id ?? throw new InvalidOperationException("App criado sem Id.");
                appClientId = created?.AppId ?? throw new InvalidOperationException("App criado sem AppId.");
                _logger.LogInformation(
                    "[Setup:{SessionId}] App Registration criado. AppId={AppId}", sessionId, appClientId);
            }

            // Garante Service Principal no tenant bootstrap
            string spId;
            var spFilterResp = await http.GetAsync(
                $"https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '{appClientId}'&$select=id");

            spFilterResp.EnsureSuccessStatusCode();
            var spFilterJson = await spFilterResp.Content.ReadAsStringAsync();
            var spFilterDoc = JsonSerializer.Deserialize<GraphList>(spFilterJson, JsonOpts);

            if (spFilterDoc?.Value?.Length > 0)
            {
                spId = spFilterDoc.Value[0].Id!;
                _logger.LogInformation("[Setup:{SessionId}] Service Principal existente. SpId={SpId}", sessionId, spId);
            }
            else
            {
                var spPayload = new { appId = appClientId };
                var spContent = new StringContent(
                    JsonSerializer.Serialize(spPayload, JsonOpts),
                    Encoding.UTF8, "application/json");

                var spResp = await http.PostAsync(
                    "https://graph.microsoft.com/v1.0/servicePrincipals", spContent);

                if (!spResp.IsSuccessStatusCode)
                {
                    var errBody = await spResp.Content.ReadAsStringAsync();
                    _logger.LogError("[Setup:{SessionId}] Falha ao criar Service Principal: {Status} {Body}",
                        sessionId, (int)spResp.StatusCode, errBody);
                    Fail(sessionId, $"Falha ao criar Service Principal ({(int)spResp.StatusCode}): {errBody}");
                    return;
                }

                var spJson = await spResp.Content.ReadAsStringAsync();
                var sp = JsonSerializer.Deserialize<GraphApp>(spJson, JsonOpts);
                spId = sp?.Id ?? throw new InvalidOperationException("SP criado sem Id.");
                _logger.LogInformation("[Setup:{SessionId}] Service Principal criado. SpId={SpId}", sessionId, spId);
            }

            // Cria client secret válido por 2 anos
            var secretPayload = new
            {
                passwordCredential = new
                {
                    displayName = "Auto-generated by IR setup",
                    endDateTime = DateTimeOffset.UtcNow.AddYears(2).ToString("o"),
                },
            };

            var secretContent = new StringContent(
                JsonSerializer.Serialize(secretPayload, JsonOpts),
                Encoding.UTF8, "application/json");

            var secretResp = await http.PostAsync(
                $"https://graph.microsoft.com/v1.0/applications/{appObjectId}/addPassword",
                secretContent);

            if (!secretResp.IsSuccessStatusCode)
            {
                var errBody = await secretResp.Content.ReadAsStringAsync();
                _logger.LogError("[Setup:{SessionId}] Falha ao criar client secret: {Status} {Body}",
                    sessionId, (int)secretResp.StatusCode, errBody);
                Fail(sessionId, $"Falha ao criar client secret ({(int)secretResp.StatusCode}): {errBody}");
                return;
            }

            var secretJson = await secretResp.Content.ReadAsStringAsync();
            var secretDoc = JsonSerializer.Deserialize<GraphSecret>(secretJson, JsonOpts);
            var clientSecret = secretDoc?.SecretText
                ?? throw new InvalidOperationException("Client secret criado sem SecretText.");

            _logger.LogInformation("[Setup:{SessionId}] Etapa 2/3 concluída. AppId={AppId}", sessionId, appClientId);

            // Etapa 3 — persistir configuração
            _logger.LogInformation("[Setup:{SessionId}] Etapa 3/3 — Persistindo configuração...", sessionId);

            var config = new AppConfig
            {
                ClientId = appClientId,
                ApplicationObjectId = appObjectId,
                ServicePrincipalId = spId,
                ClientSecret = clientSecret,
                DisplayName = AppDisplayName,
                BootstrapTenantId = tenantId,
                CreatedAt = DateTimeOffset.UtcNow,
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

    // DTOs para deserializar respostas do Graph
    private sealed class GraphList
    {
        public GraphApp[]? Value { get; set; }
    }

    private sealed class GraphApp
    {
        public string? Id { get; set; }
        public string? AppId { get; set; }
    }

    private sealed class GraphSecret
    {
        public string? SecretText { get; set; }
    }
}
