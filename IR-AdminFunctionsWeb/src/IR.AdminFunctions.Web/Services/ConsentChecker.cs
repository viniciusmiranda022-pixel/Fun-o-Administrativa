using System.Net.Http.Headers;
using System.Text.Json;
using IR.AdminFunctions.Web.Models;

namespace IR.AdminFunctions.Web.Services;

// Verifica via Microsoft Graph se o Service Principal da nossa aplicação
// já tem as permissões Basic (RoleManagement.Read.Directory) e Restore
// (RoleManagement.ReadWrite.Directory) concedidas no tenant alvo.
//
// Usa client credentials flow com o ClientId/Secret do AppConfig.
public class ConsentChecker
{
    private readonly AppConfigStore _appStore;
    private readonly IHttpClientFactory _httpClientFactory;
    private readonly ILogger<ConsentChecker> _logger;

    private const string PermissionBasicId = "483bed4a-2ad3-4361-a73b-c83ccdbdc53c";
    private const string PermissionRestoreId = "9e3f62cf-ca93-4989-b6ce-bf83c28f9fe8";

    private static readonly JsonSerializerOptions JsonOpts = new() { PropertyNameCaseInsensitive = true };

    public ConsentChecker(AppConfigStore appStore, IHttpClientFactory httpClientFactory, ILogger<ConsentChecker> logger)
    {
        _appStore = appStore;
        _httpClientFactory = httpClientFactory;
        _logger = logger;
    }

    public async Task<TenantConsentsState> CheckAsync(string tenantId, CancellationToken ct)
    {
        var cfg = _appStore.Read();
        if (!cfg.IsConfigured) return new TenantConsentsState();

        try
        {
            var token = await GetClientCredentialsTokenAsync(cfg.ClientId!, cfg.ClientSecret!, tenantId, ct);
            if (token == null) return new TenantConsentsState();

            using var http = _httpClientFactory.CreateClient();
            http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);

            // Busca o Service Principal da nossa app no tenant alvo
            var spUrl = $"https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '{cfg.ClientId}'&$select=id";
            var spResp = await http.GetAsync(spUrl, ct);
            if (!spResp.IsSuccessStatusCode)
            {
                _logger.LogWarning("Falha ao buscar SP no tenant {Tenant}: {Status}", tenantId, spResp.StatusCode);
                return new TenantConsentsState();
            }

            var spJson = await spResp.Content.ReadAsStringAsync(ct);
            var spList = JsonSerializer.Deserialize<GraphList<GraphIdItem>>(spJson, JsonOpts);
            var spId = spList?.Value?.FirstOrDefault()?.Id;
            if (spId == null) return new TenantConsentsState();

            // Busca as app role assignments do SP
            var assignUrl = $"https://graph.microsoft.com/v1.0/servicePrincipals/{spId}/appRoleAssignments";
            var assignResp = await http.GetAsync(assignUrl, ct);
            if (!assignResp.IsSuccessStatusCode) return new TenantConsentsState();

            var assignJson = await assignResp.Content.ReadAsStringAsync(ct);
            var assignments = JsonSerializer.Deserialize<GraphList<AppRoleAssignment>>(assignJson, JsonOpts);
            var roleIds = (assignments?.Value ?? [])
                .Select(a => a.AppRoleId ?? "")
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            return new TenantConsentsState
            {
                Basic = new ConsentInfo { Granted = roleIds.Contains(PermissionBasicId) },
                Restore = new ConsentInfo { Granted = roleIds.Contains(PermissionRestoreId) }
            };
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Falha ao consultar consents do tenant {Tenant}", tenantId);
            return new TenantConsentsState();
        }
    }

    // Constrói a URL de admin consent que o navegador deve abrir.
    public string BuildAdminConsentUrl(string redirectUri, string state, string? tenantHint = null)
    {
        var cfg = _appStore.Read();
        if (!cfg.IsConfigured) throw new InvalidOperationException("App não configurado. Execute o setup inicial.");

        var tenant = string.IsNullOrWhiteSpace(tenantHint) ? "common" : tenantHint;
        return $"https://login.microsoftonline.com/{tenant}/adminconsent" +
               $"?client_id={cfg.ClientId}" +
               $"&redirect_uri={Uri.EscapeDataString(redirectUri)}" +
               $"&state={Uri.EscapeDataString(state)}";
    }

    private async Task<string?> GetClientCredentialsTokenAsync(string clientId, string clientSecret, string tenantId, CancellationToken ct)
    {
        using var http = _httpClientFactory.CreateClient();
        var body = new FormUrlEncodedContent(new Dictionary<string, string>
        {
            ["grant_type"] = "client_credentials",
            ["client_id"] = clientId,
            ["client_secret"] = clientSecret,
            ["scope"] = "https://graph.microsoft.com/.default",
        });
        var resp = await http.PostAsync($"https://login.microsoftonline.com/{tenantId}/oauth2/v2.0/token", body, ct);
        if (!resp.IsSuccessStatusCode)
        {
            var err = await resp.Content.ReadAsStringAsync(ct);
            _logger.LogWarning("Falha ao obter token para tenant {Tenant}: {Err}", tenantId, err);
            return null;
        }
        var json = await resp.Content.ReadAsStringAsync(ct);
        var doc = JsonDocument.Parse(json);
        return doc.RootElement.TryGetProperty("access_token", out var t) ? t.GetString() : null;
    }

    private record GraphList<T>(List<T>? Value);
    private record GraphIdItem(string? Id);
    private record AppRoleAssignment(string? AppRoleId);
}
