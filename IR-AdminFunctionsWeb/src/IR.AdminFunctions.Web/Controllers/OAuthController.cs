using System.Collections.Concurrent;
using System.Net.Http.Headers;
using System.Text.Encodings.Web;
using System.Text.Json;
using IR.AdminFunctions.Web.Models;
using IR.AdminFunctions.Web.Services;
using Microsoft.AspNetCore.Mvc;

namespace IR.AdminFunctions.Web.Controllers;

[ApiController]
[Route("api/oauth")]
public class OAuthController : ControllerBase
{
    private static readonly ConcurrentDictionary<string, (string Hint, DateTimeOffset CreatedAt)> _states = new();
    private readonly ConsentChecker _consent;
    private readonly AppConfigStore _appStore;
    private readonly TenantStore _tenantStore;
    private readonly IHttpClientFactory _httpClientFactory;
    private readonly ILogger<OAuthController> _logger;

    private static readonly JsonSerializerOptions JsonOpts = new() { PropertyNameCaseInsensitive = true };

    public OAuthController(ConsentChecker consent, AppConfigStore appStore, TenantStore tenantStore, IHttpClientFactory httpClientFactory, ILogger<OAuthController> logger)
    {
        _consent = consent;
        _appStore = appStore;
        _tenantStore = tenantStore;
        _httpClientFactory = httpClientFactory;
        _logger = logger;
    }

    [HttpGet("start")]
    public ActionResult<ApiResponse<object>> Start([FromQuery] string? tenantHint = null)
    {
        var cfg = _appStore.Read();
        if (!cfg.IsConfigured)
            return BadRequest(ApiResponse<object>.Fail("App not configured yet. Go to /setup first."));

        // Clean up expired states (older than 10 minutes) before creating a new one
        var expiredKeys = _states
            .Where(kv => DateTimeOffset.UtcNow - kv.Value.CreatedAt > TimeSpan.FromMinutes(10))
            .Select(kv => kv.Key)
            .ToList();
        foreach (var key in expiredKeys)
            _states.TryRemove(key, out _);

        var state = Guid.NewGuid().ToString("N");
        _states[state] = (tenantHint ?? "", DateTimeOffset.UtcNow);

        var redirectUri = $"{Request.Scheme}://{Request.Host}/api/oauth/callback";
        var url = _consent.BuildAdminConsentUrl(redirectUri, state, tenantHint);

        return ApiResponse<object>.Ok(new { url, state });
    }

    [HttpGet("callback")]
    public async Task<IActionResult> Callback([FromQuery] string? tenant, [FromQuery(Name = "admin_consent")] string? adminConsent, [FromQuery] string? state, [FromQuery] string? error, [FromQuery(Name = "error_description")] string? errorDescription, CancellationToken ct)
    {
        if (!string.IsNullOrEmpty(error))
        {
            var safeError = HtmlEncoder.Default.Encode(error ?? "");
            var safeDesc = HtmlEncoder.Default.Encode(errorDescription ?? "");
            return ContentResult($"Consent error: {safeError} — {safeDesc}", isError: true);
        }

        if (string.IsNullOrEmpty(state) || !_states.TryRemove(state, out var stateEntry))
        {
            return ContentResult("Invalid state. Please try again from the Tenants page.", isError: true);
        }

        if (DateTimeOffset.UtcNow - stateEntry.CreatedAt > TimeSpan.FromMinutes(10))
        {
            return ContentResult("Consent session expired. Please start the process again.", isError: true);
        }

        if (string.IsNullOrEmpty(tenant) || !string.Equals(adminConsent, "True", StringComparison.OrdinalIgnoreCase))
        {
            return ContentResult("Consent was not completed.", isError: true);
        }

        var (displayName, domain) = await GetTenantInfoAsync(tenant, ct);

        var existing = await _tenantStore.GetAsync(tenant);
        if (existing == null)
        {
            await _tenantStore.AddAsync(new AddTenantRequest
            {
                Name = displayName ?? tenant,
                TenantId = tenant,
                ClientId = _appStore.Read().ClientId ?? string.Empty,
                CertificateThumbprint = string.Empty,
                Domain = domain
            });
        }

        var consents = await _consent.CheckAsync(tenant, ct);
        await _tenantStore.UpdateConsentsAsync(tenant, consents);

        return ContentResult($"Tenant <strong>{displayName ?? tenant}</strong> connected successfully. You may close this window.");
    }

    private async Task<(string? displayName, string? domain)> GetTenantInfoAsync(string tenantId, CancellationToken ct)
    {
        var cfg = _appStore.Read();
        if (!cfg.IsConfigured) return (null, null);

        try
        {
            // Obtains a token via client credentials for the target tenant
            using var tokenHttp = _httpClientFactory.CreateClient();
            var body = new FormUrlEncodedContent(new Dictionary<string, string>
            {
                ["grant_type"] = "client_credentials",
                ["client_id"] = cfg.ClientId!,
                ["client_secret"] = cfg.ClientSecret!,
                ["scope"] = "https://graph.microsoft.com/.default",
            });
            var tokenResp = await tokenHttp.PostAsync($"https://login.microsoftonline.com/{tenantId}/oauth2/v2.0/token", body, ct);
            if (!tokenResp.IsSuccessStatusCode) return (null, null);

            var tokenJson = await tokenResp.Content.ReadAsStringAsync(ct);
            var tokenDoc = JsonDocument.Parse(tokenJson);
            var token = tokenDoc.RootElement.TryGetProperty("access_token", out var t) ? t.GetString() : null;
            if (token == null) return (null, null);

            using var graphHttp = _httpClientFactory.CreateClient();
            graphHttp.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);

            var orgResp = await graphHttp.GetAsync("https://graph.microsoft.com/v1.0/organization?$select=displayName,verifiedDomains", ct);
            if (!orgResp.IsSuccessStatusCode) return (null, null);

            var orgJson = await orgResp.Content.ReadAsStringAsync(ct);
            var orgDoc = JsonDocument.Parse(orgJson);
            if (!orgDoc.RootElement.TryGetProperty("value", out var orgs) || orgs.GetArrayLength() == 0)
                return (null, null);

            var org = orgs[0];
            var name = org.TryGetProperty("displayName", out var dn) ? dn.GetString() : null;
            string? defaultDomain = null;
            if (org.TryGetProperty("verifiedDomains", out var domains))
            {
                foreach (var d in domains.EnumerateArray())
                {
                    if (d.TryGetProperty("isDefault", out var isDefault) && isDefault.GetBoolean())
                    {
                        defaultDomain = d.TryGetProperty("name", out var domName) ? domName.GetString() : null;
                        break;
                    }
                }
            }
            return (name, defaultDomain);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Could not retrieve tenant info for {Tenant}", tenantId);
            return (null, null);
        }
    }

    private ContentResult ContentResult(string html, bool isError = false)
    {
        var color = isError ? "#DC3545" : "#28A745";
        return new ContentResult
        {
            ContentType = "text/html",
            Content = $@"<!doctype html><html><body style='font-family:-apple-system,Segoe UI,sans-serif;padding:48px;background:#f5f5f5'>
<div style='max-width:520px;margin:0 auto;background:white;border:1px solid #DEE2E6;border-radius:8px;padding:32px;text-align:center'>
<div style='font-size:48px;color:{color};margin-bottom:16px'>{(isError ? "✗" : "✓")}</div>
<h1 style='font-size:18px;color:#222;margin:0 0 8px'>{(isError ? "Failed" : "Success")}</h1>
<p style='color:#555;margin:0 0 16px'>{html}</p>
<p style='color:#999;font-size:12px;margin:0'>You can close this window and return to the app.</p>
</div>
<script>setTimeout(()=>{{ if(window.opener){{ window.opener.postMessage({{type:'oauth-complete',ok:{(!isError).ToString().ToLower()}}},'*'); window.close(); }} }}, 1500);</script>
</body></html>"
        };
    }
}
