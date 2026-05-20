using System.Collections.Concurrent;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using IR.AdminFunctions.Web.Models;
using IR.AdminFunctions.Web.Services;
using Microsoft.AspNetCore.Mvc;

namespace IR.AdminFunctions.Web.Controllers;

[ApiController]
[Route("api/oauth")]
public class OAuthController : ControllerBase
{
    private static readonly ConcurrentDictionary<string, string> _states = new();
    private readonly ConsentChecker _consent;
    private readonly AppConfigStore _appStore;
    private readonly TenantStore _tenantStore;
    private readonly ILogger<OAuthController> _logger;

    public OAuthController(ConsentChecker consent, AppConfigStore appStore, TenantStore tenantStore, ILogger<OAuthController> logger)
    {
        _consent = consent;
        _appStore = appStore;
        _tenantStore = tenantStore;
        _logger = logger;
    }

    // O frontend chama este endpoint para obter a URL que abre o admin_consent
    // numa nova aba do navegador.
    [HttpGet("start")]
    public ActionResult<ApiResponse<object>> Start([FromQuery] string? tenantHint = null)
    {
        var cfg = _appStore.Read();
        if (!cfg.IsConfigured)
            return BadRequest(ApiResponse<object>.Fail("App ainda não configurado. Acesse /setup primeiro."));

        var state = Guid.NewGuid().ToString("N");
        _states[state] = tenantHint ?? "";

        var redirectUri = $"{Request.Scheme}://{Request.Host}/api/oauth/callback";
        var url = _consent.BuildAdminConsentUrl(redirectUri, state, tenantHint);

        return ApiResponse<object>.Ok(new { url, state });
    }

    // Callback chamado pelo Microsoft após o admin consent.
    // Query params: tenant=<tenantId>&admin_consent=True&state=<...>
    [HttpGet("callback")]
    public async Task<IActionResult> Callback([FromQuery] string? tenant, [FromQuery(Name = "admin_consent")] string? adminConsent, [FromQuery] string? state, [FromQuery] string? error, [FromQuery(Name = "error_description")] string? errorDescription, CancellationToken ct)
    {
        if (!string.IsNullOrEmpty(error))
        {
            return ContentResult($"Erro no consentimento: {error} — {errorDescription}", isError: true);
        }

        if (string.IsNullOrEmpty(state) || !_states.TryRemove(state, out _))
        {
            return ContentResult("State inválido. Tente novamente a partir da página de Tenants.", isError: true);
        }

        if (string.IsNullOrEmpty(tenant) || !string.Equals(adminConsent, "True", StringComparison.OrdinalIgnoreCase))
        {
            return ContentResult("Consentimento não foi concluído.", isError: true);
        }

        // Busca info do tenant via Graph (display name + domínio)
        var (displayName, domain) = await GetTenantInfoAsync(tenant, ct);

        var existing = await _tenantStore.GetAsync(tenant);
        if (existing == null)
        {
            await _tenantStore.AddAsync(new AddTenantRequest
            {
                Name = displayName ?? tenant,
                TenantId = tenant,
                ClientId = _appStore.Read().ClientId,
                CertificateThumbprint = null,
                Domain = domain
            });
        }

        // Atualiza consents (consulta real)
        var consents = await _consent.CheckAsync(tenant, ct);
        await _tenantStore.UpdateConsentsAsync(tenant, consents);

        return ContentResult($"Tenant <strong>{displayName ?? tenant}</strong> conectado com sucesso. Você já pode fechar esta janela.");
    }

    private async Task<(string? displayName, string? domain)> GetTenantInfoAsync(string tenantId, CancellationToken ct)
    {
        var cfg = _appStore.Read();
        try
        {
            var iss = InitialSessionState.CreateDefault();
            iss.ExecutionPolicy = Microsoft.PowerShell.ExecutionPolicy.Bypass;
            using var rs = RunspaceFactory.CreateRunspace(iss);
            rs.Open();
            using var ps = PowerShell.Create();
            ps.Runspace = rs;
            ps.AddScript($@"
                Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
                Import-Module Microsoft.Graph.Identity.DirectoryManagement -ErrorAction Stop
                $secret = ConvertTo-SecureString '{cfg.ClientSecret}' -AsPlainText -Force
                $cred = New-Object System.Management.Automation.PSCredential('{cfg.ClientId}', $secret)
                Connect-MgGraph -TenantId '{tenantId}' -ClientSecretCredential $cred -NoWelcome -ErrorAction Stop
                $org = Get-MgOrganization -ErrorAction SilentlyContinue | Select-Object -First 1
                if ($org) {{
                    Write-Output $org.DisplayName
                    $defaultDomain = $org.VerifiedDomains | Where-Object IsDefault | Select-Object -First 1 -ExpandProperty Name
                    Write-Output $defaultDomain
                }}
                Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
            ");
            var results = await Task.Run(() => ps.Invoke(), ct);
            var name = results.FirstOrDefault()?.BaseObject?.ToString();
            var domain = results.Skip(1).FirstOrDefault()?.BaseObject?.ToString();
            return (name, domain);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Não foi possível obter info do tenant {Tenant}", tenantId);
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
<h1 style='font-size:18px;color:#222;margin:0 0 8px'>{(isError ? "Falha" : "Sucesso")}</h1>
<p style='color:#555;margin:0 0 16px'>{html}</p>
<p style='color:#999;font-size:12px;margin:0'>Você pode fechar esta janela e voltar ao app.</p>
</div>
<script>setTimeout(()=>{{ if(window.opener){{ window.opener.postMessage({{type:'oauth-complete',ok:{(!isError).ToString().ToLower()}}},'*'); window.close(); }} }}, 1500);</script>
</body></html>"
        };
    }
}
