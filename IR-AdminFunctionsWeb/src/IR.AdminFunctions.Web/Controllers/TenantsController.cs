using IR.AdminFunctions.Web.Models;
using IR.AdminFunctions.Web.Services;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Options;

namespace IR.AdminFunctions.Web.Controllers;

[ApiController]
[Route("api/[controller]")]
public class TenantsController : ControllerBase
{
    private readonly TenantStore _store;
    private readonly PowerShellRunner _ps;

    public TenantsController(TenantStore store, PowerShellRunner ps)
    {
        _store = store;
        _ps = ps;
    }

    [HttpGet]
    public async Task<ActionResult<ApiResponse<IEnumerable<TenantEntry>>>> Get()
    {
        var tenants = await _store.ListAsync();
        return ApiResponse<IEnumerable<TenantEntry>>.Ok(tenants);
    }

    [HttpPost]
    public async Task<ActionResult<ApiResponse<TenantEntry>>> Add([FromBody] AddTenantRequest req)
    {
        if (string.IsNullOrWhiteSpace(req.TenantId))
            return BadRequest(ApiResponse<TenantEntry>.Fail("TenantId é obrigatório."));
        if (string.IsNullOrWhiteSpace(req.ClientId))
            return BadRequest(ApiResponse<TenantEntry>.Fail("ClientId é obrigatório."));
        if (string.IsNullOrWhiteSpace(req.CertificateThumbprint))
            return BadRequest(ApiResponse<TenantEntry>.Fail("CertificateThumbprint é obrigatório."));

        try
        {
            var entry = await _store.AddAsync(req);
            return ApiResponse<TenantEntry>.Ok(entry);
        }
        catch (InvalidOperationException ex)
        {
            return Conflict(ApiResponse<TenantEntry>.Fail(ex.Message));
        }
    }

    [HttpPut("{tenantId}/backup-config")]
    public async Task<ActionResult<ApiResponse<TenantEntry>>> UpdateBackupConfig(string tenantId, [FromBody] BackupConfig cfg)
    {
        var tenant = await _store.GetAsync(tenantId);
        if (tenant == null)
            return NotFound(ApiResponse<TenantEntry>.Fail("Tenant não encontrado."));

        cfg.LastModifiedAt = DateTimeOffset.UtcNow;
        var updated = await _store.UpdateBackupConfigAsync(tenantId, cfg);
        if (updated == null)
            return NotFound(ApiResponse<TenantEntry>.Fail("Tenant não encontrado."));
        return ApiResponse<TenantEntry>.Ok(updated);
    }

    [HttpDelete("{tenantId}")]
    public async Task<ActionResult<ApiResponse<bool>>> Remove(string tenantId)
    {
        var removed = await _store.RemoveAsync(tenantId);
        if (!removed)
            return NotFound(ApiResponse<bool>.Fail($"Tenant {tenantId} não encontrado."));
        return ApiResponse<bool>.Ok(true);
    }

    [HttpGet("{tenantId}/status")]
    public async Task<ActionResult<ApiResponse<TenantStatusResult>>> TestConnection(string tenantId, CancellationToken ct)
    {
        var tenant = await _store.GetAsync(tenantId);
        if (tenant == null)
            return NotFound(ApiResponse<TenantStatusResult>.Fail("Tenant não encontrado."));

        if (!tenant.IsConfigured)
        {
            return ApiResponse<TenantStatusResult>.Ok(new TenantStatusResult
            {
                TenantId = tenantId,
                Status = "Unconfigured",
                Message = "Preencha TenantId, ClientId e Thumbprint válidos."
            });
        }

        try
        {
            var script = $@"
                [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
                Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
                Connect-MgGraph -TenantId '{tenant.TenantId}' -ClientId '{tenant.ClientId}' -CertificateThumbprint '{tenant.CertificateThumbprint}' -NoWelcome -ErrorAction Stop
                $org = Get-MgOrganization -ErrorAction SilentlyContinue | Select-Object -First 1
                Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
                if ($org) {{ $org.DisplayName + '|' + ($org.VerifiedDomains | Where-Object IsDefault | Select-Object -First 1 -ExpandProperty Name) }}
                else {{ 'connected' }}
            ";

            var result = await _ps.RunAsync(script, null, 60, ct);

            if (!result.Success)
            {
                return ApiResponse<TenantStatusResult>.Ok(new TenantStatusResult
                {
                    TenantId = tenantId,
                    Status = "Error",
                    Message = result.Errors.FirstOrDefault() ?? "Falha na conexão."
                });
            }

            var output = result.Output.FirstOrDefault(o => !string.IsNullOrWhiteSpace(o)) ?? "connected";
            return ApiResponse<TenantStatusResult>.Ok(new TenantStatusResult
            {
                TenantId = tenantId,
                Status = "Connected",
                Message = output
            });
        }
        catch (Exception ex)
        {
            return ApiResponse<TenantStatusResult>.Ok(new TenantStatusResult
            {
                TenantId = tenantId,
                Status = "Error",
                Message = ex.Message
            });
        }
    }
}
