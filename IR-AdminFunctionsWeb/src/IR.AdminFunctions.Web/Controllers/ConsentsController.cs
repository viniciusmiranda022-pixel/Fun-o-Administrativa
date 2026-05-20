using IR.AdminFunctions.Web.Models;
using IR.AdminFunctions.Web.Services;
using Microsoft.AspNetCore.Mvc;

namespace IR.AdminFunctions.Web.Controllers;

[ApiController]
[Route("api/tenants/{tenantId}/consents")]
public class ConsentsController : ControllerBase
{
    private readonly TenantStore _tenantStore;
    private readonly ConsentChecker _checker;

    public ConsentsController(TenantStore tenantStore, ConsentChecker checker)
    {
        _tenantStore = tenantStore;
        _checker = checker;
    }

    [HttpGet]
    public async Task<ActionResult<ApiResponse<object>>> Get(string tenantId, CancellationToken ct)
    {
        var tenant = await _tenantStore.GetAsync(tenantId);
        if (tenant == null)
            return NotFound(ApiResponse<object>.Fail("Tenant não encontrado."));

        var consents = await _checker.CheckAsync(tenantId, ct);
        await _tenantStore.UpdateConsentsAsync(tenantId, consents);

        return ApiResponse<object>.Ok(new
        {
            tenantId,
            basic = new
            {
                granted = consents.Basic.Granted,
                grantedAt = consents.Basic.GrantedAt,
                permission = "RoleManagement.Read.Directory",
                description = "Permite ler role definitions e role assignments do Microsoft Entra ID (necessário para backup e comparação).",
                operations = new[] { "Backup", "Comparação", "Inventário" }
            },
            restore = new
            {
                granted = consents.Restore.Granted,
                grantedAt = consents.Restore.GrantedAt,
                permission = "RoleManagement.ReadWrite.Directory",
                description = "Permite criar e atualizar custom role definitions e role assignments. ATENÇÃO: permite alteração no RBAC do diretório.",
                operations = new[] { "Restore" },
                sensitive = true
            }
        });
    }
}
