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
            return NotFound(ApiResponse<object>.Fail("Tenant not found."));

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
                description = "Allows reading role definitions and role assignments from Microsoft Entra ID (required for backup and comparison).",
                operations = new[] { "Backup", "Comparison", "Inventory" }
            },
            restore = new
            {
                granted = consents.Restore.Granted,
                grantedAt = consents.Restore.GrantedAt,
                permission = "RoleManagement.ReadWrite.Directory",
                description = "Allows creating and updating custom role definitions and role assignments. WARNING: permits changes to directory RBAC.",
                operations = new[] { "Restore" },
                sensitive = true
            }
        });
    }
}
