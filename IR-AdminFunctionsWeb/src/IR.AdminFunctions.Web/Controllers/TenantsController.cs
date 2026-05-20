using IR.AdminFunctions.Web.Models;
using IR.AdminFunctions.Web.Services;
using Microsoft.AspNetCore.Mvc;

namespace IR.AdminFunctions.Web.Controllers;

[ApiController]
[Route("api/[controller]")]
public class TenantsController : ControllerBase
{
    private readonly SettingsReader _settings;

    public TenantsController(SettingsReader settings)
    {
        _settings = settings;
    }

    [HttpGet]
    public ActionResult<ApiResponse<IEnumerable<TenantInfo>>> Get()
    {
        var s = _settings.Read();
        if (s == null || string.IsNullOrWhiteSpace(s.TenantId))
        {
            return ApiResponse<IEnumerable<TenantInfo>>.Ok(Array.Empty<TenantInfo>());
        }

        var configured = !string.IsNullOrWhiteSpace(s.ClientId)
                         && !string.IsNullOrWhiteSpace(s.CertificateThumbprint);

        var tenant = new TenantInfo
        {
            TenantId = s.TenantId!,
            ClientId = s.ClientId,
            AppDisplayName = s.AppDisplayName,
            DisplayName = s.AppDisplayName ?? s.TenantId!,
            ConfiguredCorrectly = configured,
            ConsentStatus = configured ? "Granted" : "Pending"
        };

        return ApiResponse<IEnumerable<TenantInfo>>.Ok(new[] { tenant });
    }
}
