using IR.AdminFunctions.Web.Models;
using IR.AdminFunctions.Web.Services;
using Microsoft.AspNetCore.Mvc;

namespace IR.AdminFunctions.Web.Controllers;

[ApiController]
[Route("api/[controller]")]
public class SetupController : ControllerBase
{
    private readonly AppConfigStore _store;
    private readonly AppRegistrationService _registration;

    public SetupController(AppConfigStore store, AppRegistrationService registration)
    {
        _store = store;
        _registration = registration;
    }

    [HttpGet("state")]
    public ActionResult<ApiResponse<object>> State()
    {
        var cfg = _store.Read();
        return ApiResponse<object>.Ok(new
        {
            isConfigured = cfg.IsConfigured,
            clientId = cfg.ClientId,
            displayName = cfg.DisplayName,
            bootstrapTenantId = cfg.BootstrapTenantId,
            createdAt = cfg.CreatedAt
        });
    }

    [HttpPost("start")]
    public ActionResult<ApiResponse<object>> Start()
    {
        var sessionId = _registration.StartSetup();
        var status = _registration.GetStatus(sessionId);
        return ApiResponse<object>.Ok(new
        {
            sessionId,
            status = status.Status,
            userCode = status.UserCode,
            verificationUrl = status.VerificationUrl,
            message = status.Message
        });
    }

    [HttpGet("status/{sessionId}")]
    public ActionResult<ApiResponse<SetupStatus>> Status(string sessionId)
    {
        var status = _registration.GetStatus(sessionId);
        return ApiResponse<SetupStatus>.Ok(status);
    }
}
