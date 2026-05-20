using IR.AdminFunctions.Web.Models;
using IR.AdminFunctions.Web.Services;
using Microsoft.AspNetCore.Mvc;

namespace IR.AdminFunctions.Web.Controllers;

[ApiController]
[Route("api/dashboard")]
public class DashboardController : ControllerBase
{
    private readonly DashboardService _dashboard;

    public DashboardController(DashboardService dashboard)
    {
        _dashboard = dashboard;
    }

    [HttpGet]
    public async Task<ActionResult<ApiResponse<DashboardStats>>> Get([FromQuery] string? tenantId)
    {
        var stats = await _dashboard.BuildAsync(tenantId);
        return ApiResponse<DashboardStats>.Ok(stats);
    }
}
