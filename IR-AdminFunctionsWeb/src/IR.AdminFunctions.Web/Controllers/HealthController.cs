using IR.AdminFunctions.Web.Models;
using Microsoft.AspNetCore.Mvc;

namespace IR.AdminFunctions.Web.Controllers;

[ApiController]
[Route("api/[controller]")]
public class HealthController : ControllerBase
{
    [HttpGet]
    public ActionResult<ApiResponse<object>> Get() =>
        ApiResponse<object>.Ok(new
        {
            status = "ok",
            time = DateTimeOffset.UtcNow,
            version = typeof(HealthController).Assembly.GetName().Version?.ToString()
        });
}
