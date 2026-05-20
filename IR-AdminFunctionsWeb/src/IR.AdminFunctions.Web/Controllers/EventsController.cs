using IR.AdminFunctions.Web.Models;
using IR.AdminFunctions.Web.Services;
using Microsoft.AspNetCore.Mvc;

namespace IR.AdminFunctions.Web.Controllers;

[ApiController]
[Route("api/[controller]")]
public class EventsController : ControllerBase
{
    private readonly LogReader _logs;

    public EventsController(LogReader logs)
    {
        _logs = logs;
    }

    [HttpGet]
    public ActionResult<ApiResponse<object>> List([FromQuery] string? severity)
    {
        var events = _logs.ReadEvents(severity: severity);
        return ApiResponse<object>.Ok(new { items = events });
    }
}
