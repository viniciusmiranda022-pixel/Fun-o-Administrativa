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
    public ActionResult<ApiResponse<object>> List([FromQuery] string? severity, [FromQuery] string? since)
    {
        DateTime? sinceDate = null;
        if (!string.IsNullOrEmpty(since) && DateTime.TryParse(since, out var parsed))
            sinceDate = parsed;
        else if (string.IsNullOrEmpty(since))
            sinceDate = DateTime.UtcNow.AddDays(-30); // padrão: últimos 30 dias

        var events = _logs.ReadEvents(severity: severity, since: sinceDate);
        return ApiResponse<object>.Ok(new { items = events });
    }
}
