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

    [HttpGet("export")]
    public IActionResult Export([FromQuery] string? severity)
    {
        var events = _logs.ReadEvents(max: 5000, severity: severity).ToList();

        var sb = new System.Text.StringBuilder();
        sb.AppendLine("Time,Severity,Source,Message");
        foreach (var e in events)
        {
            sb.Append('"').Append(e.Time.ToString("yyyy-MM-dd HH:mm:ss")).Append("\",");
            sb.Append('"').Append(Csv(e.Severity)).Append("\",");
            sb.Append('"').Append(Csv(e.Source)).Append("\",");
            sb.Append('"').Append(Csv(e.Message)).Append('"').AppendLine();
        }

        var bytes = System.Text.Encoding.UTF8.GetBytes(sb.ToString());
        var name = $"events-{DateTime.UtcNow:yyyyMMdd-HHmmss}.csv";
        return File(bytes, "text/csv", name);
    }

    private static string Csv(string? v) => (v ?? string.Empty).Replace("\"", "\"\"");
}
