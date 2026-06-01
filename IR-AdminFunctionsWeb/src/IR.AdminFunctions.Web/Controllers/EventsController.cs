using IR.AdminFunctions.Web.Models;
using IR.AdminFunctions.Web.Services;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Options;

namespace IR.AdminFunctions.Web.Controllers;

[ApiController]
[Route("api/[controller]")]
public class EventsController : ControllerBase
{
    private readonly LogReader _logs;
    private readonly QuestOptions _options;

    public EventsController(LogReader logs, IOptions<QuestOptions> options)
    {
        _logs = logs;
        _options = options.Value;
    }

    [HttpDelete]
    public IActionResult Clear()
    {
        DeleteLogFiles(_options.BackupLogFolder);
        DeleteLogFiles(_options.CompareLogFolder);
        return Ok(ApiResponse<object>.Ok(new { cleared = true }));
    }

    private static void DeleteLogFiles(string folder)
    {
        if (!Directory.Exists(folder)) return;
        foreach (var f in Directory.GetFiles(folder, "*.log"))
        {
            try { System.IO.File.Delete(f); } catch { }
        }
    }

    [HttpGet]
    public ActionResult<ApiResponse<object>> List([FromQuery] string? severity, [FromQuery] string? since)
    {
        DateTime? sinceDate = null;
        if (!string.IsNullOrEmpty(since) && DateTime.TryParse(since, out var parsed))
            sinceDate = parsed;
        else if (string.IsNullOrEmpty(since))
            sinceDate = DateTime.UtcNow.AddDays(-30); // default: last 30 days

        var events = _logs.ReadEvents(severity: severity, since: sinceDate);
        return ApiResponse<object>.Ok(new { items = events });
    }
}
