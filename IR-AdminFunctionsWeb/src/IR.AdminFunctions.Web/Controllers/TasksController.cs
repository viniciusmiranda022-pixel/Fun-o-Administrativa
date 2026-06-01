using IR.AdminFunctions.Web.Models;
using IR.AdminFunctions.Web.Services;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Options;

namespace IR.AdminFunctions.Web.Controllers;

[ApiController]
[Route("api/[controller]")]
public class TasksController : ControllerBase
{
    private readonly LogReader _logs;
    private readonly JobManager _jobs;
    private readonly QuestOptions _options;

    public TasksController(LogReader logs, JobManager jobs, IOptions<QuestOptions> options)
    {
        _logs = logs;
        _jobs = jobs;
        _options = options.Value;
    }

    [HttpDelete]
    public IActionResult Clear()
    {
        _jobs.ClearAll();
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
    public ActionResult<ApiResponse<object>> List([FromQuery] string? since)
    {
        DateTime? sinceDate = null;
        if (!string.IsNullOrEmpty(since) && DateTime.TryParse(since, out var parsed))
            sinceDate = parsed;
        else if (string.IsNullOrEmpty(since))
            sinceDate = DateTime.UtcNow.AddDays(-30); // default: last 30 days

        var fromLogs = _logs.ReadTasks(since: sinceDate);
        var fromJobs = _jobs.List()
            .Where(j => !sinceDate.HasValue || (j.StartedAt?.UtcDateTime ?? j.CreatedAt.UtcDateTime) >= sinceDate.Value)
            .Select(j => new TaskEntry
            {
                Time = j.StartedAt?.UtcDateTime ?? j.CreatedAt.UtcDateTime,
                Name = j.Kind,
                Status = j.Status.ToString(),
                Type = j.Kind,
                Operation = j.Error ?? $"Job {j.Id}"
            });

        var combined = fromJobs.Concat(fromLogs)
            .OrderByDescending(t => t.Time)
            .Take(200);

        return ApiResponse<object>.Ok(new { items = combined });
    }
}
