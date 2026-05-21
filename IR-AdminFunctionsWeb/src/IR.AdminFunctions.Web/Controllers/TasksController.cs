using IR.AdminFunctions.Web.Models;
using IR.AdminFunctions.Web.Services;
using Microsoft.AspNetCore.Mvc;

namespace IR.AdminFunctions.Web.Controllers;

[ApiController]
[Route("api/[controller]")]
public class TasksController : ControllerBase
{
    private readonly LogReader _logs;
    private readonly JobManager _jobs;

    public TasksController(LogReader logs, JobManager jobs)
    {
        _logs = logs;
        _jobs = jobs;
    }

    [HttpGet]
    public ActionResult<ApiResponse<object>> List([FromQuery] string? since)
    {
        DateTime? sinceDate = null;
        if (!string.IsNullOrEmpty(since) && DateTime.TryParse(since, out var parsed))
            sinceDate = parsed;
        else if (string.IsNullOrEmpty(since))
            sinceDate = DateTime.UtcNow.AddDays(-30); // padrão: últimos 30 dias

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
