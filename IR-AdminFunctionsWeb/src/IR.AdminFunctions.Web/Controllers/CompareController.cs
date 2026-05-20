using System.Text.Json;
using IR.AdminFunctions.Web.Models;
using IR.AdminFunctions.Web.Services;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Options;

namespace IR.AdminFunctions.Web.Controllers;

[ApiController]
[Route("api/compare")]
public class CompareController : ControllerBase
{
    private readonly PowerShellRunner _runner;
    private readonly JobManager _jobs;
    private readonly SettingsReader _settings;
    private readonly QuestOptions _options;

    public CompareController(
        PowerShellRunner runner,
        JobManager jobs,
        SettingsReader settings,
        IOptions<QuestOptions> options)
    {
        _runner = runner;
        _jobs = jobs;
        _settings = settings;
        _options = options.Value;
    }

    [HttpPost("run")]
    public ActionResult<ApiResponse<Job>> Run([FromBody] CompareRunRequest? request)
    {
        var raw = _settings.ReadRaw();
        if (raw == null || string.IsNullOrWhiteSpace(raw.TenantId)
            || string.IsNullOrWhiteSpace(raw.ClientId)
            || string.IsNullOrWhiteSpace(raw.CertificateThumbprint))
        {
            return ApiResponse<Job>.Fail("settings.json incompleto (TenantId/ClientId/CertificateThumbprint).");
        }

        var input = new Dictionary<string, object?>
        {
            ["Headless"] = true,
            ["TenantId"] = raw.TenantId,
            ["ClientId"] = raw.ClientId,
            ["Thumbprint"] = raw.CertificateThumbprint,
            ["BackupId"] = request?.BackupId
        };

        var job = _jobs.Enqueue("compare", input, async (j, ct) =>
        {
            var result = await _runner.RunAsync(
                _options.CompareScript,
                input,
                _options.DefaultJobTimeoutSeconds,
                ct);

            var outputFile = result.Output.LastOrDefault(o => !string.IsNullOrWhiteSpace(o));
            return new
            {
                result.Success,
                outputFile,
                result.Errors,
                result.ElapsedMilliseconds
            };
        });

        return ApiResponse<Job>.Ok(job);
    }

    [HttpGet("results")]
    public ActionResult<ApiResponse<object>> LatestResults()
    {
        var dir = _options.CompareOutputFolder;
        if (!Directory.Exists(dir))
        {
            return ApiResponse<object>.Ok(new { items = Array.Empty<object>() });
        }

        var latest = new DirectoryInfo(dir)
            .GetFiles("compare-*.json")
            .OrderByDescending(f => f.LastWriteTimeUtc)
            .FirstOrDefault();

        if (latest == null)
        {
            return ApiResponse<object>.Ok(new { items = Array.Empty<object>() });
        }

        try
        {
            using var doc = JsonDocument.Parse(System.IO.File.ReadAllText(latest.FullName));
            return ApiResponse<object>.Ok(new
            {
                file = latest.Name,
                generatedAt = latest.LastWriteTimeUtc,
                data = doc.RootElement.Clone()
            });
        }
        catch (Exception ex)
        {
            return ApiResponse<object>.Fail("Falha lendo resultado da comparação", ex.Message);
        }
    }
}
