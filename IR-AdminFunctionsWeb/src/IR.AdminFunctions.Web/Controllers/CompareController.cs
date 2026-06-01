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
    private readonly TenantStore _tenantStore;
    private readonly QuestOptions _options;
    private readonly AppConfigStore _appConfig;

    public CompareController(
        PowerShellRunner runner,
        JobManager jobs,
        SettingsReader settings,
        TenantStore tenantStore,
        IOptions<QuestOptions> options,
        AppConfigStore appConfig)
    {
        _runner = runner;
        _jobs = jobs;
        _settings = settings;
        _tenantStore = tenantStore;
        _options = options.Value;
        _appConfig = appConfig;
    }

    [HttpPost("run")]
    public async Task<ActionResult<ApiResponse<Job>>> Run([FromBody] CompareRunRequest? request)
    {
        await _tenantStore.EnsureSettingsJsonSyncedAsync();
        var cfg = _appConfig.Read();
        var raw = _settings.ReadRaw();

        if (raw == null || string.IsNullOrWhiteSpace(raw.TenantId) || string.IsNullOrWhiteSpace(raw.ClientId))
            return ApiResponse<Job>.Fail("settings.json is incomplete (TenantId/ClientId missing). Add a tenant first.");

        var hasAuth = !string.IsNullOrWhiteSpace(raw.CertificateThumbprint) || !string.IsNullOrWhiteSpace(cfg.ClientSecret);
        if (!hasAuth)
            return ApiResponse<Job>.Fail("No credentials configured. Run the setup and grant Basic consent.");

        var input = new Dictionary<string, object?>
        {
            ["Headless"] = true,
            ["TenantId"] = raw.TenantId,
            ["ClientId"] = raw.ClientId,
            ["Thumbprint"] = raw.CertificateThumbprint,
            ["ClientSecret"] = cfg.ClientSecret,
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
            var root = doc.RootElement;

            string? backupId = null;
            JsonElement dataElement;

            if (root.ValueKind == JsonValueKind.Object && root.TryGetProperty("SnapshotFolder", out var sfEl))
            {
                var snapshotFolder = sfEl.GetString();
                backupId = snapshotFolder != null ? Path.GetFileName(snapshotFolder) : null;
                dataElement = root.TryGetProperty("Data", out var dataEl) ? dataEl.Clone() : root.Clone();
            }
            else
            {
                dataElement = root.Clone();
            }

            return ApiResponse<object>.Ok(new
            {
                file = latest.Name,
                generatedAt = latest.LastWriteTimeUtc,
                backupId,
                data = dataElement
            });
        }
        catch (Exception ex)
        {
            return ApiResponse<object>.Fail("Failed reading comparison result", ex.Message);
        }
    }
}
