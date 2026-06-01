using IR.AdminFunctions.Web.Models;
using IR.AdminFunctions.Web.Services;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Options;

namespace IR.AdminFunctions.Web.Controllers;

[ApiController]
[Route("api/[controller]")]
public class BackupsController : ControllerBase
{
    private readonly BackupReader _reader;
    private readonly PowerShellRunner _runner;
    private readonly JobManager _jobs;
    private readonly TenantStore _tenantStore;
    private readonly SettingsReader _settings;
    private readonly AppConfigStore _appConfig;
    private readonly QuestOptions _options;

    public BackupsController(
        BackupReader reader,
        PowerShellRunner runner,
        JobManager jobs,
        TenantStore tenantStore,
        SettingsReader settings,
        AppConfigStore appConfig,
        IOptions<QuestOptions> options)
    {
        _reader = reader;
        _runner = runner;
        _jobs = jobs;
        _tenantStore = tenantStore;
        _settings = settings;
        _appConfig = appConfig;
        _options = options.Value;
    }

    [HttpGet]
    public ActionResult<ApiResponse<IEnumerable<BackupSnapshot>>> List() =>
        ApiResponse<IEnumerable<BackupSnapshot>>.Ok(_reader.List());

    [HttpGet("{id}")]
    public ActionResult<ApiResponse<BackupSnapshot>> Get(string id)
    {
        var snap = _reader.Get(id);
        if (snap == null) return ApiResponse<BackupSnapshot>.Fail("Backup not found");
        return ApiResponse<BackupSnapshot>.Ok(snap);
    }

    [HttpPost("run")]
    public ActionResult<ApiResponse<Job>> Run()
    {
        var job = _jobs.Enqueue("backup", null, async (_, ct) =>
        {
            // Ensures settings.json is synchronized with appConfig before running the script
            await _tenantStore.EnsureSettingsJsonSyncedAsync();

            var raw = _settings.ReadRaw();
            var cfg = _appConfig.Read();
            var input = new Dictionary<string, object?>
            {
                ["RmadMode"] = true,
                ["TenantId"] = raw?.TenantId,
                ["ClientId"] = raw?.ClientId,
                ["Thumbprint"] = raw?.CertificateThumbprint,
                ["ClientSecret"] = cfg.ClientSecret,
            };

            var result = await _runner.RunAsync(
                _options.BackupScript,
                input,
                _options.DefaultJobTimeoutSeconds,
                ct);

            if (!result.Success)
            {
                var msg = result.Errors.Count > 0
                    ? string.Join(" | ", result.Errors.Take(3))
                    : "Script returned with errors and no details";
                throw new InvalidOperationException(msg);
            }

            return new
            {
                result.Output,
                result.Stdout,
                result.ElapsedMilliseconds
            };
        });

        return ApiResponse<Job>.Ok(job);
    }

    [HttpPost("{id}/unpack")]
    public ActionResult<ApiResponse<Job>> Unpack(string id, [FromBody] UnpackRequest? req)
    {
        var snap = _reader.Get(id);
        if (snap == null) return ApiResponse<Job>.Fail("Backup not found");

        var job = _jobs.Enqueue("unpack", new Dictionary<string, object?> { ["BackupId"] = id }, async (_, ct) =>
        {
            if (req?.Validate == true)
            {
                var required = new[] { "manifest.json", "roleDefinitions.json", "roleAssignments.json" };
                var missing = required.Where(f => !System.IO.File.Exists(Path.Combine(snap.Path, f))).ToList();
                if (missing.Count > 0)
                    throw new InvalidOperationException($"Missing file(s) in backup: {string.Join(", ", missing)}");
            }

            if (req?.RunDiff != false)
            {
                await _tenantStore.EnsureSettingsJsonSyncedAsync();
                var raw = _settings.ReadRaw();
                var cfg = _appConfig.Read();

                if (raw == null || string.IsNullOrWhiteSpace(raw.TenantId) || string.IsNullOrWhiteSpace(raw.ClientId))
                    throw new InvalidOperationException("settings.json is incomplete (TenantId/ClientId missing). Add a tenant first.");

                var hasAuth = !string.IsNullOrWhiteSpace(raw.CertificateThumbprint) || !string.IsNullOrWhiteSpace(cfg.ClientSecret);
                if (!hasAuth)
                    throw new InvalidOperationException("No credentials configured. Run the setup and grant Basic consent.");

                var input = new Dictionary<string, object?>
                {
                    ["Headless"] = true,
                    ["TenantId"] = raw.TenantId,
                    ["ClientId"] = raw.ClientId,
                    ["Thumbprint"] = raw.CertificateThumbprint,
                    ["ClientSecret"] = cfg.ClientSecret,
                    ["BackupId"] = id
                };
                var result = await _runner.RunAsync(_options.CompareScript, input, _options.DefaultJobTimeoutSeconds, ct);
                if (!result.Success)
                {
                    var msg = result.Errors.Count > 0
                        ? string.Join(" | ", result.Errors.Take(3))
                        : "Comparison failed with no details";
                    throw new InvalidOperationException(msg);
                }
            }

            return new { BackupId = id };
        });

        return ApiResponse<Job>.Ok(job);
    }
}
