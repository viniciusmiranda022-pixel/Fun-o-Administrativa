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
        if (snap == null) return ApiResponse<BackupSnapshot>.Fail("Backup não encontrado");
        return ApiResponse<BackupSnapshot>.Ok(snap);
    }

    [HttpPost("run")]
    public ActionResult<ApiResponse<Job>> Run()
    {
        var input = new Dictionary<string, object?>
        {
            ["RmadMode"] = true,
        };

        var job = _jobs.Enqueue("backup", input, async (_, ct) =>
        {
            // Garante que o settings.json está sincronizado com appConfig antes de rodar o script
            await _tenantStore.EnsureSettingsJsonSyncedAsync();

            var result = await _runner.RunAsync(
                _options.BackupScript,
                input,
                _options.DefaultJobTimeoutSeconds,
                ct);

            if (!result.Success)
            {
                var msg = result.Errors.Count > 0
                    ? string.Join(" | ", result.Errors.Take(3))
                    : "Script retornou com erros sem detalhes";
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
        if (snap == null) return ApiResponse<Job>.Fail("Backup não encontrado");

        var job = _jobs.Enqueue("unpack", new Dictionary<string, object?> { ["BackupId"] = id }, async (_, ct) =>
        {
            if (req?.Validate == true)
            {
                var required = new[] { "manifest.json", "roleDefinitions.json", "roleAssignments.json" };
                var missing = required.Where(f => !File.Exists(Path.Combine(snap.Path, f))).ToList();
                if (missing.Count > 0)
                    throw new InvalidOperationException($"Arquivo(s) ausente(s) no backup: {string.Join(", ", missing)}");
            }

            if (req?.RunDiff != false)
            {
                var raw = _settings.ReadRaw();
                var cfg = _appConfig.Read();
                var input = new Dictionary<string, object?>
                {
                    ["Headless"] = true,
                    ["TenantId"] = raw?.TenantId ?? "",
                    ["ClientId"] = raw?.ClientId ?? "",
                    ["Thumbprint"] = raw?.CertificateThumbprint ?? "",
                    ["ClientSecret"] = cfg.ClientSecret ?? "",
                    ["BackupId"] = id
                };
                var result = await _runner.RunAsync(_options.CompareScript, input, _options.DefaultJobTimeoutSeconds, ct);
                if (!result.Success)
                {
                    var msg = result.Errors.Count > 0
                        ? string.Join(" | ", result.Errors.Take(3))
                        : "Comparação falhou sem detalhes";
                    throw new InvalidOperationException(msg);
                }
            }

            return new { BackupId = id };
        });

        return ApiResponse<Job>.Ok(job);
    }
}
