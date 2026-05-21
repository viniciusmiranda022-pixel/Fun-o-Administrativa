using IR.AdminFunctions.Web.Models;
using IR.AdminFunctions.Web.Services;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Options;

namespace IR.AdminFunctions.Web.Controllers;

[ApiController]
[Route("api/restore")]
public class RestoreController : ControllerBase
{
    private readonly PowerShellRunner _runner;
    private readonly JobManager _jobs;
    private readonly SettingsReader _settings;
    private readonly BackupReader _backups;
    private readonly QuestOptions _options;
    private readonly AppConfigStore _appConfig;

    public RestoreController(
        PowerShellRunner runner,
        JobManager jobs,
        SettingsReader settings,
        BackupReader backups,
        IOptions<QuestOptions> options,
        AppConfigStore appConfig)
    {
        _runner = runner;
        _jobs = jobs;
        _settings = settings;
        _backups = backups;
        _options = options.Value;
        _appConfig = appConfig;
    }

    [HttpPost("preview")]
    public ActionResult<ApiResponse<Job>> Preview([FromBody] RestoreRequest request) =>
        Enqueue(request, "WhatIf");

    [HttpPost("apply")]
    public ActionResult<ApiResponse<Job>> Apply([FromBody] RestoreRequest request) =>
        Enqueue(request, "Apply");

    private ActionResult<ApiResponse<Job>> Enqueue(RestoreRequest request, string mode)
    {
        if (request == null || string.IsNullOrWhiteSpace(request.BackupId) || string.IsNullOrWhiteSpace(request.RoleName))
        {
            return ApiResponse<Job>.Fail("BackupId e RoleName são obrigatórios.");
        }

        var snap = _backups.Get(request.BackupId);
        if (snap == null)
        {
            return ApiResponse<Job>.Fail($"Backup {request.BackupId} não encontrado.");
        }

        var cfg = _appConfig.Read();
        var raw = _settings.ReadRaw();

        if (raw == null || string.IsNullOrWhiteSpace(raw.TenantId) || string.IsNullOrWhiteSpace(raw.ClientId))
            return ApiResponse<Job>.Fail("settings.json incompleto. Adicione um tenant primeiro.");

        var hasAuth = !string.IsNullOrWhiteSpace(raw.CertificateThumbprint) || !string.IsNullOrWhiteSpace(cfg.ClientSecret);
        if (!hasAuth)
            return ApiResponse<Job>.Fail("Nenhuma credencial configurada. Conceda o consentimento Restore primeiro.");

        var input = new Dictionary<string, object?>
        {
            ["TenantId"] = raw.TenantId,
            ["ClientId"] = raw.ClientId,
            ["CertificateThumbprint"] = raw.CertificateThumbprint ?? "",
            ["ClientSecret"] = cfg.ClientSecret ?? "",
            ["SnapshotFolder"] = snap.Path,
            ["RoleName"] = request.RoleName,
            ["Mode"] = mode,
            ["SkipConfirmation"] = mode == "Apply",
            ["RemoveExtraAssignments"] = request.RemoveExtraAssignments
        };

        var job = _jobs.Enqueue("restore-" + mode.ToLowerInvariant(), input, async (_, ct) =>
        {
            var result = await _runner.RunAsync(
                _options.RestoreScript,
                input,
                _options.DefaultJobTimeoutSeconds,
                ct);

            return new
            {
                result.Success,
                result.Output,
                result.Stdout,
                result.Errors,
                result.ElapsedMilliseconds
            };
        });

        return ApiResponse<Job>.Ok(job);
    }
}
