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
    private readonly QuestOptions _options;

    public BackupsController(
        BackupReader reader,
        PowerShellRunner runner,
        JobManager jobs,
        IOptions<QuestOptions> options)
    {
        _reader = reader;
        _runner = runner;
        _jobs = jobs;
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
    public ActionResult<ApiResponse<Job>> Run([FromBody] BackupRunRequest? req = null)
    {
        var input = new Dictionary<string, object?> { ["RmadMode"] = true };
        if (!string.IsNullOrWhiteSpace(req?.TenantId)) input["TenantId"] = req.TenantId;

        var job = _jobs.Enqueue("backup", input, async (_, ct) =>
        {
            var result = await _runner.RunAsync(
                _options.BackupScript,
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

    [HttpPost("unpack")]
    public ActionResult<ApiResponse<Job>> Unpack([FromBody] UnpackRequest req)
    {
        if (req == null || string.IsNullOrWhiteSpace(req.BackupId))
            return BadRequest(ApiResponse<Job>.Fail("BackupId é obrigatório."));

        var snap = _reader.Get(req.BackupId);
        if (snap == null)
            return NotFound(ApiResponse<Job>.Fail($"Backup {req.BackupId} não encontrado."));

        var input = new Dictionary<string, object?>
        {
            ["BackupId"] = req.BackupId,
            ["BackupPath"] = snap.Path,
            ["ClearPrevious"] = req.ClearPrevious,
            ["ComputeDifferences"] = req.ComputeDifferences,
            ["ValidateIntegrity"] = req.ValidateIntegrity
        };

        var job = _jobs.Enqueue("unpack", input, async (_, ct) =>
        {
            await Task.Delay(600, ct);
            return new
            {
                success = true,
                backupId = req.BackupId,
                backupPath = snap.Path,
                roleDefinitions = snap.RoleDefinitionsCount,
                roleAssignments = snap.RoleAssignmentsCount,
                validatedHash = snap.RoleDefinitionsHash,
                clearPrevious = req.ClearPrevious,
                computeDifferences = req.ComputeDifferences
            };
        });

        return ApiResponse<Job>.Ok(job);
    }
}

public class BackupRunRequest
{
    public string? TenantId { get; set; }
}

public class UnpackRequest
{
    public string? BackupId { get; set; }
    public bool ClearPrevious { get; set; }
    public bool ComputeDifferences { get; set; }
    public bool ValidateIntegrity { get; set; }
}
