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
    private readonly AppConfigStore _appConfig;

    public BackupsController(
        BackupReader reader,
        PowerShellRunner runner,
        JobManager jobs,
        IOptions<QuestOptions> options,
        AppConfigStore appConfig)
    {
        _reader = reader;
        _runner = runner;
        _jobs = jobs;
        _options = options.Value;
        _appConfig = appConfig;
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
        var cfg = _appConfig.Read();
        var input = new Dictionary<string, object?>
        {
            ["RmadMode"] = true,
            ["ClientSecret"] = cfg.ClientSecret ?? ""
        };

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
}
