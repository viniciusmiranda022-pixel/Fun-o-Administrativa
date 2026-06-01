using IR.AdminFunctions.Web.Models;
using IR.AdminFunctions.Web.Services;
using Microsoft.AspNetCore.Mvc;

namespace IR.AdminFunctions.Web.Controllers;

[ApiController]
[Route("api/[controller]")]
public class JobsController : ControllerBase
{
    private readonly JobManager _jobs;

    public JobsController(JobManager jobs)
    {
        _jobs = jobs;
    }

    [HttpGet]
    public ActionResult<ApiResponse<IEnumerable<Job>>> List() =>
        ApiResponse<IEnumerable<Job>>.Ok(_jobs.List());

    [HttpGet("{id}")]
    public ActionResult<ApiResponse<Job>> Get(string id)
    {
        var job = _jobs.Get(id);
        if (job == null) return ApiResponse<Job>.Fail("Job not found");
        return ApiResponse<Job>.Ok(job);
    }
}
