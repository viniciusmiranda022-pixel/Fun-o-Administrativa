using System.Collections.Concurrent;
using IR.AdminFunctions.Web.Models;

namespace IR.AdminFunctions.Web.Services;

public class JobManager
{
    private readonly ConcurrentDictionary<string, Job> _jobs = new();
    private readonly ILogger<JobManager> _logger;

    public JobManager(ILogger<JobManager> logger)
    {
        _logger = logger;
    }

    public Job Create(string kind, IDictionary<string, object?>? input = null)
    {
        var job = new Job
        {
            Id = Guid.NewGuid().ToString("N"),
            Kind = kind,
            Status = JobStatus.Queued,
            CreatedAt = DateTimeOffset.UtcNow,
            Input = input
        };
        _jobs[job.Id] = job;
        return job;
    }

    public Job? Get(string id) => _jobs.TryGetValue(id, out var job) ? job : null;

    public IEnumerable<Job> List(int take = 100) =>
        _jobs.Values.OrderByDescending(j => j.CreatedAt).Take(take);

    public void Start(string id)
    {
        if (_jobs.TryGetValue(id, out var job))
        {
            job.Status = JobStatus.Running;
            job.StartedAt = DateTimeOffset.UtcNow;
        }
    }

    public void Complete(string id, object? result)
    {
        if (_jobs.TryGetValue(id, out var job))
        {
            job.Status = JobStatus.Completed;
            job.FinishedAt = DateTimeOffset.UtcNow;
            job.Result = result;
        }
    }

    public void Fail(string id, string error, string? details = null)
    {
        if (_jobs.TryGetValue(id, out var job))
        {
            job.Status = JobStatus.Failed;
            job.FinishedAt = DateTimeOffset.UtcNow;
            job.Error = error;
            job.Details = details;
        }
    }

    public void ClearAll() => _jobs.Clear();

    public Job Enqueue(string kind, IDictionary<string, object?>? input, Func<Job, CancellationToken, Task<object?>> work)
    {
        var job = Create(kind, input);

        _ = Task.Run(async () =>
        {
            Start(job.Id);
            try
            {
                var result = await work(job, CancellationToken.None).ConfigureAwait(false);
                Complete(job.Id, result);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Job {JobId} ({Kind}) falhou", job.Id, job.Kind);
                Fail(job.Id, ex.Message, ex.StackTrace);
            }
        });

        return job;
    }
}
