using System.Collections.Concurrent;
using IR.AdminFunctions.Web.Models;

namespace IR.AdminFunctions.Web.Services;

public class JobManager
{
    private readonly ConcurrentDictionary<string, Job> _jobs = new();
    private readonly ILogger<JobManager> _logger;
    private readonly IHostApplicationLifetime _lifetime;

    public JobManager(ILogger<JobManager> logger, IHostApplicationLifetime lifetime)
    {
        _logger = logger;
        _lifetime = lifetime;
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

    /// <summary>Removes completed/failed jobs older than <paramref name="maxAgeDays"/> days.</summary>
    public int Prune(int maxAgeDays = 30)
    {
        var cutoff = DateTimeOffset.UtcNow.AddDays(-maxAgeDays);
        var toRemove = _jobs.Values
            .Where(j => j.Status is JobStatus.Completed or JobStatus.Failed
                     && (j.FinishedAt ?? j.CreatedAt) < cutoff)
            .Select(j => j.Id)
            .ToList();

        foreach (var id in toRemove)
            _jobs.TryRemove(id, out _);

        return toRemove.Count;
    }

    public Job Enqueue(string kind, IDictionary<string, object?>? input, Func<Job, CancellationToken, Task<object?>> work)
    {
        var job = Create(kind, input);
        // Use the application stopping token so jobs receive a cancellation signal on shutdown
        var ct = _lifetime.ApplicationStopping;

        _ = Task.Run(async () =>
        {
            Start(job.Id);
            try
            {
                var result = await work(job, ct).ConfigureAwait(false);
                Complete(job.Id, result);
            }
            catch (OperationCanceledException) when (ct.IsCancellationRequested)
            {
                Fail(job.Id, "Job cancelled: application is shutting down");
                _logger.LogWarning("Job {JobId} ({Kind}) cancelled due to application shutdown", job.Id, job.Kind);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Job {JobId} ({Kind}) failed", job.Id, job.Kind);
                Fail(job.Id, ex.Message, ex.StackTrace);
            }
        }, ct);

        return job;
    }
}

