namespace IR.AdminFunctions.Web.Services;

/// <summary>
/// Background service that prunes completed jobs older than 30 days every 6 hours,
/// preventing the in-memory ConcurrentDictionary from growing indefinitely.
/// </summary>
public class JobCleanupService : BackgroundService
{
    private readonly JobManager _jobs;
    private readonly ILogger<JobCleanupService> _logger;
    private static readonly TimeSpan Interval = TimeSpan.FromHours(6);
    private const int MaxAgeDays = 30;

    public JobCleanupService(JobManager jobs, ILogger<JobCleanupService> logger)
    {
        _jobs = jobs;
        _logger = logger;
    }

    protected override async Task ExecuteAsync(CancellationToken stoppingToken)
    {
        // Initial delay so startup logs are clean
        await Task.Delay(TimeSpan.FromMinutes(1), stoppingToken).ConfigureAwait(false);

        while (!stoppingToken.IsCancellationRequested)
        {
            try
            {
                var removed = _jobs.Prune(MaxAgeDays);
                if (removed > 0)
                    _logger.LogInformation("JobCleanup: removed {Count} jobs older than {Days} days", removed, MaxAgeDays);
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "JobCleanup: error during pruning");
            }

            await Task.Delay(Interval, stoppingToken).ConfigureAwait(false);
        }
    }
}
