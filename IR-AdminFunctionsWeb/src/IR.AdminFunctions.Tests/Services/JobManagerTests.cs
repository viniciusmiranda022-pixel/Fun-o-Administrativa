using IR.AdminFunctions.Web.Models;
using IR.AdminFunctions.Web.Services;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging.Abstractions;
using Moq;

namespace IR.AdminFunctions.Tests.Services;

public class JobManagerTests
{
    private static JobManager CreateManager()
    {
        var lifetime = new Mock<IHostApplicationLifetime>();
        lifetime.Setup(l => l.ApplicationStopping).Returns(CancellationToken.None);
        return new JobManager(NullLogger<JobManager>.Instance, lifetime.Object);
    }

    [Fact]
    public void Create_AssignsUniqueId()
    {
        var mgr = CreateManager();
        var j1 = mgr.Create("backup");
        var j2 = mgr.Create("backup");
        Assert.NotEqual(j1.Id, j2.Id);
    }

    [Fact]
    public void Create_SetsStatusQueued()
    {
        var mgr = CreateManager();
        var job = mgr.Create("compare");
        Assert.Equal(JobStatus.Queued, job.Status);
    }

    [Fact]
    public void Complete_SetsStatusAndResult()
    {
        var mgr = CreateManager();
        var job = mgr.Create("backup");
        mgr.Start(job.Id);
        mgr.Complete(job.Id, new { ok = true });
        var stored = mgr.Get(job.Id);
        Assert.Equal(JobStatus.Completed, stored!.Status);
        Assert.NotNull(stored.FinishedAt);
    }

    [Fact]
    public void Fail_SetsStatusAndError()
    {
        var mgr = CreateManager();
        var job = mgr.Create("backup");
        mgr.Start(job.Id);
        mgr.Fail(job.Id, "boom");
        var stored = mgr.Get(job.Id);
        Assert.Equal(JobStatus.Failed, stored!.Status);
        Assert.Equal("boom", stored.Error);
    }

    [Fact]
    public void Prune_RemovesOldCompletedJobs()
    {
        var mgr = CreateManager();
        var job = mgr.Create("backup");
        mgr.Start(job.Id);
        mgr.Complete(job.Id, null);
        // Manually age the job by directly setting FinishedAt (reflection since it's a POCO)
        typeof(IR.AdminFunctions.Web.Models.Job)
            .GetProperty("FinishedAt")!
            .SetValue(job, DateTimeOffset.UtcNow.AddDays(-31));

        var removed = mgr.Prune(30);
        Assert.Equal(1, removed);
        Assert.Null(mgr.Get(job.Id));
    }

    [Fact]
    public void Prune_KeepsRecentJobs()
    {
        var mgr = CreateManager();
        var job = mgr.Create("backup");
        mgr.Start(job.Id);
        mgr.Complete(job.Id, null);
        // recent — should NOT be pruned
        var removed = mgr.Prune(30);
        Assert.Equal(0, removed);
        Assert.NotNull(mgr.Get(job.Id));
    }

    [Fact]
    public void Prune_KeepsRunningJobs()
    {
        var mgr = CreateManager();
        var job = mgr.Create("backup");
        mgr.Start(job.Id);
        // still running — should not be pruned even if old
        typeof(IR.AdminFunctions.Web.Models.Job)
            .GetProperty("StartedAt")!
            .SetValue(job, DateTimeOffset.UtcNow.AddDays(-60));

        var removed = mgr.Prune(30);
        Assert.Equal(0, removed);
    }
}
