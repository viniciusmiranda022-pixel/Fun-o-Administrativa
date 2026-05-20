using IR.AdminFunctions.Web.Models;

namespace IR.AdminFunctions.Web.Services;

public class DashboardService
{
    private readonly TenantStore _tenants;
    private readonly BackupReader _backups;
    private readonly JobManager _jobs;
    private readonly LogReader _logs;

    public DashboardService(TenantStore tenants, BackupReader backups, JobManager jobs, LogReader logs)
    {
        _tenants = tenants;
        _backups = backups;
        _jobs = jobs;
        _logs = logs;
    }

    public async Task<DashboardStats> BuildAsync(string? tenantId)
    {
        var stats = new DashboardStats();
        var allTenants = await _tenants.ListAsync();
        var tenant = string.IsNullOrWhiteSpace(tenantId)
            ? allTenants.FirstOrDefault()
            : allTenants.FirstOrDefault(t => t.TenantId.Equals(tenantId, StringComparison.OrdinalIgnoreCase));

        stats.Tenant = tenant;

        if (tenant == null)
        {
            stats.ProtectionStatus = "NotProtected";
            return stats;
        }

        stats.ConsentBasicGranted = tenant.Consents.Basic.Granted;
        stats.ConsentRestoreGranted = tenant.Consents.Restore.Granted;

        var backups = _backups.List()
            .Where(b => string.IsNullOrEmpty(b.TenantId) || string.Equals(b.TenantId, tenant.TenantId, StringComparison.OrdinalIgnoreCase))
            .ToList();

        stats.BackupsCount = backups.Count;
        var last = backups.OrderByDescending(b => b.ModifiedAt).FirstOrDefault();
        if (last != null)
        {
            stats.LastBackupAt = last.ModifiedAt;
            stats.LastBackupId = last.Id;
            stats.ProtectedRoleDefinitionsCount = last.RoleDefinitionsCount;
            stats.ProtectedRoleAssignmentsCount = last.RoleAssignmentsCount;
        }

        var hasSchedule = tenant.BackupConfig.ScheduleEnabled;
        var hasAnyBackup = backups.Count > 0;
        var consentOk = tenant.Consents.Basic.Granted;

        if (hasSchedule && hasAnyBackup && consentOk)
            stats.ProtectionStatus = "Protected";
        else if (!consentOk)
            stats.ProtectionStatus = "NeedsAttention";
        else
            stats.ProtectionStatus = "NotProtected";

        stats.IsProtected = stats.ProtectionStatus == "Protected";

        stats.UnpackedBackupsCount = backups.Count;

        var jobs = _jobs.List(500).ToList();
        stats.TasksRunningCount = jobs.Count(j => j.Status == JobStatus.Running || j.Status == JobStatus.Queued);
        stats.TasksTotalCount = jobs.Count;

        var events = _logs.ReadEvents(500).ToList();
        stats.ErrorsCount = events.Count(e => string.Equals(e.Severity, "Error", StringComparison.OrdinalIgnoreCase))
            + jobs.Count(j => j.Status == JobStatus.Failed);
        stats.WarningsCount = events.Count(e => string.Equals(e.Severity, "Warning", StringComparison.OrdinalIgnoreCase));

        return stats;
    }
}
