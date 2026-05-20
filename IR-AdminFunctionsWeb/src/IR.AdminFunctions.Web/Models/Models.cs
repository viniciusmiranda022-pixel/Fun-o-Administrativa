using System.Text.Json;

namespace IR.AdminFunctions.Web.Models;

public class ApiResponse<T>
{
    public bool Success { get; set; }
    public T? Data { get; set; }
    public string? Error { get; set; }
    public string? Details { get; set; }

    public static ApiResponse<T> Ok(T data) => new() { Success = true, Data = data };
    public static ApiResponse<T> Fail(string error, string? details = null) =>
        new() { Success = false, Error = error, Details = details };
}

public class BackupSettings
{
    public string? TenantId { get; set; }
    public string? ClientId { get; set; }
    public string? AppDisplayName { get; set; }
    public string? CertificateThumbprint { get; set; }
    public string? CertificateStore { get; set; }
    public string? BackupRoot { get; set; }
    public string? LogRoot { get; set; }
    public string? LogoPath { get; set; }
    public int? RetentionDays { get; set; }
}

public class TenantInfo
{
    public string TenantId { get; set; } = string.Empty;
    public string? ClientId { get; set; }
    public string? AppDisplayName { get; set; }
    public string DisplayName { get; set; } = string.Empty;
    public bool ConfiguredCorrectly { get; set; }
    public string? ConsentStatus { get; set; }
}

public class TenantEntry
{
    public string Id { get; set; } = string.Empty;
    public string Name { get; set; } = string.Empty;
    public string TenantId { get; set; } = string.Empty;
    public string? ClientId { get; set; }
    public string? CertificateThumbprint { get; set; }
    public string? Domain { get; set; }
    public DateTimeOffset AddedAt { get; set; } = DateTimeOffset.UtcNow;
    public string? AddedByUser { get; set; }
    public TenantConsentsState Consents { get; set; } = new();
    public bool IsConfigured =>
        !string.IsNullOrWhiteSpace(TenantId) && !TenantId.StartsWith("PREENCHER");
}

public class TenantConsentsState
{
    public ConsentInfo Basic { get; set; } = new();
    public ConsentInfo Restore { get; set; } = new();
}

public class ConsentInfo
{
    public bool Granted { get; set; }
    public DateTimeOffset? GrantedAt { get; set; }
    public string? GrantedByUser { get; set; }
}

public class AppConfig
{
    public string? ClientId { get; set; }
    public string? ClientSecret { get; set; }
    public string? ApplicationObjectId { get; set; }
    public string? ServicePrincipalId { get; set; }
    public string? DisplayName { get; set; }
    public string? BootstrapTenantId { get; set; }
    public DateTimeOffset? CreatedAt { get; set; }

    public bool IsConfigured => !string.IsNullOrWhiteSpace(ClientId) && !string.IsNullOrWhiteSpace(ClientSecret);
}

public class SetupStartResult
{
    public string DeviceCode { get; set; } = string.Empty;
    public string UserCode { get; set; } = string.Empty;
    public string VerificationUrl { get; set; } = string.Empty;
    public string Message { get; set; } = string.Empty;
    public int ExpiresInSeconds { get; set; }
}

public class SetupStatus
{
    public string Status { get; set; } = "NotStarted"; // NotStarted | WaitingForUser | Registering | Completed | Failed
    public string? Message { get; set; }
    public string? ClientId { get; set; }
    public string? TenantId { get; set; }
    public string? UserCode { get; set; }
    public string? VerificationUrl { get; set; }
}

public class AddTenantRequest
{
    public string Name { get; set; } = string.Empty;
    public string TenantId { get; set; } = string.Empty;
    public string ClientId { get; set; } = string.Empty;
    public string CertificateThumbprint { get; set; } = string.Empty;
    public string? Domain { get; set; }
}

public class TenantStatusResult
{
    public string TenantId { get; set; } = string.Empty;
    public string Status { get; set; } = string.Empty;
    public string? Message { get; set; }
}

public class BackupSnapshot
{
    public string Id { get; set; } = string.Empty;
    public string Path { get; set; } = string.Empty;
    public DateTime CreatedAt { get; set; }
    public DateTime ModifiedAt { get; set; }
    public string? CollectedAt { get; set; }
    public string? TenantId { get; set; }
    public int? RoleDefinitionsCount { get; set; }
    public int? RoleAssignmentsCount { get; set; }
    public string? RoleDefinitionsHash { get; set; }
    public string? RoleAssignmentsHash { get; set; }
    public long SizeBytes { get; set; }
    public string? ManifestError { get; set; }

    public string SizeDisplay
    {
        get
        {
            double s = SizeBytes;
            string[] units = { "B", "KB", "MB", "GB" };
            int idx = 0;
            while (s >= 1024 && idx < units.Length - 1)
            {
                s /= 1024;
                idx++;
            }
            return $"{s:0.##} {units[idx]}";
        }
    }
}

public class UnpackedObjects
{
    public string BackupId { get; set; } = string.Empty;
    public JsonElement RoleDefinitions { get; set; }
    public JsonElement RoleAssignments { get; set; }
}

public class LogEvent
{
    public DateTime Time { get; set; }
    public string Source { get; set; } = string.Empty;
    public string Message { get; set; } = string.Empty;
    public string Severity { get; set; } = "Information";
}

public class TaskEntry
{
    public DateTime Time { get; set; }
    public string Name { get; set; } = string.Empty;
    public string Status { get; set; } = string.Empty;
    public string Type { get; set; } = string.Empty;
    public string Operation { get; set; } = string.Empty;
}

public enum JobStatus
{
    Queued,
    Running,
    Completed,
    Failed
}

public class Job
{
    public string Id { get; set; } = string.Empty;
    public string Kind { get; set; } = string.Empty;
    public JobStatus Status { get; set; }
    public DateTimeOffset CreatedAt { get; set; }
    public DateTimeOffset? StartedAt { get; set; }
    public DateTimeOffset? FinishedAt { get; set; }
    public IDictionary<string, object?>? Input { get; set; }
    public object? Result { get; set; }
    public string? Error { get; set; }
    public string? Details { get; set; }
}

public class CompareRunRequest
{
    public string? BackupId { get; set; }
}

public class RestoreRequest
{
    public string? BackupId { get; set; }
    public string? RoleName { get; set; }
    public bool RemoveExtraAssignments { get; set; }
}
