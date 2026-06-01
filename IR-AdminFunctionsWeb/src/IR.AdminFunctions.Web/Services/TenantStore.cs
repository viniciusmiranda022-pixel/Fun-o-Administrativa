using System.Security.Cryptography.X509Certificates;
using System.Text.Json;
using System.Text.Json.Serialization;
using IR.AdminFunctions.Web.Models;
using Microsoft.Extensions.Options;

namespace IR.AdminFunctions.Web.Services;

public class TenantStore
{
    private readonly string _tenantsFile;
    private readonly string _settingsFile;
    private readonly ILogger<TenantStore> _logger;
    private readonly AppConfigStore _appConfig;
    private readonly JsonSerializerOptions _json = new() { WriteIndented = true, PropertyNamingPolicy = JsonNamingPolicy.CamelCase };
    private readonly SemaphoreSlim _lock = new(1, 1);

    public TenantStore(IOptions<QuestOptions> options, AppConfigStore appConfig, ILogger<TenantStore> logger)
    {
        var configDir = Path.GetDirectoryName(options.Value.SettingsFile)!;
        _tenantsFile = Path.Combine(configDir, "tenants.json");
        _settingsFile = options.Value.SettingsFile;
        _appConfig = appConfig;
        _logger = logger;
    }

    public async Task<List<TenantEntry>> ListAsync()
    {
        await _lock.WaitAsync();
        try { return ReadFile(); }
        finally { _lock.Release(); }
    }

    public async Task<TenantEntry> AddAsync(AddTenantRequest req)
    {
        await _lock.WaitAsync();
        try
        {
            var tenants = ReadFile();
            if (tenants.Any(t => t.TenantId.Equals(req.TenantId, StringComparison.OrdinalIgnoreCase)))
                throw new InvalidOperationException($"Tenant {req.TenantId} already exists.");

            // Priority: request → settings.json → AppConfig (setup) → cert store
            var thumbprint = req.CertificateThumbprint;
            var appCfg0 = _appConfig.Read();
            if (string.IsNullOrWhiteSpace(thumbprint))
                thumbprint = ReadThumbprintFromSettingsJson(req.TenantId);
            if (string.IsNullOrWhiteSpace(thumbprint) && !string.IsNullOrWhiteSpace(appCfg0.CertificateThumbprint))
            {
                thumbprint = appCfg0.CertificateThumbprint;
                _logger.LogInformation("CertificateThumbprint imported from AppConfig (setup): {Tp}", thumbprint);
            }
            if (string.IsNullOrWhiteSpace(thumbprint))
            {
                thumbprint = AutoDetectThumbprint(appCfg0.ClientId);
                if (!string.IsNullOrWhiteSpace(thumbprint))
                    _logger.LogInformation("CertificateThumbprint automatically detected in the cert store: {Tp}", thumbprint);
            }

            var entry = new TenantEntry
            {
                Id = Guid.NewGuid().ToString("N"),
                Name = req.Name,
                TenantId = req.TenantId,
                ClientId = req.ClientId,
                CertificateThumbprint = thumbprint,
                Domain = req.Domain,
                AddedAt = DateTimeOffset.UtcNow
            };

            tenants.Add(entry);
            WriteFile(tenants);
            SyncSettingsJson(tenants);
            _logger.LogInformation("Tenant added: {Name} ({TenantId})", entry.Name, entry.TenantId);
            return entry;
        }
        finally { _lock.Release(); }
    }

    public async Task<bool> RemoveAsync(string tenantId)
    {
        await _lock.WaitAsync();
        try
        {
            var tenants = ReadFile();
            var removed = tenants.RemoveAll(t => t.TenantId.Equals(tenantId, StringComparison.OrdinalIgnoreCase));
            if (removed == 0) return false;
            WriteFile(tenants);
            SyncSettingsJson(tenants);
            _logger.LogInformation("Tenant removed: {TenantId}", tenantId);
            return true;
        }
        finally { _lock.Release(); }
    }

    public async Task<TenantEntry?> GetAsync(string tenantId)
    {
        var tenants = await ListAsync();
        return tenants.FirstOrDefault(t => t.TenantId.Equals(tenantId, StringComparison.OrdinalIgnoreCase));
    }

    public async Task<string?> AutoSyncCertificateAsync(string tenantId)
    {
        await _lock.WaitAsync();
        try
        {
            var tenants = ReadFile();
            var tenant = tenants.FirstOrDefault(t => t.TenantId.Equals(tenantId, StringComparison.OrdinalIgnoreCase));
            if (tenant == null) return null;

            if (!string.IsNullOrWhiteSpace(tenant.CertificateThumbprint))
                return tenant.CertificateThumbprint;

            var detected = AutoDetectThumbprint(_appConfig.Read().ClientId);
            if (string.IsNullOrWhiteSpace(detected)) return null;

            tenant.CertificateThumbprint = detected;
            WriteFile(tenants);
            SyncSettingsJson(tenants);
            _logger.LogInformation("CertificateThumbprint auto-detected and saved for tenant {TenantId}: {Tp}", tenantId, detected);
            return detected;
        }
        finally { _lock.Release(); }
    }

    public async Task<bool> UpdateCertificateAsync(string tenantId, string thumbprint)
    {
        await _lock.WaitAsync();
        try
        {
            var tenants = ReadFile();
            var tenant = tenants.FirstOrDefault(t => t.TenantId.Equals(tenantId, StringComparison.OrdinalIgnoreCase));
            if (tenant == null) return false;
            tenant.CertificateThumbprint = thumbprint;
            WriteFile(tenants);
            SyncSettingsJson(tenants);
            _logger.LogInformation("CertificateThumbprint updated for tenant {TenantId}", tenantId);
            return true;
        }
        finally { _lock.Release(); }
    }

    public async Task EnsureSettingsJsonSyncedAsync()
    {
        await _lock.WaitAsync();
        try { SyncSettingsJson(ReadFile()); }
        finally { _lock.Release(); }
    }

    public async Task<bool> UpdateConsentsAsync(string tenantId, TenantConsentsState consents)
    {
        await _lock.WaitAsync();
        try
        {
            var tenants = ReadFile();
            var tenant = tenants.FirstOrDefault(t => t.TenantId.Equals(tenantId, StringComparison.OrdinalIgnoreCase));
            if (tenant == null) return false;
            tenant.Consents = consents;
            WriteFile(tenants);
            return true;
        }
        finally { _lock.Release(); }
    }

    private List<TenantEntry> ReadFile()
    {
        if (!File.Exists(_tenantsFile)) return new List<TenantEntry>();
        try
        {
            var json = File.ReadAllText(_tenantsFile);
            return JsonSerializer.Deserialize<List<TenantEntry>>(json, _json) ?? new List<TenantEntry>();
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Error reading {File}", _tenantsFile);
            return new List<TenantEntry>();
        }
    }

    private void WriteFile(List<TenantEntry> tenants)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(_tenantsFile)!);
        File.WriteAllText(_tenantsFile, JsonSerializer.Serialize(tenants, _json));
    }

    private string? AutoDetectThumbprint(string? clientId)
    {
        if (string.IsNullOrWhiteSpace(clientId)) return null;
        foreach (var location in new[] { StoreLocation.LocalMachine, StoreLocation.CurrentUser })
        {
            try
            {
                using var store = new X509Store(StoreName.My, location);
                store.Open(OpenFlags.ReadOnly);
                var cert = store.Certificates
                    .Cast<X509Certificate2>()
                    .Where(c => c.NotAfter > DateTime.UtcNow &&
                                c.Subject.Contains(clientId, StringComparison.OrdinalIgnoreCase))
                    .OrderByDescending(c => c.NotAfter)
                    .FirstOrDefault();
                if (cert != null) return cert.Thumbprint;
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Error reading cert store {Location}", location);
            }
        }
        return null;
    }

    private string? ReadThumbprintFromSettingsJson(string tenantId)
    {
        try
        {
            if (!File.Exists(_settingsFile)) return null;
            var existing = JsonSerializer.Deserialize<Dictionary<string, JsonElement>>(
                File.ReadAllText(_settingsFile)) ?? new();

            // Only imports if settings.json belongs to the same tenant (or has no TenantId)
            if (existing.TryGetValue("TenantId", out var tidEl))
            {
                var tid = tidEl.GetString();
                if (!string.IsNullOrWhiteSpace(tid) &&
                    !tid.Equals(tenantId, StringComparison.OrdinalIgnoreCase))
                    return null;
            }

            if (existing.TryGetValue("CertificateThumbprint", out var tpEl))
            {
                var tp = tpEl.GetString();
                if (!string.IsNullOrWhiteSpace(tp)) return tp;
            }
            return null;
        }
        catch { return null; }
    }

    // Keeps settings.json (flat format) synchronized with the first configured tenant
    // for compatibility with existing PowerShell scripts.
    private void SyncSettingsJson(List<TenantEntry> tenants)
    {
        try
        {
            var primary = tenants.FirstOrDefault(t => t.IsConfigured) ?? tenants.FirstOrDefault();
            if (primary == null) return;

            var appCfg = _appConfig.Read();

            // Reads existing settings.json before building the flat dict to preserve extra fields
            Dictionary<string, JsonElement> existingSettings = new();
            if (File.Exists(_settingsFile))
            {
                existingSettings = JsonSerializer.Deserialize<Dictionary<string, JsonElement>>(
                    File.ReadAllText(_settingsFile)) ?? new();
            }

            // Thumbprint priority: AppConfig (current setup) → tenant → settings.json → cert store
            // AppConfig takes precedence as it is the source of truth after a new setup/reset.
            var thumbprint = !string.IsNullOrWhiteSpace(appCfg.CertificateThumbprint)
                ? appCfg.CertificateThumbprint
                : primary.CertificateThumbprint;
            if (string.IsNullOrWhiteSpace(thumbprint) &&
                existingSettings.TryGetValue("CertificateThumbprint", out var existingTpEl))
            {
                var existingTp = existingTpEl.GetString();
                if (!string.IsNullOrWhiteSpace(existingTp))
                    thumbprint = existingTp;
            }
            if (string.IsNullOrWhiteSpace(thumbprint))
            {
                var detected = AutoDetectThumbprint(appCfg.ClientId);
                if (!string.IsNullOrWhiteSpace(detected))
                {
                    thumbprint = detected;
                    _logger.LogInformation("SyncSettingsJson: CertificateThumbprint automatically detected: {Tp}", detected);
                }
            }

            var flat = new Dictionary<string, object?>
            {
                ["TenantId"] = primary.TenantId,
                ["ClientId"] = appCfg.ClientId ?? primary.ClientId,
                ["CertificateThumbprint"] = thumbprint,
                ["AppDisplayName"] = primary.Name,
                ["ClientSecret"] = appCfg.ClientSecret
            };

            // Preserves extra fields from existing settings.json (BackupRoot, LogRoot, etc.)
            foreach (var kv in existingSettings)
            {
                if (!flat.ContainsKey(kv.Key))
                    flat[kv.Key] = kv.Value;
            }

            Directory.CreateDirectory(Path.GetDirectoryName(_settingsFile)!);
            File.WriteAllText(_settingsFile, JsonSerializer.Serialize(flat,
                new JsonSerializerOptions { WriteIndented = true }));
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Failed to synchronize settings.json");
        }
    }
}
