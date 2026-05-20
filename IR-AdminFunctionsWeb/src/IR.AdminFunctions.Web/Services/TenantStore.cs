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
    private readonly JsonSerializerOptions _json = new() { WriteIndented = true, PropertyNamingPolicy = JsonNamingPolicy.CamelCase };
    private readonly SemaphoreSlim _lock = new(1, 1);

    public TenantStore(IOptions<QuestOptions> options, ILogger<TenantStore> logger)
    {
        var configDir = Path.GetDirectoryName(options.Value.SettingsFile)!;
        _tenantsFile = Path.Combine(configDir, "tenants.json");
        _settingsFile = options.Value.SettingsFile;
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
                throw new InvalidOperationException($"Tenant {req.TenantId} já existe.");

            var entry = new TenantEntry
            {
                Id = Guid.NewGuid().ToString("N"),
                Name = req.Name,
                TenantId = req.TenantId,
                ClientId = req.ClientId,
                CertificateThumbprint = req.CertificateThumbprint,
                Domain = req.Domain,
                AddedAt = DateTimeOffset.UtcNow
            };

            tenants.Add(entry);
            WriteFile(tenants);
            SyncSettingsJson(tenants);
            _logger.LogInformation("Tenant adicionado: {Name} ({TenantId})", entry.Name, entry.TenantId);
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
            _logger.LogInformation("Tenant removido: {TenantId}", tenantId);
            return true;
        }
        finally { _lock.Release(); }
    }

    public async Task<TenantEntry?> GetAsync(string tenantId)
    {
        var tenants = await ListAsync();
        return tenants.FirstOrDefault(t => t.TenantId.Equals(tenantId, StringComparison.OrdinalIgnoreCase));
    }

    public async Task<TenantEntry?> UpdateBackupConfigAsync(string tenantId, BackupConfig cfg)
    {
        await _lock.WaitAsync();
        try
        {
            var tenants = ReadFile();
            var tenant = tenants.FirstOrDefault(t => t.TenantId.Equals(tenantId, StringComparison.OrdinalIgnoreCase));
            if (tenant == null) return null;
            tenant.BackupConfig = cfg;
            WriteFile(tenants);
            _logger.LogInformation("BackupConfig atualizada para {TenantId}", tenantId);
            return tenant;
        }
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
            _logger.LogError(ex, "Erro lendo {File}", _tenantsFile);
            return new List<TenantEntry>();
        }
    }

    private void WriteFile(List<TenantEntry> tenants)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(_tenantsFile)!);
        File.WriteAllText(_tenantsFile, JsonSerializer.Serialize(tenants, _json));
    }

    // Mantém settings.json (formato plano) sincronizado com o primeiro tenant configurado
    // para compatibilidade com os scripts PowerShell existentes.
    private void SyncSettingsJson(List<TenantEntry> tenants)
    {
        try
        {
            var primary = tenants.FirstOrDefault(t => t.IsConfigured) ?? tenants.FirstOrDefault();
            if (primary == null) return;

            var flat = new Dictionary<string, object?>
            {
                ["TenantId"] = primary.TenantId,
                ["ClientId"] = primary.ClientId,
                ["CertificateThumbprint"] = primary.CertificateThumbprint,
                ["AppDisplayName"] = primary.Name
            };

            // Preserva campos extras do settings.json existente (BackupRoot, LogRoot, etc.)
            if (File.Exists(_settingsFile))
            {
                var existing = JsonSerializer.Deserialize<Dictionary<string, JsonElement>>(
                    File.ReadAllText(_settingsFile)) ?? new();
                foreach (var kv in existing)
                {
                    if (!flat.ContainsKey(kv.Key))
                        flat[kv.Key] = kv.Value;
                }
            }

            Directory.CreateDirectory(Path.GetDirectoryName(_settingsFile)!);
            File.WriteAllText(_settingsFile, JsonSerializer.Serialize(flat,
                new JsonSerializerOptions { WriteIndented = true }));
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "Falha ao sincronizar settings.json");
        }
    }
}
