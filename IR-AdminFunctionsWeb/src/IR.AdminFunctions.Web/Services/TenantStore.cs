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
                throw new InvalidOperationException($"Tenant {req.TenantId} já existe.");

            // Se o thumbprint não foi informado, tenta importar do settings.json existente
            // e em seguida do cert store do Windows (pelo ClientId do app registrado)
            var thumbprint = req.CertificateThumbprint;
            if (string.IsNullOrWhiteSpace(thumbprint))
                thumbprint = ReadThumbprintFromSettingsJson(req.TenantId);
            if (string.IsNullOrWhiteSpace(thumbprint))
            {
                thumbprint = AutoDetectThumbprint(_appConfig.Read().ClientId);
                if (!string.IsNullOrWhiteSpace(thumbprint))
                    _logger.LogInformation("CertificateThumbprint detectado automaticamente no cert store: {Tp}", thumbprint);
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
            _logger.LogInformation("CertificateThumbprint auto-detectado e salvo para tenant {TenantId}: {Tp}", tenantId, detected);
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
            _logger.LogInformation("CertificateThumbprint atualizado para tenant {TenantId}", tenantId);
            return true;
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
                _logger.LogWarning(ex, "Erro ao ler cert store {Location}", location);
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

            // Só importa se o settings.json for do mesmo tenant (ou não tiver TenantId)
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

    // Mantém settings.json (formato plano) sincronizado com o primeiro tenant configurado
    // para compatibilidade com os scripts PowerShell existentes.
    private void SyncSettingsJson(List<TenantEntry> tenants)
    {
        try
        {
            var primary = tenants.FirstOrDefault(t => t.IsConfigured) ?? tenants.FirstOrDefault();
            if (primary == null) return;

            var appCfg = _appConfig.Read();

            // Lê settings.json existente antes de montar o flat para preservar campos
            Dictionary<string, JsonElement> existingSettings = new();
            if (File.Exists(_settingsFile))
            {
                existingSettings = JsonSerializer.Deserialize<Dictionary<string, JsonElement>>(
                    File.ReadAllText(_settingsFile)) ?? new();
            }

            // Se o tenant não tem thumbprint: tenta settings.json existente, depois cert store
            var thumbprint = primary.CertificateThumbprint;
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
                    _logger.LogInformation("SyncSettingsJson: CertificateThumbprint detectado automaticamente: {Tp}", detected);
                }
            }

            var flat = new Dictionary<string, object?>
            {
                ["TenantId"] = primary.TenantId,
                ["ClientId"] = primary.ClientId ?? appCfg.ClientId,
                ["CertificateThumbprint"] = thumbprint,
                ["AppDisplayName"] = primary.Name,
                ["ClientSecret"] = appCfg.ClientSecret
            };

            // Preserva campos extras do settings.json existente (BackupRoot, LogRoot, etc.)
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
            _logger.LogWarning(ex, "Falha ao sincronizar settings.json");
        }
    }
}
