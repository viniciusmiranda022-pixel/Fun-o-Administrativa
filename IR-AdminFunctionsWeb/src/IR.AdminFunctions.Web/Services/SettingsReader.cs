using System.Text.Json;
using IR.AdminFunctions.Web.Models;
using Microsoft.Extensions.Options;

namespace IR.AdminFunctions.Web.Services;

public class SettingsReader
{
    private readonly QuestOptions _options;
    private readonly ILogger<SettingsReader> _logger;

    public SettingsReader(IOptions<QuestOptions> options, ILogger<SettingsReader> logger)
    {
        _options = options.Value;
        _logger = logger;
    }

    public BackupSettings? Read()
    {
        var path = _options.SettingsFile;
        if (!File.Exists(path))
        {
            _logger.LogWarning("settings.json não encontrado em {Path}", path);
            return null;
        }

        try
        {
            var json = File.ReadAllText(path);
            var doc = JsonDocument.Parse(json);
            var root = doc.RootElement;

            return new BackupSettings
            {
                TenantId = TryString(root, "TenantId"),
                ClientId = TryString(root, "ClientId"),
                AppDisplayName = TryString(root, "AppDisplayName"),
                CertificateThumbprint = MaskThumbprint(TryString(root, "CertificateThumbprint")),
                CertificateStore = TryString(root, "CertificateStore"),
                BackupRoot = TryString(root, "BackupRoot"),
                LogRoot = TryString(root, "LogRoot"),
                LogoPath = TryString(root, "LogoPath"),
                RetentionDays = TryInt(root, "RetentionDays"),
                ClientSecret = MaskThumbprint(TryString(root, "ClientSecret"))
            };
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Falha lendo settings.json em {Path}", path);
            return null;
        }
    }

    public BackupSettings? ReadRaw()
    {
        var path = _options.SettingsFile;
        if (!File.Exists(path))
        {
            _logger.LogWarning("settings.json não encontrado em {Path}", path);
            return null;
        }

        try
        {
            var json = File.ReadAllText(path);
            var result = JsonSerializer.Deserialize<BackupSettings>(json,
                new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
            _logger.LogInformation(
                "ReadRaw: TenantId='{TenantId}' ClientId='{ClientId}' HasThumbprint={HasTp} HasSecret={HasSecret}",
                result?.TenantId ?? "(null)",
                result?.ClientId ?? "(null)",
                !string.IsNullOrWhiteSpace(result?.CertificateThumbprint),
                !string.IsNullOrWhiteSpace(result?.ClientSecret));
            return result;
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Falha lendo settings.json em {Path}", path);
            return null;
        }
    }

    private static string? TryString(JsonElement root, string name) =>
        root.TryGetProperty(name, out var v) && v.ValueKind == JsonValueKind.String ? v.GetString() : null;

    private static int? TryInt(JsonElement root, string name) =>
        root.TryGetProperty(name, out var v) && v.ValueKind == JsonValueKind.Number ? v.GetInt32() : null;

    private static string? MaskThumbprint(string? thumb)
    {
        if (string.IsNullOrEmpty(thumb) || thumb.Length < 8) return thumb;
        return string.Concat(thumb.AsSpan(0, 4), "...", thumb.AsSpan(thumb.Length - 4, 4));
    }
}
