using System.Text.Json;
using IR.AdminFunctions.Web.Models;
using Microsoft.Extensions.Options;

namespace IR.AdminFunctions.Web.Services;

public class AppConfigStore
{
    private readonly string _configFile;
    private readonly ILogger<AppConfigStore> _logger;
    private readonly object _lock = new();
    private readonly JsonSerializerOptions _json = new() { WriteIndented = true, PropertyNamingPolicy = JsonNamingPolicy.CamelCase };

    public AppConfigStore(IOptions<QuestOptions> options, ILogger<AppConfigStore> logger)
    {
        var dir = Path.Combine(options.Value.WebLogRoot, "..");
        _configFile = Path.GetFullPath(Path.Combine(dir, "app-config.json"));
        _logger = logger;
    }

    public AppConfig Read()
    {
        lock (_lock)
        {
            if (!File.Exists(_configFile)) return new AppConfig();
            try
            {
                var json = File.ReadAllText(_configFile);
                return JsonSerializer.Deserialize<AppConfig>(json, _json) ?? new AppConfig();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Erro lendo {File}", _configFile);
                return new AppConfig();
            }
        }
    }

    public void Write(AppConfig config)
    {
        lock (_lock)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(_configFile)!);
            File.WriteAllText(_configFile, JsonSerializer.Serialize(config, _json));
            _logger.LogInformation("AppConfig salvo em {File}", _configFile);
        }
    }

    public void Reset()
    {
        lock (_lock)
        {
            if (File.Exists(_configFile))
            {
                File.Delete(_configFile);
                _logger.LogInformation("AppConfig removido: {File}", _configFile);
            }
        }
    }
}
