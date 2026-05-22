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
                _logger.LogError(ex, "Error reading {File}", _configFile);
                return new AppConfig();
            }
        }
    }

    public void Write(AppConfig config)
    {
        lock (_lock)
        {
            var dir = Path.GetDirectoryName(_configFile)!;
            Directory.CreateDirectory(dir);
            var tmp = _configFile + ".tmp";
            File.WriteAllText(tmp, JsonSerializer.Serialize(config, _json));
            File.Move(tmp, _configFile, overwrite: true);
            _logger.LogInformation("AppConfig saved to {File}", _configFile);
        }
    }

    public void Reset()
    {
        lock (_lock)
        {
            if (File.Exists(_configFile))
            {
                File.Delete(_configFile);
                _logger.LogInformation("AppConfig removed: {File}", _configFile);
            }
        }
    }
}
