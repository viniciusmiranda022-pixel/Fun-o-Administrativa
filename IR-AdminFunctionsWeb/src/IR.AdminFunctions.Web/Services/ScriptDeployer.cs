using IR.AdminFunctions.Web.Models;
using Microsoft.Extensions.Options;

namespace IR.AdminFunctions.Web.Services;

public class ScriptDeployer : IHostedService
{
    private readonly QuestOptions _options;
    private readonly ILogger<ScriptDeployer> _logger;

    public ScriptDeployer(IOptions<QuestOptions> options, ILogger<ScriptDeployer> logger)
    {
        _options = options.Value;
        _logger = logger;
    }

    public Task StartAsync(CancellationToken cancellationToken)
    {
        var baseScriptsDir = Path.Combine(AppContext.BaseDirectory, "Scripts");
        if (!Directory.Exists(baseScriptsDir))
        {
            _logger.LogWarning("Pasta de scripts embutidos não encontrada: {Dir} — scripts não serão atualizados automaticamente.", baseScriptsDir);
            return Task.CompletedTask;
        }

        var deployed = 0;

        deployed += DeployFolder(
            Path.Combine(baseScriptsDir, "Backup"),
            Path.Combine(_options.BackupRoot, "Scripts"));

        deployed += DeployFolder(
            Path.Combine(baseScriptsDir, "Compare"),
            Path.Combine(_options.CompareRoot, "Scripts"));

        _logger.LogInformation("ScriptDeployer: {Count} script(s) implantado(s)", deployed);
        return Task.CompletedTask;
    }

    private int DeployFolder(string sourceDir, string destDir)
    {
        if (!Directory.Exists(sourceDir)) return 0;
        Directory.CreateDirectory(destDir);
        var count = 0;
        foreach (var src in Directory.GetFiles(sourceDir))
        {
            File.Copy(src, Path.Combine(destDir, Path.GetFileName(src)), overwrite: true);
            count++;
        }
        return count;
    }

    public Task StopAsync(CancellationToken cancellationToken) => Task.CompletedTask;
}
