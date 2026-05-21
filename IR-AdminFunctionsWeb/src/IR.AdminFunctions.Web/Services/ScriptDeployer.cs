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
        var sourceDir = Path.Combine(AppContext.BaseDirectory, "Scripts");
        if (!Directory.Exists(sourceDir))
        {
            _logger.LogWarning("Pasta de scripts embutidos não encontrada: {Dir} — scripts não serão atualizados automaticamente.", sourceDir);
            return Task.CompletedTask;
        }

        var destDir = Path.Combine(_options.BackupRoot, "Scripts");
        Directory.CreateDirectory(destDir);

        var deployed = 0;
        foreach (var src in Directory.GetFiles(sourceDir))
        {
            var dest = Path.Combine(destDir, Path.GetFileName(src));
            File.Copy(src, dest, overwrite: true);
            deployed++;
        }

        _logger.LogInformation("ScriptDeployer: {Count} script(s) implantado(s) em {Dest}", deployed, destDir);
        return Task.CompletedTask;
    }

    public Task StopAsync(CancellationToken cancellationToken) => Task.CompletedTask;
}
