using System.Text.Json;
using Microsoft.Extensions.Options;

namespace IR.AdminFunctions.Web.Services;

public class ProvisioningService : IHostedService
{
    private readonly QuestOptions _options;
    private readonly ILogger<ProvisioningService> _logger;
    private readonly IHostEnvironment _env;

    public ProvisioningService(IOptions<QuestOptions> options, ILogger<ProvisioningService> logger, IHostEnvironment env)
    {
        _options = options.Value;
        _logger = logger;
        _env = env;
    }

    public Task StartAsync(CancellationToken cancellationToken)
    {
        try
        {
            EnsureModule(_options.BackupModule, new[] { "Scripts", "Config", "Logs", "Backups" });
            EnsureModule(_options.CompareModule, new[] { "Scripts", "Xaml", "Logs", "Output" });
            EnsureModule(_options.RestoreModule, new[] { "Scripts", "Logs" });
            EnsureSettingsTemplate();
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failure during auto-provisioning of files. The application will continue.");
        }

        // Installing Microsoft.Graph in background to avoid blocking startup
        _ = Task.Run(() => EnsurePowerShellDependencies(cancellationToken), cancellationToken);

        return Task.CompletedTask;
    }

    public Task StopAsync(CancellationToken cancellationToken) => Task.CompletedTask;

    private void EnsurePowerShellDependencies(CancellationToken ct)
    {
        try
        {
            using var rs = System.Management.Automation.Runspaces.RunspaceFactory.CreateRunspace();
            rs.Open();

            if (IsMicrosoftGraphAvailable(rs))
            {
                _logger.LogInformation("Microsoft.Graph is already available — no installation needed.");
                return;
            }

            _logger.LogInformation("Microsoft.Graph not found. Installing NuGet provider and Microsoft.Graph (Scope CurrentUser)...");

            RunScript(rs, @"
                [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
                Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser -ErrorAction Stop
            ", "Install-PackageProvider NuGet");

            if (ct.IsCancellationRequested) return;

            RunScript(rs, @"
                Install-Module -Name Microsoft.Graph -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
            ", "Install-Module Microsoft.Graph");

            _logger.LogInformation("Microsoft.Graph successfully installed for CurrentUser.");
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed to install Microsoft.Graph. Run manually: " +
                "Install-Module Microsoft.Graph -Scope CurrentUser -Force");
        }
    }

    private static bool IsMicrosoftGraphAvailable(System.Management.Automation.Runspaces.Runspace rs)
    {
        using var ps = System.Management.Automation.PowerShell.Create();
        ps.Runspace = rs;
        ps.AddScript("(Get-Module -ListAvailable -Name Microsoft.Graph) -ne $null");
        var result = ps.Invoke();
        return result.Count > 0 && result[0]?.BaseObject is true;
    }

    private void RunScript(System.Management.Automation.Runspaces.Runspace rs, string script, string label)
    {
        using var ps = System.Management.Automation.PowerShell.Create();
        ps.Runspace = rs;
        ps.AddScript(script);
        ps.Invoke();

        if (ps.HadErrors)
        {
            var errors = string.Join("; ", ps.Streams.Error.Select(e => e.ToString()));
            throw new InvalidOperationException($"{label} failed: {errors}");
        }

        _logger.LogInformation("{Label} completed.", label);
    }

    private void EnsureModule(string moduleName, string[] subFolders)
    {
        var moduleRoot = Path.Combine(_options.ProgramDataRoot, moduleName);
        Directory.CreateDirectory(moduleRoot);

        foreach (var sub in subFolders)
        {
            Directory.CreateDirectory(Path.Combine(moduleRoot, sub));
        }

        var sourceModule = FindSourceModule(moduleName);
        if (sourceModule == null)
        {
            _logger.LogWarning("Source for module {Module} not found — folders created, but scripts were not copied.", moduleName);
            return;
        }

        var copiedAny = false;
        foreach (var subDir in new[] { "Scripts", "Xaml" })
        {
            var src = Path.Combine(sourceModule, subDir);
            if (!Directory.Exists(src)) continue;

            var dst = Path.Combine(moduleRoot, subDir);
            Directory.CreateDirectory(dst);

            foreach (var file in Directory.EnumerateFiles(src))
            {
                var target = Path.Combine(dst, Path.GetFileName(file));
                if (!File.Exists(target))
                {
                    File.Copy(file, target);
                    copiedAny = true;
                    _logger.LogInformation("Provisioned: {Source} → {Target}", file, target);
                }
            }
        }

        if (copiedAny)
        {
            _logger.LogInformation("Module {Module} provisioned from {Source}", moduleName, sourceModule);
        }
    }

    private string? FindSourceModule(string moduleName)
    {
        var candidates = new[]
        {
            Path.Combine(AppContext.BaseDirectory, "PowerShellModules", moduleName),
            Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "..", moduleName),
            Path.Combine(_env.ContentRootPath, "..", "..", "..", moduleName),
            Path.Combine(_env.ContentRootPath, "..", "..", moduleName)
        };

        foreach (var candidate in candidates)
        {
            var full = Path.GetFullPath(candidate);
            if (Directory.Exists(Path.Combine(full, "Scripts")))
            {
                return full;
            }
        }

        return null;
    }

    private void EnsureSettingsTemplate()
    {
        var path = _options.SettingsFile;
        if (File.Exists(path)) return;

        Directory.CreateDirectory(Path.GetDirectoryName(path)!);

        var template = new
        {
            Tenants = new[]
            {
                new
                {
                    Name = "FILL-IN-TenantName",
                    TenantId = "FILL-IN-TenantId",
                    ClientId = "FILL-IN-ClientId",
                    Thumbprint = "FILL-IN-Thumbprint"
                }
            },
            Defaults = new
            {
                RetentionDays = 30,
                ParallelDegree = 4
            }
        };

        var json = JsonSerializer.Serialize(template, new JsonSerializerOptions { WriteIndented = true });
        File.WriteAllText(path, json);
        _logger.LogWarning("settings.json did not exist — template created at {Path}. Edit the FILL-IN-* fields before use.", path);
    }
}
