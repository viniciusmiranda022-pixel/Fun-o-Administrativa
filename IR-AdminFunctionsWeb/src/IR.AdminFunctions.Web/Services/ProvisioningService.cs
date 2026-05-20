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
            _logger.LogError(ex, "Falha no auto-provisionamento de arquivos. A aplicação continuará.");
        }

        // Instalação do Microsoft.Graph em background para não bloquear o startup
        _ = Task.Run(() => EnsurePowerShellDependencies(cancellationToken), cancellationToken);

        return Task.CompletedTask;
    }

    public Task StopAsync(CancellationToken cancellationToken) => Task.CompletedTask;

    private void EnsurePowerShellDependencies(CancellationToken ct)
    {
        try
        {
            var iss = System.Management.Automation.Runspaces.InitialSessionState.CreateDefault();
            iss.ExecutionPolicy = Microsoft.PowerShell.ExecutionPolicy.Bypass;
            using var rs = System.Management.Automation.Runspaces.RunspaceFactory.CreateRunspace(iss);
            rs.Open();

            if (IsMicrosoftGraphAvailable(rs))
            {
                _logger.LogInformation("Microsoft.Graph já está disponível — nenhuma instalação necessária.");
                return;
            }

            _logger.LogInformation("Microsoft.Graph não encontrado. Instalando NuGet provider e Microsoft.Graph (Scope CurrentUser)...");

            RunScript(rs, @"
                [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
                Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser -ErrorAction Stop
            ", "Install-PackageProvider NuGet");

            if (ct.IsCancellationRequested) return;

            RunScript(rs, @"
                Install-Module -Name Microsoft.Graph -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
            ", "Install-Module Microsoft.Graph");

            _logger.LogInformation("Microsoft.Graph instalado com sucesso em CurrentUser.");
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Falha ao instalar Microsoft.Graph. Execute manualmente: " +
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
            throw new InvalidOperationException($"{label} falhou: {errors}");
        }

        _logger.LogInformation("{Label} concluído.", label);
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
            _logger.LogWarning("Origem do módulo {Module} não encontrada — pastas criadas, mas scripts não foram copiados.", moduleName);
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
                    _logger.LogInformation("Provisionado: {Source} → {Target}", file, target);
                }
            }
        }

        if (copiedAny)
        {
            _logger.LogInformation("Módulo {Module} provisionado a partir de {Source}", moduleName, sourceModule);
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
                    Name = "PREENCHER-NomeDoTenant",
                    TenantId = "PREENCHER-TenantId",
                    ClientId = "PREENCHER-ClientId",
                    Thumbprint = "PREENCHER-Thumbprint"
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
        _logger.LogWarning("settings.json não existia — template criado em {Path}. Edite os campos PREENCHER-* antes de usar.", path);
    }
}
