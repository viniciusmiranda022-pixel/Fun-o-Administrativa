using System.Diagnostics;
using System.Text;

namespace IR.AdminFunctions.Web.Services;

public class PowerShellRunner
{
    private readonly ILogger<PowerShellRunner> _logger;

    public PowerShellRunner(ILogger<PowerShellRunner> logger)
    {
        _logger = logger;
    }

    public async Task<PowerShellResult> RunAsync(
        string scriptPath,
        IDictionary<string, object?>? parameters = null,
        int timeoutSeconds = 900,
        CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(scriptPath))
            throw new ArgumentException("scriptPath obrigatório", nameof(scriptPath));

        if (!File.Exists(scriptPath))
            throw new FileNotFoundException($"Script PowerShell não encontrado: {scriptPath}", scriptPath);

        var command = BuildPsCommand(scriptPath, parameters);

        var psi = new ProcessStartInfo
        {
            FileName = "powershell.exe",
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true,
        };
        psi.ArgumentList.Add("-ExecutionPolicy");
        psi.ArgumentList.Add("Bypass");
        psi.ArgumentList.Add("-NonInteractive");
        psi.ArgumentList.Add("-Command");
        psi.ArgumentList.Add(command);

        using var process = new Process { StartInfo = psi };

        var sw = Stopwatch.StartNew();
        _logger.LogInformation("Executando script {Script} com timeout {Timeout}s", scriptPath, timeoutSeconds);

        process.Start();

        var stdoutTask = process.StandardOutput.ReadToEndAsync(cancellationToken);
        var stderrTask = process.StandardError.ReadToEndAsync(cancellationToken);

        using var timeoutCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        timeoutCts.CancelAfter(TimeSpan.FromSeconds(timeoutSeconds));

        try
        {
            await process.WaitForExitAsync(timeoutCts.Token);
        }
        catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested)
        {
            try { process.Kill(entireProcessTree: true); } catch { }
            _logger.LogError("Script {Script} excedeu timeout de {Timeout}s", scriptPath, timeoutSeconds);
            throw new TimeoutException($"Script {scriptPath} excedeu {timeoutSeconds}s");
        }

        var stdoutText = await stdoutTask;
        var stderrText = await stderrTask;
        sw.Stop();

        var stdout = SplitLines(stdoutText);
        var stderr = SplitLines(stderrText);
        var success = process.ExitCode == 0;

        _logger.LogInformation(
            "Script {Script} concluído em {Elapsed}ms. ExitCode={ExitCode}",
            scriptPath, sw.ElapsedMilliseconds, process.ExitCode);

        if (!success)
        {
            if (stderr.Count > 0)
            {
                foreach (var err in stderr)
                    _logger.LogError("PS stderr: {Error}", err);
            }
            else
            {
                _logger.LogWarning(
                    "Script falhou (ExitCode={ExitCode}) sem saída em stderr. Stdout ({Lines} linhas): {Preview}",
                    process.ExitCode, stdout.Count,
                    string.Join(" | ", stdout.TakeLast(5)));
            }
        }

        var errors = stderr.Count > 0
            ? stderr
            : (!success ? new List<string> { $"Script falhou com ExitCode={process.ExitCode}" } : new List<string>());

        return new PowerShellResult
        {
            Success = success,
            Output = stdout,
            Stdout = stdout,
            Errors = errors,
            Warnings = new List<string>(),
            ElapsedMilliseconds = sw.ElapsedMilliseconds
        };
    }

    private static string BuildPsCommand(string scriptPath, IDictionary<string, object?>? parameters)
    {
        var sb = new StringBuilder();
        sb.Append("& '");
        sb.Append(scriptPath.Replace("'", "''"));
        sb.Append('\'');

        if (parameters != null)
        {
            foreach (var (key, value) in parameters)
            {
                if (value is null) continue;
                if (value is bool b)
                    sb.Append($" -{key}:{(b ? "$true" : "$false")}");
                else if (value is string s)
                    sb.Append($" -{key} '{s.Replace("'", "''")}'");
                else
                    sb.Append($" -{key} {value}");
            }
        }

        sb.Append("; exit $LASTEXITCODE");
        return sb.ToString();
    }

    private static List<string> SplitLines(string text) =>
        text.Split('\n', StringSplitOptions.RemoveEmptyEntries)
            .Select(l => l.TrimEnd('\r'))
            .Where(l => !string.IsNullOrWhiteSpace(l))
            .ToList();
}

public class PowerShellResult
{
    public bool Success { get; set; }
    public List<string> Output { get; set; } = new();
    public List<string> Stdout { get; set; } = new();
    public List<string> Errors { get; set; } = new();
    public List<string> Warnings { get; set; } = new();
    public long ElapsedMilliseconds { get; set; }
}
