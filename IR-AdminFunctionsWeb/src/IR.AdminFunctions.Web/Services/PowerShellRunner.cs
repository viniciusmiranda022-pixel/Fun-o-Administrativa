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
            throw new ArgumentException("scriptPath is required", nameof(scriptPath));

        if (!File.Exists(scriptPath))
            throw new FileNotFoundException($"PowerShell script not found: {scriptPath}", scriptPath);

        var command = BuildPsCommand(scriptPath, parameters);
        _logger.LogInformation("PS command: {Command}", command);

        var psi = new ProcessStartInfo
        {
            FileName = "powershell.exe",
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true,
            StandardOutputEncoding = Encoding.UTF8,
            StandardErrorEncoding = Encoding.UTF8,
        };
        psi.ArgumentList.Add("-ExecutionPolicy");
        psi.ArgumentList.Add("Bypass");
        psi.ArgumentList.Add("-NonInteractive");
        psi.ArgumentList.Add("-Command");
        psi.ArgumentList.Add(command);

        using var process = new Process { StartInfo = psi };

        var sw = Stopwatch.StartNew();
        _logger.LogInformation("Running script {Script} with timeout {Timeout}s", scriptPath, timeoutSeconds);

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
            _logger.LogError("Script {Script} exceeded timeout of {Timeout}s", scriptPath, timeoutSeconds);
            throw new TimeoutException($"Script {scriptPath} exceeded {timeoutSeconds}s timeout");
        }

        var stdoutText = await stdoutTask;
        var stderrText = await stderrTask;
        sw.Stop();

        var stdout = SplitLines(stdoutText);
        var stderr = SplitLines(stderrText);
        var success = process.ExitCode == 0;

        _logger.LogInformation(
            "Script {Script} completed in {Elapsed}ms. ExitCode={ExitCode}",
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
                    "Script failed (ExitCode={ExitCode}) with no stderr output. Stdout ({Lines} lines): {Preview}",
                    process.ExitCode, stdout.Count,
                    string.Join(" | ", stdout.TakeLast(5)));
            }
        }

        if (stdout.Count > 0)
        {
            var preview = stdout.TakeLast(30).ToList();
            foreach (var line in preview)
                _logger.LogInformation("PS stdout: {Line}", line);
        }

        var errors = stderr.Count > 0
            ? stderr
            : (!success ? new List<string> { $"Script failed with ExitCode={process.ExitCode}" } : new List<string>());

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
                if (value is string s0 && string.IsNullOrEmpty(s0)) continue;
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
