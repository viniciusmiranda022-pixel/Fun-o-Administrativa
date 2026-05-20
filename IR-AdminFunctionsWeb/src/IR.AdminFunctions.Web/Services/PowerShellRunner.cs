using System.Diagnostics;
using System.Management.Automation;
using System.Management.Automation.Runspaces;

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
        {
            throw new ArgumentException("scriptPath obrigatório", nameof(scriptPath));
        }

        if (!File.Exists(scriptPath))
        {
            throw new FileNotFoundException($"Script PowerShell não encontrado: {scriptPath}", scriptPath);
        }

        var initial = InitialSessionState.CreateDefault();
        initial.ExecutionPolicy = Microsoft.PowerShell.ExecutionPolicy.Bypass;

        using var runspace = RunspaceFactory.CreateRunspace(initial);
        runspace.Open();

        using var ps = PowerShell.Create();
        ps.Runspace = runspace;
        ps.AddCommand(scriptPath);

        if (parameters != null)
        {
            foreach (var (key, value) in parameters)
            {
                if (value is null)
                {
                    continue;
                }
                ps.AddParameter(key, value);
            }
        }

        var stdout = new List<string>();
        var errors = new List<string>();
        var warnings = new List<string>();

        ps.Streams.Information.DataAdded += (s, e) =>
        {
            var record = ((PSDataCollection<InformationRecord>)s!)[e.Index];
            if (record?.MessageData != null)
            {
                stdout.Add(record.MessageData.ToString() ?? string.Empty);
            }
        };
        ps.Streams.Error.DataAdded += (s, e) =>
        {
            var record = ((PSDataCollection<ErrorRecord>)s!)[e.Index];
            errors.Add(record?.ToString() ?? string.Empty);
        };
        ps.Streams.Warning.DataAdded += (s, e) =>
        {
            var record = ((PSDataCollection<WarningRecord>)s!)[e.Index];
            warnings.Add(record?.Message ?? string.Empty);
        };

        var sw = Stopwatch.StartNew();
        _logger.LogInformation("Executando script {Script} com timeout {Timeout}s", scriptPath, timeoutSeconds);

        PSDataCollection<PSObject>? results = null;
        var invokeTask = Task.Run(() => results = ps.Invoke(), cancellationToken);

        using var timeoutCts = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        var delayTask = Task.Delay(TimeSpan.FromSeconds(timeoutSeconds), timeoutCts.Token);

        var completed = await Task.WhenAny(invokeTask, delayTask).ConfigureAwait(false);

        if (completed != invokeTask)
        {
            ps.Stop();
            _logger.LogError("Script {Script} excedeu timeout de {Timeout}s", scriptPath, timeoutSeconds);
            throw new TimeoutException($"Script {scriptPath} excedeu {timeoutSeconds}s");
        }

        timeoutCts.Cancel();
        await invokeTask.ConfigureAwait(false);
        sw.Stop();

        _logger.LogInformation(
            "Script {Script} concluído em {Elapsed}ms. HadErrors={HadErrors}",
            scriptPath, sw.ElapsedMilliseconds, ps.HadErrors);

        return new PowerShellResult
        {
            Success = !ps.HadErrors,
            Output = results?.Select(r => r?.BaseObject?.ToString() ?? string.Empty).ToList() ?? new List<string>(),
            Stdout = stdout,
            Errors = errors,
            Warnings = warnings,
            ElapsedMilliseconds = sw.ElapsedMilliseconds
        };
    }
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
