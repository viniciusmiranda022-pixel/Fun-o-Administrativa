using System.Text.RegularExpressions;
using IR.AdminFunctions.Web.Models;
using Microsoft.Extensions.Options;

namespace IR.AdminFunctions.Web.Services;

public class LogReader
{
    private static readonly Regex LineRegex =
        new(@"^\[(?<ts>\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})\]\s*(?<msg>.*)$", RegexOptions.Compiled);

    private readonly QuestOptions _options;
    private readonly ILogger<LogReader> _logger;

    public LogReader(IOptions<QuestOptions> options, ILogger<LogReader> logger)
    {
        _options = options.Value;
        _logger = logger;
    }

    public IEnumerable<LogEvent> ReadEvents(int max = 500, string? severity = null, DateTime? since = null)
    {
        var roots = new[]
        {
            _options.BackupLogFolder,
            _options.CompareLogFolder
        };

        var events = new List<LogEvent>();

        foreach (var root in roots)
        {
            if (!Directory.Exists(root)) continue;

            var files = new DirectoryInfo(root)
                .GetFiles("*.log", SearchOption.TopDirectoryOnly)
                .OrderByDescending(f => f.LastWriteTimeUtc)
                .Take(20);

            foreach (var file in files)
            {
                try
                {
                    foreach (var line in File.ReadLines(file.FullName))
                    {
                        var ev = Parse(line, file.Name);
                        if (ev == null) continue;
                        if (severity != null && !string.Equals(ev.Severity, severity, StringComparison.OrdinalIgnoreCase)) continue;
                        if (since.HasValue && ev.Time < since.Value) continue;
                        events.Add(ev);
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogWarning(ex, "Failed reading log {Path}", file.FullName);
                }
            }
        }

        return events
            .OrderByDescending(e => e.Time)
            .Take(max);
    }

    public IEnumerable<TaskEntry> ReadTasks(int max = 200, DateTime? since = null)
    {
        var events = ReadEvents(max * 2, since: since);

        return events
            .Select(e => new TaskEntry
            {
                Time = e.Time,
                Name = e.Source,
                Status = InferStatus(e.Message),
                Type = InferType(e.Source),
                Operation = e.Message
            })
            .Take(max);
    }

    private static LogEvent? Parse(string line, string source)
    {
        var match = LineRegex.Match(line);
        if (!match.Success)
        {
            return null;
        }

        if (!DateTime.TryParse(match.Groups["ts"].Value, out var ts))
        {
            return null;
        }

        var msg = match.Groups["msg"].Value;
        return new LogEvent
        {
            Time = ts,
            Source = source,
            Message = msg,
            Severity = InferSeverity(msg)
        };
    }

    private static string InferSeverity(string msg)
    {
        if (msg.StartsWith("ERRO", StringComparison.OrdinalIgnoreCase) ||
            msg.Contains("error", StringComparison.OrdinalIgnoreCase) ||
            msg.StartsWith("ERROR", StringComparison.OrdinalIgnoreCase))
        {
            return "Error";
        }
        if (msg.StartsWith("WARNING", StringComparison.OrdinalIgnoreCase) ||
            msg.Contains("warn", StringComparison.OrdinalIgnoreCase))
        {
            return "Warning";
        }
        return "Information";
    }

    private static string InferStatus(string msg)
    {
        if (msg.Contains("success", StringComparison.OrdinalIgnoreCase) ||
            msg.Contains("complet", StringComparison.OrdinalIgnoreCase))
        {
            return "Completed";
        }
        if (msg.Contains("fail", StringComparison.OrdinalIgnoreCase) ||
            msg.Contains("error", StringComparison.OrdinalIgnoreCase))
        {
            return "Failed";
        }
        if (msg.Contains("inici", StringComparison.OrdinalIgnoreCase) ||
            msg.Contains("start", StringComparison.OrdinalIgnoreCase))
        {
            return "Running";
        }
        return "Information";
    }

    private static string InferType(string source)
    {
        if (source.Contains("Export", StringComparison.OrdinalIgnoreCase) ||
            source.Contains("Backup", StringComparison.OrdinalIgnoreCase) ||
            source.Contains("rmad", StringComparison.OrdinalIgnoreCase))
        {
            return "Backup";
        }
        if (source.Contains("Compare", StringComparison.OrdinalIgnoreCase))
        {
            return "Compare";
        }
        if (source.Contains("Restore", StringComparison.OrdinalIgnoreCase))
        {
            return "Restore";
        }
        return "Other";
    }
}
