using System.Text.Json;
using IR.AdminFunctions.Web.Models;
using Microsoft.Extensions.Options;

namespace IR.AdminFunctions.Web.Services;

public class BackupReader
{
    private readonly QuestOptions _options;
    private readonly SettingsReader _settings;
    private readonly ILogger<BackupReader> _logger;

    private static readonly string[] RequiredFiles =
    {
        "manifest.json",
        "roleDefinitions.json",
        "roleAssignments.json"
    };

    public BackupReader(IOptions<QuestOptions> options, SettingsReader settings, ILogger<BackupReader> logger)
    {
        _options = options.Value;
        _settings = settings;
        _logger = logger;
    }

    public string ResolveBackupRoot()
    {
        var raw = _settings.ReadRaw();
        if (!string.IsNullOrWhiteSpace(raw?.BackupRoot) && Directory.Exists(raw.BackupRoot))
        {
            return raw.BackupRoot!;
        }
        return _options.BackupsFolder;
    }

    public IEnumerable<BackupSnapshot> List()
    {
        var root = ResolveBackupRoot();
        if (!Directory.Exists(root))
        {
            _logger.LogWarning("Backups folder does not exist: {Root}", root);
            yield break;
        }

        var dirs = new DirectoryInfo(root)
            .GetDirectories()
            .OrderByDescending(d => d.LastWriteTimeUtc);

        foreach (var dir in dirs)
        {
            if (!IsValid(dir.FullName)) continue;
            yield return Load(dir);
        }
    }

    public BackupSnapshot? Get(string id)
    {
        var root = ResolveBackupRoot();
        var path = Path.Combine(root, id);
        if (!Directory.Exists(path) || !IsValid(path)) return null;
        return Load(new DirectoryInfo(path));
    }

    public UnpackedObjects? GetUnpacked(string id)
    {
        var snap = Get(id);
        if (snap == null) return null;

        var defsPath = Path.Combine(snap.Path, "roleDefinitions.json");
        var assignmentsPath = Path.Combine(snap.Path, "roleAssignments.json");

        try
        {
            var defs = JsonDocument.Parse(File.ReadAllText(defsPath));
            var assigns = JsonDocument.Parse(File.ReadAllText(assignmentsPath));
            return new UnpackedObjects
            {
                BackupId = id,
                RoleDefinitions = defs.RootElement.Clone(),
                RoleAssignments = assigns.RootElement.Clone()
            };
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Failed reading objects from backup {Id}", id);
            return null;
        }
    }

    private static bool IsValid(string path) =>
        RequiredFiles.All(f => File.Exists(Path.Combine(path, f)));

    private BackupSnapshot Load(DirectoryInfo dir)
    {
        var snap = new BackupSnapshot
        {
            Id = dir.Name,
            Path = dir.FullName,
            CreatedAt = dir.CreationTimeUtc,
            ModifiedAt = dir.LastWriteTimeUtc
        };

        var manifestPath = Path.Combine(dir.FullName, "manifest.json");
        if (File.Exists(manifestPath))
        {
            try
            {
                using var doc = JsonDocument.Parse(File.ReadAllText(manifestPath));
                var root = doc.RootElement;
                snap.CollectedAt = TryString(root, "CollectedAt");
                snap.TenantId = TryString(root, "TenantId");
                snap.RoleDefinitionsCount = TryInt(root, "RoleDefinitionsCount");
                snap.RoleAssignmentsCount = TryInt(root, "RoleAssignmentsCount");
                snap.RoleDefinitionsHash = TryString(root, "RoleDefinitionsHash");
                snap.RoleAssignmentsHash = TryString(root, "RoleAssignmentsHash");
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "Invalid manifest at {Path}", manifestPath);
                snap.ManifestError = ex.Message;
            }
        }

        try
        {
            long size = 0;
            foreach (var f in dir.GetFiles("*", SearchOption.AllDirectories))
            {
                size += f.Length;
            }
            snap.SizeBytes = size;
        }
        catch { }

        return snap;
    }

    private static string? TryString(JsonElement root, string name) =>
        root.TryGetProperty(name, out var v) && v.ValueKind == JsonValueKind.String ? v.GetString() : null;

    private static int? TryInt(JsonElement root, string name) =>
        root.TryGetProperty(name, out var v) && v.ValueKind == JsonValueKind.Number ? v.GetInt32() : null;
}
