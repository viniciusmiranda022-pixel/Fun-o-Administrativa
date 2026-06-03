using IR.AdminFunctions.Web;
using IR.AdminFunctions.Web.Services;
using Microsoft.Extensions.Logging.Abstractions;
using Microsoft.Extensions.Options;
using Moq;

namespace IR.AdminFunctions.Tests.Services;

public class BackupReaderPathTraversalTests
{
    /// <summary>
    /// Creates a BackupReader whose ResolveBackupRoot() returns <paramref name="backupsFolder"/>.
    /// We achieve this by mocking SettingsReader.ReadRaw() to return null so the reader
    /// falls back to QuestOptions.BackupsFolder.  Because BackupsFolder is a computed property
    /// (ProgramDataRoot + BackupModule + "Backups") we set ProgramDataRoot and BackupModule
    /// so the concatenation yields the desired path.
    /// </summary>
    private static BackupReader CreateReader(string backupsFolder)
    {
        // backupsFolder = <parent>/<module>/Backups  →  ProgramDataRoot=<parent>, BackupModule=<module>
        // When no clean split is available, derive parent and module name from the path.
        var parent = Path.GetDirectoryName(backupsFolder) ?? backupsFolder;
        var module = Path.GetFileName(parent);
        var programDataRoot = Path.GetDirectoryName(parent) ?? parent;

        var opts = Options.Create(new QuestOptions
        {
            ProgramDataRoot = programDataRoot,
            BackupModule = module
        });

        var settings = new Mock<SettingsReader>(opts, NullLogger<SettingsReader>.Instance);
        // SettingsReader.ReadRaw() returns null → falls back to options.BackupsFolder
        settings.Setup(s => s.ReadRaw()).Returns((IR.AdminFunctions.Web.Models.BackupSettings?)null);
        return new BackupReader(opts, settings.Object, NullLogger<BackupReader>.Instance);
    }

    /// <summary>
    /// Builds a BackupReader rooted at a freshly-created temp folder so that
    /// <c>QuestOptions.BackupsFolder</c> resolves to a real directory we control.
    /// </summary>
    private static (BackupReader reader, string root) CreateReaderWithTempRoot()
    {
        // Structure: <tmp>/<guid>/IR-AdministrativeFunctionBackup/Backups
        var tmpBase = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
        var module = "IR-AdministrativeFunctionBackup";
        var backupsFolder = Path.Combine(tmpBase, module, "Backups");
        Directory.CreateDirectory(backupsFolder);

        var opts = Options.Create(new QuestOptions
        {
            ProgramDataRoot = tmpBase,
            BackupModule = module
        });

        var settings = new Mock<SettingsReader>(opts, NullLogger<SettingsReader>.Instance);
        settings.Setup(s => s.ReadRaw()).Returns((IR.AdminFunctions.Web.Models.BackupSettings?)null);

        return (new BackupReader(opts, settings.Object, NullLogger<BackupReader>.Instance), backupsFolder);
    }

    [Theory]
    [InlineData("../../../etc/passwd")]
    [InlineData("..\\..\\Windows\\System32")]
    [InlineData("%2e%2e%2fetc")]
    public void Get_RejectsPathTraversal(string maliciousId)
    {
        var (reader, _) = CreateReaderWithTempRoot();
        var result = reader.Get(maliciousId);
        Assert.Null(result);
    }

    [Fact]
    public void Get_ReturnsNullForMissingBackup()
    {
        var (reader, _) = CreateReaderWithTempRoot();
        var result = reader.Get("nonexistent-backup-id-xyz");
        Assert.Null(result);
    }
}
