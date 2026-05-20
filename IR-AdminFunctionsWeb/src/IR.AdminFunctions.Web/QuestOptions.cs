namespace IR.AdminFunctions.Web;

public class QuestOptions
{
    public string ProgramDataRoot { get; set; } = @"C:\ProgramData\Quest";
    public string BackupModule { get; set; } = "IR-AdministrativeFunctionBackup";
    public string CompareModule { get; set; } = "IR-AdministrativeFunctionCompare";
    public string RestoreModule { get; set; } = "IR-AdministrativeFunctionRestore";
    public string WebLogRoot { get; set; } = @"C:\ProgramData\Quest\IR-AdministrativeFunctionWeb\Logs";
    public int DefaultJobTimeoutSeconds { get; set; } = 900;

    public string BackupRoot => Path.Combine(ProgramDataRoot, BackupModule);
    public string CompareRoot => Path.Combine(ProgramDataRoot, CompareModule);
    public string RestoreRoot => Path.Combine(ProgramDataRoot, RestoreModule);

    public string BackupScript => Path.Combine(BackupRoot, "Scripts", "Run-AdministrativeFunctionsBackup.ps1");
    public string CompareScript => Path.Combine(CompareRoot, "Scripts", "Start-AdministrativeFunctionCompare.ps1");
    public string RestoreScript => Path.Combine(RestoreRoot, "Scripts", "Restore-CustomRole-Full.ps1");

    public string BackupsFolder => Path.Combine(BackupRoot, "Backups");
    public string SettingsFile => Path.Combine(BackupRoot, "Config", "settings.json");
    public string BackupLogFolder => Path.Combine(BackupRoot, "Logs");
    public string CompareOutputFolder => Path.Combine(CompareRoot, "Output");
    public string CompareLogFolder => Path.Combine(CompareRoot, "Logs");
}
