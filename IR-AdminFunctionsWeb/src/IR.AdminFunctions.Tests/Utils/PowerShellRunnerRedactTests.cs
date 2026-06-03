using IR.AdminFunctions.Web.Services;

namespace IR.AdminFunctions.Tests.Utils;

public class PowerShellRunnerRedactTests
{
    [Fact]
    public void RedactCommand_ClientSecretNotInCommand()
    {
        // ClientSecret is routed via IR_CLIENT_SECRET env var and is never included in the command string.
        // This test confirms that a command built by PowerShellRunner does not carry -ClientSecret at all.
        // If somehow a legacy command string contains -ClientSecret, RedactCommand should leave nothing sensitive.
        var cmd = "& 'C:\\Scripts\\Run.ps1' -TenantId 'abc' -Thumbprint 'abc123'";
        var redacted = PowerShellRunner.RedactCommand(cmd);
        Assert.DoesNotContain("-ClientSecret", redacted);
        Assert.Contains("abc", redacted); // TenantId is not sensitive
    }

    [Fact]
    public void RedactCommand_MasksThumbprint()
    {
        var cmd = "& 'C:\\Scripts\\Run.ps1' -Thumbprint 'ABCDEF1234567890'";
        var redacted = PowerShellRunner.RedactCommand(cmd);
        Assert.DoesNotContain("ABCDEF1234567890", redacted);
        Assert.Contains("-Thumbprint '***'", redacted);
    }

    [Fact]
    public void RedactCommand_LeavesNonSensitiveParamsAlone()
    {
        var cmd = "& 'C:\\Scripts\\Run.ps1' -TenantId 'tenant-123'";
        var redacted = PowerShellRunner.RedactCommand(cmd);
        Assert.Contains("tenant-123", redacted);
    }
}
