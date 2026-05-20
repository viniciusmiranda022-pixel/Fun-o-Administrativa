using System.Management.Automation;
using System.Management.Automation.Runspaces;
using IR.AdminFunctions.Web.Models;

namespace IR.AdminFunctions.Web.Services;

// Verifica via Microsoft Graph se o Service Principal da nossa aplicação
// já tem as permissões Basic (RoleManagement.Read.Directory) e Restore
// (RoleManagement.ReadWrite.Directory) concedidas no tenant alvo.
//
// Usa client credentials flow com o ClientId/Secret do AppConfig.
public class ConsentChecker
{
    private readonly AppConfigStore _appStore;
    private readonly ILogger<ConsentChecker> _logger;

    private const string PermissionBasicId = "483bed4a-2ad3-4361-a73b-c83ccdbdc53c";       // RoleManagement.Read.Directory
    private const string PermissionRestoreId = "9e3f62cf-ca93-4989-b6ce-bf83c28f9fe8";    // RoleManagement.ReadWrite.Directory

    public ConsentChecker(AppConfigStore appStore, ILogger<ConsentChecker> logger)
    {
        _appStore = appStore;
        _logger = logger;
    }

    public async Task<TenantConsentsState> CheckAsync(string tenantId, CancellationToken ct)
    {
        var cfg = _appStore.Read();
        if (!cfg.IsConfigured)
        {
            return new TenantConsentsState();
        }

        try
        {
            using var rs = RunspaceFactory.CreateRunspace();
            rs.Open();
            using var ps = PowerShell.Create();
            ps.Runspace = rs;

            ps.AddScript($@"
                [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
                Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
                Import-Module Microsoft.Graph.Applications -ErrorAction Stop

                $secret = ConvertTo-SecureString '{cfg.ClientSecret}' -AsPlainText -Force
                $cred = New-Object System.Management.Automation.PSCredential('{cfg.ClientId}', $secret)
                Connect-MgGraph -TenantId '{tenantId}' -ClientSecretCredential $cred -NoWelcome -ErrorAction Stop

                $sp = Get-MgServicePrincipal -Filter ""appId eq '{cfg.ClientId}'"" -ErrorAction SilentlyContinue | Select-Object -First 1
                if (-not $sp) {{
                    Write-Output 'NO_SERVICE_PRINCIPAL'
                }} else {{
                    $assignments = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id -ErrorAction SilentlyContinue
                    foreach ($a in $assignments) {{
                        Write-Output ('ROLE:' + $a.AppRoleId)
                    }}
                }}
                Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
            ");

            var results = await Task.Run(() => ps.Invoke(), ct);
            if (ps.HadErrors)
            {
                var err = string.Join("; ", ps.Streams.Error.Select(e => e.ToString()));
                _logger.LogWarning("Falha ao consultar consents do tenant {Tenant}: {Err}", tenantId, err);
                return new TenantConsentsState();
            }

            var lines = results.Select(r => r?.BaseObject?.ToString() ?? "").ToList();
            if (lines.Any(l => l == "NO_SERVICE_PRINCIPAL"))
                return new TenantConsentsState();

            var roleIds = lines.Where(l => l.StartsWith("ROLE:")).Select(l => l[5..]).ToHashSet(StringComparer.OrdinalIgnoreCase);
            return new TenantConsentsState
            {
                Basic = new ConsentInfo { Granted = roleIds.Contains(PermissionBasicId) },
                Restore = new ConsentInfo { Granted = roleIds.Contains(PermissionRestoreId) }
            };
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "Erro consultando consents do tenant {Tenant}", tenantId);
            return new TenantConsentsState();
        }
    }

    // Constrói a URL de admin consent que o navegador deve abrir.
    // Após aceite, a Microsoft redireciona para o redirectUri com tenant=<id>&admin_consent=True&state=<...>.
    public string BuildAdminConsentUrl(string redirectUri, string state, string? tenantHint = null)
    {
        var cfg = _appStore.Read();
        if (!cfg.IsConfigured) throw new InvalidOperationException("App não configurado. Execute o setup inicial.");

        var tenant = string.IsNullOrWhiteSpace(tenantHint) ? "organizations" : tenantHint;
        var url = $"https://login.microsoftonline.com/{tenant}/v2.0/adminconsent" +
                  $"?client_id={cfg.ClientId}" +
                  $"&redirect_uri={Uri.EscapeDataString(redirectUri)}" +
                  $"&scope={Uri.EscapeDataString("https://graph.microsoft.com/.default")}" +
                  $"&state={Uri.EscapeDataString(state)}";
        return url;
    }
}
