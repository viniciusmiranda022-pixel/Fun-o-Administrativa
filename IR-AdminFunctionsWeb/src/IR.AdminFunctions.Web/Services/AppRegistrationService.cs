using System.Collections.Concurrent;
using System.Net.Http.Headers;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Text.Json;
using Microsoft.Identity.Client;
using IR.AdminFunctions.Web.Models;

namespace IR.AdminFunctions.Web.Services;

// Registers the multi-tenant application in Azure AD using MSAL device code flow.
// The device code callback provides UserCode and VerificationUri directly,
// without relying on PowerShell output parsing.
public class AppRegistrationService
{
    private readonly AppConfigStore _store;
    private readonly ILogger<AppRegistrationService> _logger;
    private readonly IHttpClientFactory _httpClientFactory;
    private readonly ConcurrentDictionary<string, SetupStatus> _states = new();

    private const string AppDisplayName = "IR Administrative Function Recovery";

    // Public Client ID of Microsoft Graph PowerShell (well-known, multi-tenant)
    private const string PublicClientId = "14d82eec-204b-4c2f-b7e8-296a70dab67e";
    private const string Authority = "https://login.microsoftonline.com/organizations";

    // Default Redirect URIs for the tenant consent OAuth flow
    private static readonly string[] DefaultRedirectUris =
    [
        "http://localhost:8080/api/oauth/callback",
        "https://localhost:8080/api/oauth/callback",
    ];

    // Microsoft Graph permission IDs (global)
    private const string GraphResourceAppId = "00000003-0000-0000-c000-000000000000";
    private const string RoleManagementReadDirectory = "483bed4a-2ad3-4361-a73b-c83ccdbdc53c";
    private const string RoleManagementReadWriteDirectory = "9e3f62cf-ca93-4989-b6ce-bf83c28f9fe8";

    private static readonly string[] GraphScopes =
    [
        "https://graph.microsoft.com/Application.ReadWrite.All",
        "https://graph.microsoft.com/Directory.ReadWrite.All",
        "https://graph.microsoft.com/User.Read",
    ];

    private static readonly JsonSerializerOptions JsonOpts = new()
    {
        PropertyNameCaseInsensitive = true,
        WriteIndented = false,
    };

    public AppRegistrationService(
        AppConfigStore store,
        ILogger<AppRegistrationService> logger,
        IHttpClientFactory httpClientFactory)
    {
        _store = store;
        _logger = logger;
        _httpClientFactory = httpClientFactory;
    }

    public SetupStatus GetStatus(string sessionId) =>
        _states.TryGetValue(sessionId, out var s) ? s : new SetupStatus { Status = "NotStarted" };

    public string StartSetup()
    {
        var sessionId = Guid.NewGuid().ToString("N");
        _states[sessionId] = new SetupStatus { Status = "WaitingForUser", Message = "Waiting for user to authenticate..." };
        _logger.LogInformation("[Setup:{SessionId}] Setup started.", sessionId);
        _ = Task.Run(() => RunSetupAsync(sessionId));
        return sessionId;
    }

    private async Task RunSetupAsync(string sessionId)
    {
        try
        {
            // Step 1 — authentication via device code (MSAL)
            _logger.LogInformation("[Setup:{SessionId}] Step 1/3 — Starting authentication via device code (MSAL)...", sessionId);

            var msalApp = PublicClientApplicationBuilder
                .Create(PublicClientId)
                .WithAuthority(Authority)
                .Build();

            AuthenticationResult tokenResult;
            try
            {
                tokenResult = await msalApp
                    .AcquireTokenWithDeviceCode(GraphScopes, deviceCode =>
                    {
                        _logger.LogInformation(
                            "[Setup:{SessionId}] Device code received. UserCode={Code}, Url={Url}",
                            sessionId, deviceCode.UserCode, deviceCode.VerificationUrl);

                        UpdateStatus(sessionId, s =>
                        {
                            s.UserCode = deviceCode.UserCode;
                            s.VerificationUrl = deviceCode.VerificationUrl;
                            s.Message = $"Go to {deviceCode.VerificationUrl} and enter the code: {deviceCode.UserCode}";
                        });

                        return Task.CompletedTask;
                    })
                    .ExecuteAsync();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "[Setup:{SessionId}] MSAL authentication failed.", sessionId);
                Fail(sessionId, $"Authentication failed: {ex.Message}");
                return;
            }

            var tenantId = tokenResult.TenantId;
            var account = tokenResult.Account?.Username;

            _logger.LogInformation(
                "[Setup:{SessionId}] Step 1/3 completed. TenantId={TenantId}, Account={Account}",
                sessionId, tenantId, account);

            if (string.IsNullOrEmpty(tenantId))
            {
                Fail(sessionId, "Could not retrieve the TenantId from the token.");
                return;
            }

            UpdateStatus(sessionId, s =>
            {
                s.Status = "Registering";
                s.Message = $"Connected as {account}. Creating App Registration...";
                s.TenantId = tenantId;
            });

            // Step 2 — create App Registration via Microsoft Graph REST
            _logger.LogInformation("[Setup:{SessionId}] Step 2/3 — Creating App Registration via Graph API...", sessionId);

            using var http = _httpClientFactory.CreateClient();
            http.DefaultRequestHeaders.Authorization =
                new AuthenticationHeaderValue("Bearer", tokenResult.AccessToken);
            http.DefaultRequestHeaders.Accept.Add(
                new MediaTypeWithQualityHeaderValue("application/json"));

            // Checks whether the app already exists
            string appObjectId;
            string appClientId;

            var filterResp = await http.GetAsync(
                $"https://graph.microsoft.com/v1.0/applications?$filter=displayName eq '{Uri.EscapeDataString(AppDisplayName)}'&$select=id,appId");

            filterResp.EnsureSuccessStatusCode();
            var filterJson = await filterResp.Content.ReadAsStringAsync();
            var filterDoc = JsonSerializer.Deserialize<GraphList>(filterJson, JsonOpts);

            if (filterDoc?.Value?.Length > 0)
            {
                appObjectId = filterDoc.Value[0].Id!;
                appClientId = filterDoc.Value[0].AppId!;
                _logger.LogInformation(
                    "[Setup:{SessionId}] Existing App Registration found. AppId={AppId}", sessionId, appClientId);
            }
            else
            {
                // Creates the App Registration
                var appPayload = new
                {
                    displayName = AppDisplayName,
                    signInAudience = "AzureADMultipleOrgs",
                    requiredResourceAccess = new[]
                    {
                        new
                        {
                            resourceAppId = GraphResourceAppId,
                            resourceAccess = new[]
                            {
                                new { id = RoleManagementReadDirectory, type = "Role" },
                                new { id = RoleManagementReadWriteDirectory, type = "Role" },
                            },
                        },
                    },
                    web = new
                    {
                        redirectUris = DefaultRedirectUris,
                    },
                };

                var createContent = new StringContent(
                    JsonSerializer.Serialize(appPayload, JsonOpts),
                    Encoding.UTF8, "application/json");

                var createResp = await http.PostAsync(
                    "https://graph.microsoft.com/v1.0/applications", createContent);

                if (!createResp.IsSuccessStatusCode)
                {
                    var errBody = await createResp.Content.ReadAsStringAsync();
                    _logger.LogError("[Setup:{SessionId}] Failed to create App Registration: {Status} {Body}",
                        sessionId, (int)createResp.StatusCode, errBody);
                    Fail(sessionId, $"Failed to create App Registration ({(int)createResp.StatusCode}): {errBody}");
                    return;
                }

                var createJson = await createResp.Content.ReadAsStringAsync();
                var created = JsonSerializer.Deserialize<GraphApp>(createJson, JsonOpts);
                appObjectId = created?.Id ?? throw new InvalidOperationException("App created without Id.");
                appClientId = created?.AppId ?? throw new InvalidOperationException("App created without AppId.");
                _logger.LogInformation(
                    "[Setup:{SessionId}] App Registration created. AppId={AppId}", sessionId, appClientId);
            }

            // Ensures redirect URIs on the app (required for the OAuth consent flow)
            await EnsureRedirectUrisAsync(http, sessionId, appObjectId);

            // Ensures the Service Principal exists in the bootstrap tenant
            string spId;
            var spFilterResp = await http.GetAsync(
                $"https://graph.microsoft.com/v1.0/servicePrincipals?$filter=appId eq '{appClientId}'&$select=id");

            spFilterResp.EnsureSuccessStatusCode();
            var spFilterJson = await spFilterResp.Content.ReadAsStringAsync();
            var spFilterDoc = JsonSerializer.Deserialize<GraphList>(spFilterJson, JsonOpts);

            if (spFilterDoc?.Value?.Length > 0)
            {
                spId = spFilterDoc.Value[0].Id!;
                _logger.LogInformation("[Setup:{SessionId}] Existing Service Principal found. SpId={SpId}", sessionId, spId);
            }
            else
            {
                var spPayload = new { appId = appClientId };
                var spContent = new StringContent(
                    JsonSerializer.Serialize(spPayload, JsonOpts),
                    Encoding.UTF8, "application/json");

                var spResp = await http.PostAsync(
                    "https://graph.microsoft.com/v1.0/servicePrincipals", spContent);

                if (!spResp.IsSuccessStatusCode)
                {
                    var errBody = await spResp.Content.ReadAsStringAsync();
                    _logger.LogError("[Setup:{SessionId}] Failed to create Service Principal: {Status} {Body}",
                        sessionId, (int)spResp.StatusCode, errBody);
                    Fail(sessionId, $"Failed to create Service Principal ({(int)spResp.StatusCode}): {errBody}");
                    return;
                }

                var spJson = await spResp.Content.ReadAsStringAsync();
                var sp = JsonSerializer.Deserialize<GraphApp>(spJson, JsonOpts);
                spId = sp?.Id ?? throw new InvalidOperationException("SP created without Id.");
                _logger.LogInformation("[Setup:{SessionId}] Service Principal created. SpId={SpId}", sessionId, spId);
            }

            // Creates client secret valid for 2 years
            var secretPayload = new
            {
                passwordCredential = new
                {
                    displayName = "Auto-generated by IR setup",
                    endDateTime = DateTimeOffset.UtcNow.AddYears(2).ToString("o"),
                },
            };

            var secretContent = new StringContent(
                JsonSerializer.Serialize(secretPayload, JsonOpts),
                Encoding.UTF8, "application/json");

            var secretResp = await http.PostAsync(
                $"https://graph.microsoft.com/v1.0/applications/{appObjectId}/addPassword",
                secretContent);

            if (!secretResp.IsSuccessStatusCode)
            {
                var errBody = await secretResp.Content.ReadAsStringAsync();
                _logger.LogError("[Setup:{SessionId}] Failed to create client secret: {Status} {Body}",
                    sessionId, (int)secretResp.StatusCode, errBody);
                Fail(sessionId, $"Failed to create client secret ({(int)secretResp.StatusCode}): {errBody}");
                return;
            }

            var secretJson = await secretResp.Content.ReadAsStringAsync();
            var secretDoc = JsonSerializer.Deserialize<GraphSecret>(secretJson, JsonOpts);
            var clientSecret = secretDoc?.SecretText
                ?? throw new InvalidOperationException("Client secret created without SecretText.");

            _logger.LogInformation("[Setup:{SessionId}] Step 2/3 completed. AppId={AppId}", sessionId, appClientId);

            // Step 2.5 — create authentication certificate and register in Azure AD
            string? certThumbprint = null;
            try
            {
                UpdateStatus(sessionId, s => s.Message = "Creating authentication certificate...");
                certThumbprint = await CreateAndRegisterCertificateAsync(http, sessionId, appObjectId, appClientId);
                _logger.LogInformation("[Setup:{SessionId}] Certificate created. Thumbprint={Tp}", sessionId, certThumbprint);
            }
            catch (Exception ex)
            {
                _logger.LogWarning(ex, "[Setup:{SessionId}] Failed to create certificate. Setup continues without a certificate.", sessionId);
            }

            // Step 3 — persist configuration
            _logger.LogInformation("[Setup:{SessionId}] Step 3/3 — Persisting configuration...", sessionId);

            var config = new AppConfig
            {
                ClientId = appClientId,
                ApplicationObjectId = appObjectId,
                ServicePrincipalId = spId,
                ClientSecret = clientSecret,
                CertificateThumbprint = certThumbprint,
                DisplayName = AppDisplayName,
                BootstrapTenantId = tenantId,
                CreatedAt = DateTimeOffset.UtcNow,
            };

            _store.Write(config);
            _logger.LogInformation("[Setup:{SessionId}] Setup completed. ClientId={ClientId}", sessionId, config.ClientId);

            UpdateStatus(sessionId, s =>
            {
                s.Status = "Completed";
                s.Message = "App registered successfully. You can now add tenants.";
                s.ClientId = config.ClientId;
            });
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "[Setup:{SessionId}] Unhandled exception during setup.", sessionId);
            Fail(sessionId, ex.Message);
        }
    }

    private async Task<string?> CreateAndRegisterCertificateAsync(
        HttpClient http, string sessionId, string appObjectId, string appClientId)
    {
        // Creates a self-signed certificate in memory
        using var rsa = RSA.Create(2048);
        var req = new CertificateRequest(
            $"CN={appClientId}",
            rsa,
            HashAlgorithmName.SHA256,
            RSASignaturePadding.Pkcs1);
        req.CertificateExtensions.Add(
            new X509KeyUsageExtension(X509KeyUsageFlags.DigitalSignature, false));

        using var cert = req.CreateSelfSigned(
            DateTimeOffset.UtcNow.AddDays(-1),
            DateTimeOffset.UtcNow.AddYears(2));

        // Exports as PFX and re-imports with persistence flag in the cert store
        var pfxBytes = cert.Export(X509ContentType.Pkcs12);
        string thumbprint;

        // Tries LocalMachine first (required for Windows services)
        try
        {
            var flags = X509KeyStorageFlags.MachineKeySet | X509KeyStorageFlags.PersistKeySet | X509KeyStorageFlags.Exportable;
            using var persistent = new X509Certificate2(pfxBytes, (string?)null, flags);
            using var store = new X509Store(StoreName.My, StoreLocation.LocalMachine);
            store.Open(OpenFlags.ReadWrite);
            store.Add(persistent);
            thumbprint = persistent.Thumbprint;
            _logger.LogInformation("[Setup:{SessionId}] Cert installed in LocalMachine\\My: {Tp}", sessionId, thumbprint);
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "[Setup:{SessionId}] Could not install in LocalMachine\\My, falling back to CurrentUser\\My.", sessionId);
            var flags = X509KeyStorageFlags.UserKeySet | X509KeyStorageFlags.PersistKeySet | X509KeyStorageFlags.Exportable;
            using var persistent = new X509Certificate2(pfxBytes, (string?)null, flags);
            using var store = new X509Store(StoreName.My, StoreLocation.CurrentUser);
            store.Open(OpenFlags.ReadWrite);
            store.Add(persistent);
            thumbprint = persistent.Thumbprint;
            _logger.LogInformation("[Setup:{SessionId}] Cert installed in CurrentUser\\My: {Tp}", sessionId, thumbprint);
        }

        // Registers the public key in Azure AD via PATCH /applications/{id}
        var pubKeyDer = cert.Export(X509ContentType.Cert);
        var keyCredPayload = new
        {
            keyCredentials = new[]
            {
                new
                {
                    type = "AsymmetricX509Cert",
                    usage = "Verify",
                    key = Convert.ToBase64String(pubKeyDer),
                    displayName = "IR Setup Certificate",
                }
            }
        };

        var patchContent = new StringContent(
            JsonSerializer.Serialize(keyCredPayload, JsonOpts),
            Encoding.UTF8, "application/json");
        var patchReq = new HttpRequestMessage(HttpMethod.Patch,
            $"https://graph.microsoft.com/v1.0/applications/{appObjectId}")
        { Content = patchContent };

        var patchResp = await http.SendAsync(patchReq);
        if (!patchResp.IsSuccessStatusCode)
        {
            var errBody = await patchResp.Content.ReadAsStringAsync();
            _logger.LogWarning("[Setup:{SessionId}] Failed to register cert in Azure AD: {Status} {Body}",
                sessionId, (int)patchResp.StatusCode, errBody);
        }
        else
        {
            _logger.LogInformation("[Setup:{SessionId}] Certificate successfully registered in Azure AD.", sessionId);
        }

        return thumbprint;
    }

    private async Task EnsureRedirectUrisAsync(HttpClient http, string sessionId, string appObjectId)
    {
        try
        {
            var getResp = await http.GetAsync(
                $"https://graph.microsoft.com/v1.0/applications/{appObjectId}?$select=web");
            getResp.EnsureSuccessStatusCode();
            var getJson = await getResp.Content.ReadAsStringAsync();
            var appFull = JsonSerializer.Deserialize<GraphAppFull>(getJson, JsonOpts);

            var existing = appFull?.Web?.RedirectUris ?? [];
            var missing = DefaultRedirectUris.Except(existing).ToArray();

            if (missing.Length == 0)
            {
                _logger.LogInformation("[Setup:{SessionId}] Redirect URIs already present on the app.", sessionId);
                return;
            }

            var merged = existing.Concat(missing).ToArray();
            var patchPayload = new { web = new { redirectUris = merged } };
            var patchContent = new StringContent(
                JsonSerializer.Serialize(patchPayload, JsonOpts),
                Encoding.UTF8, "application/json");

            var patchReq = new HttpRequestMessage(HttpMethod.Patch,
                $"https://graph.microsoft.com/v1.0/applications/{appObjectId}")
            { Content = patchContent };

            var patchResp = await http.SendAsync(patchReq);
            if (patchResp.IsSuccessStatusCode)
                _logger.LogInformation("[Setup:{SessionId}] Redirect URIs updated: {Uris}", sessionId, string.Join(", ", missing));
            else
            {
                var err = await patchResp.Content.ReadAsStringAsync();
                _logger.LogWarning("[Setup:{SessionId}] Could not update redirect URIs: {Status} {Body}", sessionId, (int)patchResp.StatusCode, err);
            }
        }
        catch (Exception ex)
        {
            _logger.LogWarning(ex, "[Setup:{SessionId}] Error ensuring redirect URIs.", sessionId);
        }
    }

    private void UpdateStatus(string sessionId, Action<SetupStatus> mutator)
    {
        _states.AddOrUpdate(sessionId,
            _ => { var s = new SetupStatus(); mutator(s); return s; },
            (_, s) => { mutator(s); return s; });
    }

    private void Fail(string sessionId, string reason)
    {
        UpdateStatus(sessionId, s =>
        {
            s.Status = "Failed";
            s.Message = reason;
        });
    }

    // DTOs for deserializing Graph responses
    private sealed class GraphList
    {
        public GraphApp[]? Value { get; set; }
    }

    private sealed class GraphApp
    {
        public string? Id { get; set; }
        public string? AppId { get; set; }
    }

    private sealed class GraphAppFull
    {
        public GraphWebSection? Web { get; set; }
    }

    private sealed class GraphWebSection
    {
        public string[]? RedirectUris { get; set; }
    }

    private sealed class GraphSecret
    {
        public string? SecretText { get; set; }
    }
}
