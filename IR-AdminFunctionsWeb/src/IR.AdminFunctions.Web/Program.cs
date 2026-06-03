using System.Threading.RateLimiting;
using IR.AdminFunctions.Web;
using IR.AdminFunctions.Web.Services;
using Microsoft.AspNetCore.RateLimiting;
using Microsoft.AspNetCore.ResponseCompression;
using Microsoft.Extensions.Hosting.WindowsServices;
using Serilog;

var builder = WebApplication.CreateBuilder(new WebApplicationOptions
{
    Args = args,
    ContentRootPath = WindowsServiceHelpers.IsWindowsService()
        ? AppContext.BaseDirectory
        : default
});

var questOptions = builder.Configuration.GetSection("Quest").Get<QuestOptions>() ?? new QuestOptions();
Directory.CreateDirectory(questOptions.WebLogRoot);

Log.Logger = new LoggerConfiguration()
    .ReadFrom.Configuration(builder.Configuration)
    .Enrich.FromLogContext()
    .WriteTo.Console()
    .WriteTo.File(
        Path.Combine(questOptions.WebLogRoot, "webapp-.log"),
        rollingInterval: RollingInterval.Day,
        retainedFileCountLimit: 30,
        outputTemplate: "{Timestamp:yyyy-MM-dd HH:mm:ss.fff zzz} [{Level:u3}] {Message:lj} {Properties:j}{NewLine}{Exception}")
    .CreateLogger();

builder.Host.UseSerilog();

builder.Host.UseWindowsService(options =>
{
    options.ServiceName = "IR-AdminFunctionsWeb";
});

builder.WebHost.UseUrls(builder.Configuration["Urls"] ?? "http://localhost:8080");

builder.Services.Configure<QuestOptions>(builder.Configuration.GetSection("Quest"));
builder.Services.AddSingleton<PowerShellRunner>();
builder.Services.AddSingleton<BackupReader>();
builder.Services.AddSingleton<SettingsReader>();
builder.Services.AddSingleton<LogReader>();
builder.Services.AddSingleton<JobManager>();
builder.Services.AddSingleton<TenantStore>();
builder.Services.AddSingleton<AppConfigStore>();
builder.Services.AddSingleton<AppRegistrationService>();
builder.Services.AddSingleton<ConsentChecker>();
builder.Services.AddHttpClient();
builder.Services.AddHostedService<ScriptDeployer>();
builder.Services.AddHostedService<ProvisioningService>();
builder.Services.AddHostedService<JobCleanupService>();

// Response compression (gzip/brotli) — reduces JSON payload size by 5-10x
builder.Services.AddResponseCompression(opts =>
{
    opts.EnableForHttps = true;
    opts.Providers.Add<BrotliCompressionProvider>();
    opts.Providers.Add<GzipCompressionProvider>();
    opts.MimeTypes = ResponseCompressionDefaults.MimeTypes.Concat(
        ["application/json", "text/html", "application/javascript"]);
});

// Rate limiting — prevents abuse of job-triggering endpoints
builder.Services.AddRateLimiter(opts =>
{
    // General: 120 requests per minute per IP
    opts.AddFixedWindowLimiter("general", o =>
    {
        o.PermitLimit = 120;
        o.Window = TimeSpan.FromMinutes(1);
        o.QueueProcessingOrder = QueueProcessingOrder.OldestFirst;
        o.QueueLimit = 10;
    });
    // Jobs: max 10 job triggers per minute per IP (backup/restore/compare)
    opts.AddFixedWindowLimiter("jobs", o =>
    {
        o.PermitLimit = 10;
        o.Window = TimeSpan.FromMinutes(1);
        o.QueueProcessingOrder = QueueProcessingOrder.OldestFirst;
        o.QueueLimit = 2;
    });
    opts.RejectionStatusCode = StatusCodes.Status429TooManyRequests;
});

builder.Services.AddControllers()
    .AddJsonOptions(opts =>
    {
        opts.JsonSerializerOptions.PropertyNamingPolicy = System.Text.Json.JsonNamingPolicy.CamelCase;
        opts.JsonSerializerOptions.Converters.Add(new System.Text.Json.Serialization.JsonStringEnumConverter());
    });

var app = builder.Build();

// Security headers on every response
app.Use(async (ctx, next) =>
{
    ctx.Response.Headers["X-Content-Type-Options"] = "nosniff";
    ctx.Response.Headers["X-Frame-Options"] = "DENY";
    ctx.Response.Headers["X-XSS-Protection"] = "1; mode=block";
    ctx.Response.Headers["Referrer-Policy"] = "strict-origin-when-cross-origin";
    ctx.Response.Headers["Permissions-Policy"] = "camera=(), microphone=(), geolocation=()";
    ctx.Response.Headers["X-DNS-Prefetch-Control"] = "off";
    await next(ctx);
});

app.UseResponseCompression();
app.UseRateLimiter();
app.UseSerilogRequestLogging();
app.UseDefaultFiles();
app.UseStaticFiles();
app.MapControllers().RequireRateLimiting("general");
app.MapFallbackToFile("index.html");

try
{
    Log.Information("IR-AdminFunctionsWeb starting on {Urls}", builder.Configuration["Urls"] ?? "http://localhost:8080");
    app.Run();
}
catch (Exception ex)
{
    Log.Fatal(ex, "Fatal failure during host initialization");
    throw;
}
finally
{
    Log.CloseAndFlush();
}
