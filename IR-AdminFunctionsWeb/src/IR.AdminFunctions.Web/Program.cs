using IR.AdminFunctions.Web;
using IR.AdminFunctions.Web.Services;
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

builder.Services.AddControllers()
    .AddJsonOptions(opts =>
    {
        opts.JsonSerializerOptions.PropertyNamingPolicy = System.Text.Json.JsonNamingPolicy.CamelCase;
    });

var app = builder.Build();

app.UseSerilogRequestLogging();
app.UseDefaultFiles();
app.UseStaticFiles();
app.MapControllers();
app.MapFallbackToFile("index.html");

try
{
    Log.Information("IR-AdminFunctionsWeb iniciando em {Urls}", builder.Configuration["Urls"]);
    app.Run();
}
catch (Exception ex)
{
    Log.Fatal(ex, "Falha fatal na inicialização do host");
    throw;
}
finally
{
    Log.CloseAndFlush();
}
