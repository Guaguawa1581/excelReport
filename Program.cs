using excelReport.Services;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
builder.Services.AddControllersWithViews();
builder.Services.AddMemoryCache();
builder.Services.AddHttpClient("ReportDataSource");

builder.Services.AddSingleton<IReportConfigStore, ReportConfigStore>();
builder.Services.AddSingleton<IReportTemplateScanner, ReportTemplateScanner>();
builder.Services.AddScoped<IDataSourceClient, DataSourceClient>();
builder.Services.AddScoped<IReportEngine, ReportEngine>();

var app = builder.Build();

// 啟動時自動 seed 示範範本與示範設定，讓 dotnet run 後即可直接操作。
var env = app.Services.GetRequiredService<IWebHostEnvironment>();
var templatesDir = Path.Combine(env.ContentRootPath, "App_Data", "templates");
TemplateSeeder.EnsureSampleTemplate(templatesDir, "sample_inspection.xlsx");

using (var scope = app.Services.CreateScope())
{
    var configStore = scope.ServiceProvider.GetRequiredService<IReportConfigStore>();
    var baseUrl = Environment.GetEnvironmentVariable("ASPNETCORE_URLS")
        ?? builder.Configuration["urls"]
        ?? app.Urls.FirstOrDefault()
        ?? "http://localhost:5047";
    ConfigSeeder.EnsureSampleConfig(configStore, baseUrl.Split(';')[0].TrimEnd('/'));
}

// Configure the HTTP request pipeline.
if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Home/Error");
    app.UseHsts();
}

app.UseHttpsRedirection();
app.UseStaticFiles();

app.UseRouting();

app.UseAuthorization();

app.MapControllerRoute(
    name: "default",
    pattern: "{controller=Report}/{action=Index}/{id?}");

app.Run();
