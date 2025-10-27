using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Configuration;
using Serilog;
using HTML2EXCEL.Application.DTOs;
using HTML2EXCEL.Application.Handlers;
using HTML2EXCEL.Domain.Interfaces;
using HTML2EXCEL.Infrastructure.Config;
using HTML2EXCEL.Infrastructure.Services;
using Microsoft.Extensions.Logging;
using HTML2EXCEL.Infrastructure.Logging;

class Program
{
    static async Task Main(string[] args)
    {
        var host = Host.CreateDefaultBuilder(args)
            .UseSerilog((context, services, loggerConfig) =>
            {
                SerilogConfig.ConfigureSerilog(context, loggerConfig);
            })
            .ConfigureAppConfiguration((context, config) =>
            {
                config.AddJsonFile("appsettings.json", optional: false, reloadOnChange: true);
            })
            .ConfigureServices((context, services) =>
            {
                var apiSettings = context.Configuration.GetSection("ApiSettings").Get<ApiSettings>() ?? new ApiSettings();

                services.AddSingleton(apiSettings);
                services.AddHttpClient();
                services.AddTransient<IAuthService, AuthService>();
                services.AddTransient<IApiService, ApiService>();
                services.AddTransient<IHtmlParser, HtmlParser>();
                services.AddTransient<IExcelExporter, ExcelExporter>();
                services.AddTransient<GenerateExcelHandler>();
            })
            .Build();

        var logger = host.Services.GetRequiredService<ILogger<Program>>();
        logger.LogInformation("Application started.");

        try
        {
            var handler = host.Services.GetRequiredService<GenerateExcelHandler>();
            var request = new HtmlToExcelRequest
            {
                Username = "test",
                Password = "1234",
                OutputPath = "output/result.xlsx"
            };

            var result = await handler.HandleAsync(request);
            logger.LogInformation(result.Message);
        }
        catch (Exception ex)
        {
            logger.LogError(ex, "Unhandled exception occurred.");
        }
        finally
        {
            Log.CloseAndFlush();
        }
    }
}
