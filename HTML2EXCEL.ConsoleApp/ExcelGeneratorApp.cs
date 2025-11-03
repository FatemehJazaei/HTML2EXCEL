using DocumentFormat.OpenXml.Drawing.Charts;
using HTML2EXCEL.Application.DTOs;
using HTML2EXCEL.Application.Handlers;
using HTML2EXCEL.Application.Settings;
using HTML2EXCEL.Domain.Entities;
using HTML2EXCEL.Domain.Interfaces;
using HTML2EXCEL.Infrastructure.Config;
using HTML2EXCEL.Infrastructure.Data;
using HTML2EXCEL.Infrastructure.Logging;
using HTML2EXCEL.Infrastructure.Repositories;
using HTML2EXCEL.Infrastructure.Services;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Identity.Client;
using Serilog;
using Serilog.Events;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HTML2EXCEL.ConsoleApp
{
    public class ExcelGeneratorApp
    {
        public static async Task RunAsync()
        {
            using IHost host = Host.CreateDefaultBuilder()
                .UseSerilog((context, loggerConfig) =>
                {
                    SerilogConfig.ConfigureSerilog(context, loggerConfig);
                })
                .ConfigureAppConfiguration((context, config) =>
                {
                    config.AddJsonFile("appsettings.json", optional: false, reloadOnChange: true);
                    config.AddUserSecrets<Program>(optional: true);
                })
                .ConfigureServices((context, services) =>
                {
                    services.AddDbContext<AppDbContext>(options =>
                            options.UseSqlServer(
                                context.Configuration.GetConnectionString("DefaultConnection")
                            ));
                    var apiSettings = context.Configuration.GetSection("ApiSettings").Get<ApiSettings>()!;
                    var authSettings = context.Configuration.GetSection("AuthSettings").Get<AuthSettings>()!;
                    var outputSettings = context.Configuration.GetSection("OutputSettings").Get<OutputSettings>()!;

                    services.AddSingleton(apiSettings);
                    services.AddSingleton(authSettings);
                    services.AddSingleton(outputSettings);

                    services.AddHttpClient<IAuthService, AuthService>();
                    services.AddHttpClient<IApiService, ApiService>();
                    services.AddSingleton<IHtmlParser, HtmlParser>();
                    services.AddSingleton<IExcelExporter, ExcelExporter>();
                    services.AddScoped<IHtmlRepository, HtmlRepository>();

                    services.AddSingleton<GenerateExcelHandler>();

                    services.Configure<HtmlProcessingSettings>(
                          context.Configuration.GetSection("HtmlProcessingSettings"));
                })
                .Build();

            var handler = host.Services.GetRequiredService<GenerateExcelHandler>();
            var appConfig = host.Services.GetRequiredService<AuthSettings>();
            var appPath = host.Services.GetRequiredService<OutputSettings>();

            Console.WriteLine("Starting HTML to Excel process...");

            var request = new HtmlToExcelRequest
            {
                Username = appConfig.Username,
                Password = appConfig.Password,
                OutputPath = appPath.ExcelPath,
                CompanyId = appConfig.CompanyId,
                PeriodId = appConfig.PeriodId
            };

            var result = await handler.HandleAsync(request);

            if (result.Success)
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine($"Success: {result.Message}");
                Console.WriteLine($"File saved at: {result.OutputPath}");
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"Error: {result.Message}");
            }

            Console.ResetColor();
        }
    }
}
