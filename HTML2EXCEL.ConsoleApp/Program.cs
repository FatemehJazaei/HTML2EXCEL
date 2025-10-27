using System;
using System.IO;
using System.Threading.Tasks;
using HTML2EXCEL.Infrastructure.Config;
using HTML2EXCEL.Application.Interfaces;
using HTML2EXCEL.Infrastructure.Services;
using HTML2EXCEL.Application.DTOs;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Http;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Configuration;
using HTML2EXCEL.Application.Handlers;

namespace HTML2EXCEL.ConsoleApp
{
    internal class Program
    {
        static async Task Main(string[] args)
        {

            using IHost host = Host.CreateDefaultBuilder(args)
                .ConfigureAppConfiguration((context, config) =>
                {
                    config.SetBasePath(Directory.GetCurrentDirectory());
                    config.AddJsonFile("appsettings.json", optional: false, reloadOnChange: true);
                    config.AddEnvironmentVariables();
                })
                .ConfigureServices((context, services) =>
                {
                    //  appsettings.json
                    var apiSettingsSection = context.Configuration.GetSection("ApiSettings");
                    var authSettingsSection = context.Configuration.GetSection("AuthSettings");

                    services.Configure<ApiSettings>(apiSettingsSection);
                    services.Configure<AuthSettings>(authSettingsSection);

                    var apiSettings = apiSettingsSection.Get<ApiSettings>();
                    var authSettings = authSettingsSection.Get<AuthSettings>();

                    services.AddHttpClient<IApiService, ApiService>(client =>
                    {
                        client.BaseAddress = new Uri(apiSettings.BaseUrl);
                        client.Timeout = TimeSpan.FromSeconds(30);
                    });

                    services.AddHttpClient<IAuthService, AuthService>(client =>
                    {
                        client.BaseAddress = new Uri(apiSettings.BaseUrl);
                    });


                    services.AddSingleton<IHtmlParser, HtmlParser>();
                    services.AddSingleton<IExcelExporter, ExcelExporter>();

                    // ---  Handler  ---
                    services.AddTransient<GenerateExcelHandler>();

                })
                .Build();

            // --- Resolve Handler ---
            var handler = host.Services.GetRequiredService<GenerateExcelHandler>();

            Console.WriteLine("Starting Excel generation process...");

            try
            {
                var request = new HtmlToExcelRequest
                {
                    Username = "your_username",
                    Password = "your_password",
                    OutputPath = "output/result.xlsx"
                };

                var result = await handler.HandleAsync(request);

                Console.WriteLine(result.Message);
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($" Error: {ex.Message}");
                Console.ResetColor();
            }

            await host.StopAsync();
        }
    }
}