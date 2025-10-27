using Microsoft.Extensions.Hosting;
using Serilog;
using Serilog.Events;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HTML2EXCEL.Infrastructure.Logging
{
    public static class SerilogConfig
    {
        public static void ConfigureSerilog(HostBuilderContext context, LoggerConfiguration loggerConfiguration)
        {
            var env = context.HostingEnvironment;

            loggerConfiguration
                .ReadFrom.Configuration(context.Configuration)
                .Enrich.FromLogContext()
                .Enrich.WithProperty("Environment", env.EnvironmentName)
                .MinimumLevel.Override("Microsoft", LogEventLevel.Warning)
                .WriteTo.Console()
                .WriteTo.File(
                    path: "logs/log-.txt",
                    rollingInterval: RollingInterval.Day,
                    restrictedToMinimumLevel: LogEventLevel.Information,
                    retainedFileCountLimit: 10,
                    shared: true);
        }
    }
}
