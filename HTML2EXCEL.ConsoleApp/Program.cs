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
using System;
using HTML2EXCEL.Infrastructure.Data;
using HTML2EXCEL.Infrastructure.Repositories;
using Microsoft.EntityFrameworkCore;
using HTML2EXCEL.Domain.Entities;
using HTML2EXCEL.ConsoleApp;

class Program
{
    public static async Task Main(string[] args)
    {
        await ExcelGeneratorApp.RunAsync();
        Log.CloseAndFlush();
    }
}
