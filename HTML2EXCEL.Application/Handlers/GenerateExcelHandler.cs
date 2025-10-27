using HTML2EXCEL.Application.DTOs;
using HTML2EXCEL.Application.Interfaces;
using HTML2EXCEL.Domain.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HTML2EXCEL.Application.Handlers
{
    public class GenerateExcelHandler
    {
        private readonly IAuthService _authService;
        private readonly IApiService _apiService;
        private readonly IHtmlParser _htmlParser;
        private readonly IExcelExporter _excelExporter;

        public GenerateExcelHandler(
            IAuthService authService,
            IApiService apiService,
            IHtmlParser htmlParser,
            IExcelExporter excelExporter)
        {
            _authService = authService;
            _apiService = apiService;
            _htmlParser = htmlParser;
            _excelExporter = excelExporter;
        }

        /// <summary>
        /// Executes the HTML -> Excel use case.
        /// </summary>
        public async Task<HtmlToExcelResult> HandleAsync(HtmlToExcelRequest request)
        {
            try
            {
                // Authentication
                var token = await _authService.GetAccessTokenAsync(request.Username, request.Password);

                // Call API to get data key
                var dataKey = await _apiService.GetDataKeyAsync(token);

                // Get Excel URL (or HTML)
                var excelUrl = await _apiService.GetExcelUrlAsync(token, dataKey);


                var htmlContent = await new System.Net.Http.HttpClient().GetStringAsync(excelUrl);

                // Parse HTML tables
                List<TableData> tables = await _htmlParser.ParseTablesAsync(htmlContent);

                if (tables == null || tables.Count == 0)
                {
                    return new HtmlToExcelResult
                    {
                        Success = false,
                        Message = "No tables found in HTML."
                    };
                }

                // Export to Excel
                await _excelExporter.ExportAsync(tables, request.OutputPath);

                return new HtmlToExcelResult
                {
                    Success = true,
                    Message = "Excel file generated successfully.",
                    OutputPath = request.OutputPath
                };
            }
            catch (Exception ex)
            {
                return new HtmlToExcelResult
                {
                    Success = false,
                    Message = $"Error: {ex.Message}"
                };
            }
        }
    }
}
