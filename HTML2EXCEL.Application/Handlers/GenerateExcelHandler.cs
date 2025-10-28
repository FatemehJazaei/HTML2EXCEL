using HTML2EXCEL.Application.DTOs;
using HTML2EXCEL.Domain.Entities;
using HTML2EXCEL.Domain.Interfaces;
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
                var token = await _authService.GetAccessTokenAsync(request.Username, request.Password, request.CompanyId, request.PeriodId);
                var model = await _apiService.GetModelAsync(token, request.TableTemplateId);
                var path = await _apiService.GetFilePathAsync(token, model);

                var fileBytes = await _apiService.DownloadExcelFileAsync(token, path);

                if (fileBytes == null )
                {
                    return new HtmlToExcelResult
                    {
                        Success = false,
                        Message = "No tables found in HTML."
                    };
                }

                // Export to Excel
                await _excelExporter.ExportAsync(fileBytes, request.OutputPath);

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
