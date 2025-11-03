using HTML2EXCEL.Application.DTOs;
using HTML2EXCEL.Domain.Entities;
using HTML2EXCEL.Domain.Interfaces;
using Microsoft.Extensions.Options;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Runtime;
using System.Text;
using System.Threading.Tasks;

namespace HTML2EXCEL.Application.Handlers
{
    public class GenerateExcelHandler
    {
        private readonly IHtmlRepository _htmlRepository;
        private readonly IAuthService _authService;
        private readonly IApiService _apiService;
        private readonly IHtmlParser _htmlParser;
        private readonly IExcelExporter _excelExporter;
        private readonly HtmlProcessingSettings _settings;

        public GenerateExcelHandler(
            IAuthService authService,
            IApiService apiService,
            IHtmlParser htmlParser,
            IExcelExporter excelExporter,
            IHtmlRepository htmlRepository,
            IOptions<HtmlProcessingSettings> options)
        {
            _authService = authService;
            _apiService = apiService;
            _htmlParser = htmlParser;
            _excelExporter = excelExporter;
            _htmlRepository = htmlRepository;
            _settings = options.Value;
        }

        /// <summary>
        /// Executes the HTML -> Excel use case.
        /// </summary>
        public async Task<HtmlToExcelResult> HandleAsync(HtmlToExcelRequest request)
        {
            try
            {
                var htmlContent = await _htmlRepository.GetHtmlContentAsync(_settings.TargetId);
     
                var token = await _authService.GetAccessTokenAsync(
                    request.Username, request.Password, request.CompanyId, request.PeriodId);

                var excelStream = await _excelExporter.CreateWorkbookAsync();


                excelStream = await _htmlParser.ParseTablesAsync(htmlContent, token,  _apiService, _excelExporter);

 
                await _excelExporter.SaveAsync(excelStream, request.OutputPath);

                return new HtmlToExcelResult
                {
                    Success = true,
                    Message = "Excel generated successfully.",
                    OutputPath = request.OutputPath
                };
            }
            catch (Exception ex)
            {
                return new HtmlToExcelResult
                {
                    Success = false,
                    Message = ex.Message
                };
            }
        }
    }
}
