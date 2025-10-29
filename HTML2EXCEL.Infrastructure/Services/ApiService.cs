using HTML2EXCEL.Domain.Interfaces;
using HTML2EXCEL.Infrastructure.Config;
using System;
using System.Net.Http;
using System.Net.Http.Json;
using System.Threading.Tasks;
using Polly;
using Polly.Retry;
using System.Net.Http.Headers;
using System.Text.Json;
using HTML2EXCEL.Domain.Entities;

namespace HTML2EXCEL.Infrastructure.Services
{
    public class ApiService : IApiService
    {
        private readonly HttpClient _httpClient;
        private readonly ApiSettings _settings;

        public ApiService(HttpClient httpClient, ApiSettings settings)
        {
            _httpClient = httpClient;
            _settings = settings;
        }

        public async Task<string> GetModelAsync(string token, int tableTemplateId)
        {
            var url = $"{_settings.BaseUrl}/{_settings.ModelEndpoint}";
            _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);

            var payload = new
            {
                controllerName = _settings.ControllerName, 
                inputData = new
                {
                    tableTemplateId = tableTemplateId,
                    idList = Array.Empty<int>()
                }
            };

            var response = await _httpClient.PostAsJsonAsync(url, payload);
            response.EnsureSuccessStatusCode();

            var json = await response.Content.ReadAsStringAsync();
            var obj = JsonDocument.Parse(json);
            return obj.RootElement.GetProperty("model").GetString()!;
        }

        public async Task<string> GetFilePathAsync(string token, string model)
        {
            var url = $"{_settings.BaseUrl}/{_settings.PathEndpoint}/{model}";
            _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);

            var response = await _httpClient.GetAsync(url);
            response.EnsureSuccessStatusCode();

            var json = await response.Content.ReadAsStringAsync();
            var obj = JsonDocument.Parse(json);
            return obj.RootElement.GetProperty("path").GetString()!;
        }

        public async Task<byte[]> DownloadExcelFileAsync(string token, string filePath)
        {
            var url = $"{_settings.BaseUrl}/{filePath}";
            _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);

            var response = await _httpClient.GetAsync(url);
            response.EnsureSuccessStatusCode();

            return await response.Content.ReadAsByteArrayAsync();
        }
    }
}
