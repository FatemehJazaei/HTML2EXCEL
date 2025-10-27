using HTML2EXCEL.Application.Interfaces;
using HTML2EXCEL.Infrastructure.Config;
using System;
using System.Net.Http;
using System.Net.Http.Json;
using System.Threading.Tasks;
using Polly;
using Polly.Retry;

namespace HTML2EXCEL.Infrastructure.Services
{
    public class ApiService : IApiService
    {
        private readonly HttpClient _httpClient;
        private readonly ApiSettings _apiSettings;
        private readonly AsyncRetryPolicy _retryPolicy;

        public ApiService(HttpClient httpClient, ApiSettings apiSettings)
        {
            _httpClient = httpClient;
            _apiSettings = apiSettings;

            _retryPolicy = Policy
                .Handle<HttpRequestException>()
                .WaitAndRetryAsync(
                    retryCount: _apiSettings.RetryCount,
                    sleepDurationProvider: attempt => TimeSpan.FromSeconds(Math.Pow(_apiSettings.RetryDelaySeconds, attempt))
                );
        }

        public async Task<string> GetDataKeyAsync(string token)
        {
            return await _retryPolicy.ExecuteAsync(async () =>
            {
                using var request = new HttpRequestMessage(HttpMethod.Get, _apiSettings.DataKeyEndpoint);
                request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);

                var response = await _httpClient.SendAsync(request);
                response.EnsureSuccessStatusCode();

                var result = await response.Content.ReadFromJsonAsync<DataKeyResponse>();
                return result?.DataKey ?? throw new Exception("Data key not returned.");
            });
        }

        public async Task<string> GetExcelUrlAsync(string token, string dataKey)
        {
            return await _retryPolicy.ExecuteAsync(async () =>
            {
                var endpoint = $"{_apiSettings.ExcelUrlEndpoint}?key={dataKey}";
                using var request = new HttpRequestMessage(HttpMethod.Get, endpoint);
                request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);

                var response = await _httpClient.SendAsync(request);
                response.EnsureSuccessStatusCode();

                var result = await response.Content.ReadFromJsonAsync<ExcelUrlResponse>();
                return result?.Url ?? throw new Exception("Excel URL not returned.");
            });
        }

        private record DataKeyResponse(string DataKey);
        private record ExcelUrlResponse(string Url);
    }
}
