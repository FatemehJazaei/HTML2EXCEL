using DocumentFormat.OpenXml.Drawing.Charts;
using HTML2EXCEL.Domain.Interfaces;
using HTML2EXCEL.Infrastructure.Config;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http.Json;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace HTML2EXCEL.Infrastructure.Services
{
    public class AuthService : IAuthService
    {
        private readonly HttpClient _httpClient;
        private readonly ApiSettings _settings;

        public AuthService(HttpClient httpClient, ApiSettings settings)
        {
            _httpClient = httpClient;
            _settings = settings;
        }

        public async Task<string> GetAccessTokenAsync(string username, string password,int companyId ,int periodId)
        {
            var url = $"{_settings.BaseUrl}/{_settings.AuthEndpoint}";

            var payload = new
            {
                userName = username,
                password = password,
                companyId = companyId,
                periodId = periodId
            };

            var response = await _httpClient.PostAsJsonAsync(url, payload);
            response.EnsureSuccessStatusCode();

            var json = await response.Content.ReadAsStringAsync();
            var token = JsonSerializer.Deserialize<string>(json);

            return token!;
        }
    }
}
