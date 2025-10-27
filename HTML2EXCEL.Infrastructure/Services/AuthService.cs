using HTML2EXCEL.Application.Interfaces;
using HTML2EXCEL.Infrastructure.Config;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http.Json;
using System.Text;
using System.Threading.Tasks;

namespace HTML2EXCEL.Infrastructure.Services
{
    public class AuthService : IAuthService
    {
        private readonly HttpClient _httpClient;
        private readonly ApiSettings _apiSettings;

        private string? _cachedToken;

        public AuthService(HttpClient httpClient, ApiSettings apiSettings)
        {
            _httpClient = httpClient;
            _apiSettings = apiSettings;
        }

        public async Task<string> GetAccessTokenAsync(string username, string password)
        {

            if (!string.IsNullOrEmpty(_cachedToken))
                return _cachedToken;

            var request = new
            {
                username,
                password
            };

            var response = await _httpClient.PostAsJsonAsync(_apiSettings.AuthEndpoint, request);
            response.EnsureSuccessStatusCode();

            var result = await response.Content.ReadFromJsonAsync<AuthResponse>();

            if (result == null || string.IsNullOrEmpty(result.Token))
                throw new System.Exception("Failed to get auth token.");

            _cachedToken = result.Token;

            return _cachedToken;
        }

        private class AuthResponse
        {
            public string Token { get; set; } = string.Empty;
            public int ExpiresIn { get; set; }
        }
    }
}
