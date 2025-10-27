using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HTML2EXCEL.Application.Interfaces
{
    public interface IAuthService
    {
        /// <summary>
        /// Authenticate with API using credentials and get access token.
        /// </summary>
        /// <param name="username">Username or API key</param>
        /// <param name="password">Password or secret</param>
        /// <returns>Access token string</returns>
        Task<string> GetAccessTokenAsync(string username, string password);
    }
}
