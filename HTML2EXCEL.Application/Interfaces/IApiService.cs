using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HTML2EXCEL.Application.Interfaces
{
    public interface IApiService
    {
        /// <summary>
        /// Calls the API to fetch a data ID or key from the server.
        /// </summary>
        /// <param name="token">Authentication token</param>
        /// <returns>Generated data ID or key</returns>
        Task<string> GetDataKeyAsync(string token);

        /// <summary>
        /// Uses the given key to get a downloadable Excel file URL.
        /// </summary>
        /// <param name="token">Authentication token</param>
        /// <param name="dataKey">Previously obtained key</param>
        /// <returns>Excel file URL</returns>
        Task<string> GetExcelUrlAsync(string token, string dataKey);
    }
}
