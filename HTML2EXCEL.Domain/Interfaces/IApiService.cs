using HTML2EXCEL.Domain.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HTML2EXCEL.Domain.Interfaces
{
    public interface IApiService
    {
        /// <summary>
        /// Calls the API to fetch a data ID or key from the server.
        /// </summary>
        /// <param name="token">Authentication token</param>
        /// <param name="tableTemplateId">Authentication token</param>
        /// <returns>Generated data ID or key</returns>
        Task<string> GetModelAsync(string token, int tableTemplateId);

        /// <summary>
        /// Uses the given key to get a downloadable Excel file Path.
        /// </summary>
        /// <param name="token">Authentication token</param>
        /// <param name="model">Previously obtained key</param>
        /// <returns>Excel file Path</returns>
        Task<string> GetFilePathAsync(string token, string model);
        /// <summary>
        /// Uses the given Path to get a downloadable Excel file .
        /// </summary>
        /// <param name="token">Authentication token</param>
        /// <param name="filePath">Excel file Path</param>
        /// <returns>Excel file </returns>
        Task<byte[]> DownloadExcelFileAsync(string token, string filePath);

    }
}
