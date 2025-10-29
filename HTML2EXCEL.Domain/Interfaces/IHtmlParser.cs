using HTML2EXCEL.Domain.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HTML2EXCEL.Domain.Interfaces
{
    public interface IHtmlParser
    {
        /// <summary>
        /// Parse HTML string and extract all tables as TableData entities.
        /// </summary>
        /// <param name="htmlContent">Raw HTML string</param>
        /// <returns>List of TableData entities</returns>
        Task<MemoryStream> ParseTablesAsync(string htmlContent, string token, IApiService apiService,
        IExcelExporter excelExporter);
    }
}
