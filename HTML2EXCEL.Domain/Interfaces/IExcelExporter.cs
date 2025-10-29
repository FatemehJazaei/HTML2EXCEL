using HTML2EXCEL.Domain.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HTML2EXCEL.Domain.Interfaces
{
    public interface IExcelExporter
    {
        Task<MemoryStream> CreateWorkbookAsync();
        Task<int> WriteTextAsync(string text, int row);
        Task<int> WriteTableAsync(List<List<string>> data, int startRow);

        /// <summary>
        /// Create an Excel file containing multiple tables.
        /// </summary>
        /// <param name="tables">List of TableData to export</param>
        /// <param name="outputPath">Full file path for Excel output</param>
        Task SaveAsync(MemoryStream stream, string path);
    }
}
