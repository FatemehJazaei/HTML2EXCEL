using HTML2EXCEL.Domain.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;

namespace HTML2EXCEL.Domain.Interfaces
{
    public interface IExcelExporter
    {
        XLWorkbook GetWorkbook();
        Task<MemoryStream> CreateWorkbookAsync();

        /// <summary>
        /// Create an Excel file containing multiple tables.
        /// </summary>
        /// <param name="tables">List of TableData to export</param>
        /// <param name="outputPath">Full file path for Excel output</param>
  
        Task SaveAsync(MemoryStream stream, string path);
    }
}
