using ClosedXML.Excel;
using HTML2EXCEL.Domain.Interfaces;
using HTML2EXCEL.Domain.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HTML2EXCEL.Infrastructure.Services
{
    public class ExcelExporter : IExcelExporter
    {
        private readonly XLWorkbook _workbook;

        public ExcelExporter()
        {
            _workbook = new XLWorkbook();
        }

        public XLWorkbook GetWorkbook() => _workbook;


        public Task<MemoryStream> CreateWorkbookAsync()
            => Task.FromResult(new MemoryStream());

        public async Task SaveAsync(MemoryStream stream, string path)
        {
            _workbook.SaveAs(path);
            await Task.CompletedTask;
        }
    }
}
