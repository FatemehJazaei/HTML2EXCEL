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
        private readonly IXLWorksheet _sheet;

        public ExcelExporter()
        {
            _workbook = new XLWorkbook();
            _sheet = _workbook.Worksheets.Add("Report");
        }

        public Task<MemoryStream> CreateWorkbookAsync()
            => Task.FromResult(new MemoryStream());

        public async Task<int> WriteTextAsync(string text, int row)
        {
            _sheet.Cell(row, 1).Value = text;
            _sheet.Cell(row, 1).Style.Font.Bold = false;
            return await Task.FromResult(row + 1);
        }

        public async Task<int> WriteTableAsync(List<List<string>> data, int startRow)
        {
            int r = startRow;
            foreach (var row in data)
            {
                for (int c = 0; c < row.Count; c++)
                    _sheet.Cell(r, c + 1).Value = row[c];
                r++;
            }
            return await Task.FromResult(r + 1);
        }

        public async Task SaveAsync(MemoryStream stream, string path)
        {
            _workbook.SaveAs(path);
            await Task.CompletedTask;
        }
    }
}
