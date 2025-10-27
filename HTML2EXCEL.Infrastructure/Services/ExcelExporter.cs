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
        public async Task ExportAsync(List<TableData> tables, string outputPath)
        {
            if (tables == null || tables.Count == 0)
                throw new ArgumentException("No table data provided for export.");

            var directory = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                Directory.CreateDirectory(directory);

            using var workbook = new XLWorkbook();

            foreach (var table in tables)
            {
                var sheetName = GetValidSheetName(table.Name);
                var worksheet = workbook.Worksheets.Add(sheetName);

                int currentRow = 1;

                if (table.Headers != null && table.Headers.Count > 0)
                {
                    for (int i = 0; i < table.Headers.Count; i++)
                    {
                        worksheet.Cell(currentRow, i + 1).Value = table.Headers[i];
                        worksheet.Cell(currentRow, i + 1).Style.Font.Bold = true;
                        worksheet.Cell(currentRow, i + 1).Style.Fill.BackgroundColor = XLColor.LightGray;
                    }
                    currentRow++;
                }

                if (table.Rows != null)
                {
                    foreach (var row in table.Rows)
                    {
                        for (int i = 0; i < row.Count; i++)
                        {
                            worksheet.Cell(currentRow, i + 1).Value = row[i];
                        }
                        currentRow++;
                    }
                }

                worksheet.Columns().AdjustToContents();
            }

            using var stream = new FileStream(outputPath, FileMode.Create, FileAccess.Write, FileShare.None, 4096, useAsync: true);
            workbook.SaveAs(stream);

            await stream.FlushAsync();
        }

        private string GetValidSheetName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
                name = "Sheet";

            name = name.Length > 31 ? name.Substring(0, 31) : name;


            foreach (var invalidChar in Path.GetInvalidFileNameChars())
                name = name.Replace(invalidChar.ToString(), "_");

            return name;
        }
    }
}
