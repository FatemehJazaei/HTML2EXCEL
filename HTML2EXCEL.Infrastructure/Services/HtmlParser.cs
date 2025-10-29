using HTML2EXCEL.Domain.Interfaces;
using HTML2EXCEL.Domain.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using HtmlAgilityPack;
using ClosedXML.Excel;
using System.Text.RegularExpressions;

namespace HTML2EXCEL.Infrastructure.Services
{
    public class HtmlParser : IHtmlParser
    {
        private static readonly Regex PxRegex = new(@"(\d+)", RegexOptions.Compiled);
        private static readonly Regex RgbRegex = new(@"(\d+)", RegexOptions.Compiled);

        public async Task<MemoryStream> ParseTablesAsync(string htmlContent, string token, IApiService _apiService,
        IExcelExporter excelExporter)
        {
            var doc = new HtmlDocument();
            doc.LoadHtml(htmlContent);

            using var workbook = new XLWorkbook();

            var pages = doc.DocumentNode.SelectNodes("//div[contains(@class, 'page') and contains(@class, 'tx-frame')]");
            if (pages == null || pages.Count == 0)
                throw new Exception("No pages found in HTML.");

            int pageIndex = 1;
            foreach (var page in pages)
            {
                var sheet = workbook.Worksheets.Add($"صفحه {pageIndex}");
                sheet.RightToLeft = true;

                ReadDOM(sheet, page, token, _apiService);
                pageIndex++;
            }

            var stream = new MemoryStream();
            workbook.SaveAs(stream);
            stream.Position = 0;
            return await Task.FromResult(stream);
        }

        private async void ReadDOM(IXLWorksheet sheet, HtmlNode page, string token, IApiService _apiService)
        {
            int currentRow = 1;
            double baseTop = 0;

            foreach (var div in page.SelectNodes(".//div") ?? Enumerable.Empty<HtmlNode>())
            {
                var styleDiv = ParseStyle(GetStyle(div));

                if (styleDiv.TryGetValue("padding-top", out var paddingTop) && paddingTop.Contains("mm"))
                {
                    baseTop += double.Parse(paddingTop.Replace("mm", "").Trim());
                    currentRow += MmToRow(baseTop);
                }

                foreach (var element in div.ChildNodes.Where(n => n.NodeType == HtmlNodeType.Element))
                {
                    if (element.Name == "table")
                    {
                        var tableNode = element.SelectSingleNode(".//div[contains(@class, 'table-templateId')]");
                        if (tableNode != null)
                        {
                            var idValue = tableNode.GetAttributeValue("data-id", null);
                            if (int.TryParse(idValue, out int tableTemplateId))
                            {
                                var model = await _apiService.GetModelAsync(token, tableTemplateId);
                                var filePath = await _apiService.GetFilePathAsync(token, model);
                                var excelBytes = await _apiService.DownloadExcelFileAsync(token, filePath);
                                var tableData = await ExtractTopLeftCells(excelBytes);
                                WriteTableToExcel(tableData, sheet, ref currentRow);
                            }
                        }

                    }
                    else if (element.Name == "p")
                    {
                        var text = string.Join("", element.Descendants("span")
                            .Select(s => s.InnerText.Trim()));

                        var styleSpan = ParseStyle(GetStyle(element.Descendants("span").FirstOrDefault()));
                        var styleP = ParseStyle(GetStyle(element));
                        var combinedStyle = styleDiv
                            .Concat(styleP)
                            .Concat(styleSpan)
                            .ToDictionary(k => k.Key, v => v.Value);

                        WriteExcel(combinedStyle, text, sheet, ref currentRow);
                    }
                    currentRow++;
                }

                currentRow++;
            }
        }

        private void WriteExcel(Dictionary<string, string> style, string text, IXLWorksheet sheet, ref int currentRow)
        {
            if (string.IsNullOrWhiteSpace(text))
                return;

            int col = MmToCol(ParseMm(style, "left") + ParseMm(style, "margin-left") + ParseMm(style, "padding-left"));

            string color = FixColor(style.TryGetValue("color", out var c) ? c : null);
            double fontSize = PxToPt(ExtractNumber(style.TryGetValue("font-size", out var fs) ? fs : "12px"));
            bool bold = style.TryGetValue("font-weight", out var fw) && fw.Contains("bold", StringComparison.OrdinalIgnoreCase);
            string align = style.TryGetValue("text-align", out var a) ? a.ToLower() : "right";

            var cell = sheet.Cell(currentRow, col);
            cell.Value = text;
            cell.Style.Font.FontName = "B Nazanin";
            cell.Style.Font.FontSize = fontSize;
            cell.Style.Font.Bold = bold;
            cell.Style.Font.FontColor = XLColor.FromHtml($"#{color.Substring(2)}");

            cell.Style.Alignment.Horizontal = align == "center"
                ? XLAlignmentHorizontalValues.Center
                : XLAlignmentHorizontalValues.Right;
            cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            cell.Style.Alignment.WrapText = true;

            currentRow++;
        }

        // --- Helper methods ---

        private static string GetStyle(HtmlNode node)
            => node?.GetAttributeValue("style", "") ?? "";

        private static Dictionary<string, string> ParseStyle(string styleStr)
            => styleStr?
                .Split(';', StringSplitOptions.RemoveEmptyEntries)
                .Select(part => part.Split(':'))
                .Where(kv => kv.Length == 2)
                .ToDictionary(kv => kv[0].Trim().ToLower(), kv => kv[1].Trim())
                ?? new Dictionary<string, string>();

        private static double PxToPt(double px) => px * 0.75;
        private static double ParseMm(Dictionary<string, string> styles, string key)
            => styles.TryGetValue(key, out var v) && v.Contains("mm") ? double.Parse(v.Replace("mm", "")) : 0;
        private static int MmToRow(double mm) => (int)(mm / 7.5) + 1;
        private static int MmToCol(double mm) => (int)(mm / 10.5) + 1;
        private static double ExtractNumber(string input)
            => double.TryParse(Regex.Match(input, @"\d+").Value, out var num) ? num : 0;

        private static string FixColor(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return "FF000000";
            value = value.Trim().ToLower();

            if (value.StartsWith("rgb"))
            {
                var nums = Regex.Matches(value, @"\d+").Select(m => int.Parse(m.Value)).ToArray();
                if (nums.Length == 3)
                    return $"FF{nums[0]:X2}{nums[1]:X2}{nums[2]:X2}";
            }
            if (value.StartsWith("#"))
            {
                var val = value[1..];
                if (val.Length == 3)
                    val = string.Concat(val.Select(c => $"{c}{c}"));
                if (val.Length == 6)
                    return $"FF{val.ToUpper()}";
            }
            return "FF000000";
        }


        private async Task<List<List<string>>> ExtractTopLeftCells(byte[] excelBytes)
        {
            using var ms = new MemoryStream(excelBytes);
            using var workbook = new XLWorkbook(ms);
            var ws = workbook.Worksheets.First();

            var result = new List<List<string>>();
            for (int r = 1; r <= 10; r++)
            {
                var row = new List<string>();
                for (int c = 1; c <= 10; c++)
                {
                    row.Add(ws.Cell(r, c).GetValue<string>());
                }
                result.Add(row);
            }
            return result;
        }

        private void WriteTableToExcel(List<List<string>> table, IXLWorksheet sheet, ref int currentRow)
        {
            foreach (var row in table)
            {
                int currentCol = 1;
                foreach (var cellValue in row)
                {
                    sheet.Cell(currentRow, currentCol).Value = cellValue;
                    currentCol++;
                }
                currentRow++;
            }
        }


    }

}
