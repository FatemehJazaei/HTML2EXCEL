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
using Microsoft.IdentityModel.Tokens;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Wordprocessing;
using Azure;
using DocumentFormat.OpenXml.Office2010.Excel;
using HTML2EXCEL.Infrastructure.Repositories;
using DocumentFormat.OpenXml.Spreadsheet;

namespace HTML2EXCEL.Infrastructure.Services
{
    public class HtmlParser : IHtmlParser
    {
        private static readonly Regex PxRegex = new(@"(\d+)", RegexOptions.Compiled);
        private static readonly Regex RgbRegex = new(@"(\d+)", RegexOptions.Compiled);

        public async Task<MemoryStream> ParseTablesAsync(string htmlContent, string token, IApiService _apiService,
        IExcelExporter excelExporter, ITableTemplateRepository _tableTemplateRepository)
        {
            var doc = new HtmlDocument();
            doc.LoadHtml(htmlContent);

            var workbook = excelExporter.GetWorkbook();

            var pages = doc.DocumentNode.SelectNodes(
                "/html/body/div[contains(concat(' ', normalize-space(@class), ' '), ' page ') " +
                "and contains(concat(' ', normalize-space(@class), ' '), ' tx-frame ')]"
            ).ToList() ?? Enumerable.Empty<HtmlNode>().ToList();

            if (pages == null || pages.Count == 0)
                throw new Exception("No pages found in HTML.");

            int pageIndex = 1;
            foreach (var page in pages)
            {
                Console.WriteLine(pageIndex);
                var sheet = workbook.Worksheets.Add($"صفحه {pageIndex}");
                sheet.RightToLeft = true;

                await ReadDOM(sheet, page, token, _apiService, _tableTemplateRepository);
                pageIndex++;
            }

            var stream = new MemoryStream();
            workbook.SaveAs(stream);
            stream.Position = 0;
            return await Task.FromResult(stream);
        }
        private async Task ReadDOM(IXLWorksheet sheet, HtmlNode page, string token, IApiService _apiService, ITableTemplateRepository _tableTemplateRepository)
        {
            int currentRow = 1;
            var styleDiv = ParseStyle(GetStyle(page));
            await ProcessDiv(page, sheet, currentRow, token, _apiService, styleDiv, _tableTemplateRepository);
        }

        private async Task ProcessDiv(HtmlNode div, IXLWorksheet sheet, int currentRow, string token, IApiService _apiService, Dictionary<string, string> parentStyle, ITableTemplateRepository _tableTemplateRepository)
        {
            var styleDiv = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var childStyle = ParseStyle(GetStyle(div));
            if (parentStyle != null)
                foreach (var kv in parentStyle)
                    styleDiv[kv.Key] = kv.Value;

            if (childStyle != null)
                foreach (var kv in childStyle)
                    styleDiv[kv.Key] = kv.Value;

            foreach (var element in div.ChildNodes.Where(n => n.NodeType == HtmlNodeType.Element))
            {
                
                if (element.GetAttributeValue("class", "").Contains("table-templateId"))
                {
                    var tableNode = element;
                    var idValue = tableNode.GetAttributeValue("data-id", null);
                    if (int.TryParse(idValue, out int tableTemplateId))
                    {
                        var model = await _apiService.GetModelAsync(token, tableTemplateId);
                        var filePath = await _apiService.GetFilePathAsync(token, model);
                        var (rows, cols) = await _tableTemplateRepository.GetRowANDColCountAsync(tableTemplateId);
                        var excelBytes = await _apiService.DownloadExcelFileAsync(token, filePath);
                        var tableData = await ExtractTopLeftCells(excelBytes, rows, cols);
                        WriteTableToExcel(tableData, sheet, ref currentRow);
                    }
                }

                else if (element.Name == "p")
                {
                    var combined = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    var styleP = ParseStyle(GetStyle(element));

                    if (styleDiv != null)
                        foreach (var kv in styleDiv)
                            combined[kv.Key] = kv.Value;

                    if (styleP != null)
                        foreach (var kv in styleP)
                            combined[kv.Key] = kv.Value;

                    foreach (var span in element.Descendants("span").Where(s => s.ParentNode.Name != "span"))
                    {
                        ProcessSpan(span, sheet, ref currentRow, combined);
                    }


                    currentRow++;
                }

                else if (element.Name == "div")
                {
                    await ProcessDiv(element, sheet, currentRow, token, _apiService, styleDiv, _tableTemplateRepository);
                }

            }
        }

        /*
        private async Task ReadDOM(IXLWorksheet sheet, HtmlNode page, string token, IApiService _apiService)
        {
            int currentRow = 1;

            foreach (var div in page.SelectNodes(".//div") ?? Enumerable.Empty<HtmlNode>())
            {
                var styleDiv = ParseStyle(GetStyle(div));
                foreach (var element in div.ChildNodes.Where(n => n.NodeType == HtmlNodeType.Element))
                {
                    if (element.GetAttributeValue("class", "").Contains("table-templateId"))
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
                        var combined = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                        var styleP = ParseStyle(GetStyle(element));

                        if (styleDiv != null)
                            foreach (var kv in styleDiv)
                                combined[kv.Key] = kv.Value;

                        if (styleP != null)
                            foreach (var kv in styleP)
                                combined[kv.Key] = kv.Value;

                        foreach (var span in element.Descendants("span").Where(s => s.ParentNode.Name != "span"))
                        {
                            ProcessSpan(span, sheet, ref currentRow, combined);
                        }
                    }
                    currentRow++;
                }
                
            }
            
        }
        */

        private void WriteExcel(Dictionary<string, string> style, string text, IXLWorksheet sheet, ref int currentRow)
        {
            if (string.IsNullOrWhiteSpace(text))
                return;

            int col = 1;
            // int col = MmToCol(ParseMm(style, "left") + ParseMm(style, "margin-left") + ParseMm(style, "padding-left"));

            string color = FixColor(style.TryGetValue("color", out var c) ? c : null);
            double fontSize = PxToPt(ExtractNumber(style.TryGetValue("font-size", out var fs) ? fs : "12px"));
            bool bold = style.TryGetValue("font-weight", out var fw) && fw.Contains("bold", StringComparison.OrdinalIgnoreCase);
            string align = style.TryGetValue("text-align", out var a) ? a.ToLower() : "right";


            
            var range = sheet.Range(currentRow, 1, currentRow, 9);
            range.Merge();

            var cell = sheet.Cell(currentRow, col);

            string existingText = cell.Value.ToString() ?? string.Empty;
            if (!string.IsNullOrEmpty(existingText))
                text = existingText + " " + text;

            cell.Value = text;

            cell.Style.Font.FontName = "B Nazanin";
            cell.Style.Font.FontSize = fontSize;
            cell.Style.Font.Bold = bold;
            cell.Style.Font.FontColor = XLColor.FromHtml($"#{color.Substring(2)}");


            cell.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
            cell.Style.Alignment.WrapText = true;

            if (align == "center")
                cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            else
                cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;

            double dynamicHeight = 20;

            if (Math.Ceiling(text.Length / 100.0) > 1) dynamicHeight = Math.Ceiling(text.Length / 100.0) * 20;

            sheet.Row(currentRow).Height = dynamicHeight;
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


        private async Task<List<List<string>>> ExtractTopLeftCells(byte[] excelBytes, int rows, int cols)
        {
            using var ms = new MemoryStream(excelBytes);
            using var workbook = new XLWorkbook(ms);
            var ws = workbook.Worksheets.First();

            var result = new List<List<string>>();
            for (int r = 1; r <= rows; r++)
            {
                var row = new List<string>();
                for (int c = 1; c <= cols; c++)
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

        void ProcessSpan(HtmlNode span, IXLWorksheet sheet, ref int currentRow, Dictionary<string, string> parentStyle)
        {
            var styleSpan = ParseStyle(GetStyle(span));
            var combined = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            if (parentStyle != null)
                foreach (var kv in parentStyle)
                    combined[kv.Key] = kv.Value;

            if (styleSpan != null)
                foreach (var kv in styleSpan)
                    combined[kv.Key] = kv.Value;

            foreach (var node in span.ChildNodes)
            {
                if (node.NodeType == HtmlNodeType.Text)
                {
                    var text = System.Net.WebUtility.HtmlDecode(node.InnerText.Trim());
                    text = text.Replace("\u00A0", "").Trim();

                    if (!string.IsNullOrEmpty(text) && !string.IsNullOrWhiteSpace(text))
                        WriteExcel(combined, text, sheet, ref currentRow);
                }
                else if (node.Name == "br")
                {
                    currentRow++;
                }
                else if (node.Name == "span"){
                    ProcessSpan(node, sheet, ref currentRow, combined);
                }
            }

            /*
              
            if (span.Descendants("span").Any())
            {
                foreach (var innerSpan in span.ChildNodes.Where(n => n.Name == "span"))
                {
                    ProcessSpan(innerSpan, sheet, ref currentRow, combined);
                }
            }
            else
            {
                var text = string.Join("", span.ChildNodes
                    .Where(n => n.NodeType == HtmlNodeType.Text || n.Name == "br")
                    .Select(n =>
                        n.Name == "br" ? "\n" : n.InnerText.Trim()
                    ));

                if (!string.IsNullOrEmpty(text))
                {
                    text = System.Net.WebUtility.HtmlDecode(text);
                    WriteExcel(combined, text, sheet, ref currentRow);
                }
            }
            */
        }
    }

}
