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
        IExcelExporter excelExporter)
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
                int percent = (int)((double)pageIndex / pages.Count * 100);
                Console.Clear();
                Console.WriteLine(percent + "%");

                var sheet = workbook.Worksheets.Add($"صفحه {pageIndex}");
                sheet.RightToLeft = true;

                await ReadDOM(sheet, page, token, _apiService);
                pageIndex++;
            }

            var stream = new MemoryStream();
            workbook.SaveAs(stream);
            stream.Position = 0;
            return await Task.FromResult(stream);
        }
        private async Task ReadDOM(IXLWorksheet sheet, HtmlNode page, string token, IApiService _apiService)
        {
            int currentRow = 1;
            var styleDiv = ParseStyle(GetStyle(page));
            _ = await ProcessDiv(page, sheet, currentRow, token, _apiService, styleDiv);
        }

        private async Task<int> ProcessDiv(HtmlNode div, IXLWorksheet sheet, int currentRow, string token, IApiService _apiService, Dictionary<string, string> parentStyle)
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
                        var excelBytes = await _apiService.DownloadExcelFileAsync(token, filePath);
                        WriteTableToExcel(excelBytes, sheet, ref currentRow );
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
                    /*
                    foreach (var span in element.Descendants("span").Where(s => s.ParentNode.Name != "span" && s.ParentNode.Name != "u"))
                    {
                        ProcessSpan(span, sheet, ref currentRow, combined);
                    }
                    */
                    foreach (var child in element.ChildNodes)
                    {
                        if (child.Name == "span" || child.Name == "u")
                        {
                            ProcessSpan(child, sheet, ref currentRow, combined);
                        }
                        
                    }
                    currentRow += 1;
                }

                else if (element.Name == "div")
                {
                    currentRow = await ProcessDiv(element, sheet, currentRow, token, _apiService, styleDiv);
                }

            }
            return currentRow;
        }  

        private void WriteExcel(Dictionary<string, string> style, string text, IXLWorksheet sheet, ref int currentRow)
        {
            if (string.IsNullOrWhiteSpace(text))
                return;

            int col = 1;

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

        private void WriteTableToExcel(byte[] excelBytes, IXLWorksheet targetSheet, ref int currentRow)
        {
            
            using var ms = new MemoryStream(excelBytes);
            using var workbook = new XLWorkbook(ms);
            var sourceSheet = workbook.Worksheets.First();

            int rows = sourceSheet.LastRowUsed().RowNumber();
            int cols = sourceSheet.LastColumnUsed().ColumnNumber();

            var sourceRange = sourceSheet.Range(1, 1, cols, rows);
            var targetCell = targetSheet.Cell(currentRow, 1);
            sourceRange.CopyTo(targetCell);

            int destRowStart = currentRow;
            int destColStart = 1;

            for (int r = 0; r < rows; r++)
            {

                targetSheet.Row(destRowStart + r).Height = 25;
            }

            for (int c = 0; c < cols; c++)
            {
                double width = sourceSheet.Column(c + 1).Width;
                targetSheet.Column(destColStart + c).Width = width;
            }

            currentRow += rows;
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
                   // currentRow++;
                }
                else if (node.Name == "span" || node.Name == "u")
                {
                    ProcessSpan(node, sheet, ref currentRow, combined);
                }
            }

        }
    }

}
