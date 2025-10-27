using HTML2EXCEL.Domain.Interfaces;
using HTML2EXCEL.Domain.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using HtmlAgilityPack;

namespace HTML2EXCEL.Infrastructure.Services
{
    public class HtmlParser : IHtmlParser
    {
        public async Task<List<TableData>> ParseTablesAsync(string htmlContent)
        {
            var tables = new List<TableData>();


            var htmlDoc = new HtmlDocument();
            htmlDoc.LoadHtml(htmlContent);


            var tableNodes = htmlDoc.DocumentNode.SelectNodes("//table");

            if (tableNodes == null || tableNodes.Count == 0)
                return tables;

            foreach (var table in tableNodes)
            {
                var tableData = new TableData();
                tableData.Name = GetTableName(table);

                var headerRow = table.SelectSingleNode(".//thead/tr") ?? table.SelectSingleNode(".//tr[1]");
                var headers = new List<string>();

                if (headerRow != null)
                {
                    var headerCells = headerRow.SelectNodes(".//th|.//td");
                    if (headerCells != null)
                    {
                        foreach (var cell in headerCells)
                            headers.Add(CleanText(cell.InnerText));
                    }
                }

                tableData.Headers = headers;

                var rows = new List<List<string>>();
                var rowNodes = table.SelectNodes(".//tbody/tr") ?? table.SelectNodes(".//tr[position()>1]");

                if (rowNodes != null)
                {
                    foreach (var row in rowNodes)
                    {
                        var cells = row.SelectNodes(".//td");
                        if (cells == null) continue;

                        var rowData = new List<string>();
                        foreach (var cell in cells)
                            rowData.Add(CleanText(cell.InnerText));

                        rows.Add(rowData);
                    }
                }

                tableData.Rows = rows;
                tables.Add(tableData);
            }

            return await Task.FromResult(tables);
        }

        private string GetTableName(HtmlNode tableNode)
        {

            var caption = tableNode.SelectSingleNode(".//caption");
            if (caption != null)
                return CleanText(caption.InnerText);

            var id = tableNode.GetAttributeValue("id", "");
            if (!string.IsNullOrEmpty(id))
                return id;

            var cls = tableNode.GetAttributeValue("class", "");
            if (!string.IsNullOrEmpty(cls))
                return cls;

            return "Table_" + Guid.NewGuid().ToString("N").Substring(0, 6);
        }

        private string CleanText(string text)
        {
            return HtmlEntity.DeEntitize(text)
                .Trim()
                .Replace("\n", " ")
                .Replace("\r", " ")
                .Replace("\t", " ");
        }
    }
}
