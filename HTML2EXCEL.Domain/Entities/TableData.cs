using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HTML2EXCEL.Domain.Entities
{
    public class TableData
    {
        public string Name { get; set; } = string.Empty;
        public List<string> Headers { get; set; } = new();
        public List<List<string>> Rows { get; set; } = new();

        public TableData() { }

        public TableData(string name, List<string> headers, List<List<string>> rows)
        {
            Name = name;
            Headers = headers;
            Rows = rows;
        }

        public bool IsEmpty() => Rows.Count == 0 || Headers.Count == 0;
    }
}
