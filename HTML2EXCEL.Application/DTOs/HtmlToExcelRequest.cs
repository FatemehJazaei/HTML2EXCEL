using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HTML2EXCEL.Application.DTOs
{
    public class HtmlToExcelRequest
    {
        public string HtmlContent { get; set; } = string.Empty;
        public string OutputFilePath { get; set; } = string.Empty;
        public string Token { get; set; } = string.Empty;
    }
}
