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
        public string OutputPath { get; set; } = "output.xlsx";
        public string Username { get; set; } = string.Empty;
        public string Password { get; set; } = string.Empty;
        public int CompanyId { get; set; } = 1;
        public int PeriodId { get; set; } = 1;
        public int TableTemplateId { get; set; }
    }
}
