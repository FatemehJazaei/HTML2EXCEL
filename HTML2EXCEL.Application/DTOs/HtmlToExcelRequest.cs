using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HTML2EXCEL.Application.DTOs
{
    public class HtmlToExcelRequest
    {
        public string OutputPath { get; set; } 
        public string Username { get; set; } 
        public string Password { get; set; } 
        public int CompanyId { get; set; }
        public int PeriodId { get; set; }
    }
}
