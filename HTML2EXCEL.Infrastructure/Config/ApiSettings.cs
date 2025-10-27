using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HTML2EXCEL.Infrastructure.Config
{
    public class ApiSettings
    {
        public string BaseUrl { get; set; } = string.Empty;
        public string AuthEndpoint { get; set; } = "/auth/token";
        public string DataKeyEndpoint { get; set; } = "/api/data/key";
        public string ExcelUrlEndpoint { get; set; } = "/api/data/excel";
        public int RetryCount { get; set; } = 3;
        public int RetryDelaySeconds { get; set; } = 2;
    }
}
