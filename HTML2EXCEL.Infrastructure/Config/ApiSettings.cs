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
        public string AuthEndpoint { get; set; } = string.Empty;
        public string ModelEndpoint { get; set; } = string.Empty;
        public string PathEndpoint { get; set; } = string.Empty;

        public int ControllerName { get; set; } = 20;
        public int RetryCount { get; set; } = 3;
        public int RetryDelaySeconds { get; set; } = 2;
    }
}
