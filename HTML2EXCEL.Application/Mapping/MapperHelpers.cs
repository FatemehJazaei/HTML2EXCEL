using HTML2EXCEL.Domain.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HTML2EXCEL.Application.Mapping
{
    public static class MapperHelpers
    {
        public static string ToStringRepresentation(this TableData table)
        {
            var headers = string.Join("\t", table.Headers);
            var rows = string.Join("\n", table.Rows.Select(r => string.Join("\t", r)));
            return $"{table.Name}\n{headers}\n{rows}";
        }
    }
}
