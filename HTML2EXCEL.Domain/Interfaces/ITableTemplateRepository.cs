using HTML2EXCEL.Domain.Entities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HTML2EXCEL.Domain.Interfaces
{
    public interface ITableTemplateRepository
    {
        Task<(int Rows, int Cols)> GetRowANDColCountAsync(int id);
    }

}
