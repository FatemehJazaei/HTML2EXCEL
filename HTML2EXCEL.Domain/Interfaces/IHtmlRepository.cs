using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HTML2EXCEL.Domain.Interfaces
{
    public interface IHtmlRepository
    {
        Task<string?> GetHtmlContentAsync(int id);
    }
}
