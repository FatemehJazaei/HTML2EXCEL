using HTML2EXCEL.Domain.Interfaces;
using HTML2EXCEL.Infrastructure.Data;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HTML2EXCEL.Infrastructure.Repositories
{
    public class HtmlRepository : IHtmlRepository
    {
        private readonly AppDbContext _context;

        public HtmlRepository(AppDbContext context)
        {
            _context = context;
        }

        public async Task<string?> GetHtmlContentAsync(int id)
        {
            var entity = await _context.HtmlContents
                .AsNoTracking()
                .FirstOrDefaultAsync(x => x.Id == id);

            return entity?.Html;
        }
    }
}
