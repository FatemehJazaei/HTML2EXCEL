using HTML2EXCEL.Domain.Entities;
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
    public class TableTemplateRepository : ITableTemplateRepository
    {
        private readonly AppDbContext _context;

        public TableTemplateRepository(AppDbContext context)
        {
            _context = context;
        }

        public async Task<(int Rows, int Cols)> GetRowANDColCountAsync(int id)
        {
            var entity = await _context.TableTemplates
                .AsNoTracking()
                .FirstOrDefaultAsync(x => x.Id == id);

            return (entity.Rows, entity.Cols);
        }
    }
}
