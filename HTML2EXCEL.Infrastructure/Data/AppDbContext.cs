using HTML2EXCEL.Domain.Entities;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HTML2EXCEL.Infrastructure.Data
{
    public class AppDbContext : DbContext
    {
        public DbSet<HtmlContentEntity> HtmlContents { get; set; }

        public AppDbContext(DbContextOptions<AppDbContext> options)
            : base(options)
        {
        }

        public override int SaveChanges() =>
            throw new InvalidOperationException("This context is read-only.");

        public override Task<int> SaveChangesAsync(CancellationToken cancellationToken = default) =>
            throw new InvalidOperationException("This context is read-only.");

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<HtmlContentEntity>()
                        .ToTable("FinancialStatements")          
                        .HasKey(x => x.Id);               

            modelBuilder.Entity<HtmlContentEntity>()
                        .Property(x => x.Html)
                        .HasColumnName("Pages")           
                        .IsRequired();
        }
    }
}
