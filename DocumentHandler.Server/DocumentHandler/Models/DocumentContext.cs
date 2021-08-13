using DocumentHandler.Models;
using Microsoft.EntityFrameworkCore;

namespace DocumentHandler
{
    public class DocumentContext : DbContext
    {
        public DbSet<Document> Documents { get; set; }

        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            optionsBuilder.UseSqlServer("Server=.\\SQLEXPRESS;Database=BFDdb;Trusted_Connection=True;");
        }
    }
}