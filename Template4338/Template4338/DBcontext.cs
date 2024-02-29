using Microsoft.EntityFrameworkCore;
using System;

namespace Template4338
{
    internal class DBcontext : DbContext
    {
        private const string ConnectionString =
            "Data Source=(localdb)\\mssqllocaldb;" +
            "Initial Catalog=LLLL;" +
            "Integrated Security=True;";

        public DbSet<Model> Users { get; set; }

        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            if (!optionsBuilder.IsConfigured)
            {
                optionsBuilder.UseSqlServer(ConnectionString);
            }
        }

        public void EnsureDatabaseCreated()
        {
            try
            {
                Database.EnsureCreated();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при создании базы данных: {ex.Message}");
            }
        }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<Model>().HasKey(e => e.Id);
        }
    }
}
