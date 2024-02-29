using Microsoft.EntityFrameworkCore;


namespace Template4338
{
    internal class DBcontext : DbContext
    {
        private const string ConnectionString = "" +
         "Data Source=(localdb)\\mssqllocaldb;" +
         "Initial Catalog=Librar;" +
         "Integrated Security=True;";

        public DbSet<Model> Users { get; set; }

        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            optionsBuilder.UseSqlServer(ConnectionString);
        }

        public void EnsureDatabaseCreated()
        {
            Database.EnsureCreated();
        }

    }
}
