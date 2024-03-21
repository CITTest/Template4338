using System.Data.Entity;

namespace Template4338
{
    public class MyDbContext : DbContext
    {
        public DbSet<Table> Tables { get; set; }
        public DbSet<TableJSON> TablesJSON { get; set; }
    }

    public class Table
    {
        public int Id { get; set; }
        public string FullName { get; set; }
        public string ClientId { get; set; }
        public string BirthDate { get; set; }
        public string Index { get; set; }
        public string City { get; set; }
        public string Street { get; set; }
        public string Home { get; set; }
        public string Apartment { get; set; }
        public string Email { get; set; }
    }

    public class TableJSON
    {
        public int Id { get; set; }
        public string FullName { get; set; }
        public string CodeClient { get; set; }
        public string BirthDate { get; set; }
        public string Index { get; set; }
        public string City { get; set; }
        public string Street { get; set; }
        public int Home { get; set; }
        public int Kvartira { get; set; }
        public string E_mail { get; set; }

        public TableJSON()
        {
            
        }
        public TableJSON(int id, string fullname, string codeclient, string birthdate, string index, string city, string street, int home, int kvartira, string e_mail)
        {
            Id = id;
            FullName = fullname;
            CodeClient = codeclient;
            BirthDate = birthdate;
            Index = index;
            City = city;
            Street = street;
            Home = home;
            Kvartira = kvartira;
            E_mail = e_mail;

        }
    }
    public class StreetTable
    {
        public string Street { get; set; }
    }
}
