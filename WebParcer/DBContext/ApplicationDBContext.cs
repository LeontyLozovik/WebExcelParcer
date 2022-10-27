using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Internal;
using WebParcer.Models.TableModels;

namespace WebParcer.DBContext
{
    public class ApplicationDBContext : DbContext   //Структура базы данных
    {
        public ApplicationDBContext(DbContextOptions contextOptions)
            : base(contextOptions)
        {
            Database.EnsureDeleted();
            Database.EnsureCreated();
        }
        public DbSet<Class1> Class1s { get; set; }
        public DbSet<Class2> Class2s { get; set; }
        public DbSet<Class3> Class3s { get; set; }
        public DbSet<Class4> Class4s { get; set; }
        public DbSet<Class5> Class5s { get; set; }
        public DbSet<Class6> Class6s { get; set; }
        public DbSet<Class7> Class7s { get; set; }
        public DbSet<Class8> Class8s { get; set; }
        public DbSet<Class9> Class9s { get; set; }
    }
}
