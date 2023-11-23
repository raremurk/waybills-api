using Microsoft.EntityFrameworkCore;
using WaybillsAPI.Models;

namespace WaybillsAPI.Context
{
    public class WaybillsContext : DbContext
    {
        public DbSet<Transport> Transport { get; set; } = null!;
        public DbSet<Driver> Drivers { get; set; } = null!;
        public DbSet<Waybill> Waybills { get; set; } = null!;
        public DbSet<Operation> Operations { get; set; } = null!;
        public DbSet<Calculation> Calculations { get; set; } = null!;

        public WaybillsContext(DbContextOptions<WaybillsContext> options) : base(options)
        {
            Database.EnsureCreated();
        }

        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            optionsBuilder.UseSqlite("Filename=Waybills.db");
        }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<Transport>().HasData(Data.Transports());
            modelBuilder.Entity<Driver>().HasData(Data.Drivers());

            modelBuilder.Entity<Waybill>()
                .HasMany(j => j.Operations)
                .WithOne(j => j.Waybill)
                .OnDelete(DeleteBehavior.Cascade);

            modelBuilder.Entity<Waybill>()
                .HasMany(j => j.Calculations)
                .WithOne(j => j.Waybill)
                .OnDelete(DeleteBehavior.Cascade);

            modelBuilder.Entity<Waybill>()
                .HasOne(j => j.Driver)
                .WithMany(j => j.Waybills)
                .OnDelete(DeleteBehavior.SetNull);

            modelBuilder.Entity<Waybill>()
                .HasOne(j => j.Transport)
                .WithMany(j => j.Waybills)
                .OnDelete(DeleteBehavior.SetNull);

        }
    }
}
