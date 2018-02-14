namespace officedemo.Migrations
{
    using officedemo.Models;
    using System;
    using System.Data.Entity;
    using System.Data.Entity.Migrations;
    using System.Linq;

    internal sealed class Configuration : DbMigrationsConfiguration<officedemo.Models.officedemoContext>
    {
        public Configuration()
        {
            AutomaticMigrationsEnabled = false;
        }

        protected override void Seed(officedemo.Models.officedemoContext context)
        {
            context.Months.AddOrUpdate
                (m => m.Id,
                new Month { Id = 1, Name = "januari" },
                new Month { Id = 2, Name = "februari" },
                new Month { Id = 3, Name = "mars" },
                new Month { Id = 4, Name = "april" },
                new Month { Id = 5, Name = "maj" },
                new Month { Id = 6, Name = "juni" },
                new Month { Id = 7, Name = "juli" },
                new Month { Id = 8, Name = "augusti" },
                new Month { Id = 9, Name = "september" },
                new Month { Id = 10, Name = "oktober" },
                new Month { Id = 11, Name = "november" },
                new Month { Id = 12, Name = "december" });


            context.Resellers.AddOrUpdate(
                r => r.Id,
                new Reseller { Id = 1, Name = "Edsviken Bil AB" },
                new Reseller { Id = 2, Name = "Thorvalds Fordon" },
                new Reseller { Id = 3, Name = "FM Automobil" },
                new Reseller { Id = 4, Name = "Silverdalen AB" },
                new Reseller { Id = 5, Name = "Kista Limousiner" },
                new Reseller { Id = 6, Name = "Lennartssons Motor" },
                new Reseller { Id = 7, Name = "Marcus Racing" },
                new Reseller { Id = 8, Name = "Pettersson Bil AB" }
                );

            context.SaveChanges();
        }
    }
}
