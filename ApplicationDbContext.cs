using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Data.SqlClient;
using Microsoft.EntityFrameworkCore;
using Trello.Models;

namespace Trello
{
    public class ApplicationDbContext : DbContext
    {
        public ApplicationDbContext(DbContextOptions<ApplicationDbContext> options) : base(options) { }
        public DbSet<ListModel>? Lists { get; set; }
        public DbSet<CardModel>? Cards { get; set; }
        public DbSet<Completed>? Completed { get; set; }

        public void CreateTable(string tableName) 
        {
            switch (tableName) 
            {
                case "Lists":
                    Database.ExecuteSqlRaw(@"
                    CREATE TABLE Lists (
                        Id NVARCHAR(100) PRIMARY KEY,
                        Name NVARCHAR(100) NOT NULL
                    )");
                    break;
                case "Cards":
                    Database.ExecuteSqlRaw(@"
                    CREATE TABLE Cards (
                        Id NVARCHAR(100) PRIMARY KEY,
                        Name NVARCHAR(100) NOT NULL,
                        Date DATE,
                        Weight INT,
                        IsComplete bit NOT NULL,
                        ListId NVARCHAR(100) NOT NULL,
                        FOREIGN KEY (ListId) REFERENCES Lists(Id)
                    )");
                    break;
                case "Completed":
                    Database.ExecuteSqlRaw(@"
                    CREATE TABLE Completed (
                        Date DATE PRIMARY KEY,
                        AllCompleted INT,
                        ShoperiaAllCompleted INT,
                        Shoperia_W1 INT,
                        Shoperia_W2 INT,
                        Shoperia_W3 INT,
                        Shoperia_UnWeighted INT,
                        XpressAllCompleted INT,
                        Xpress_W1 INT,
                        Xpress_W2 INT,
                        Xpress_W3 INT,
                        Xpress_UnWeighted INT,
                        Home12AllCompleted INT,
                        Home12_W1 INT,
                        Home12_W2 INT,
                        Home12_W3 INT,
                        Home12_UnWeighted INT,
                        MatebikeAllCompleted INT,
                        Matebike_W1 INT,
                        Matebike_W2 INT,
                        Matebike_W3 INT,
                        Matebike_UnWeighted INT
                    )");
                    break;
            }
        }
        public Completed AddDateToDB(int year, int month) 
        {
            Database.ExecuteSqlRaw("INSERT INTO Completed (Date, AllCompleted, " +
                                    "ShoperiaAllCompleted, Shoperia_W1, Shoperia_W2, Shoperia_W3, Shoperia_UnWeighted, " +
                                    "XpressAllCompleted, Xpress_W1, Xpress_W2, Xpress_W3, Xpress_UnWeighted, " +
                                    "Home12AllCompleted, Home12_W1, Home12_W2, Home12_W3, Home12_UnWeighted, " +
                                    "MatebikeAllCompleted, Matebike_W1, Matebike_W2, Matebike_W3, Matebike_UnWeighted)" +
                        "VALUES (@dateParameter, 0, " +
                        "0, 0, 0, 0, 0, " +
                        "0, 0, 0, 0, 0, " +
                        "0, 0, 0, 0, 0, " +
                        "0, 0, 0, 0, 0)",
                        new Microsoft.Data.SqlClient.SqlParameter("@dateParameter", new DateTime(year, month, 1)));

            var temp = Completed?.FirstOrDefault(e => e.Date == new DateTime(year, month, 1));
            if (temp == null)
                throw new Exception("AddDateToDB error");
            return temp;

        }

        public void DeleteCard(CardModel card) 
        {
            SqlParameter cardId = new SqlParameter("@cardId", card.Id);
            string sql = "DELETE FROM Cards WHERE Id = @cardId";
            Database.ExecuteSqlRaw(sql, cardId);
            //CACHE TÖRLÉSE
            Entry(card).State = EntityState.Detached;
        }
    }
}
