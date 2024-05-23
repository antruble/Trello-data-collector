using Manatee.Trello;
using Microsoft.EntityFrameworkCore;
using System;
using System.Data;
using System.Linq.Expressions;
using Trello;
using Trello.Models;

var settings = Settings.LoadSettings();


if (settings.BoardId == null || settings.ConnectionString == null)
{
    throw new Exception("settings can't be null");
}
TrelloAuthorization.Default.AppKey = settings.ApiKey;
TrelloAuthorization.Default.UserToken = settings.UserSecret;

// LISTÁK ELLENŐRZÉSE
await CheckLists(settings.BoardId, settings.ConnectionString);

//EXCELBE ÍRÁS
Excel.WriteToExcel();


/// <summary>
/// Ellenőrzi és szinkronizálja a Trello listákat és kártyákat az SQL adatbázissal.
/// </summary>
/// <param name="boardId">A Trello board id-ja.</param>
/// <param name="connectionstring">Az SQL adatbázis connection stringje.</param>
static async Task CheckLists(string boardId, string connectionstring)
{
    // <trello beállítása>
    ITrelloFactory factory = new TrelloFactory();
    var board = factory.Board(boardId);
    await board.Lists.Refresh();
    var trelloLists = board.Lists;
    // </trello beállítása>

    // SQL KAPCSOLAT LÉTREHOZÁSA
    var dbOptions = new DbContextOptionsBuilder<ApplicationDbContext>().UseSqlServer(connectionstring).Options;
    using (var dbContext = new ApplicationDbContext(dbOptions))
    {
        try
        {
            // ADATBÁZIS TÁBLÁK ELLENŐRZÉSE, HA VALAMELYIK HIÁNYZIK -> LÉTREHOZÁS
            CheckDatabaseTables(dbContext);

            // ADATBÁZISBAN LÉTEZŐ LISTÁK KIGYŰJTÉSE A KÁRTYÁIKKAL
            var dbLists = dbContext.Lists?.Include(a => a.Cards).ToList();

            // EGYESÉVEL MINDEN TRELLO LISTA ELLENŐRZÉSE
            foreach (var trelloList in trelloLists)
            {
                await trelloList.Refresh();
                // ADOT TRELLO LISTA MEGKERESÉSE A DB-BEN
                var dbList = dbLists?.Find(l => l.Id?.Trim() == trelloList.Id.Trim());
                // NEM SZEREPEL A DB-BEN -> LÉTRE KELL HOZNI
                if (dbList == null)
                {
                    dbList = new ListModel
                    {
                        Name = trelloList.Name,
                        Id = trelloList.Id,
                        Cards = new List<CardModel>(),
                    };
                    dbContext.Lists?.Add(dbList);
                    dbContext.SaveChanges();
                }

                var trelloCards = trelloList.Cards;
                // A LISTÁBAN SZEREPLŐ KÁRTYÁK VIZSGÁLATA EGYESÉVEL
                foreach (var trelloCard in trelloCards)
                {
                    // HA ARCHIVÁLVA VAN NEM VIZSGÁLJUK
                    if (trelloCard.IsArchived == true)
                        break;

                    // A DBLIST A JELENLEG VIZSGÁLT TRELLO LISTA ELTÁROLVA AZ ADATBÁZISBAN. MEGPRÓBÁLJUK MEGKERESNI AZ ADATBÁZIS LISTÁBAN A VIZSGÁLT TRELLO KÁRTYÁT
                    var dbCard = dbList.Cards?.FirstOrDefault(c => c.Id?.Trim() == trelloCard.Id.Trim());
                    // A TRELLÓBÓL LEKÉRT TASK LISTÁJÁT MEGVIZSGÁLJUK AZ ADATBÁZISBAN, ÉS HA NEM TALÁLJUK BENNE A VIZSGÁLT KÁRTYÁT ->
                    // ( A ) ÁTKERÜLT MÁSIK LISTÁBA
                    // ( B ) MÉG NEM LÉTEZIK
                    if (dbCard == null)
                    {
                        DateTime? oldDate = null;
                        // TRELLOKÁRTYA MEGKERESÉSE A DB ÖSSZES KÁRTYÁI KÖZÜL
                        var cardInDB = dbContext.Cards?.FirstOrDefault(c => c.Id.Trim() == trelloCard.Id.Trim());
                        // HA A cardInDB NEM NULL, AKKOR A KÁRTYA MÁR RÖGZÍTVE VAN -> ÚJ LISTÁBA KERÜLT -> FRISSÍTENI KELL
                        if (cardInDB != null)
                        {
                            // HA A KÁRTYA EL VOLT FOGADVA, ÉS ÚGY LETT ELMOZGATVA
                            // ˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇ
                            // (1) AZ ELŐZŐ LISTA SZÁMLÁLÓJÁBÓL KI KELL SZEDNI
                            // (2) AZ ELFOGADÁS DÁTUMÁT EL KELL TÁROLNI, AMIT MÁR CSAK AZ ADATBÁZIS TÁROL
                            //     (A MÁSIK LISTÁBAN EZT A HÓNAPOT KELL FRISSÍTENI)
                            // ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
                            if (cardInDB.IsComplete == true) // (1) HA A TASK AZ ADATBÁZISBAN EL VAN FOGADVA, AKKOR KI KELL TÖRÖLNI A RÉGI LISTÁJÁHOZ TARTOZÓ BOLT SZÁMLÁLÓJÁBÓL
                                Utilities.UpdateTes(dbContext, cardInDB, "REMOVE");

                            // NEM SZERETNÉNK FRISSÍTENI A TASKHOZ TARTOZÓ DÁTUMOT, KIVÉVE, HA MOZGATÁS UTÁN LETT ELFOGADVA (A NULL ÉRTÉK JELZI, HOGY DÁTUMOT KELL MAJD ÁLLÍTANUNK)
                            oldDate = cardInDB.IsComplete == false && trelloCard.IsComplete == true ? null : cardInDB.Date;

                            // A TASK KITÖRLÉSE AZ ADATBÁZISBÓL, MIVEL ELAVULT LISTÁBAN VAN
                            dbContext.DeleteCard(cardInDB);
                            dbContext.SaveChanges();
                            // MIUTÁN TÖRÖLTÜK AZ ADATBÁZISBÓL AZ ELAVULT ADATOT, EZÉRT MOSTMÁR AKÁR ÚJ A KÁRTYA, AKÁR MÁR LÉTEZETT, UGYANÚGY BÁNUNK VELE
                        }
                        // SÚLYOZÁS KISZÁMOLÁSA
                        int weight;
                        if (trelloCard.Labels == null)
                            weight = 0;
                        else
                        {
                            List<string> labelIDs = new List<string>();
                            foreach (var label in trelloCard.Labels)
                                labelIDs.Add(label.Id);

                            weight = Utilities.GetWeightFromLabels(labelIDs);
                        };
                        // TRELLOHOZ SZÜKSÉGES MODEL LÉTREHOZÁSA
                        var newCard = new CardModel
                        {
                            Id = trelloCard.Id,
                            Date = oldDate != null ? (DateTime) oldDate : (DateTime) trelloCard.LastActivity,
                            Name = trelloCard.Name,
                            Weight = weight,
                            IsComplete = trelloCard.IsComplete,
                            List = dbList,
                            ListId = trelloList.Id,
                        };
                        // A KÁRTYA MÁR EL LETT FOGADVA MIELŐTT BEKERÜLT VOLNA A DB-BE -> DOKUMENTÁLNI KELL
                        if (newCard.IsComplete == true)
                            Utilities.UpdateTes(dbContext, newCard);
                        dbList.Cards?.Add(newCard);
                        dbContext.SaveChanges();
                    }
                    // LÉTEZIK A KÁRTYA ÉS JÓ LISTÁBAN SZEREPEL
                    else 
                    {
                        // KISZÁMOLJUK A SÚLYOZÁSÁT
                        int weight;
                        if (trelloCard.Labels == null)
                            weight = 0;
                        else
                        {
                            List<string> labelIDs = new List<string>();
                            foreach (var label in trelloCard.Labels)
                                labelIDs.Add(label.Id);
                            weight = Utilities.GetWeightFromLabels(labelIDs);
                        }

                        // HA A TRELLOBAN ELTÉR A TASK ELFOGADÁSI STÁTUSZA AZ ADATBÁZISHOZ KÉPEST
                        // (A) TRELLOBAN EL LETT FOGADVA  => NÖVELNI KELL A SZÁMLÁLÓT
                        // (B) TRELLOBAN ÚJRA LETT NYITVA => CSÖKKENTENI KELL A SZÁMLÁLÓT
                        if (trelloCard.IsComplete != dbCard.IsComplete)
                        {
                            dbCard.IsComplete = trelloCard.IsComplete;
                            // (A) EL LETT FOGADVA
                            if (trelloCard.IsComplete == true) {
                                dbCard.Date = (DateTime) trelloCard.LastActivity;
                                Utilities.UpdateTes(dbContext, dbCard);
                            }
                            // (B) ÚJRA LETT NYITVA
                            else
                                Utilities.UpdateTes(dbContext, dbCard, "REMOVE");

                            dbContext.SaveChanges();
                        }
                        
                        // VÁLTOZOTT SÚLYOZÁS
                        if (weight != dbCard.Weight) 
                        {
                            if (trelloCard.IsComplete == true)
                            {
                                if (dbCard.IsComplete == true)
                                {
                                    Utilities.UpdateTes(dbContext, dbCard, "REMOVE");
                                    dbCard.Weight = weight;
                                    Utilities.UpdateTes(dbContext, dbCard);
                                }
                            }
                            dbCard.Weight = weight;
                            dbContext.SaveChanges();
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("Database save error: " + ex.Message);
        }
    }
}

/// <summary>
/// Ellenőrzi az adatbázis tábláit, és ha valamelyik hiányzik, létrehozza azokat.
/// </summary>
/// <param name="dbContext">Az adatbázis kapcsolat változója.</param>
static void CheckDatabaseTables(ApplicationDbContext dbContext)
{
    try
    {
        // Kapcsolat létrehozása az adatbázishoz
        var connection = dbContext.Database.GetDbConnection();
        connection.Open();
        var command = connection.CreateCommand();

        // Lekérdezés az 'Lists' tábla létezésére, ha nem létezik false-szal tér vissza
        command.CommandText = "SELECT COUNT(*) FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'Lists'";
        bool isListsTableExist = (Convert.ToInt32(command.ExecuteScalar()) <= 0);

        // Lekérdezés a 'Cards' tábla létezésére, ha nem létezik false-szal tér vissza
        command.CommandText = "SELECT COUNT(*) FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'Cards'";
        bool isCardsTableExist = (Convert.ToInt32(command.ExecuteScalar()) <= 0);

        // Lekérdezés a 'Completed' tábla létezésére, ha nem létezik false-szal tér vissza
        command.CommandText = "SELECT COUNT(*) FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'Completed'";
        bool isCompletedTableExist = (Convert.ToInt32(command.ExecuteScalar()) <= 0);
        connection.Close();

        // Ha nem létezik 'Lists' tábla az adatbázisban -> létre kell hozni
        if (isListsTableExist)
            dbContext.CreateTable("Lists");

        // Ha nem létezik 'Cards' tábla az adatbázisban -> létre kell hozni
        if (isCardsTableExist)
            dbContext.CreateTable("Cards");

        // Ha nem létezik 'Completed' tábla az adatbázisban -> létre kell hozni
        if (isCompletedTableExist)
            dbContext.CreateTable("Completed");
    }
    catch (Exception ex)
    {
        // Hiba esetén hibaüzenet kiírása
        Console.WriteLine("Adatbázis tábla létrehozási hiba: " + ex.Message);
    }
}




