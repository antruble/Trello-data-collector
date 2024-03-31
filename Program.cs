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

                    // ADOTT TRELLO KÁRTYA MEGKERESÉSE LISTÁBAN
                    var dbCard = dbList.Cards?.FirstOrDefault(c => c.Id?.Trim() == trelloCard.Id.Trim());
                    // NEM LÉTEZIK A KÁRTYA AZ ADOTT LISTÁBAN -> ( A ) ÁTKERÜLT MÁSIK LISTÁBA ( B ) MÉG NEM LÉTEZIK
                    if (dbCard == null)
                    {
                        DateTime? oldDate = null;
                        // TRELLOKÁRTYA MEGKERESÉSE A DB ÖSSZES KÁRTYA LISTÁJÁBÓL
                        var temp = dbContext.Cards?.FirstOrDefault(c => c.Id.Trim() == trelloCard.Id.Trim());
                        // HA A TEMP NEM NULL, AKKOR A KÁRTYA MÁR RÖGZÍTVE VAN -> ÚJ LISTÁBA KERÜLT -> FRISSÍTENI KELL
                        if (temp != null)
                        {
                            // HA A KÁRTYA EL VAN FOGADVA, ÉS ÚGY LETT ELMOZGATVA
                            // ˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇˇ
                            // (1) AZ ELŐZŐ LISTA SZÁMLÁLÓJÁBÓL KI KELL SZEDNI
                            // (2) AZ ELFOGADÁS DÁTUMÁT EL KELL TÁROLNI, AMIT MÁR CSAK AZ ADATBÁZIS TÁROL
                            //     (A MÁSIK LISTÁBAN EZT A HÓNAPOT KELL FRISSÍTENI)
                            // ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
                            if (temp.IsComplete == true)
                                Utilities.UpdateTes(dbContext, temp, "REMOVE");

                            // NEM SZERETNÉNK FRISSÍTENI A DÁTUMOT, KIVÉVE, HA MOZGATÁS KÖZBEN EL IS LETT FOGADVA
                            oldDate = temp.IsComplete == false && trelloCard.IsComplete == true ? null : temp.Date;

                            // KÁRTYA KITÖRLÉSE DB-BŐL
                            dbContext.DeleteCard(temp);
                            dbContext.SaveChanges();
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
                        if (trelloCard.IsComplete == true)
                            Utilities.UpdateTes(dbContext, newCard);
                        dbList.Cards?.Add(newCard);
                        dbContext.SaveChanges();
                    }
                    // LÉTEZIK A KÁRTYA 
                    else 
                    {
                        // A KÁRTYÁN LÉVŐ LABELEK KÖZÜL KIVÁLASZTJUK A LEGNAGYOBB FONTOSSÁGI SÚLYÚT
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

                        // [(A) TRELLOBAN EL LETT FOGADVA || (B) TRELLOBAN ÚJRA LETT NYITVA] => DOKUMENTÁLNI KELL
                        if (trelloCard.IsComplete != dbCard.IsComplete)
                        {
                            dbCard.IsComplete = trelloCard.IsComplete;
                            // (A) EL LETT FOGADVA
                            if (trelloCard.IsComplete == true)
                                Utilities.UpdateTes(dbContext, dbCard);
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

static void CheckDatabaseTables(ApplicationDbContext dbContext) 
{
    try 
    {
        var connection = dbContext.Database.GetDbConnection();
        connection.Open();
        var command = connection.CreateCommand();

        // TÁBLÁS LEKÉRDEZÉSE, HA NEM LÉTEZIK FALSE-SZAL TÉR VISSZA
        command.CommandText = "SELECT COUNT(*) FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'Lists'";
        bool isListsTableExist = (Convert.ToInt32(command.ExecuteScalar()) <= 0);
        command.CommandText = "SELECT COUNT(*) FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'Cards'";
        bool isCardsTableExist = (Convert.ToInt32(command.ExecuteScalar()) <= 0);
        command.CommandText = "SELECT COUNT(*) FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'Completed'";
        bool isCompletedTableExist = (Convert.ToInt32(command.ExecuteScalar()) <= 0);
        connection.Close();

        // NEM LÉTEZIK LISTS TÁBLA AZ ADATBÁZISBAN -> LÉTRE KELL HOZNI
        if (isListsTableExist)
            dbContext.CreateTable("Lists");
        // NEM LÉTEZIK CARDS TÁBLA AZ ADATBÁZISBAN -> LÉTRE KELL HOZNI
        if (isCardsTableExist)
            dbContext.CreateTable("Cards");
        // NEM LÉTEZIK COMPLETED TÁBLA AZ ADATBÁZISBAN -> LÉTRE KELL HOZNI
        if (isCompletedTableExist)
            dbContext.CreateTable("Completed");
    } 
    catch (Exception ex)
    {
        Console.WriteLine("Database table create error: " + ex.Message);
    }

    
}



